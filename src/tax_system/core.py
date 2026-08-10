from __future__ import annotations

import csv
import hashlib
import json
import os
import re
import shutil
import sqlite3
import tempfile
from contextlib import closing
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any, Iterable

from openpyxl import load_workbook


LEDGER_COLUMNS = ["日時", "名前", "ふりがな", "生年月日", "住所", "電話番号", "商品名", "個数", "単価", "金額", "備考"]
INVENTORY_COLUMNS = ["商品名", "仕入れ原価", "在庫数"]


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as source:
        for chunk in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def json_text(value: Any) -> str:
    return json.dumps(value, ensure_ascii=False, separators=(",", ":"), default=str)


def _to_number(value: Any) -> float | None:
    if value in (None, ""):
        return None
    try:
        return float(str(value).replace(",", "").strip())
    except ValueError:
        return None


def _to_date(value: Any) -> date | None:
    if value in (None, ""):
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    match = re.match(r"(\d{4})[-/](\d{1,2})[-/](\d{1,2})", str(value).strip())
    if not match:
        return None
    year, month, day = (int(part) for part in match.groups())
    try:
        return date(year, month, day)
    except ValueError:
        return None


def _index_of(headers: list[Any], name: str, contains: bool = False) -> int | None:
    for index, header in enumerate(headers):
        if header is None:
            continue
        if contains and name in str(header):
            return index
        if not contains and header == name:
            return index
    return None


@dataclass
class CheckResult:
    code: str
    level: str
    message: str
    sheet: str | None = None
    row_no: int | None = None


class TaxSystem:
    def __init__(self, home: str | Path = ".tax-system") -> None:
        self.home = Path(home).resolve()
        self.templates = self.home / "templates"
        self.db_path = self.home / "tax-system.sqlite3"

    def initialize(self) -> None:
        self.templates.mkdir(parents=True, exist_ok=True)
        with closing(self.connect()) as db, db:
            db.executescript(
                """
                PRAGMA journal_mode=WAL;
                CREATE TABLE IF NOT EXISTS template_versions(
                  id INTEGER PRIMARY KEY, report_type TEXT NOT NULL, name TEXT NOT NULL,
                  version TEXT NOT NULL, effective_from TEXT NOT NULL, format TEXT NOT NULL,
                  source_name TEXT NOT NULL, stored_path TEXT NOT NULL, sha256 TEXT NOT NULL,
                  settings_json TEXT NOT NULL, created_at TEXT NOT NULL,
                  UNIQUE(report_type, name, version)
                );
                CREATE TABLE IF NOT EXISTS imports(
                  id INTEGER PRIMARY KEY, kind TEXT NOT NULL, source_name TEXT NOT NULL,
                  source_sha256 TEXT NOT NULL, imported_at TEXT NOT NULL,
                  metadata_json TEXT NOT NULL
                );
                CREATE TABLE IF NOT EXISTS records(
                  id INTEGER PRIMARY KEY, import_id INTEGER NOT NULL REFERENCES imports(id),
                  sheet_name TEXT, row_no INTEGER NOT NULL, data_json TEXT NOT NULL
                );
                CREATE TABLE IF NOT EXISTS exports(
                  id INTEGER PRIMARY KEY, import_id INTEGER NOT NULL, template_id INTEGER,
                  mode TEXT NOT NULL, output_name TEXT NOT NULL, output_sha256 TEXT NOT NULL,
                  checks_json TEXT NOT NULL, created_at TEXT NOT NULL
                );
                CREATE TABLE IF NOT EXISTS allocations(
                  id INTEGER PRIMARY KEY, sale_import_id INTEGER NOT NULL, sale_sheet TEXT NOT NULL,
                  sale_row_no INTEGER NOT NULL, ledger_record_id INTEGER, status TEXT NOT NULL,
                  candidates_json TEXT NOT NULL, note TEXT, created_at TEXT NOT NULL,
                  UNIQUE(sale_import_id, sale_sheet, sale_row_no)
                );
                CREATE TABLE IF NOT EXISTS ledger_items(
                  id INTEGER PRIMARY KEY, ledger_import_id INTEGER NOT NULL, ledger_row_no INTEGER NOT NULL,
                  product TEXT NOT NULL, qty REAL NOT NULL, unit_cost REAL NOT NULL, amount REAL NOT NULL,
                  source TEXT NOT NULL, created_at TEXT NOT NULL
                );
                """
            )

    def connect(self) -> sqlite3.Connection:
        self.home.mkdir(parents=True, exist_ok=True)
        db = sqlite3.connect(self.db_path)
        db.row_factory = sqlite3.Row
        return db

    def register_template(self, report_type: str, source: str | Path, name: str,
                          version: str, effective_from: str) -> int:
        self.initialize()
        source = Path(source).resolve(strict=True)
        suffix = source.suffix.lower()
        if suffix not in {".xlsx", ".csv"}:
            raise ValueError("テンプレートは .xlsx または .csv のみ登録できます")
        target_dir = self.templates / report_type / name / version
        target_dir.mkdir(parents=True, exist_ok=False)
        target = target_dir / ("template" + suffix)
        shutil.copy2(source, target)
        settings = self._template_settings(target)
        now = datetime.now().isoformat(timespec="seconds")
        with closing(self.connect()) as db, db:
            cur = db.execute(
                "INSERT INTO template_versions(report_type,name,version,effective_from,format,source_name,stored_path,sha256,settings_json,created_at) VALUES(?,?,?,?,?,?,?,?,?,?)",
                (report_type, name, version, effective_from, suffix[1:], source.name,
                 str(target), sha256(target), json_text(settings), now),
            )
            return int(cur.lastrowid)

    def _template_settings(self, path: Path) -> dict[str, Any]:
        if path.suffix.lower() == ".csv":
            raw = path.read_bytes()
            encoding = "utf-8-sig" if raw.startswith(b"\xef\xbb\xbf") else "utf-8"
            try:
                text = raw.decode(encoding)
            except UnicodeDecodeError:
                encoding, text = "cp932", raw.decode("cp932")
            dialect = csv.Sniffer().sniff(text[:8192])
            rows = list(csv.reader(text.splitlines(), dialect))
            return {"encoding": encoding, "bom": raw.startswith(b"\xef\xbb\xbf"),
                    "delimiter": dialect.delimiter, "lineterminator": "CRLF" if b"\r\n" in raw else "LF",
                    "header": rows[0] if rows else []}
        wb = load_workbook(path, read_only=False, data_only=False)
        return {"sheets": [{"name": ws.title, "max_row": ws.max_row, "max_column": ws.max_column,
                            "headers": [ws.cell(2, col).value for col in range(1, ws.max_column + 1)],
                            "merged": [str(r) for r in ws.merged_cells.ranges]}
                           for ws in wb.worksheets]}

    def import_ledger(self, source: str | Path) -> int:
        self.initialize()
        source = Path(source).resolve(strict=True)
        raw = source.read_bytes()
        try:
            text, encoding = raw.decode("utf-8-sig"), "utf-8-sig"
        except UnicodeDecodeError:
            text, encoding = raw.decode("cp932"), "cp932"
        reader = csv.DictReader(text.splitlines())
        if reader.fieldnames != LEDGER_COLUMNS:
            raise ValueError(f"古物台帳の列が既存仕様と一致しません: {reader.fieldnames}")
        rows = list(reader)
        return self._save_import("ledger", source, rows, {"encoding": encoding, "columns": reader.fieldnames})

    def import_comparison(self, source: str | Path) -> int:
        self.initialize()
        source = Path(source).resolve(strict=True)
        wb = load_workbook(source, read_only=False, data_only=False)
        records: list[dict[str, Any]] = []
        sheets: list[dict[str, Any]] = []
        for ws in wb.worksheets:
            headers = [ws.cell(2, c).value for c in range(1, ws.max_column + 1)]
            split = next((i for i, h in enumerate(headers) if i > 0 and h == "年月日"), None)
            if split is None:
                raise ValueError(f"{ws.title}: 受入れ・払出しの境界を判定できません")
            sheets.append({"name": ws.title, "headers": headers, "max_column": ws.max_column})
            for row_no in range(3, ws.max_row + 1):
                cells = [ws.cell(row_no, c) for c in range(1, ws.max_column + 1)]
                values = [c.value for c in cells]
                # Existing sheets contain formulas in otherwise unused rows. Dates are
                # the authoritative indication that a purchase/sale row is populated.
                if values[0] in (None, "") and values[split] in (None, ""):
                    continue
                records.append({"sheet": ws.title, "row_no": row_no, "values": values})
        return self._save_import("comparison", source, records, {"sheets": sheets})

    def import_inventory(self, source: str | Path) -> int:
        self.initialize()
        source = Path(source).resolve(strict=True)
        raw = source.read_bytes()
        try:
            text, encoding = raw.decode("utf-8-sig"), "utf-8-sig"
        except UnicodeDecodeError:
            text, encoding = raw.decode("cp932"), "cp932"
        reader = csv.DictReader(text.splitlines())
        if reader.fieldnames != INVENTORY_COLUMNS:
            raise ValueError(f"期末在庫表の列が想定と一致しません（{INVENTORY_COLUMNS}）: {reader.fieldnames}")
        rows = list(reader)
        return self._save_import("inventory", source, rows, {"encoding": encoding, "columns": reader.fieldnames})

    def _save_import(self, kind: str, source: Path, rows: Iterable[dict[str, Any]], metadata: dict[str, Any]) -> int:
        now = datetime.now().isoformat(timespec="seconds")
        rows = list(rows)
        with closing(self.connect()) as db, db:
            cur = db.execute("INSERT INTO imports(kind,source_name,source_sha256,imported_at,metadata_json) VALUES(?,?,?,?,?)",
                             (kind, source.name, sha256(source), now, json_text(metadata)))
            import_id = int(cur.lastrowid)
            payload = []
            for index, row in enumerate(rows, 2):
                sheet = row.get("sheet")
                row_no = int(row.get("row_no", index))
                payload.append((import_id, sheet, row_no, json_text(row)))
            db.executemany("INSERT INTO records(import_id,sheet_name,row_no,data_json) VALUES(?,?,?,?)", payload)
        return import_id

    def list_templates(self) -> list[dict[str, Any]]:
        with closing(self.connect()) as db, db:
            rows = db.execute(
                "SELECT id, report_type, name, version, effective_from, format, created_at FROM template_versions ORDER BY report_type, name, effective_from DESC"
            ).fetchall()
        return [dict(row) for row in rows]

    def list_imports(self) -> list[dict[str, Any]]:
        with closing(self.connect()) as db, db:
            rows = db.execute(
                "SELECT id, kind, source_name, imported_at FROM imports ORDER BY id DESC"
            ).fetchall()
        return [dict(row) for row in rows]

    def get_import(self, import_id: int) -> dict[str, Any] | None:
        with closing(self.connect()) as db, db:
            row = db.execute(
                "SELECT id, kind, source_name, imported_at, metadata_json FROM imports WHERE id=?", (import_id,)
            ).fetchone()
        if not row:
            return None
        data = dict(row)
        data["metadata"] = json.loads(data.pop("metadata_json"))
        return data

    def get_records(self, import_id: int, sheet: str | None = None, offset: int = 0,
                    limit: int = 100) -> tuple[list[dict[str, Any]], int]:
        where = "import_id=?"
        params: list[Any] = [import_id]
        if sheet is not None:
            where += " AND sheet_name=?"
            params.append(sheet)
        with closing(self.connect()) as db, db:
            total = db.execute(f"SELECT COUNT(*) FROM records WHERE {where}", params).fetchone()[0]
            rows = db.execute(
                f"SELECT * FROM records WHERE {where} ORDER BY sheet_name, row_no LIMIT ? OFFSET ?",
                params + [limit, offset],
            ).fetchall()
        results = []
        for row in rows:
            item = dict(row)
            item["data"] = json.loads(item.pop("data_json"))
            results.append(item)
        return results, total

    def list_exports(self) -> list[dict[str, Any]]:
        with closing(self.connect()) as db, db:
            rows = db.execute(
                "SELECT id, import_id, template_id, mode, output_name, created_at FROM exports ORDER BY id DESC"
            ).fetchall()
        return [dict(row) for row in rows]

    def validate(self, import_id: int) -> list[CheckResult]:
        with closing(self.connect()) as db, db:
            imp = db.execute("SELECT * FROM imports WHERE id=?", (import_id,)).fetchone()
            if not imp:
                raise ValueError("取込IDが見つかりません")
            records = db.execute("SELECT * FROM records WHERE import_id=? ORDER BY sheet_name,row_no", (import_id,)).fetchall()
        if imp["kind"] == "ledger":
            return self._validate_ledger(records)
        if imp["kind"] == "inventory":
            return self._validate_inventory(records)
        checks = self._validate_comparison(records, json.loads(imp["metadata_json"]))
        checks.extend(self._validate_allocations(import_id))
        return checks

    def _validate_inventory(self, records: Iterable[sqlite3.Row]) -> list[CheckResult]:
        results: list[CheckResult] = []
        for record in records:
            data = json.loads(record["data_json"])
            row_no = record["row_no"]
            for key in INVENTORY_COLUMNS:
                if data.get(key) in (None, ""):
                    results.append(CheckResult("REQUIRED", "error", f"必須項目「{key}」が空欄です", row_no=row_no))
            if data.get("仕入れ原価") not in (None, "") and _to_number(data.get("仕入れ原価")) is None:
                results.append(CheckResult("NUMBER_FORMAT", "error", "仕入れ原価を数値として解釈できません", row_no=row_no))
            if data.get("在庫数") not in (None, "") and _to_number(data.get("在庫数")) is None:
                results.append(CheckResult("NUMBER_FORMAT", "error", "在庫数を数値として解釈できません", row_no=row_no))
        return results

    def _validate_ledger(self, records: Iterable[sqlite3.Row]) -> list[CheckResult]:
        results: list[CheckResult] = []
        seen: set[tuple[Any, ...]] = set()
        for record in records:
            data = json.loads(record["data_json"])
            row_no = record["row_no"]
            for key in ("日時", "名前", "住所", "商品名", "個数", "金額"):
                if data.get(key) in (None, ""):
                    results.append(CheckResult("REQUIRED", "error", f"必須項目「{key}」が空欄です", row_no=row_no))
            try:
                qty, unit, amount = int(data.get("個数", 0)), float(str(data.get("単価", 0)).replace(",", "")), float(str(data.get("金額", 0)).replace(",", ""))
                if abs(qty * unit - amount) > 0.5:
                    results.append(CheckResult("AMOUNT_MISMATCH", "error", "数量×単価と金額が一致しません", row_no=row_no))
            except (TypeError, ValueError):
                results.append(CheckResult("NUMBER_FORMAT", "error", "数量・単価・金額を数値として解釈できません", row_no=row_no))
            key = tuple(data.get(k) for k in ("日時", "名前", "商品名", "個数", "金額"))
            if key in seen:
                results.append(CheckResult("DUPLICATE", "error", "同一取引の可能性があります", row_no=row_no))
            seen.add(key)
        return results

    def _validate_comparison(self, records: Iterable[sqlite3.Row], metadata: dict[str, Any]) -> list[CheckResult]:
        results: list[CheckResult] = []
        sheet_headers = {s["name"]: s["headers"] for s in metadata["sheets"]}
        for record in records:
            item = json.loads(record["data_json"]); values = item["values"]
            headers = sheet_headers[record["sheet_name"]]
            split = next((i for i, h in enumerate(headers) if i > 0 and h == "年月日"), None)
            if split is None:
                results.append(CheckResult("LAYOUT", "error", "受入れ・払出しの境界を判定できません", record["sheet_name"], record["row_no"])); continue
            purchase, sale = values[:split], values[split:]
            if sale[0] not in (None, ""):
                purchase_qty_index = next((i for i, h in enumerate(headers[:split]) if h == "数量"), None)
                if purchase_qty_index is None or purchase[purchase_qty_index] in (None, ""):
                    results.append(CheckResult("PURCHASE_QTY", "error", "対応仕入数量が空欄です", record["sheet_name"], record["row_no"]))
                sale_qty_index = next((i for i, h in enumerate(headers[split:]) if h == "数量"), None)
                if purchase_qty_index is not None and sale_qty_index is not None:
                    pq, sq = purchase[purchase_qty_index], sale[sale_qty_index]
                    if not isinstance(pq, str) and not isinstance(sq, str) and pq != sq:
                        results.append(CheckResult("QTY_MISMATCH", "error", "販売数量と対応仕入数量が一致しません", record["sheet_name"], record["row_no"]))
        return results

    def _has_ledger_data(self) -> bool:
        with closing(self.connect()) as db, db:
            return db.execute("SELECT 1 FROM imports WHERE kind='ledger' LIMIT 1").fetchone() is not None

    def _validate_allocations(self, comparison_import_id: int) -> list[CheckResult]:
        if not self._has_ledger_data():
            return [CheckResult("LEDGER_MISSING", "error", "古物台帳が未取込のため、仕入との対応を確認できません")]
        checks: list[CheckResult] = []
        for result in self.allocate(comparison_import_id):
            if result["status"] == "matched" and result["note"]:
                checks.append(CheckResult("ALLOCATION_MISMATCH", "error", f"対応する仕入記録と{result['note']}",
                                          result["sheet"], result["row_no"]))
            elif result["status"] == "ambiguous":
                checks.append(CheckResult("ALLOCATION_AMBIGUOUS", "error",
                                          result["note"] or "対応する仕入の候補が複数あります。手動で選択してください",
                                          result["sheet"], result["row_no"]))
            elif result["status"] == "not_found":
                checks.append(CheckResult("ALLOCATION_NOT_FOUND", "error",
                                          result["note"] or "対応する仕入が古物台帳に見つかりません",
                                          result["sheet"], result["row_no"]))
        return checks

    def allocate(self, comparison_import_id: int) -> list[dict[str, Any]]:
        """既存の相対表データ（仕入・販売が同じ行に書かれている）の仕入側について、
        対応する古物台帳の購入記録を商品名・数量・日付・取引相手で自動的に探し、割り当てる。
        1件の古物台帳記録は、どこかひとつの販売にしか対応付けない（使い回さない）。
        """
        with closing(self.connect()) as db, db:
            imp = db.execute("SELECT * FROM imports WHERE id=?", (comparison_import_id,)).fetchone()
            if not imp or imp["kind"] != "comparison":
                raise ValueError("相対表の取込IDを指定してください")
            metadata = json.loads(imp["metadata_json"])
            records = db.execute(
                "SELECT * FROM records WHERE import_id=? ORDER BY sheet_name, row_no", (comparison_import_id,)
            ).fetchall()
            ledger_rows = db.execute(
                "SELECT r.id, r.data_json FROM records r JOIN imports i ON r.import_id = i.id WHERE i.kind='ledger'"
            ).fetchall()
            existing = {
                (row["sale_sheet"], row["sale_row_no"]): row
                for row in db.execute(
                    "SELECT * FROM allocations WHERE sale_import_id=?", (comparison_import_id,)
                ).fetchall()
            }
            other_consumed = {
                row["ledger_record_id"]
                for row in db.execute(
                    "SELECT ledger_record_id FROM allocations WHERE ledger_record_id IS NOT NULL "
                    "AND status IN ('matched','manual') AND sale_import_id != ?",
                    (comparison_import_id,),
                ).fetchall()
            }

        sheet_headers = {s["name"]: s["headers"] for s in metadata["sheets"]}
        ledger_by_id: dict[int, dict[str, Any]] = {}
        for row in ledger_rows:
            data = json.loads(row["data_json"])
            ledger_by_id[row["id"]] = {
                "id": row["id"],
                "name": (data.get("名前") or "").strip(),
                "product": (data.get("商品名") or "").strip(),
                "qty": _to_number(data.get("個数")),
                "date": _to_date(data.get("日時")),
                "amount": _to_number(data.get("金額")),
            }

        consumed = set(other_consumed)
        for row in existing.values():
            if row["status"] == "manual" and row["ledger_record_id"] is not None:
                consumed.add(row["ledger_record_id"])

        results: list[dict[str, Any]] = []
        upserts: list[tuple[Any, ...]] = []
        now = datetime.now().isoformat(timespec="seconds")

        for record in records:
            headers = sheet_headers.get(record["sheet_name"], [])
            split = next((i for i, h in enumerate(headers) if i > 0 and h == "年月日"), None)
            if split is None:
                continue
            values = json.loads(record["data_json"])["values"]
            purchase, sale = values[:split], values[split:]
            if sale[0] in (None, ""):
                continue
            purchase_headers = headers[:split]

            prior = existing.get((record["sheet_name"], record["row_no"]))
            if prior and prior["status"] == "manual":
                results.append({
                    "row_no": record["row_no"], "sheet": record["sheet_name"], "status": "manual",
                    "ledger_record_id": prior["ledger_record_id"],
                    "ledger": ledger_by_id.get(prior["ledger_record_id"]),
                    "product": None, "note": "手動で割当済み", "candidates": [],
                })
                continue

            product_idx = _index_of(purchase_headers, "品目")
            qty_idx = _index_of(purchase_headers, "数量")
            name_idx = _index_of(purchase_headers, "相手方名")
            amount_idx = _index_of(purchase_headers, "代価", contains=True)
            product = str(purchase[product_idx]).strip() if product_idx is not None and purchase[product_idx] not in (None, "") else None
            qty = _to_number(purchase[qty_idx]) if qty_idx is not None else None
            purchase_date = _to_date(purchase[0])
            counterparty = str(purchase[name_idx]).strip() if name_idx is not None and purchase[name_idx] not in (None, "") else None
            amount = _to_number(purchase[amount_idx]) if amount_idx is not None else None

            candidates = [c for c in ledger_by_id.values() if c["id"] not in consumed and product and c["product"] == product]
            strict = [c for c in candidates if c["qty"] == qty and c["date"] == purchase_date]
            pool = strict or candidates
            if counterparty:
                named = [c for c in pool if c["name"] == counterparty]
                if named:
                    pool = named

            status: str; ledger_id: int | None = None; note: str | None; candidate_ids: list[int] = []
            if not product:
                status, note = "not_found", "品目が空欄です"
            elif len(pool) == 1:
                status, ledger_id = "matched", pool[0]["id"]
                consumed.add(ledger_id)
                mismatches = []
                if qty is not None and pool[0]["qty"] != qty:
                    mismatches.append("数量")
                if purchase_date is not None and pool[0]["date"] != purchase_date:
                    mismatches.append("日付")
                if amount is not None and pool[0]["amount"] is not None and abs(pool[0]["amount"] - amount) > 0.5:
                    mismatches.append("金額")
                note = ("・".join(mismatches) + "が一致しません") if mismatches else None
            elif len(pool) > 1:
                status, note, candidate_ids = "ambiguous", f"候補{len(pool)}件から選択してください", [c["id"] for c in pool]
            else:
                status, note, candidate_ids = "not_found", "一致する古物台帳の記録がありません", [c["id"] for c in candidates]

            upserts.append((comparison_import_id, record["sheet_name"], record["row_no"], ledger_id, status,
                            json_text(candidate_ids), note, now))
            results.append({
                "row_no": record["row_no"], "sheet": record["sheet_name"], "status": status,
                "ledger_record_id": ledger_id, "ledger": ledger_by_id.get(ledger_id), "product": product,
                "note": note, "candidates": [ledger_by_id[cid] for cid in candidate_ids],
            })

        with closing(self.connect()) as db, db:
            db.executemany(
                """INSERT INTO allocations(sale_import_id, sale_sheet, sale_row_no, ledger_record_id, status, candidates_json, note, created_at)
                   VALUES(?,?,?,?,?,?,?,?)
                   ON CONFLICT(sale_import_id, sale_sheet, sale_row_no) DO UPDATE SET
                     ledger_record_id=excluded.ledger_record_id, status=excluded.status,
                     candidates_json=excluded.candidates_json, note=excluded.note, created_at=excluded.created_at
                """,
                upserts,
            )
        return results

    def search_ledger(self, query: str, limit: int = 30) -> list[dict[str, Any]]:
        query = query.strip()
        if not query:
            return []
        with closing(self.connect()) as db, db:
            rows = db.execute(
                "SELECT r.id, r.data_json FROM records r JOIN imports i ON r.import_id=i.id WHERE i.kind='ledger' ORDER BY r.id"
            ).fetchall()
            consumed = {
                row["ledger_record_id"]
                for row in db.execute(
                    "SELECT ledger_record_id FROM allocations WHERE ledger_record_id IS NOT NULL AND status IN ('matched','manual')"
                ).fetchall()
            }
        result = []
        for row in rows:
            if row["id"] in consumed:
                continue
            data = json.loads(row["data_json"])
            haystack = f"{data.get('商品名', '')} {data.get('名前', '')}"
            if query not in haystack:
                continue
            result.append({"id": row["id"], "name": data.get("名前"), "product": data.get("商品名"),
                           "qty": data.get("個数"), "date": data.get("日時"), "amount": data.get("金額")})
            if len(result) >= limit:
                break
        return result

    def set_manual_allocation(self, comparison_import_id: int, sheet: str, row_no: int,
                              ledger_record_id: int | None) -> None:
        now = datetime.now().isoformat(timespec="seconds")
        with closing(self.connect()) as db, db:
            db.execute(
                """INSERT INTO allocations(sale_import_id, sale_sheet, sale_row_no, ledger_record_id, status, candidates_json, note, created_at)
                   VALUES(?,?,?,?,?,?,?,?)
                   ON CONFLICT(sale_import_id, sale_sheet, sale_row_no) DO UPDATE SET
                     ledger_record_id=excluded.ledger_record_id, status='manual', candidates_json='[]',
                     note=NULL, created_at=excluded.created_at
                """,
                (comparison_import_id, sheet, row_no, ledger_record_id, "manual", "[]", None, now),
            )

    def _latest_inventory(self) -> dict[str, float]:
        with closing(self.connect()) as db, db:
            imp = db.execute(
                "SELECT id FROM imports WHERE kind='inventory' ORDER BY imported_at DESC LIMIT 1"
            ).fetchone()
            if not imp:
                return {}
            rows = db.execute("SELECT data_json FROM records WHERE import_id=?", (imp["id"],)).fetchall()
        costs: dict[str, float] = {}
        for row in rows:
            data = json.loads(row["data_json"])
            product = (data.get("商品名") or "").strip()
            cost = _to_number(data.get("仕入れ原価"))
            if product and cost is not None:
                costs[product] = cost
        return costs

    def search_inventory(self, query: str, limit: int = 30) -> list[dict[str, Any]]:
        query = query.strip()
        if not query:
            return []
        result = []
        for product, cost in self._latest_inventory().items():
            if query in product:
                result.append({"product": product, "unit_cost": cost})
                if len(result) >= limit:
                    break
        return result

    def propose_ledger_breakdown(self, ledger_import_id: int) -> list[dict[str, Any]]:
        """商品名が空欄の古物台帳の行について、相対表から判明している内訳（確定）と、
        まだ手動で埋めていない残額（要確認）を計算する。金額の内訳は書き換えず、
        常にこの場で再計算するため、相対表・在庫表を取込み直すと結果も更新される。
        """
        with closing(self.connect()) as db, db:
            imp = db.execute("SELECT * FROM imports WHERE id=?", (ledger_import_id,)).fetchone()
            if not imp or imp["kind"] != "ledger":
                raise ValueError("古物台帳の取込IDを指定してください")
            ledger_records = db.execute(
                "SELECT * FROM records WHERE import_id=? ORDER BY row_no", (ledger_import_id,)
            ).fetchall()
            comparison_rows = db.execute(
                "SELECT r.import_id, r.sheet_name, r.data_json FROM records r "
                "JOIN imports i ON r.import_id = i.id WHERE i.kind='comparison'"
            ).fetchall()
            comparison_metadata = {
                row["id"]: json.loads(row["metadata_json"])
                for row in db.execute("SELECT id, metadata_json FROM imports WHERE kind='comparison'").fetchall()
            }
            manual_items = db.execute(
                "SELECT * FROM ledger_items WHERE ledger_import_id=? ORDER BY id", (ledger_import_id,)
            ).fetchall()

        inventory = self._latest_inventory()

        comparison_by_key: dict[tuple[str, date | None], list[dict[str, Any]]] = {}
        for row in comparison_rows:
            metadata = comparison_metadata.get(row["import_id"], {})
            headers = next((s["headers"] for s in metadata.get("sheets", []) if s["name"] == row["sheet_name"]), None)
            if not headers:
                continue
            split = next((i for i, h in enumerate(headers) if i > 0 and h == "年月日"), None)
            if split is None:
                continue
            values = json.loads(row["data_json"])["values"]
            purchase, sale = values[:split], values[split:]
            if sale[0] in (None, ""):
                continue
            purchase_headers = headers[:split]
            product_idx = _index_of(purchase_headers, "品目")
            qty_idx = _index_of(purchase_headers, "数量")
            name_idx = _index_of(purchase_headers, "相手方名")
            if product_idx is None or name_idx is None:
                continue
            product = str(purchase[product_idx]).strip() if purchase[product_idx] not in (None, "") else None
            counterparty = str(purchase[name_idx]).strip() if purchase[name_idx] not in (None, "") else None
            qty = _to_number(purchase[qty_idx]) if qty_idx is not None else None
            purchase_date = _to_date(purchase[0])
            if not product or not counterparty:
                continue
            comparison_by_key.setdefault((counterparty, purchase_date), []).append({"product": product, "qty": qty or 0})

        manual_by_row: dict[int, list[dict[str, Any]]] = {}
        for item in manual_items:
            manual_by_row.setdefault(item["ledger_row_no"], []).append(dict(item))

        results: list[dict[str, Any]] = []
        for record in ledger_records:
            data = json.loads(record["data_json"])
            if data.get("商品名") not in (None, ""):
                continue
            row_no = record["row_no"]
            name = (data.get("名前") or "").strip()
            purchase_date = _to_date(data.get("日時"))
            total = _to_number(data.get("金額"))

            known = []
            for entry in comparison_by_key.get((name, purchase_date), []):
                cost = inventory.get(entry["product"])
                qty = entry["qty"]
                known.append({
                    "product": entry["product"], "qty": qty, "unit_cost": cost,
                    "amount": (cost * qty) if cost is not None else None, "source": "comparison",
                })
            manual = [{
                "id": item["id"], "product": item["product"], "qty": item["qty"],
                "unit_cost": item["unit_cost"], "amount": item["amount"], "source": "manual",
            } for item in manual_by_row.get(row_no, [])]

            accounted = sum((i["amount"] or 0) for i in known + manual)
            remainder = (total - accounted) if total is not None else None
            results.append({
                "row_no": row_no, "name": name, "date": purchase_date, "total": total,
                "known_items": known, "manual_items": manual, "remainder": remainder,
                "resolved": remainder is not None and abs(remainder) < 0.5,
            })
        return results

    def add_ledger_item(self, ledger_import_id: int, row_no: int, product: str, qty: float) -> None:
        cost = self._latest_inventory().get(product)
        if cost is None:
            raise ValueError(f"期末在庫表に「{product}」の仕入れ原価が見つかりません")
        now = datetime.now().isoformat(timespec="seconds")
        with closing(self.connect()) as db, db:
            db.execute(
                "INSERT INTO ledger_items(ledger_import_id, ledger_row_no, product, qty, unit_cost, amount, source, created_at) "
                "VALUES(?,?,?,?,?,?,?,?)",
                (ledger_import_id, row_no, product, qty, cost, cost * qty, "manual", now),
            )

    def remove_ledger_item(self, item_id: int) -> None:
        with closing(self.connect()) as db, db:
            db.execute("DELETE FROM ledger_items WHERE id=? AND source='manual'", (item_id,))

    def export_completed_ledger(self, ledger_import_id: int, output: str | Path) -> Path:
        breakdown = self.propose_ledger_breakdown(ledger_import_id)
        unresolved = [b for b in breakdown if not b["resolved"]]
        if unresolved:
            raise ValueError(f"{len(unresolved)}件の行で内訳の金額が合計と一致していません。先に確定してください")
        breakdown_by_row = {b["row_no"]: b for b in breakdown}
        with closing(self.connect()) as db, db:
            records = db.execute(
                "SELECT * FROM records WHERE import_id=? ORDER BY row_no", (ledger_import_id,)
            ).fetchall()
        output = Path(output).resolve()
        output.parent.mkdir(parents=True, exist_ok=True)
        with output.open("w", encoding="utf-8-sig", newline="") as fh:
            writer = csv.DictWriter(fh, fieldnames=LEDGER_COLUMNS, lineterminator="\r\n", quoting=csv.QUOTE_MINIMAL)
            writer.writeheader()
            for record in records:
                data = json.loads(record["data_json"])
                entry = breakdown_by_row.get(record["row_no"])
                if not entry:
                    writer.writerow(data)
                    continue
                original_total = data.get("金額")
                for item in entry["known_items"] + entry["manual_items"]:
                    row = dict(data)
                    row["商品名"] = item["product"]
                    row["個数"] = item["qty"]
                    row["単価"] = item["unit_cost"]
                    row["金額"] = item["amount"]
                    note = f"内訳復元（元合計{original_total}円・{'相対表' if item['source'] == 'comparison' else '手動'}）"
                    row["備考"] = f"{data.get('備考') or ''} {note}".strip()
                    writer.writerow(row)
        return output

    def export(self, import_id: int, output: str | Path, template_id: int | None = None,
               preview: bool = False) -> Path:
        checks = self.validate(import_id)
        if any(c.level == "error" for c in checks) and not preview:
            raise ValueError("検証エラーがあるため正式出力できません。--preview を指定すると確認用出力が可能です")
        with closing(self.connect()) as db, db:
            imp = db.execute("SELECT * FROM imports WHERE id=?", (import_id,)).fetchone()
            records = db.execute("SELECT * FROM records WHERE import_id=? ORDER BY sheet_name,row_no", (import_id,)).fetchall()
            template = db.execute("SELECT * FROM template_versions WHERE id=?", (template_id,)).fetchone() if template_id else None
        output = Path(output).resolve()
        output.parent.mkdir(parents=True, exist_ok=True)
        if imp["kind"] == "ledger":
            self._export_ledger(records, output)
        else:
            if not template: raise ValueError("相対表出力には --template-id が必要です")
            self._export_comparison(records, Path(template["stored_path"]), output, preview)
        with closing(self.connect()) as db, db:
            db.execute("INSERT INTO exports(import_id,template_id,mode,output_name,output_sha256,checks_json,created_at) VALUES(?,?,?,?,?,?,?)",
                       (import_id, template_id, "preview" if preview else "formal", output.name, sha256(output),
                        json_text([c.__dict__ for c in checks]), datetime.now().isoformat(timespec="seconds")))
        return output

    def _export_ledger(self, records: Iterable[sqlite3.Row], output: Path) -> None:
        with output.open("w", encoding="utf-8-sig", newline="") as fh:
            writer = csv.DictWriter(fh, fieldnames=LEDGER_COLUMNS, lineterminator="\r\n", quoting=csv.QUOTE_MINIMAL)
            writer.writeheader()
            for record in records: writer.writerow(json.loads(record["data_json"]))

    def _export_comparison(self, records: Iterable[sqlite3.Row], template: Path, output: Path, preview: bool) -> None:
        wb = load_workbook(template, read_only=False, data_only=False)
        grouped: dict[str, list[dict[str, Any]]] = {}
        for record in records: grouped.setdefault(record["sheet_name"], []).append(json.loads(record["data_json"]))
        for ws in wb.worksheets:
            items = grouped.get(ws.title, [])
            for r in range(3, ws.max_row + 1):
                for c in range(1, ws.max_column + 1): ws.cell(r, c).value = None
            for index, item in enumerate(items, 3):
                for col, value in enumerate(item["values"], 1): ws.cell(index, col).value = value
            if preview:
                ws.oddHeader.center.text = "確認用（正式帳票ではありません）"
        fd, temp_name = tempfile.mkstemp(suffix=".xlsx", dir=output.parent); os.close(fd)
        try:
            wb.save(temp_name); os.replace(temp_name, output)
        finally:
            if os.path.exists(temp_name): os.unlink(temp_name)
