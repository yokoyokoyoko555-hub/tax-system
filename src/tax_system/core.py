from __future__ import annotations

import csv
import hashlib
import json
import os
import shutil
import sqlite3
import tempfile
from contextlib import closing
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable

from openpyxl import load_workbook


LEDGER_COLUMNS = ["日時", "名前", "ふりがな", "生年月日", "住所", "電話番号", "商品名", "個数", "単価", "金額", "備考"]


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as source:
        for chunk in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def json_text(value: Any) -> str:
    return json.dumps(value, ensure_ascii=False, separators=(",", ":"), default=str)


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
            row = db.execute("SELECT id, kind, source_name, imported_at FROM imports WHERE id=?", (import_id,)).fetchone()
        return dict(row) if row else None

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
        return self._validate_ledger(records) if imp["kind"] == "ledger" else self._validate_comparison(records, json.loads(imp["metadata_json"]))

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
