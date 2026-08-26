from __future__ import annotations

import os
import shutil
import tempfile
from datetime import datetime
from pathlib import Path

from flask import Flask, abort, flash, redirect, render_template, request, send_file, session, url_for
from werkzeug.utils import secure_filename

from .core import EXPORT_DATA_COLUMNS, INVENTORY_DISPLAY_COLUMNS, LEDGER_COLUMNS, TaxSystem, translate_ja_to_en

KIND_LABELS = {"ledger": "古物台帳", "comparison": "相対表", "inventory": "期末在庫表", "export_data": "輸出データ"}
FLAT_COLUMNS = {"ledger": LEDGER_COLUMNS, "inventory": INVENTORY_DISPLAY_COLUMNS, "export_data": EXPORT_DATA_COLUMNS}

OUTPUT_DIR = Path("outputs").resolve()


def safe_upload_filename(original_name: str, fallback: str) -> str:
    """secure_filename() は非ASCII文字だけの名前（日本語のみのファイル名など）を
    拡張子ごと消してしまい、拡張子で種類判定するimport_autoが「対応していない
    ファイル形式」として弾いてしまう。拡張子は元のファイル名からそのまま残し、
    ベース名だけ安全化する（安全化した結果が空ならfallbackを使う）。
    """
    suffix = Path(original_name).suffix
    base = secure_filename(Path(original_name).stem) or fallback
    return base + suffix


def format_number(value):
    """表示用: 整数値のfloatは末尾の.0を取り除く（10800.0 -> 10800）。
    実際にfloat型の値だけを対象にする（電話番号・郵便番号など数字に見える
    文字列は対象外にして、先頭の0が消えるような事故を防ぐ）。計算結果には影響しない。
    Excelの数式結果は浮動小数点誤差で 1799.9999999999998 のようになることがあるため、
    丸めた後にあらためて整数判定する。
    """
    if isinstance(value, bool):
        return value
    if isinstance(value, float):
        rounded = round(value, 2)
        return int(rounded) if rounded == int(rounded) else rounded
    return value


def create_app(home: str | Path | None = None) -> Flask:
    app = Flask(__name__)
    app.secret_key = os.environ.get("TAX_SYSTEM_SECRET", "local-only-tax-system")
    app.config["TAX_SYSTEM_HOME"] = home or os.environ.get("TAX_SYSTEM_HOME", ".tax-system")
    app.jinja_env.globals["kind_labels"] = KIND_LABELS
    app.jinja_env.filters["num"] = format_number

    def system() -> TaxSystem:
        ts = TaxSystem(app.config["TAX_SYSTEM_HOME"])
        ts.initialize()
        return ts

    @app.route("/")
    def index():
        ts = system()
        return render_template(
            "index.html",
            templates=ts.list_templates(),
            imports=ts.list_imports(),
            exports=ts.list_exports(),
        )

    @app.route("/library/<kind>")
    def library_view(kind: str):
        if kind not in KIND_LABELS:
            abort(404)
        ts = system()
        imports = [imp for imp in ts.list_imports() if imp["kind"] == kind]
        if kind == "inventory":
            imports.sort(key=lambda imp: (imp["as_of"] is None, imp["as_of"] or ""))
        merged = ts.latest_merged_ledger_import() if kind == "ledger" else None
        months = ts.month_summary(merged["id"]) if merged else []
        ledger_duplicates_all = ts.find_ledger_duplicates(merged["id"]) if merged else None
        ledger_duplicates = None
        ledger_dup_page = ledger_dup_total_pages = 1
        if ledger_duplicates_all is not None:
            dup_page_size = 30
            ledger_dup_page = max(1, request.args.get("dup_page", 1, type=int))
            ledger_dup_total_pages = max(1, -(-len(ledger_duplicates_all) // dup_page_size))
            start = (ledger_dup_page - 1) * dup_page_size
            ledger_duplicates = ledger_duplicates_all[start:start + dup_page_size]
        comparison_months = None
        duplicates = None
        purchase_reuse = None
        if kind == "comparison":
            comparison_months = ts.comparison_month_summary()
            duplicates = ts.find_comparison_duplicates()
            purchase_reuse = ts.find_comparison_purchase_reuse()
            for group in purchase_reuse:
                by_file: dict[tuple[int, str], list[int]] = {}
                for occ in group["occurrences"]:
                    by_file.setdefault((occ["import_id"], occ["sheet"]), []).append(occ["row_no"])
                group["view_links"] = [
                    {"import_id": import_id, "sheet": sheet, "row_nos": row_nos}
                    for (import_id, sheet), row_nos in by_file.items()
                ]
        return render_template(
            "library.html", kind=kind, imports=imports, merged=merged, months=months,
            ledger_duplicates=ledger_duplicates, ledger_dup_total=len(ledger_duplicates_all or []),
            ledger_dup_page=ledger_dup_page, ledger_dup_total_pages=ledger_dup_total_pages,
            comparison_months=comparison_months, duplicates=duplicates, purchase_reuse=purchase_reuse,
        )

    @app.route("/library/comparison/month/<month>")
    def comparison_month_view(month: str):
        ts = system()
        per_page = 100
        page = max(1, request.args.get("page", 1, type=int))
        rows, total = ts.get_comparison_month_records(month, offset=(page - 1) * per_page, limit=per_page)
        total_pages = max(1, -(-total // per_page))
        templates = [t for t in ts.list_templates() if t["report_type"] == "comparison"]
        checks = ts.validate_comparison_month(month)
        return render_template(
            "comparison_month.html", month=month, rows=rows, total=total, page=page, total_pages=total_pages,
            templates=templates, checks=checks, has_errors=any(c.level == "error" for c in checks),
        )

    @app.route("/library/comparison/month/<month>/export", methods=["POST"])
    def comparison_month_export(month: str):
        ts = system()
        preview = request.form.get("preview") == "on"
        template_id = request.form.get("template_id") or None
        if not template_id:
            flash("テンプレートを選択してください", "error")
            return redirect(url_for("comparison_month_view", month=month))
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        prefix = "確認用" if preview else "正式"
        output = OUTPUT_DIR / f"{prefix}_相対表_{month}_{stamp}.xlsx"
        try:
            ts.export_comparison_month(month, int(template_id), output, preview)
        except ValueError as exc:
            flash(f"出力に失敗しました: {exc}", "error")
            return redirect(url_for("comparison_month_view", month=month))
        return redirect(url_for("download", filename=output.name))

    @app.route("/comparison-duplicates/delete", methods=["POST"])
    def comparison_duplicate_delete():
        ts = system()
        import_id = int(request.form["import_id"])
        sheet = request.form["sheet"]
        row_no = int(request.form["row_no"])
        ts.delete_comparison_record(import_id, sheet, row_no)
        flash("行を削除しました", "success")
        return redirect(url_for("library_view", kind="comparison"))

    @app.route("/comparison/fill-purchase", methods=["POST"])
    def comparison_fill_purchase():
        ts = system()
        month = request.form.get("month") or None
        result = ts.fill_comparison_purchase_from_ledger(month=month)
        if result["filled"]:
            flash(
                f"古物台帳から仕入側を{result['filled']}件埋めました"
                f"（一致する古物台帳が見つからなかったのは{result['not_found']}件です）", "success",
            )
        else:
            flash(f"埋められる行はありませんでした（{result['not_found']}件は一致する古物台帳が見つかりませんでした）", "success")
        if month:
            return redirect(url_for("comparison_month_view", month=month))
        return redirect(url_for("library_view", kind="comparison"))

    @app.route("/comparison-duplicates/delete-all", methods=["POST"])
    def comparison_duplicate_delete_all():
        ts = system()
        count = ts.delete_recommended_comparison_duplicates()
        if count:
            flash(f"削除推奨（古い方）の行を{count}件まとめて削除しました", "success")
        else:
            flash("削除対象の重複行はありませんでした", "success")
        return redirect(url_for("library_view", kind="comparison"))

    @app.route("/ledger-duplicates/delete", methods=["POST"])
    def ledger_duplicate_delete():
        ts = system()
        import_id = int(request.form["import_id"])
        row_no = int(request.form["row_no"])
        ts.delete_ledger_record(import_id, row_no)
        flash("行を削除しました", "success")
        return redirect(url_for("library_view", kind="ledger"))

    @app.route("/ledger-duplicates/delete-all", methods=["POST"])
    def ledger_duplicate_delete_all():
        ts = system()
        import_id = int(request.form["import_id"])
        count = ts.delete_recommended_ledger_duplicates(import_id)
        if count:
            flash(f"削除推奨（情報が少ない方）の行を{count}件まとめて削除しました", "success")
        else:
            flash("削除対象の重複行はありませんでした", "success")
        return redirect(url_for("library_view", kind="ledger"))

    @app.route("/ledger-duplicates/dismiss", methods=["POST"])
    def ledger_duplicate_dismiss():
        ts = system()
        import_id = int(request.form["import_id"])
        row_no = int(request.form["row_no"])
        ts.dismiss_ledger_duplicate(import_id, row_no)
        flash("重複ではないものとして、この組み合わせを一覧から外しました", "success")
        return redirect(url_for("library_view", kind="ledger"))

    @app.route("/templates/new", methods=["GET", "POST"])
    def new_template():
        if request.method == "POST":
            upload = request.files.get("file")
            if not upload or not upload.filename:
                flash("ファイルを選択してください", "error")
                return redirect(url_for("new_template"))
            with tempfile.TemporaryDirectory() as tmp:
                path = Path(tmp) / safe_upload_filename(upload.filename, "template")
                upload.save(path)
                try:
                    system().register_template(
                        request.form["report_type"], path, request.form["name"],
                        request.form["version"], request.form["effective_from"],
                    )
                except (ValueError, FileExistsError) as exc:
                    flash(f"登録に失敗しました: {exc}", "error")
                    return redirect(url_for("new_template"))
            flash("テンプレートを登録しました", "success")
            return redirect(url_for("index"))
        return render_template("new_template.html")

    @app.route("/import", methods=["GET", "POST"])
    def new_import():
        if request.method == "POST":
            uploads = [u for u in request.files.getlist("files") if u and u.filename]
            if not uploads:
                flash("ファイルを選択してください", "error")
                return redirect(url_for("new_import"))
            ts = system()
            inventory_as_of = request.form.get("inventory_as_of") or None
            results = []
            with tempfile.TemporaryDirectory() as tmp:
                for index, upload in enumerate(uploads):
                    path = Path(tmp) / safe_upload_filename(upload.filename, f"upload-{index}")
                    upload.save(path)
                    try:
                        result = ts.import_auto(path, inventory_as_of=inventory_as_of)
                        ts.rename_import_source(result["import_id"], upload.filename)
                        results.append({"filename": upload.filename, "ok": True, **result})
                    except ValueError as exc:
                        results.append({"filename": upload.filename, "ok": False, "error": str(exc)})
            if any(r.get("ok") and r.get("kind") == "ledger" for r in results):
                try:
                    merge_result = ts.auto_merge_ledger_imports()
                except ValueError:
                    pass  # fewer than 2 raw ledger imports exist yet — nothing to merge
                else:
                    if merge_result["duplicates"]:
                        flash(
                            f"古物台帳を結合しました（重複の可能性がある行が{len(merge_result['duplicates'])}件あります。"
                            "内容を確認してください）", "error",
                        )
            return render_template("import_result.html", results=results)
        return render_template("new_import.html")

    @app.route("/validate/<int:import_id>")
    def validate_import(import_id: int):
        ts = system()
        imp = ts.get_import(import_id)
        if not imp:
            abort(404)
        checks = ts.validate(import_id)
        templates = [t for t in ts.list_templates() if t["report_type"] == imp["kind"]] if imp["kind"] == "comparison" else []
        return render_template("validate.html", imp=imp, checks=checks, templates=templates,
                               has_errors=any(c.level == "error" for c in checks))

    @app.route("/records/<int:import_id>")
    def records_view(import_id: int):
        ts = system()
        imp = ts.get_import(import_id)
        if not imp:
            abort(404)
        per_page = 100
        page = max(1, request.args.get("page", 1, type=int))
        sheets = [s["name"] for s in imp["metadata"].get("sheets", [])] if imp["kind"] == "comparison" else []
        sheet = request.args.get("sheet") or (sheets[0] if sheets else None)
        if sheets and sheet not in sheets:
            sheet = sheets[0]
        months = ts.list_months(import_id) if imp["kind"] in ("ledger", "export_data", "comparison") else []
        month = request.args.get("month") if months else None
        sort = request.args.get("sort") or None
        sort_dir = "desc" if request.args.get("dir") == "desc" else "asc"
        row_no_list = request.args.getlist("row_no", type=int)
        row_no = row_no_list[0] if len(row_no_list) == 1 else (row_no_list or None)
        if row_no is None and months and month not in months:
            # only force a default month when browsing by month; a row_no view is
            # pinpointing specific rows and must not be narrowed by month as well
            month = months[-1]
        rows, total = ts.get_records(
            import_id, sheet=sheet, month=month, sort=sort, sort_dir=sort_dir, row_no=row_no,
            offset=(page - 1) * per_page, limit=per_page,
        )
        headers = next((s["headers"] for s in imp["metadata"].get("sheets", []) if s["name"] == sheet), None)
        display_headers = [h for h in headers if h] if headers else None
        if display_headers is not None:
            for row in rows:
                row["cells"] = [v for h, v in zip(headers, row["data"]["values"]) if h]
        total_pages = max(1, -(-total // per_page))
        sibling_imports = []
        if imp["kind"] == "comparison":
            sibling_imports = sorted(
                (i for i in ts.list_imports() if i["kind"] == "comparison"), key=lambda i: i["id"]
            )
            for i in sibling_imports:
                i["label"] = Path(i["source_name"]).stem
        purchase_blank = (
            imp["kind"] == "comparison" and isinstance(row_no, int)
            and len(rows) == 1 and rows[0].get("cells") and rows[0]["cells"][0] in (None, "")
        )
        query = request.args.get("q", "").strip()
        search_results = ts.search_ledger(query) if purchase_blank and query else []
        return render_template(
            "records.html", imp=imp, rows=rows, headers=display_headers,
            flat_columns=FLAT_COLUMNS.get(imp["kind"]),
            sheets=sheets, sheet=sheet, months=months, month=month,
            sort=sort, sort_dir=sort_dir, sibling_imports=sibling_imports, row_no=row_no,
            purchase_blank=purchase_blank, query=query, search_results=search_results,
            page=page, total=total, total_pages=total_pages,
        )

    @app.route("/comparison/link-purchase", methods=["POST"])
    def comparison_link_purchase():
        ts = system()
        import_id = int(request.form["import_id"])
        sheet = request.form["sheet"]
        row_no = int(request.form["row_no"])
        ledger_record_id = int(request.form["ledger_record_id"])
        try:
            qty = float(request.form["qty"])
            amount = float(request.form["amount"])
            ts.link_comparison_purchase_manually(import_id, sheet, row_no, ledger_record_id, qty, amount)
        except (ValueError, TypeError) as exc:
            flash(f"紐づけに失敗しました: {exc}", "error")
        else:
            flash("古物台帳の記録を手動で紐づけました", "success")
        return redirect(url_for("records_view", import_id=import_id, sheet=sheet, row_no=row_no))

    @app.route("/ledger-completion/<int:import_id>")
    def ledger_completion_view(import_id: int):
        ts = system()
        imp = ts.get_import(import_id)
        if not imp or imp["kind"] != "ledger":
            abort(404)
        month = request.args.get("month") or None
        months = ts.list_months(import_id)
        breakdown = ts.propose_ledger_breakdown(import_id, month=month)
        search_row = request.args.get("search_row", type=int)
        query = request.args.get("q", "").strip()
        search_row_month = next((b["month"] for b in breakdown if b["row_no"] == search_row), None)
        search_results = ts.search_inventory(query, search_row_month) if query and search_row_month else []
        all_resolved = bool(breakdown) and all(b["resolved"] for b in breakdown)
        suggestion = session.get("ledger_suggestion")
        if suggestion and suggestion.get("import_id") != import_id:
            suggestion = None
        return render_template(
            "ledger_completion.html", imp=imp, breakdown=breakdown, month=month, months=months,
            search_row=search_row, query=query, search_results=search_results, all_resolved=all_resolved,
            suggestion=suggestion,
        )

    @app.route("/ledger-completion/<int:import_id>/<int:row_no>/suggest", methods=["POST"])
    def ledger_completion_suggest(import_id: int, row_no: int):
        ts = system()
        month = request.form.get("month") or None
        items = ts.suggest_ledger_completion(import_id, row_no)
        if not items:
            flash("AIによる提案を取得できませんでした（APIキー未設定か、提案できる組み合わせが見つかりませんでした）", "error")
        session["ledger_suggestion"] = {"import_id": import_id, "row_no": row_no, "items": items or []}
        return redirect(url_for("ledger_completion_view", import_id=import_id, month=month))

    @app.route("/ledger-completion/<int:import_id>/<int:row_no>/add", methods=["POST"])
    def ledger_completion_add(import_id: int, row_no: int):
        ts = system()
        month = request.form.get("month") or None
        product = request.form.get("product", "").strip()
        qty = request.form.get("qty", "").strip()
        try:
            ts.add_ledger_item(import_id, row_no, product, float(qty))
        except (ValueError, TypeError) as exc:
            flash(f"追加に失敗しました: {exc}", "error")
        return redirect(url_for("ledger_completion_view", import_id=import_id, month=month))

    @app.route("/ledger-completion/<int:import_id>/<int:row_no>/remove/<int:item_id>", methods=["POST"])
    def ledger_completion_remove(import_id: int, row_no: int, item_id: int):
        ts = system()
        month = request.form.get("month") or None
        ts.remove_ledger_item(item_id)
        return redirect(url_for("ledger_completion_view", import_id=import_id, month=month))

    @app.route("/ledger-completion/<int:import_id>/export", methods=["POST"])
    def ledger_completion_export(import_id: int):
        ts = system()
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        output = OUTPUT_DIR / f"内訳復元_ledger_{import_id}_{stamp}.csv"
        try:
            ts.export_completed_ledger(import_id, output)
        except ValueError as exc:
            flash(f"出力に失敗しました: {exc}", "error")
            return redirect(url_for("ledger_completion_view", import_id=import_id))
        return redirect(url_for("download", filename=output.name))

    @app.route("/ledger-merge", methods=["POST"])
    def ledger_merge():
        ts = system()
        try:
            result = ts.auto_merge_ledger_imports()
        except ValueError as exc:
            flash(f"結合に失敗しました: {exc}", "error")
            return redirect(url_for("index"))
        if result["duplicates"]:
            flash(f"結合しました（{result['total']}件・重複の可能性がある行が{len(result['duplicates'])}件あります。内容を確認してください）", "error")
        else:
            flash(f"結合しました（{result['total']}件）", "success")
        return redirect(url_for("ledger_completion_view", import_id=result["import_id"]))

    @app.route("/export-entry", methods=["GET"])
    def export_entry():
        draft = session.get("export_draft")
        if not draft:
            return render_template("export_entry_start.html")
        return render_template("export_entry_items.html", draft=draft)

    @app.route("/export-entry/start", methods=["POST"])
    def export_entry_start():
        session["export_draft"] = {
            "date": request.form.get("date", "").strip(),
            "customer": request.form.get("customer", "").strip(),
            "payment": request.form.get("payment", "").strip(),
            "currency": request.form.get("currency", "JPY").strip(),
            "items": [],
        }
        return redirect(url_for("export_entry"))

    @app.route("/export-entry/add-item", methods=["POST"])
    def export_entry_add_item():
        draft = session.get("export_draft")
        if not draft:
            return redirect(url_for("export_entry"))
        product = request.form.get("product", "").strip()
        try:
            qty = float(request.form.get("qty", ""))
            price = float(request.form.get("price", ""))
        except ValueError:
            flash("数量・単価は数値で入力してください", "error")
            return redirect(url_for("export_entry"))
        if not product:
            flash("品名を入力してください", "error")
            return redirect(url_for("export_entry"))
        eng = translate_ja_to_en(product)
        if eng is None:
            flash("英語名の自動翻訳に失敗しました。英語名は手動で入力してください", "error")
        draft["items"].append({"product": product, "eng": eng or "", "qty": qty, "price": price, "subtotal": qty * price})
        session["export_draft"] = draft
        return redirect(url_for("export_entry"))

    @app.route("/export-entry/edit-item/<int:index>", methods=["POST"])
    def export_entry_edit_item(index: int):
        draft = session.get("export_draft")
        if not draft or not (0 <= index < len(draft["items"])):
            return redirect(url_for("export_entry"))
        draft["items"][index]["eng"] = request.form.get("eng", "").strip()
        session["export_draft"] = draft
        return redirect(url_for("export_entry"))

    @app.route("/export-entry/remove-item/<int:index>", methods=["POST"])
    def export_entry_remove_item(index: int):
        draft = session.get("export_draft")
        if draft and 0 <= index < len(draft["items"]):
            draft["items"].pop(index)
            session["export_draft"] = draft
        return redirect(url_for("export_entry"))

    @app.route("/export-entry/cancel", methods=["POST"])
    def export_entry_cancel():
        session.pop("export_draft", None)
        return redirect(url_for("index"))

    @app.route("/export-entry/finish", methods=["POST"])
    def export_entry_finish():
        draft = session.get("export_draft")
        if not draft or not draft["items"]:
            flash("商品を1件以上追加してください", "error")
            return redirect(url_for("export_entry"))
        ts = system()
        rows = [{
            "年月日": draft["date"], "品名": item["product"], "金額": item["price"],
            "数量": item["qty"], "小計": item["subtotal"], "相手方名": draft["customer"],
            "支払方法": draft["payment"], "通貨": draft["currency"], "英語名": item["eng"],
        } for item in draft["items"]]
        source_name = f"手入力（{draft['date']}・{draft['customer']}）"
        import_id = ts.record_export_entry(rows, source_name)
        session.pop("export_draft", None)
        return redirect(url_for("validate_import", import_id=import_id))

    @app.route("/build-comparison/<int:import_id>", methods=["GET", "POST"])
    def build_comparison_view(import_id: int):
        ts = system()
        imp = ts.get_import(import_id)
        if not imp or imp["kind"] != "export_data":
            abort(404)
        templates = [t for t in ts.list_templates() if t["report_type"] == "comparison"]
        if request.method == "POST":
            template_id = request.form.get("template_id")
            if not template_id:
                flash("テンプレートを選択してください", "error")
                return redirect(url_for("build_comparison_view", import_id=import_id))
            try:
                result = ts.build_comparison(import_id, int(template_id))
            except ValueError as exc:
                flash(f"組み立てに失敗しました: {exc}", "error")
                return redirect(url_for("build_comparison_view", import_id=import_id))
            flash(f"相対表を組み立てました（全{result['total']}件中、対応する仕入が見つからなかった行: {result['unmatched']}件）",
                  "success" if result["unmatched"] == 0 else "error")
            return redirect(url_for("validate_import", import_id=result["import_id"]))
        return render_template("build_comparison.html", imp=imp, templates=templates)

    @app.route("/allocations/<int:import_id>")
    def allocations_view(import_id: int):
        ts = system()
        imp = ts.get_import(import_id)
        if not imp or imp["kind"] != "comparison":
            abort(404)
        results = ts.allocate(import_id)
        search_sheet = request.args.get("search_sheet")
        search_row = request.args.get("search_row", type=int)
        query = request.args.get("q", "").strip()
        search_results = ts.search_ledger(query) if query else []
        return render_template(
            "allocations.html", imp=imp, results=results,
            search_sheet=search_sheet, search_row=search_row, query=query, search_results=search_results,
        )

    @app.route("/allocations/<int:import_id>/<sheet>/<int:row_no>", methods=["POST"])
    def allocations_set(import_id: int, sheet: str, row_no: int):
        ts = system()
        raw = request.form.get("ledger_record_id")
        ts.set_manual_allocation(import_id, sheet, row_no, int(raw) if raw else None)
        flash("仕入との対応を保存しました", "success")
        return redirect(url_for("allocations_view", import_id=import_id))

    @app.route("/export/<int:import_id>", methods=["POST"])
    def export_import(import_id: int):
        ts = system()
        imp = ts.get_import(import_id)
        if not imp:
            abort(404)
        preview = request.form.get("preview") == "on"
        template_id = request.form.get("template_id") or None
        month = request.form.get("month") or None
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        suffix = ".csv" if imp["kind"] in ("ledger", "inventory") else ".xlsx"
        prefix = "確認用" if preview else "正式"
        name_part = f"{imp['kind']}_{import_id}" + (f"_{month}" if month else "")
        output = OUTPUT_DIR / f"{prefix}_{name_part}_{stamp}{suffix}"
        try:
            ts.export(import_id, output, int(template_id) if template_id else None, preview, month=month)
        except ValueError as exc:
            flash(f"出力に失敗しました: {exc}", "error")
            if month:
                return redirect(url_for("records_view", import_id=import_id, month=month))
            return redirect(url_for("validate_import", import_id=import_id))
        return redirect(url_for("download", filename=output.name))

    @app.route("/imports/<int:import_id>/delete", methods=["GET", "POST"])
    def delete_import_view(import_id: int):
        ts = system()
        imp = ts.get_import(import_id)
        if not imp:
            abort(404)
        if request.method == "POST":
            ts.delete_import(import_id)
            flash(f"取込ID {import_id}（{imp['source_name']}）を削除しました", "success")
            return redirect(url_for("index"))
        _, total = ts.get_records(import_id, limit=1)
        return render_template("delete_import.html", imp=imp, total=total)

    @app.route("/imports/delete-many/confirm", methods=["POST"])
    def delete_imports_confirm():
        ts = system()
        import_ids = [int(v) for v in request.form.getlist("import_ids")]
        kind = request.form.get("kind", "")
        imports = [imp for imp in (ts.get_import(i) for i in import_ids) if imp]
        back_url = url_for("library_view", kind=kind) if kind in KIND_LABELS else url_for("index")
        if not imports:
            flash("削除する取込が選択されていません", "error")
            return redirect(back_url)
        return render_template("delete_imports.html", imports=imports, kind=kind, back_url=back_url)

    @app.route("/imports/delete-many", methods=["POST"])
    def delete_imports_view():
        ts = system()
        import_ids = [int(v) for v in request.form.getlist("import_ids")]
        kind = request.form.get("kind", "")
        count = ts.delete_imports(import_ids)
        flash(f"{count}件の取込を削除しました", "success")
        return redirect(url_for("library_view", kind=kind) if kind in KIND_LABELS else url_for("index"))

    @app.route("/download/<path:filename>")
    def download(filename: str):
        target = (OUTPUT_DIR / filename).resolve()
        if OUTPUT_DIR not in target.parents or not target.is_file():
            abort(404)
        return send_file(target, as_attachment=True)

    return app


def main() -> None:
    import argparse

    parser = argparse.ArgumentParser(prog="tax-system-web", description="tax-system ローカルWeb画面")
    parser.add_argument("--home", default=os.environ.get("TAX_SYSTEM_HOME", ".tax-system"))
    parser.add_argument("--port", type=int, default=5000)
    args = parser.parse_args()
    app = create_app(args.home)
    print(f"ブラウザで http://127.0.0.1:{args.port} を開いてください")
    app.run(host="127.0.0.1", port=args.port, debug=False)


if __name__ == "__main__":
    main()
