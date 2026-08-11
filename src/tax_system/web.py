from __future__ import annotations

import os
import shutil
import tempfile
from datetime import datetime
from pathlib import Path

from flask import Flask, abort, flash, redirect, render_template, request, send_file, url_for
from werkzeug.utils import secure_filename

from .core import EXPORT_DATA_COLUMNS, INVENTORY_COLUMNS, LEDGER_COLUMNS, TaxSystem

KIND_LABELS = {"ledger": "古物台帳", "comparison": "相対表", "inventory": "期末在庫表", "export_data": "輸出データ"}
FLAT_COLUMNS = {"ledger": LEDGER_COLUMNS, "inventory": INVENTORY_COLUMNS, "export_data": EXPORT_DATA_COLUMNS}

OUTPUT_DIR = Path("outputs").resolve()


def create_app(home: str | Path | None = None) -> Flask:
    app = Flask(__name__)
    app.secret_key = os.environ.get("TAX_SYSTEM_SECRET", "local-only-tax-system")
    app.config["TAX_SYSTEM_HOME"] = home or os.environ.get("TAX_SYSTEM_HOME", ".tax-system")
    app.jinja_env.globals["kind_labels"] = KIND_LABELS

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

    @app.route("/templates/new", methods=["GET", "POST"])
    def new_template():
        if request.method == "POST":
            upload = request.files.get("file")
            if not upload or not upload.filename:
                flash("ファイルを選択してください", "error")
                return redirect(url_for("new_template"))
            with tempfile.TemporaryDirectory() as tmp:
                path = Path(tmp) / secure_filename(upload.filename)
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
            upload = request.files.get("file")
            if not upload or not upload.filename:
                flash("ファイルを選択してください", "error")
                return redirect(url_for("new_import"))
            report_type = request.form["report_type"]
            with tempfile.TemporaryDirectory() as tmp:
                path = Path(tmp) / secure_filename(upload.filename)
                upload.save(path)
                try:
                    ts = system()
                    if report_type == "ledger":
                        import_id = ts.import_ledger(path)
                    elif report_type == "inventory":
                        import_id = ts.import_inventory(path)
                    elif report_type == "export_data":
                        import_id = ts.import_export_data(path)
                    else:
                        import_id = ts.import_comparison(path)
                except ValueError as exc:
                    flash(f"取込に失敗しました: {exc}", "error")
                    return redirect(url_for("new_import"))
            if report_type == "inventory":
                return redirect(url_for("index"))
            return redirect(url_for("validate_import", import_id=import_id))
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
        rows, total = ts.get_records(import_id, sheet=sheet, offset=(page - 1) * per_page, limit=per_page)
        headers = next((s["headers"] for s in imp["metadata"].get("sheets", []) if s["name"] == sheet), None)
        display_headers = [h for h in headers if h] if headers else None
        if display_headers is not None:
            for row in rows:
                row["cells"] = [v for h, v in zip(headers, row["data"]["values"]) if h]
        total_pages = max(1, -(-total // per_page))
        return render_template(
            "records.html", imp=imp, rows=rows, headers=display_headers,
            flat_columns=FLAT_COLUMNS.get(imp["kind"]),
            sheets=sheets, sheet=sheet, page=page, total=total, total_pages=total_pages,
        )

    @app.route("/ledger-completion/<int:import_id>")
    def ledger_completion_view(import_id: int):
        ts = system()
        imp = ts.get_import(import_id)
        if not imp or imp["kind"] != "ledger":
            abort(404)
        breakdown = ts.propose_ledger_breakdown(import_id)
        search_row = request.args.get("search_row", type=int)
        query = request.args.get("q", "").strip()
        search_results = ts.search_inventory(query) if query else []
        all_resolved = bool(breakdown) and all(b["resolved"] for b in breakdown)
        return render_template(
            "ledger_completion.html", imp=imp, breakdown=breakdown,
            search_row=search_row, query=query, search_results=search_results, all_resolved=all_resolved,
        )

    @app.route("/ledger-completion/<int:import_id>/<int:row_no>/add", methods=["POST"])
    def ledger_completion_add(import_id: int, row_no: int):
        ts = system()
        product = request.form.get("product", "").strip()
        qty = request.form.get("qty", "").strip()
        try:
            ts.add_ledger_item(import_id, row_no, product, float(qty))
        except (ValueError, TypeError) as exc:
            flash(f"追加に失敗しました: {exc}", "error")
        return redirect(url_for("ledger_completion_view", import_id=import_id))

    @app.route("/ledger-completion/<int:import_id>/<int:row_no>/remove/<int:item_id>", methods=["POST"])
    def ledger_completion_remove(import_id: int, row_no: int, item_id: int):
        ts = system()
        ts.remove_ledger_item(item_id)
        return redirect(url_for("ledger_completion_view", import_id=import_id))

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
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        suffix = ".csv" if imp["kind"] == "ledger" else ".xlsx"
        prefix = "確認用" if preview else "正式"
        output = OUTPUT_DIR / f"{prefix}_{imp['kind']}_{import_id}_{stamp}{suffix}"
        try:
            ts.export(import_id, output, int(template_id) if template_id else None, preview)
        except ValueError as exc:
            flash(f"出力に失敗しました: {exc}", "error")
            return redirect(url_for("validate_import", import_id=import_id))
        return redirect(url_for("download", filename=output.name))

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
