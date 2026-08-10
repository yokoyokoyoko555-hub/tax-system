from __future__ import annotations

import os
import shutil
import tempfile
from datetime import datetime
from pathlib import Path

from flask import Flask, abort, flash, redirect, render_template, request, send_file, url_for
from werkzeug.utils import secure_filename

from .core import TaxSystem

OUTPUT_DIR = Path("outputs").resolve()


def create_app(home: str | Path | None = None) -> Flask:
    app = Flask(__name__)
    app.secret_key = os.environ.get("TAX_SYSTEM_SECRET", "local-only-tax-system")
    app.config["TAX_SYSTEM_HOME"] = home or os.environ.get("TAX_SYSTEM_HOME", ".tax-system")

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
                    import_id = ts.import_ledger(path) if report_type == "ledger" else ts.import_comparison(path)
                except ValueError as exc:
                    flash(f"取込に失敗しました: {exc}", "error")
                    return redirect(url_for("new_import"))
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
