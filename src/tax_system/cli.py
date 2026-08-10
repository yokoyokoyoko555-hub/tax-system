from __future__ import annotations

import argparse
import json
import os
import sys

from .core import TaxSystem


def parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(prog="tax-system", description="古物台帳・相対表 帳票管理")
    p.add_argument("--home", default=os.environ.get("TAX_SYSTEM_HOME", ".tax-system"),
                   help="機密データ保存先（Git管理外）。環境変数 TAX_SYSTEM_HOME でも指定可")
    sub = p.add_subparsers(dest="command", required=True)
    sub.add_parser("init")
    reg = sub.add_parser("register-template")
    reg.add_argument("--type", choices=["ledger", "comparison"], required=True); reg.add_argument("--file", required=True)
    reg.add_argument("--name", required=True); reg.add_argument("--version", required=True); reg.add_argument("--effective-from", required=True)
    imp = sub.add_parser("import")
    imp.add_argument("--type", choices=["ledger", "comparison"], required=True); imp.add_argument("--file", required=True)
    val = sub.add_parser("validate"); val.add_argument("import_id", type=int)
    exp = sub.add_parser("export"); exp.add_argument("import_id", type=int); exp.add_argument("--output", required=True)
    exp.add_argument("--template-id", type=int); exp.add_argument("--preview", action="store_true")
    return p


def main() -> None:
    args = parser().parse_args(); app = TaxSystem(args.home)
    try:
        if args.command == "init": app.initialize(); print("初期化しました")
        elif args.command == "register-template": print(app.register_template(args.type, args.file, args.name, args.version, args.effective_from))
        elif args.command == "import": print(app.import_ledger(args.file) if args.type == "ledger" else app.import_comparison(args.file))
        elif args.command == "validate":
            checks = app.validate(args.import_id); print(json.dumps([c.__dict__ for c in checks], ensure_ascii=False, indent=2)); sys.exit(1 if any(c.level == "error" for c in checks) else 0)
        elif args.command == "export": print(app.export(args.import_id, args.output, args.template_id, args.preview))
    except Exception as exc:
        print(f"エラー: {exc}", file=sys.stderr); sys.exit(2)


if __name__ == "__main__": main()

