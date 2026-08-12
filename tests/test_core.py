import csv
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

from tax_system.core import (
    EXPORT_DATA_COLUMNS,
    INVENTORY_COLUMNS,
    LEDGER_COLUMNS,
    LEDGER_IDENTITY_COLUMNS,
    LEDGER_POS_COLUMNS,
    TaxSystem,
)


class LedgerTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)
        self.app = TaxSystem(self.root / "runtime")

    def tearDown(self):
        self.tmp.cleanup()

    def write_csv(self, row):
        path = self.root / "ledger.csv"
        with path.open("w", encoding="utf-8-sig", newline="") as fh:
            writer = csv.DictWriter(fh, fieldnames=LEDGER_COLUMNS, lineterminator="\r\n")
            writer.writeheader(); writer.writerow(row)
        return path

    def test_valid_ledger_roundtrip(self):
        row = dict(zip(LEDGER_COLUMNS, ["2026-01-01", "テスト", "てすと", "2000-01-01", "匿名住所", "000", "商品", "2", "100", "200", ""])); path = self.write_csv(row)
        import_id = self.app.import_ledger(path)
        self.assertEqual([], self.app.validate(import_id))
        output = self.root / "out.csv"; self.app.export(import_id, output)
        self.assertTrue(output.read_bytes().startswith(b"\xef\xbb\xbf"))
        self.assertIn(b"\r\n", output.read_bytes())

    def test_error_blocks_formal_export(self):
        row = dict(zip(LEDGER_COLUMNS, ["2026-01-01", "", "", "", "", "", "商品", "2", "100", "999", ""])); path = self.write_csv(row)
        import_id = self.app.import_ledger(path)
        with self.assertRaises(ValueError): self.app.export(import_id, self.root / "blocked.csv")
        self.app.export(import_id, self.root / "preview.csv", preview=True)


class AllocationTests(unittest.TestCase):
    PURCHASE_HEADERS = ["年月日", "品目", "数量", "相手方名", "代価"]
    SALE_HEADERS = ["年月日", "数量"]

    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)
        self.app = TaxSystem(self.root / "runtime")

    def tearDown(self):
        self.tmp.cleanup()

    def write_ledger(self, row):
        path = self.root / "ledger.csv"
        with path.open("w", encoding="utf-8-sig", newline="") as fh:
            writer = csv.DictWriter(fh, fieldnames=LEDGER_COLUMNS, lineterminator="\r\n")
            writer.writeheader(); writer.writerow(row)
        return path

    def write_comparison(self, purchase_row, sale_row, name="comparison.xlsx"):
        wb = Workbook()
        ws = wb.active
        ws.title = "輸出販売"
        headers = self.PURCHASE_HEADERS + self.SALE_HEADERS
        for col, value in enumerate(headers, 1):
            ws.cell(2, col).value = value
        for col, value in enumerate(purchase_row + sale_row, 1):
            ws.cell(3, col).value = value
        path = self.root / name
        wb.save(path)
        return path

    def test_matching_purchase_is_found_automatically(self):
        ledger_row = dict(zip(LEDGER_COLUMNS, [
            "2026-05-01", "山田太郎", "やまだたろう", "1990-01-01", "匿名住所", "000",
            "テストカード", "2", "10000", "20000", "",
        ]))
        self.app.import_ledger(self.write_ledger(ledger_row))
        comparison_id = self.app.import_comparison(
            self.write_comparison(["2026-05-01", "テストカード", 2, "山田太郎", 20000], ["2026-06-01", 2])
        )
        self.assertEqual([], self.app.validate(comparison_id))

    def test_missing_purchase_blocks_export(self):
        ledger_row = dict(zip(LEDGER_COLUMNS, [
            "2026-05-01", "山田太郎", "やまだたろう", "1990-01-01", "匿名住所", "000",
            "テストカード", "2", "10000", "20000", "",
        ]))
        self.app.import_ledger(self.write_ledger(ledger_row))
        comparison_id = self.app.import_comparison(
            self.write_comparison(["2026-05-01", "別の商品", 2, "山田太郎", 20000], ["2026-06-01", 2])
        )
        checks = self.app.validate(comparison_id)
        self.assertTrue(any(c.code == "ALLOCATION_NOT_FOUND" for c in checks))
        with self.assertRaises(ValueError):
            self.app.export(comparison_id, self.root / "blocked.xlsx", template_id=None)

    def test_ledger_record_is_not_reused_for_two_sales(self):
        ledger_row = dict(zip(LEDGER_COLUMNS, [
            "2026-05-01", "山田太郎", "やまだたろう", "1990-01-01", "匿名住所", "000",
            "テストカード", "2", "10000", "20000", "",
        ]))
        self.app.import_ledger(self.write_ledger(ledger_row))
        first_id = self.app.import_comparison(
            self.write_comparison(["2026-05-01", "テストカード", 2, "山田太郎", 20000], ["2026-06-01", 2], name="first.xlsx")
        )
        second_id = self.app.import_comparison(
            self.write_comparison(["2026-05-01", "テストカード", 2, "山田太郎", 20000], ["2026-06-02", 2], name="second.xlsx")
        )
        self.assertEqual([], self.app.validate(first_id))
        checks = self.app.validate(second_id)
        self.assertTrue(any(c.code == "ALLOCATION_NOT_FOUND" for c in checks))


class LedgerCompletionTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)
        self.app = TaxSystem(self.root / "runtime")

    def tearDown(self):
        self.tmp.cleanup()

    def write_inventory(self):
        path = self.root / "inventory.csv"
        with path.open("w", encoding="utf-8-sig", newline="") as fh:
            writer = csv.DictWriter(fh, fieldnames=INVENTORY_COLUMNS, lineterminator="\r\n")
            writer.writeheader()
            writer.writerow({"商品名": "テストカードA", "仕入れ原価": "5000", "在庫数": "10"})
            writer.writerow({"商品名": "テストカードB", "仕入れ原価": "3000", "在庫数": "5"})
        return path

    def write_ledger_lump(self, total):
        path = self.root / "ledger.csv"
        row = dict(zip(LEDGER_COLUMNS, [
            "2026-05-01", "鈴木一郎", "すずきいちろう", "1985-01-01", "匿名住所", "000",
            "", "", "", str(total), "",
        ]))
        with path.open("w", encoding="utf-8-sig", newline="") as fh:
            writer = csv.DictWriter(fh, fieldnames=LEDGER_COLUMNS, lineterminator="\r\n")
            writer.writeheader(); writer.writerow(row)
        return path

    def write_comparison(self, rows, name="comparison.xlsx"):
        wb = Workbook()
        ws = wb.active
        ws.title = "輸出販売"
        headers = ["年月日", "品目", "数量", "相手方名", "代価", "年月日", "数量"]
        for col, value in enumerate(headers, 1):
            ws.cell(2, col).value = value
        for r, (purchase_row, sale_row) in enumerate(rows, 3):
            for col, value in enumerate(purchase_row + sale_row, 1):
                ws.cell(r, col).value = value
        path = self.root / name
        wb.save(path)
        return path

    def test_breakdown_uses_comparison_and_manual_fill(self):
        self.app.import_inventory(self.write_inventory())
        ledger_id = self.app.import_ledger(self.write_ledger_lump(11000))
        self.app.import_comparison(self.write_comparison([
            (["2026-05-01", "テストカードA", 1, "鈴木一郎", 5000], ["2026-06-01", 1]),
            (["2026-05-01", "テストカードB", 1, "鈴木一郎", 3000], ["2026-06-02", 1]),
        ]))

        breakdown = self.app.propose_ledger_breakdown(ledger_id)
        self.assertEqual(1, len(breakdown))
        entry = breakdown[0]
        self.assertEqual(2, len(entry["known_items"]))
        self.assertAlmostEqual(8000, sum(i["amount"] for i in entry["known_items"]))
        self.assertFalse(entry["resolved"])
        self.assertAlmostEqual(3000, entry["remainder"])

        self.app.add_ledger_item(ledger_id, entry["row_no"], "テストカードB", 1)
        resolved_breakdown = self.app.propose_ledger_breakdown(ledger_id)
        self.assertTrue(resolved_breakdown[0]["resolved"])

        output = self.root / "completed.csv"
        self.app.export_completed_ledger(ledger_id, output)
        with output.open(encoding="utf-8-sig") as fh:
            rows = list(csv.DictReader(fh))
        self.assertEqual(3, len(rows))
        self.assertAlmostEqual(11000, sum(float(r["金額"]) for r in rows))

    def test_export_blocked_when_unresolved(self):
        self.app.import_inventory(self.write_inventory())
        ledger_id = self.app.import_ledger(self.write_ledger_lump(11000))
        self.app.import_comparison(self.write_comparison([
            (["2026-05-01", "テストカードA", 1, "鈴木一郎", 5000], ["2026-06-01", 1]),
        ]))
        with self.assertRaises(ValueError):
            self.app.export_completed_ledger(ledger_id, self.root / "blocked.csv")


class BuildComparisonTests(unittest.TestCase):
    HEADERS = ["年月日", "品目", "数量", "相手方名", "代価", "年月日", "数量"]

    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)
        self.app = TaxSystem(self.root / "runtime")

    def tearDown(self):
        self.tmp.cleanup()

    def write_ledger(self, rows):
        path = self.root / "ledger.csv"
        with path.open("w", encoding="utf-8-sig", newline="") as fh:
            writer = csv.DictWriter(fh, fieldnames=LEDGER_COLUMNS, lineterminator="\r\n")
            writer.writeheader()
            for row in rows:
                writer.writerow(dict(zip(LEDGER_COLUMNS, row)))
        return path

    def write_export_data(self, rows):
        path = self.root / "export_data.csv"
        with path.open("w", encoding="utf-8-sig", newline="") as fh:
            writer = csv.DictWriter(fh, fieldnames=EXPORT_DATA_COLUMNS, lineterminator="\r\n")
            writer.writeheader()
            for row in rows:
                writer.writerow(dict(zip(EXPORT_DATA_COLUMNS, row)))
        return path

    def register_template(self):
        wb = Workbook()
        ws = wb.active
        ws.title = "輸出販売"
        for col, value in enumerate(self.HEADERS, 1):
            ws.cell(2, col).value = value
        path = self.root / "template.xlsx"
        wb.save(path)
        return self.app.register_template("comparison", path, "test", "v1", "2026-01-01")

    def test_fifo_matches_oldest_purchase_first(self):
        self.app.import_ledger(self.write_ledger([
            ["2026-01-01", "Aさん", "えー", "1990-01-01", "匿名住所", "000", "カードX", "1", "1000", "1000", ""],
            ["2026-02-01", "Bさん", "びー", "1991-01-01", "匿名住所", "000", "カードX", "1", "1200", "1200", ""],
        ]))
        export_id = self.app.import_export_data(self.write_export_data([
            ["2026-03-01", "カードX", "2000", "1", "2000", "海外顧客1", "銀行振込", "JPY"],
            ["2026-03-02", "カードX", "2200", "1", "2200", "海外顧客2", "銀行振込", "JPY"],
        ]))
        template_id = self.register_template()

        result = self.app.build_comparison(export_id, template_id)
        self.assertEqual(2, result["total"])
        self.assertEqual(0, result["unmatched"])

        self.assertEqual([], self.app.validate(result["import_id"]))

        allocations = self.app.allocate(result["import_id"])
        by_row = {a["row_no"]: a for a in allocations}
        self.assertEqual("Aさん", by_row[2]["ledger"]["name"])
        self.assertEqual("Bさん", by_row[3]["ledger"]["name"])

    def test_unmatched_row_flagged_for_manual_allocation(self):
        self.app.import_ledger(self.write_ledger([
            ["2026-01-01", "Aさん", "えー", "1990-01-01", "匿名住所", "000", "カードY", "1", "1000", "1000", ""],
        ]))
        export_id = self.app.import_export_data(self.write_export_data([
            ["2026-03-01", "カードZ", "2000", "1", "2000", "海外顧客1", "銀行振込", "JPY"],
        ]))
        template_id = self.register_template()

        result = self.app.build_comparison(export_id, template_id)
        self.assertEqual(1, result["unmatched"])

        checks = self.app.validate(result["import_id"])
        self.assertTrue(any(c.code == "ALLOCATION_NOT_FOUND" for c in checks))

    def test_manual_export_entry_feeds_build_comparison(self):
        self.app.import_ledger(self.write_ledger([
            ["2026-01-01", "Aさん", "えー", "1990-01-01", "匿名住所", "000", "カードX", "1", "1000", "1000", ""],
        ]))
        export_id = self.app.record_export_entry([{
            "年月日": "2026-03-01", "品名": "カードX", "金額": 2000, "数量": 1, "小計": 2000,
            "相手方名": "海外顧客1", "支払方法": "銀行振込", "通貨": "JPY", "英語名": "Card X",
        }], "手入力（テスト）")
        template_id = self.register_template()

        result = self.app.build_comparison(export_id, template_id)
        self.assertEqual(1, result["total"])
        self.assertEqual(0, result["unmatched"])
        self.assertEqual([], self.app.validate(result["import_id"]))


class RawLedgerImportTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)
        self.app = TaxSystem(self.root / "runtime")

    def tearDown(self):
        self.tmp.cleanup()

    def write_pos_csv(self, rows):
        path = self.root / "pos.csv"
        with path.open("w", encoding="utf-8-sig", newline="") as fh:
            writer = csv.DictWriter(fh, fieldnames=LEDGER_POS_COLUMNS, lineterminator="\r\n")
            writer.writeheader()
            for row in rows:
                writer.writerow(dict(zip(LEDGER_POS_COLUMNS, row)))
        return path

    def write_identity_xlsx(self, rows):
        wb = Workbook()
        ws = wb.active
        for col, value in enumerate(LEDGER_IDENTITY_COLUMNS, 1):
            ws.cell(1, col).value = value
        for r, row in enumerate(rows, 2):
            for col, value in enumerate(row, 1):
                ws.cell(r, col).value = value
        path = self.root / "identity.xlsx"
        wb.save(path)
        return path

    def test_pos_import_maps_fields_and_skips_unapproved(self):
        result = self.app.import_ledger_pos(self.write_pos_csv([
            ["1", "承認済み", "2026-04-01 10:00:00", "1", "テスト太郎", "明細なし", "カードA", "1", "1000", "1000", "備考1", "全体1"],
            ["2", "却下", "2026-04-02 10:00:00", "1", "テスト次郎", "明細なし", "カードB", "1", "500", "500", "", ""],
        ]))
        self.assertEqual(1, result["imported"])
        self.assertEqual(1, result["skipped"])
        records, _ = self.app.get_records(result["import_id"])
        data = records[0]["data"]
        self.assertEqual("テスト太郎", data["名前"])
        self.assertEqual("カードA", data["商品名"])
        self.assertEqual("", data["ふりがな"])
        self.assertIn("備考1", data["備考"])
        self.assertIn("全体1", data["備考"])

    def test_identity_import_maps_fields_and_ignores_trailing_blank_column(self):
        result = self.app.import_ledger_identity(self.write_identity_xlsx([
            ["2026-04-01", "テスト太郎", "てすとたろう", "1990-01-01", "匿名住所", "000-0000", 5000],
        ]))
        self.assertEqual(1, result["imported"])
        records, _ = self.app.get_records(result["import_id"])
        data = records[0]["data"]
        self.assertEqual("てすとたろう", data["ふりがな"])
        self.assertEqual("", data["商品名"])
        self.assertEqual(5000, data["金額"])


class AutoImportTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)
        self.app = TaxSystem(self.root / "runtime")

    def tearDown(self):
        self.tmp.cleanup()

    def write_csv(self, columns, rows, name):
        path = self.root / name
        with path.open("w", encoding="utf-8-sig", newline="") as fh:
            writer = csv.DictWriter(fh, fieldnames=columns, lineterminator="\r\n")
            writer.writeheader()
            for row in rows:
                writer.writerow(dict(zip(columns, row)))
        return path

    def test_detects_ledger_csv(self):
        path = self.write_csv(LEDGER_COLUMNS, [
            ["2026-04-01", "A", "えー", "1990-01-01", "住所A", "000", "商品A", "1", "1000", "1000", ""],
        ], "unknown.csv")
        result = self.app.import_auto(path)
        self.assertEqual("ledger", result["kind"])

    def test_detects_inventory_csv(self):
        path = self.write_csv(INVENTORY_COLUMNS, [["商品X", "1000", "5"]], "unknown2.csv")
        result = self.app.import_auto(path)
        self.assertEqual("inventory", result["kind"])

    def test_detects_export_data_csv(self):
        path = self.write_csv(EXPORT_DATA_COLUMNS, [
            ["2026-05-01", "品X", "1000", "1", "1000", "海外顧客", "銀行振込", "JPY"],
        ], "unknown3.csv")
        result = self.app.import_auto(path)
        self.assertEqual("export_data", result["kind"])

    def test_detects_pos_csv(self):
        path = self.write_csv(LEDGER_POS_COLUMNS, [
            ["1", "承認済み", "2026-04-01 10:00:00", "1", "テスト太郎", "明細なし", "カードA", "1", "1000", "1000", "", ""],
        ], "unknown4.csv")
        result = self.app.import_auto(path)
        self.assertEqual("ledger", result["kind"])

    def test_unrecognized_csv_raises(self):
        path = self.write_csv(["列A", "列B"], [["1", "2"]], "unknown5.csv")
        with self.assertRaises(ValueError):
            self.app.import_auto(path)


class MergeLedgerTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)
        self.app = TaxSystem(self.root / "runtime")

    def tearDown(self):
        self.tmp.cleanup()

    def write_ledger(self, rows, name):
        path = self.root / name
        with path.open("w", encoding="utf-8-sig", newline="") as fh:
            writer = csv.DictWriter(fh, fieldnames=LEDGER_COLUMNS, lineterminator="\r\n")
            writer.writeheader()
            for row in rows:
                writer.writerow(dict(zip(LEDGER_COLUMNS, row)))
        return path

    def test_merge_allowed_even_with_unresolved_errors(self):
        first = self.app.import_ledger(self.write_ledger([
            ["2026-04-01", "A", "えー", "1990-01-01", "住所A", "000", "商品A", "1", "1000", "1000", ""],
        ], "a.csv"))
        second = self.app.import_ledger(self.write_ledger([
            ["2026-04-02", "", "", "", "", "", "商品B", "1", "2000", "2000", ""],
        ], "b.csv"))
        result = self.app.merge_ledger_imports([first, second])
        self.assertEqual(2, result["total"])
        checks = self.app.validate(result["import_id"])
        self.assertTrue(any(c.code == "REQUIRED" for c in checks))
        self.assertEqual(3, len(self.app.list_imports()))  # originals kept, plus the new merged one

    def test_merge_combines_rows_and_flags_duplicates(self):
        first = self.app.import_ledger(self.write_ledger([
            ["2026-04-01", "A", "えー", "1990-01-01", "住所A", "000", "商品A", "1", "1000", "1000", ""],
        ], "a.csv"))
        second = self.app.import_ledger(self.write_ledger([
            ["2026-04-01", "A", "えー", "1990-01-01", "住所A", "000", "商品A", "1", "1000", "1000", ""],
            ["2026-04-03", "B", "びー", "1991-01-01", "住所B", "000", "商品C", "1", "3000", "3000", ""],
        ], "b.csv"))
        result = self.app.merge_ledger_imports([first, second])
        self.assertEqual(3, result["total"])
        self.assertEqual(1, len(result["duplicates"]))

        # the intentional duplicate row is also caught by the normal ledger validation
        checks = self.app.validate(result["import_id"])
        self.assertTrue(any(c.code == "DUPLICATE" for c in checks))

        output = self.app.export(result["import_id"], self.root / "merged.csv", preview=True)
        with output.open(encoding="utf-8-sig") as fh:
            rows = list(csv.DictReader(fh))
        self.assertEqual(3, len(rows))


if __name__ == "__main__": unittest.main()

