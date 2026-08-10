import csv
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

from tax_system.core import INVENTORY_COLUMNS, LEDGER_COLUMNS, TaxSystem


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


if __name__ == "__main__": unittest.main()

