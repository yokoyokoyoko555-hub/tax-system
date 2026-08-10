import csv
import tempfile
import unittest
from pathlib import Path

from tax_system.core import LEDGER_COLUMNS, TaxSystem


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


if __name__ == "__main__": unittest.main()

