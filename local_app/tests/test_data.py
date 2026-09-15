import csv
import io
import json
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from data import identifier, load_selection, numeric_field, records, export_zip, ad_records


class Sheet:
    def __init__(self, rows):
        self.rows = rows

    def iter_rows(self, **kwargs):
        return iter(self.rows)


class ImportTests(unittest.TestCase):
    def test_identifiers(self):
        self.assertEqual(identifier(123.0), "123")
        self.assertEqual(identifier("123.0"), "123")
        self.assertEqual(identifier("000123"), "000123")

    def test_duplicate_header_does_not_confuse_class_with_money(self):
        sheet = Sheet([["Group", None, None], ["Артикул", "Выручка", "Выручка"], ["A1", "A", 234.5]])
        rows = list(records(sheet, {"артикул"}))
        self.assertEqual(rows[0][0], 3)
        self.assertEqual(numeric_field(rows[0][1], "Выручка"), 234.5)
        self.assertEqual(rows[0][1]["Выручка"], "A")

    def test_ambiguous_numeric_column(self):
        self.assertIsNone(numeric_field({"Выручка": 2, "Выручка [3]": 5}, "Выручка"))

    def test_ad_groups_preserve_metric_and_window(self):
        sheet = Sheet([["info", None, None, None, None, None, None, "ОДРР", None],
                       ["селлер ску товара", "Ozon SKU ID", "name", "model", "url", "url2", "url3", "Сегодня", "30 дней"],
                       ["A", "123", "x", "9", "", "", "", 2, 5]])
        row = list(ad_records(sheet))[0][1]
        self.assertEqual(row["ОДРР / Сегодня"], 2)
        self.assertEqual(row["ОДРР / 30 дней"], 5)

    def test_selection_dedup_and_conflict(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "skus.csv"
            path.write_text("SKU,Артикул\n123,A\n123,A\n456,B\n", encoding="utf-8-sig")
            self.assertEqual(len(load_selection(path)), 2)
            path.write_text("SKU,Артикул\n123,A\n123,B\n", encoding="utf-8-sig")
            with self.assertRaises(ValueError):
                load_selection(path)

    def test_package_keeps_missing_and_prevents_csv_formula(self):
        package = {"products": [{"sku": "123", "article": "A", "name": "=1+1", "issues": [], "metrics": {"profit": None}}], "warnings": ["Unknown period"]}
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "output.zip"
            export_zip(package, path)
            with zipfile.ZipFile(path) as archive:
                self.assertIsNone(json.loads(archive.read("data.json"))["products"][0]["metrics"]["profit"])
                row = next(csv.DictReader(io.StringIO(archive.read("metrics.csv").decode("utf-8-sig"))))
                self.assertEqual(row["name"], "'=1+1")
                self.assertEqual(row["profit"], "")


if __name__ == "__main__":
    unittest.main()
