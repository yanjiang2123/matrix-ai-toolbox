import csv
import tempfile
import unittest
from pathlib import Path

import openpyxl

import convert
import excel_diff
import matrix_core
import sql_tools


class DifferenceArchiveTests(unittest.TestCase):
    def test_difference_moves_through_fixed_and_reappeared_states(self):
        diff = {"pk": "1", "col": "amount", "a": "10", "b": "11",
                "kind": "字段值差异"}
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "archive.xlsx"

            first = matrix_core.accumulate_diffs(path, [diff], "第一轮", ["1"])
            self.assertEqual((1, 0), (first["added"], first["fixed"]))
            self.assertEqual("未修复", self._status(path))

            clean = matrix_core.accumulate_diffs(path, [], "第二轮", ["1"])
            self.assertEqual(1, clean["fixed"])
            self.assertEqual("已修复", self._status(path))

            again = matrix_core.accumulate_diffs(path, [diff], "第三轮", ["1"])
            self.assertEqual(1, again["updated"])
            self.assertEqual("又出现", self._status(path))

            fixed_again = matrix_core.accumulate_diffs(path, [], "第四轮", ["1"])
            self.assertEqual(1, fixed_again["fixed"])
            self.assertEqual("已修复", self._status(path))

    @staticmethod
    def _status(path):
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        try:
            return wb.active.cell(2, 8).value
        finally:
            wb.close()


class ExportSafetyTests(unittest.TestCase):
    def test_xlsx_and_csv_escape_formula_like_values(self):
        with tempfile.TemporaryDirectory() as tmp:
            xlsx_path = Path(tmp) / "safe.xlsx"
            csv_path = Path(tmp) / "safe.csv"
            matrix_core.write_xlsx(
                xlsx_path, {"数据": (["=危险表头"], [["=1+1"], [" +cmd"]])}
            )
            matrix_core.write_csv(csv_path, ["=危险表头"], [["=1+1"], [" +cmd"]])

            wb = openpyxl.load_workbook(xlsx_path, data_only=False)
            try:
                ws = wb["数据"]
                self.assertEqual("'=危险表头", ws["A1"].value)
                self.assertEqual("'=1+1", ws["A2"].value)
                self.assertEqual("s", ws["A2"].data_type)
                self.assertEqual("' +cmd", ws["A3"].value)
            finally:
                wb.close()

            with csv_path.open(encoding="utf-8-sig", newline="") as handle:
                rows = list(csv.reader(handle))
            self.assertEqual([["'=危险表头"], ["'=1+1"], ["' +cmd"]], rows)


class ReportAndSqlTests(unittest.TestCase):
    def test_same_display_names_keep_both_side_sheets(self):
        result = excel_diff.compare_tables(
            ["id", "amount"], [["A", 1]],
            ["id", "amount"], [["B", 2]],
            ["id"],
        )
        report = excel_diff.build_report(result, "same.xlsx", "same.xlsx")
        self.assertIn("仅A有", report)
        self.assertIn("仅B有", report)
        self.assertEqual("A", report["仅A有"][1][0][0])
        self.assertEqual("B", report["仅B有"][1][0][0])

    def test_insert_quotes_each_table_identifier(self):
        sql = convert._wrap_insert("sales.order", "`id`", ["(1)"])
        self.assertTrue(sql.startswith("INSERT INTO `sales`.`order`"))

    def test_comparison_returns_all_checked_primary_keys(self):
        result = sql_tools.compare_details(
            ["id", "value"], [["1", "same"], ["2", "A"]],
            ["id", "value"], [["1", "same"], ["2", "B"]],
            ["id"],
        )
        self.assertEqual(["1", "2"], result["pks_in_scope"])

    def test_truncated_differences_are_excluded_from_archive_scope(self):
        value_columns = [f"c{i}" for i in range(2001)]
        headers = ["id", *value_columns]
        result = sql_tools.compare_details(
            headers, [["1", *(["A"] * len(value_columns))]],
            headers, [["1", *(["B"] * len(value_columns))]],
            ["id"],
        )
        self.assertTrue(result["archive_scope_truncated"])
        self.assertNotIn("1", result["pks_in_scope"])


if __name__ == "__main__":
    unittest.main()
