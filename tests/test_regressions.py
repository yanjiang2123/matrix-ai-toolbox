import csv
import tempfile
import unittest
import zipfile
from datetime import date, datetime
from pathlib import Path
from unittest.mock import patch

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


class TextExcelConversionTests(unittest.TestCase):
    def test_quoted_delimiter_and_internal_blank_row_are_preserved(self):
        rows = matrix_core.text_to_rows(
            '\ufeff编号,备注,金额\n1,"包含,逗号",12.50\n\n2,"普通文本",8', ","
        )
        self.assertEqual(["编号", "备注", "金额"], rows[0])
        self.assertEqual(["1", "包含,逗号", "12.50"], rows[1])
        self.assertEqual(["", "", ""], rows[2])
        self.assertEqual(["2", "普通文本", "8"], rows[3])

    def test_xlsx_has_business_types_and_readable_formatting(self):
        long_note = "用于检查自动换行与行高的中文长文本。" * 12
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "formatted.xlsx"
            matrix_core.write_xlsx(path, {"数据": (
                ["编号", "日期", "金额", "比例", "备注", "命令"],
                [
                    ["001", "2026-08-17", "1234.50", "12.5%", long_note, "=1+1"],
                    ["002", "", "-45.67", "0%", "短文本", " +cmd"],
                ],
            )})

            wb = openpyxl.load_workbook(path, data_only=False)
            try:
                ws = wb["数据"]
                self.assertEqual("001", ws["A2"].value)
                self.assertEqual("s", ws["A2"].data_type)
                self.assertEqual("@", ws["A2"].number_format)
                self.assertIsInstance(ws["B2"].value, date)
                self.assertEqual("yyyy-mm-dd", ws["B2"].number_format)
                self.assertEqual(1234.5, ws["C2"].value)
                self.assertEqual("#,##0.00", ws["C2"].number_format)
                self.assertEqual(-45.67, ws["C3"].value)
                self.assertEqual(0.125, ws["D2"].value)
                self.assertEqual("0.0%", ws["D2"].number_format)
                self.assertEqual("'=1+1", ws["F2"].value)
                self.assertEqual("' +cmd", ws["F3"].value)
                self.assertEqual("微软雅黑", ws["E2"].font.name)
                self.assertTrue(ws["E2"].alignment.wrap_text)
                self.assertEqual("thin", ws["E2"].border.left.style)
                self.assertGreater(ws.row_dimensions[2].height, 20)
                self.assertLessEqual(ws.column_dimensions["E"].width, 42)
                self.assertEqual("A2", ws.freeze_panes)
                self.assertEqual("A1:F3", ws.auto_filter.ref)
                self.assertEqual(1, ws.page_setup.fitToWidth)
                self.assertIn("A1:F3", str(ws.print_area).replace("$", ""))
            finally:
                wb.close()


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

    def test_ddl_special_identifier_and_decimal_scale_are_preserved(self):
        ddl = "CREATE TABLE demo (`金额（元）` decimal(18,2), `业务日期` date)"
        parsed = convert.parse_ddl(ddl)
        self.assertEqual(["金额（元）", "业务日期"],
                         [col["name"] for col in parsed["columns"]])
        self.assertEqual(2, parsed["columns"][0]["scale"])
        self.assertEqual("6172.80", convert._sql_literal(
            6172.799999999999, parsed["columns"][0]))
        self.assertEqual("12.00", convert._sql_literal(
            12.0, parsed["columns"][0]))
        self.assertEqual("'2026-08-01'", convert._sql_literal(
            datetime(2026, 8, 1), parsed["columns"][1]))

    def test_word_has_readable_header_pagination_and_display_formats(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "source.xlsx"
            output = Path(tmp) / "result.docx"
            matrix_core.write_xlsx(source, {"数据": (
                ["日期", "金额", "完成率", "状态", "长说明"],
                [[date(2026, 8, 1), 1280.5, 0.125, True, "中文说明" * 30]],
            )})
            result = convert.excel_to_word(source, output, title="业务验收")
            self.assertEqual(1, result["rows"])
            self.assertTrue(output.exists())

            from docx import Document
            doc = Document(output)
            table = doc.tables[0]
            self.assertEqual("FFFFFF", str(
                table.rows[0].cells[0].paragraphs[0].runs[0].font.color.rgb))
            self.assertEqual("2026-08-01", table.rows[1].cells[0].text)
            self.assertEqual("1,280.50", table.rows[1].cells[1].text)
            self.assertEqual("12.50%", table.rows[1].cells[2].text)
            self.assertEqual("是", table.rows[1].cells[3].text)
            with zipfile.ZipFile(output) as archive:
                xml = archive.read("word/document.xml").decode("utf-8")
            self.assertIn("<w:tblHeader", xml)
            self.assertIn("<w:cantSplit", xml)
            self.assertIn('<w:tblLayout w:type="fixed"', xml)
            self.assertIn("<w:tcMar", xml)

    def test_pdf_can_use_available_local_font_and_wrap_long_text(self):
        if convert._pick_font() is None:
            self.skipTest("本机没有可供 PDF 使用的字体")
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "source.xlsx"
            output = Path(tmp) / "result.pdf"
            matrix_core.write_xlsx(source, {"数据": (
                ["日期", "金额", "说明"],
                [[date(2026, 8, 1), 1280.5, "需要自动换行的业务说明" * 8]],
            )})
            result = convert.excel_to_pdf(source, output, title="业务验收")
            self.assertEqual(1, result["rows"])
            self.assertEqual(0, result["clipped_cells"])
            self.assertEqual(0, result["clipped_headers"])
            self.assertGreater(output.stat().st_size, 1000)

    def test_native_ocr_without_confidence_is_labeled_for_manual_review(self):
        blocks = [
            {"text": "编号", "x": 0.1, "y": 0.1, "w": 0.1, "h": 0.05},
            {"text": "001", "x": 0.1, "y": 0.2, "w": 0.1, "h": 0.05},
        ]
        with patch.object(convert, "ocr_blocks", return_value=blocks):
            result = convert.image_to_rows(Path("unused.png"))
        self.assertFalse(result["confidence_available"])
        self.assertIsNone(result["avg_conf"])
        self.assertIn("人工核对", result["warning"])

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
