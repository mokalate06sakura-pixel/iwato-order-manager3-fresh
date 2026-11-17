from copy import copy
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

import openpyxl
import pandas as pd
from openpyxl.styles import Alignment, Border, Font, Side

import create_order_form_maruhachi as maru


def make_masters(root):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    names = ["わ定番", "あ定番"] + [f"固定{i:02d}" for i in range(2, 16)]
    for name in (maru.TEMPLATE_SHEET_NAME_TOKUYOU, maru.TEMPLATE_SHEET_NAME_YUHOUSE):
        ws = wb.create_sheet(name)
        ws.merge_cells("A1:I1")
        ws["A1"] = "丸八発注書"
        for i, food in enumerate(names):
            row = 6 + i
            ws.cell(row, 1, "元テンプレートの日付")
            ws.cell(row, 2, f"F{i:02d}")
            ws.cell(row, 4, food)
            ws.cell(row, 5, "定番規格")
        for row in range(6, 29):
            ws.row_dimensions[row].height = 25
            for col in range(1, 10):
                cell = ws.cell(row, col)
                cell.font = Font(name="游ゴシック", size=11)
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = Border(bottom=Side(style="thin"))
        for col in "ABCDEFGHI":
            ws.column_dimensions[col].width = 13.25
        ws.print_area = "A1:I30"
        ws.page_setup.orientation = "landscape"
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 1
        ws.oddFooter.left.text = "保持するフッター"
    tags = wb.create_sheet("タグ")
    tags.append(["コード", "丸八食品名", "規格", "食品名"])
    for i, name in enumerate(names):
        tags.append([f"F{i:02d}", name, "定番規格", name])
    tags["Z1"] = "保持するタグ"
    template, codes = root / "template.xlsm", root / "codes.xlsm"
    wb.save(template)
    wb.save(codes)
    wb.close()
    return template, codes


class MaruhachiUseDateTests(unittest.TestCase):
    def outputs(self, items):
        directory = tempfile.TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        root = Path(directory.name)
        template, codes = make_masters(root)
        data = []
        for date, food, resident, staff, yuhouse in items:
            data.append({"仕入先": maru.SUPPLIER_NAME, "使用日": date, "食品名": food,
                         "換算値": "入力規格", "特養入所者": resident,
                         "特養職員": staff, "ユーハウス": yuhouse})
        inspection = root / "inspection.xlsx"
        pd.DataFrame(data).to_excel(inspection, index=False)
        loader = openpyxl.load_workbook
        opened = []
        def track(*args, **kwargs):
            book = loader(*args, **kwargs)
            opened.append(book)
            return book
        try:
            with patch.object(openpyxl, "load_workbook", side_effect=track):
                paths = maru.generate_maruhachi_order_forms_both_facilities(inspection, template, codes, root)
        finally:
            for book in opened:
                book.close()
                if book.vba_archive is not None:
                    book.vba_archive.close()
        results = []
        for path in paths:
            book = openpyxl.load_workbook(path, keep_vba=True)
            self.addCleanup(book.vba_archive.close)
            self.addCleanup(book.close)
            sheets = [ws for ws in book if ws.sheet_state == "visible" and ws.title != "タグ"]
            results.append((book, sheets))
        return results

    @staticmethod
    def labels(ws):
        return [(r, ws.cell(r, 1).value) for r in range(6, 29) if ws.cell(r, 1).value is not None]

    def test_first_visible_fixed_order_keeps_date_across_fixed_and_append_items(self):
        for tokuyou, (book, sheets) in zip((True, False), self.outputs([
            ("11/2月", "あ定番", 2, 1, 4),
            ("11/2月", "わ定番", 5, 0, 6),
            ("11/2月", "追加品", 3, 0.5, 7),
        ])):
            self.assertEqual(len(sheets), 1)
            ws = sheets[0]
            self.assertEqual(self.labels(ws), [(6, "11/2月")])
            self.assertEqual(ws["F6"].value, 5 if tokuyou else 6)
            self.assertEqual(ws["F7"].value, 2 if tokuyou else 4)
            self.assertEqual(ws["F22"].value, 3 if tokuyou else 7)
            self.assertEqual(ws["G7"].value, 1 if tokuyou else None)
            self.assertEqual(ws["G22"].value, 0.5 if tokuyou else None)
            self.assertEqual(ws["D22"].value, "追加品")
            self.assertEqual(ws["E22"].value, "入力規格")
            self.assertEqual(ws["A7"].font.name, "游ゴシック")
            self.assertEqual(ws["A7"].border.bottom.style, "thin")
            self.assertEqual(book["タグ"]["Z1"].value, "保持するタグ")
            unused = maru.TEMPLATE_SHEET_NAME_YUHOUSE if tokuyou else maru.TEMPLATE_SHEET_NAME_TOKUYOU
            self.assertEqual(book[unused]["A6"].value, "元テンプレートの日付")

    def test_date_is_at_first_ordered_row_not_an_unordered_fixed_item(self):
        for _, sheets in self.outputs([
            ("11/2月", "わ定番", 0, 0, 0),
            ("11/2月", "あ定番", 0, 2, 3),
        ]):
            self.assertEqual(self.labels(sheets[0]), [(7, "11/2月")])

    def test_last_fixed_row_then_append_item_still_has_one_date(self):
        for _, sheets in self.outputs([
            ("11/2月", "固定15", 1, 0, 1),
            ("11/2月", "追加品", 1, 0, 1),
        ]):
            self.assertEqual(self.labels(sheets[0]), [(21, "11/2月")])

    def test_append_only_continuation_pages_keep_their_first_date(self):
        for count, expected_pages in ((7, 1), (8, 2), (15, 3)):
            for _, sheets in self.outputs([
                ("11/2月", f"追加品{i:02d}", 1.25, 0.5, 2.75) for i in range(count)
            ]):
                with self.subTest(count=count):
                    self.assertEqual(len(sheets), expected_pages)
                    self.assertEqual(sum(ws.cell(r, 6).value is not None for ws in sheets for r in range(6, 29)), count)
                    for ws in sheets:
                        self.assertEqual(self.labels(ws), [(22, "11/2月")])
                        self.assertEqual(ws.print_area.split("!", 1)[1], "$A$1:$I$30")
                        self.assertEqual(ws.row_dimensions[22].height, 25)
                        self.assertEqual(ws.column_dimensions["A"].width, 13.25)
                        self.assertEqual(ws.oddFooter.left.text, "保持するフッター")

    def test_dates_remain_separate_and_zero_quantity_does_not_add_date(self):
        for _, sheets in self.outputs([
            ("11/2月", "あ定番", 1, 0, 1),
            ("11/3火", "追加品", 1, 0, 1),
            ("11/4水", "追加品", 0, 0, 0),
        ]):
            self.assertEqual(len(sheets), 3)
            self.assertEqual([self.labels(ws) for ws in sheets], [[(7, "11/2月")], [(22, "11/3火")], []])


if __name__ == "__main__":
    unittest.main()
