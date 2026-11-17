"""日別献立の行選択・印刷設定・画像保護を検証する。"""
import base64
from copy import copy
from io import BytesIO
import unittest
from zipfile import ZipFile

import openpyxl
from openpyxl.chart import BarChart, Reference
from openpyxl.drawing.image import Image
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, TwoCellAnchor
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from split_menu_daily import cleanup_daily_sheets, list_menu_sheets, split_menu_daily

PNG = base64.b64decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+aN1sAAAAASUVORK5CYII=")


def menu_fixture():
    book = openpyxl.Workbook()
    source = book.active
    source.title = "月間献立"
    for row in range(1, 31):
        source.row_dimensions[row].height = 15 + row / 10
        source.cell(row, 36, row)
    for col in ("A", "J", "AJ"):
        source.column_dimensions[col].width = 4.25
    for row, text in {2: "現場：いわと", 3: "年月日：2026年11月1日（日）", 5: "朝食", 6: "朝ごはん", 8: "昼食", 9: "<A>", 10: "通常食", 11: "<B>", 12: "セレクト食", 13: "夕食", 14: "夕ごはん", 17: "現場：いわと", 18: "年月日：2026年11月2日（月）", 20: "朝食", 21: "朝ごはん2", 23: "昼食", 24: "昼ごはん2", 26: "夕食", 27: "夕ごはん2"}.items():
        source.cell(row, 1, text)
    source["J11"] = "<B>"
    source["S11"] = "<B>"  # 同じ帯の複数の食種マーカーは1か所として扱う。
    source.merge_cells("A3:AI3")
    source.merge_cells("A18:AI18")
    source.merge_cells("B12:H12")
    source["B12"] = "選べる主菜"
    source["B12"].font = Font(name="游ゴシック", size=12, bold=True)
    source["B12"].fill = PatternFill("solid", fgColor="FFFF00")
    source["J21"] = "=AJ21+1"
    source.oddHeader.center.text = "月間献立"
    source.page_margins.left = 0.2
    for anchor in ("AH2", "B7", "AH8"):
        image = Image(BytesIO(PNG))
        image.anchor = anchor
        source.add_image(image)
    chart = BarChart()
    chart.add_data(Reference(source, min_col=36, min_row=1, max_row=5))
    source.add_chart(chart, "AH2")  # スタンプの位置でもグラフは残す。
    other = book.create_sheet("別シート")
    other["A1"] = "保持"
    old = book.create_sheet("日別_古い結果")
    old["A1"] = "前回の結果"
    old.print_area = "A1:B2"
    book.active = 0
    output = BytesIO()
    book.save(output)
    book.close()
    return output.getvalue()


def v6_fixture():
    book = openpyxl.load_workbook(BytesIO(menu_fixture()))
    ws = book["月間献立"]
    ws["A4"], ws["J4"], ws["S4"], ws["AB4"] = "いわと常食", "常食", "軟菜きざみ", "ブレンダー食"
    for column, label in ((10, "常食主菜"), (19, "軟菜主菜"), (28, "ブレンダー主菜")):
        ws.cell(6, column, label)
        ws.cell(6, column + 8, 10 + column)
        ws.merge_cells(start_row=6, start_column=column, end_row=6, end_column=column + 7)
        ws.cell(12, column, label + "B")
    ws["J21"] = "=SUM($AJ$21,AJ21,'別シート'!AJ21)"
    ws["A2"].font = Font(name="游ゴシック", size=12, bold=True)
    ws["A2"].alignment = Alignment(horizontal="left")
    ws["J2"].border = Border(bottom=Side(style="thin"))
    for column in range(1, 37):
        ws.column_dimensions[get_column_letter(column)].width = 4 + column / 10
    for row in range(1, 31):
        ws.row_dimensions[row].height = 90
    buffer = BytesIO()
    book.save(buffer)
    book.close()
    return buffer.getvalue()


def month_fixture():
    fixture = openpyxl.Workbook()
    source = fixture.active
    source.title = "12月31日分"
    for day in range(1, 32):
        top = (day - 1) * 12 + 1
        source.cell(top + 1, 1, "現場：いわと")
        source.cell(top + 2, 1, f"年月日：2026年12月{day}日")
        source.cell(top + 3, 1, "いわと常食")
        source.cell(top + 3, 10, "常食")
        source.cell(top + 4, 1, "朝食")
        source.cell(top + 5, 10, f"主菜{day}")
        source.cell(top + 5, 18, day)
        source.cell(top + 6, 1, "昼食")
        if day in (15, 31):
            source.cell(top + 7, 10, "<B>")
            source.cell(top + 8, 10, f"選択主菜{day}")
            source.cell(top + 8, 18, day + 100)
        source.cell(top + 9, 1, "夕食")
        source.cell(top + 10, 36, day + 200)
    data = BytesIO()
    fixture.save(data)
    fixture.close()
    return data.getvalue()


class DailyMenuTests(unittest.TestCase):
    def test_split_select_menu_and_print_layout(self):
        data = menu_fixture()
        result, report = split_menu_daily(data, "月間献立")
        book = openpyxl.load_workbook(BytesIO(result))
        original = openpyxl.load_workbook(BytesIO(data))
        self.assertEqual(report, {"days": 2, "select_days": 1, "stamps_removed": 1, "old_sheets_removed": 1,
                                 "stamps_kept": 0, "stamp_copies": 0, "include_stamps": False,
                                 "remove_iwato": False, "equalize_widths": False,
                                 "sheets": ["日別_11月1日", "日別_11月1日_B", "日別_11月2日"]})
        self.assertEqual(book.sheetnames, ["月間献立", "別シート"] + report["sheets"])
        self.assertEqual(book.active.title, "日別_11月1日")
        main, alternate, second = [book[n] for n in report["sheets"]]
        self.assertNotIn("セレクト食", [c.value for row in main for c in row])
        self.assertEqual(alternate["A5"].value, "昼食")
        self.assertEqual(alternate["A6"].value, "<B>")
        self.assertEqual(alternate["A7"].value, "セレクト食")
        self.assertEqual(alternate["B7"].value, "選べる主菜")
        self.assertEqual(copy(alternate["B7"].font), copy(original["月間献立"]["B12"].font))
        self.assertEqual(alternate.row_dimensions[7].height, original["月間献立"].row_dimensions[12].height)
        self.assertIn("B7:H7", str(alternate.merged_cells))
        self.assertEqual(main["A11"].value, "夕食")
        self.assertEqual(second["A3"].value, "年月日：2026年11月2日（月）")
        self.assertEqual(second["J6"].value, "=AJ6+1")
        for ws in (main, alternate, second):
            self.assertEqual(str(ws.page_setup.paperSize), ws.PAPERSIZE_A3)
            self.assertIsNone(ws.page_setup.usePrinterDefaults)
            self.assertEqual(ws.page_setup.orientation, "portrait")
            self.assertEqual((ws.page_setup.fitToWidth, ws.page_setup.fitToHeight), (1, 1))
            self.assertTrue(ws.sheet_properties.pageSetUpPr.fitToPage)
            self.assertEqual(str(ws.print_area), f"'{ws.title}'!$A$1:$AJ${ws.max_row}")
            self.assertEqual(ws.column_dimensions["AJ"].width, 4.25)
            self.assertEqual(ws.oddHeader.center.text, "月間献立")
            self.assertEqual(ws.page_margins.left, 0.2)
            self.assertEqual(len(ws._images), 0)
        self.assertEqual(len(book["月間献立"]._images), 2)
        self.assertEqual(len(book["月間献立"]._charts), 1)
        for row in original["月間献立"]:
            for cell in row:
                self.assertEqual(book["月間献立"][cell.coordinate].value, cell.value)
        with ZipFile(BytesIO(data)) as before, ZipFile(BytesIO(result)) as after:
            for name in before.namelist():
                if name.startswith(("xl/media/", "xl/charts/")):
                    self.assertEqual(before.read(name), after.read(name), name)
        book.close()
        original.close()

    def test_compact_layout_keeps_three_blocks_headers_and_quantities(self):
        data = v6_fixture()
        for equalize in (False, True):
            for include in (False, True):
                with self.subTest(equalize=equalize, include=include):
                    result, report = split_menu_daily(data, "月間献立", include_stamps=include, remove_iwato=True, equalize_widths=equalize)
                    book = openpyxl.load_workbook(BytesIO(result))
                    source = openpyxl.load_workbook(BytesIO(data))
                    main, alternate, second = [book[n] for n in report["sheets"]]
                    self.assertEqual([main.cell(4, c).value for c in (1, 10, 19)], ["常食", "軟菜きざみ", "ブレンダー食"])
                    self.assertEqual([main.cell(6, c).value for c in (1, 10, 19)], ["常食主菜", "軟菜主菜", "ブレンダー主菜"])
                    self.assertEqual([main.cell(6, c).value for c in (9, 18, 27)], [20, 29, 38])
                    self.assertEqual(main["A2"].value, "現場：いわと")
                    self.assertEqual(main["A3"].value, "年月日：2026年11月1日（日）")
                    self.assertEqual(main["A5"].value, "朝食")
                    self.assertEqual(main["A11"].value, "夕食")
                    self.assertEqual(copy(main["A2"].font), copy(source["月間献立"]["A2"].font))
                    self.assertEqual(copy(main["A2"].border), copy(source["月間献立"]["J2"].border))
                    self.assertIn("A6:H6", str(main.merged_cells))
                    self.assertEqual(alternate["A5"].value, "昼食")
                    self.assertEqual(alternate["A6"].value, "<B>")
                    self.assertEqual(alternate["A7"].value, "常食主菜B")
                    self.assertEqual(second["A6"].value, "=SUM($AA$21,AA6,'別シート'!AJ6)")
                    for ws in (main, alternate, second):
                        self.assertEqual(ws.max_column, 27)
                        self.assertEqual(str(ws.print_area), f"'{ws.title}'!$A$1:$AA${ws.max_row}")
                        self.assertEqual(ws.page_setup.orientation, "portrait")
                        widths = [ws.column_dimensions[get_column_letter(c)].width for c in range(1, 28)]
                        if equalize:
                            blocks = [sum(widths[i:i+9]) for i in (0, 9, 18)]
                            self.assertLess(max(blocks) - min(blocks), 0.04)
                            original_ratio = source["月間献立"].column_dimensions["J"].width / source["月間献立"].column_dimensions["R"].width
                            self.assertAlmostEqual(widths[0] / widths[8], original_ratio, places=2)
                        else:
                            self.assertEqual(widths, [source["月間献立"].column_dimensions[get_column_letter(c)].width for c in range(10, 37)])
                    if equalize:
                        self.assertGreater(main.column_dimensions["A"].width, alternate.column_dimensions["A"].width)
                    self.assertEqual(report["stamps_removed"], 0)
                    self.assertEqual(len(book["月間献立"]._images), 3)
                    self.assertEqual(len(main._images), int(include))
                    self.assertEqual(len(alternate._images), int(include))
                    if include:
                        self.assertEqual(main._images[0].anchor._from.row, 0)
                        self.assertEqual(main._images[0].anchor._from.col, 26)
                    for row in source["月間献立"]:
                        for cell in row:
                            self.assertEqual(book["月間献立"][cell.coordinate].value, cell.value)
                    source.close()
                    book.close()

    def test_width_adjustment_requires_three_blocks(self):
        with self.assertRaisesRegex(ValueError, "いわと常食"):
            split_menu_daily(menu_fixture(), "月間献立", equalize_widths=True)

    def test_full_31_day_month_with_select_menu_on_last_day(self):
        data = BytesIO(month_fixture())
        for compact in (False, True):
            with self.subTest(compact=compact):
                result, report = split_menu_daily(data.getvalue(), "12月31日分", remove_iwato=compact, equalize_widths=compact)
                self.assertEqual((report["days"], report["select_days"]), (31, 2))
                book = openpyxl.load_workbook(BytesIO(result))
                self.assertEqual(len(report["sheets"]), 33)
                self.assertEqual(report["sheets"][-2:], ["日別_12月31日", "日別_12月31日_B"])
                total = 0
                for day in range(1, 32):
                    ws = book[f"日別_12月{day}日"]
                    self.assertEqual(ws["A3"].value, f"年月日：2026年12月{day}日")
                    total += ws.cell(6, 9 if compact else 18).value
                    self.assertEqual((ws.page_setup.fitToWidth, ws.page_setup.fitToHeight), (1, 1))
                    self.assertEqual(ws.page_setup.orientation, "portrait")
                self.assertEqual(total, sum(range(1, 32)))
                final_b = book["日別_12月31日_B"]
                self.assertEqual(final_b.cell(7, 1 if compact else 10).value, "選択主菜31")
                self.assertEqual(final_b.cell(7, 9 if compact else 18).value, 131)
                book.close()

    def test_include_stamps_on_daily_and_b_sheets_with_rebased_position(self):
        data = menu_fixture()
        fixture = openpyxl.load_workbook(BytesIO(data))
        image = Image(BytesIO(PNG))
        image.anchor = TwoCellAnchor(_from=AnchorMarker(col=33, row=16), to=AnchorMarker(col=34, row=18))
        fixture["月間献立"].add_image(image)
        buffer = BytesIO()
        fixture.save(buffer)
        fixture.close()
        data = buffer.getvalue()
        result, report = split_menu_daily(data, "月間献立", include_stamps=True)
        self.assertEqual(report["stamps_removed"], 0)
        self.assertEqual(report["stamps_kept"], 2)
        self.assertEqual(report["stamp_copies"], 3)
        book = openpyxl.load_workbook(BytesIO(result))
        self.assertEqual(len(book["月間献立"]._images), 4)
        for name in report["sheets"]:
            ws = book[name]
            self.assertEqual(len(ws._images), 1, name)
            self.assertEqual(ws._images[0].anchor._from.row, 1, name)
            self.assertEqual(ws._images[0].anchor._from.col, 33, name)
            self.assertEqual(len(ws._charts), 0)
        self.assertEqual(book["日別_11月2日"]._images[0].anchor.to.row, 3)
        repeated, second_report = split_menu_daily(result, "月間献立", include_stamps=True)
        self.assertEqual(second_report["stamp_copies"], 3)
        no_stamps, disabled = split_menu_daily(repeated, "月間献立", include_stamps=False)
        self.assertEqual(disabled["stamps_removed"], 2)
        without = openpyxl.load_workbook(BytesIO(no_stamps))
        for name in report["sheets"]:
            self.assertEqual(len(without[name]._images), 0)
        self.assertEqual(len(without["月間献立"]._images), 2)
        without.close()
        book.close()

    def test_rerun_replaces_daily_and_cleanup_preserves_original(self):
        data = menu_fixture()
        first, _ = split_menu_daily(data, "月間献立")
        second, report = split_menu_daily(first, "月間献立")
        self.assertEqual(report["old_sheets_removed"], 3)
        self.assertEqual(report["stamps_removed"], 0)
        self.assertEqual(openpyxl.load_workbook(BytesIO(first)).sheetnames,
                         openpyxl.load_workbook(BytesIO(second)).sheetnames)
        cleaned, removed = cleanup_daily_sheets(data)
        self.assertEqual(removed, 1)
        book = openpyxl.load_workbook(BytesIO(cleaned))
        self.assertEqual(book.sheetnames, ["月間献立", "別シート"])
        self.assertEqual(len(book["月間献立"]._images), 3)
        book.close()

    def test_sheet_selection_and_invalid_input(self):
        data = menu_fixture()
        self.assertEqual(list_menu_sheets(data), (["月間献立", "別シート"], "月間献立"))
        with self.assertRaisesRegex(ValueError, "日付行"):
            split_menu_daily(data, "別シート")
        with self.assertRaisesRegex(ValueError, "元の献立"):
            split_menu_daily(data, "日別_古い結果")
        with self.assertRaises(ValueError):
            split_menu_daily(b"invalid workbook", "月間献立")


if __name__ == "__main__":
    unittest.main()
