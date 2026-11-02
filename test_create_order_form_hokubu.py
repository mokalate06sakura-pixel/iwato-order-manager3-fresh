"""北部市場専用の合成テンプレート・検収簿を使う回帰テスト。

実行: python -m unittest discover -s . -p test_create_order_form_hokubu.py -v
VBAは保存の検証用バイト列であり、実マクロの実行確認は含まない。
"""
from copy import copy
from decimal import Decimal
from io import BytesIO
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from zipfile import ZipFile

import openpyxl
import pandas as pd
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.pagebreak import Break

import create_order_form_hokubu as hokubu


VBA_BYTES = b"HOKUBU_TEST_VBA_PRESERVATION_ONLY"
LONG_FOOD = "長い食品名・国産冷凍カット野菜ミックス（にんじん・玉ねぎ・じゃがいも）"


def make_template(path):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for mode, name in ((True, hokubu.TOKUYOU_SHEET), (False, hokubu.YUHOUSE_SHEET)):
        ws = wb.create_sheet(name)
        ws.merge_cells("A1:K1")
        ws["A1"] = "発 注 書"
        ws.merge_cells("A2:G2")
        ws["A2"] = "特養 いわと" if mode else "ユーハウス いわと"
        ws.merge_cells("A4:D4")
        ws["A4"] = hokubu.SUPPLIER_NAME
        ws["E4"] = "御 中"
        date = hokubu.TOKUYOU_DELIVERY_CELL if mode else hokubu.YUHOUSE_DELIVERY_CELL
        ws.merge_cells(f"{date}:K4")
        ws[date] = "納品日"
        for address, value in {"A6": "使用日", "B6": "食品名", "D6": "入居者", "E6": "職員", "F6": "合計数量", "H6": "その他"}.items():
            ws[address] = value
        for row in range(7, 19):
            ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=3)
            ws.cell(row, 2, "置換前の食品名")
            ws.row_dimensions[row].height = 31 + row % 3
        ws["A20"] = "備考欄"
        for row in ws:
            for cell in row:
                if isinstance(cell, MergedCell):
                    continue
                cell.font = Font(name="ＭＳ ゴシック", size=8 + cell.column % 4, bold=True, scheme="minor")
                cell.alignment = Alignment(horizontal="left", vertical="top", indent=0)
                cell.border = Border(**{side: Side(style="thin") for side in ("left", "right", "top", "bottom")})
                cell.fill = PatternFill("solid", fgColor="FFFFEE")
                cell.number_format = "0.0"
        for col in "ABCDEFGHIJK":
            ws.column_dimensions[col].width = 12.5 if col != "B" else 38
        ws.row_dimensions[1].height = 28
        ws.page_setup.orientation = "landscape"
        ws.page_setup.paperSize = ws.PAPERSIZE_A3
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 1
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.print_area = "A1:K20"
        ws.print_title_rows = "1:6"
        ws.print_title_cols = "A:B"
        ws.row_breaks.append(Break(id=20))
    other = wb.create_sheet("他業者テンプレート")
    other["A1"] = "他業者用・変更禁止"
    other["A1"].font = Font(name="Arial", size=19, bold=True)
    buffer = BytesIO()
    wb.save(buffer)
    archive = ZipFile(buffer, "a")
    archive.writestr("xl/vbaProject.bin", VBA_BYTES)
    wb.vba_archive = archive
    wb.save(path)
    archive.close()


def make_data(path):
    rows = []
    def add(food, res=1, staff=0, yuhouse=1, unit="kg", delivery="11/2月", use="11/3火"):
        rows.append({"仕入先": hokubu.SUPPLIER_NAME, "納品日": delivery, "使用日": use,
                     "食品名": food, "単位": unit, "特養入所者": res,
                     "特養職員": staff, "ユーハウス": yuhouse})
    add("誤差数量", 3.2, 1.4, 4.6000000000000005)
    add("同一食品集約", 0.1, 0.2, 0.1)
    add("同一食品集約", 0.2, 0.1, 0.2)
    add("パック", 5.0, 0, 5.0, "パック")
    add("本数", 50.0, 0, 50.0, "本")
    add("精度保持", 0.00000123456789, 0, 0.00000123456789)
    add(LONG_FOOD)
    for i in range(7):
        add(f"通常食品{i}")
    add("別の使用日", use="11/4水")
    add("別の納品日", delivery="11/3火", use="11/4水")
    add("数量ゼロ", 0, 0, 0)
    excluded = dict(rows[0], 仕入先="他業者", 食品名="他業者の食品")
    pd.DataFrame([*rows, excluded]).to_excel(path, index=False)


class QuantityTest(unittest.TestCase):
    def test_noise_trailing_zeros_and_precision(self):
        for qty, unit, expected in [
            (4.6000000000000005, "kg", "4.6kg"),
            (4.600000000000001, "kg", "4.6kg"),
            (3.1000000000000001, "kg", "3.1kg"),
            (0.1 + 0.2, "kg", "0.3kg"),
            (5.0, "パック", "5パック"), (50.0, "本", "50本"),
            (1.2345678901234567, "kg", "1.2345678901234567kg"),
            (0.00000123456789, "kg", "0.00000123456789kg"),
            (Decimal("1.0000000000000002"), "kg", "1.0000000000000002kg"),
            ("1.0000000000000002", "kg", "1.0000000000000002kg"),
            (-4.6000000000000005, "kg", "-4.6kg"),
            (0, "kg", ""), (None, "kg", ""), (float("nan"), "kg", ""),
        ]:
            with self.subTest(qty=qty):
                self.assertEqual(hokubu._format_qty_with_unit(qty, unit), expected)

    def test_row_total_uses_decimal_arithmetic(self):
        ws = openpyxl.Workbook().active
        hokubu._write_row_tokuyou(ws, 7, "11/3", "野菜", 3.2, 1.4, "kg")
        self.assertEqual(ws["F7"].value, "4.6kg")
        self.assertEqual(ws["D7"].value + ws["E7"].value, Decimal("4.6"))


class WorkbookTest(unittest.TestCase):
    def test_both_facilities_layout_values_and_vba(self):
        with TemporaryDirectory() as directory:
            root = Path(directory)
            template, data = root / "template.xlsm", root / "input.xlsx"
            make_template(template)
            make_data(data)
            reference = openpyxl.load_workbook(template, keep_vba=True)
            self.addCleanup(reference.vba_archive.close)
            self.addCleanup(reference.close)
            outputs = hokubu.generate_hokubu_order_forms_both_facilities(data, template, root)
            for is_tokuyou, output in zip((True, False), outputs):
                with self.subTest(facility=output.name):
                    wb = openpyxl.load_workbook(output, keep_vba=True)
                    self.addCleanup(wb.vba_archive.close)
                    self.addCleanup(wb.close)
                    base_name = hokubu.TOKUYOU_SHEET if is_tokuyou else hokubu.YUHOUSE_SHEET
                    untouched = hokubu.YUHOUSE_SHEET if is_tokuyou else hokubu.TOKUYOU_SHEET
                    base = reference[base_name]
                    generated = [ws for ws in wb if ws.title not in (untouched, "他業者テンプレート")]
                    self.assertEqual(len(generated), 4)
                    self.assertNotIn(base_name, wb.sheetnames)
                    self.assertEqual([sum(ws.cell(r, 2).value is not None for r in range(7, 19)) for ws in generated], [12, 1, 1, 1])
                    for name in (untouched, "他業者テンプレート"):
                        for row in reference[name]:
                            for cell in row:
                                actual = wb[name][cell.coordinate]
                                self.assertEqual(actual.value, cell.value)
                                self.assertEqual(actual._style, cell._style)
                    records = {}
                    dates = [("11/3火", "11月2日納品分"), ("11/3火", "11月2日納品分"),
                             ("11/4水", "11月2日納品分"), ("11/4水", "11月3日納品分")]
                    for ws, (use_date, delivery_date) in zip(generated, dates):
                        self.assertEqual(ws.merged_cells, base.merged_cells)
                        self.assertEqual(ws.page_setup, base.page_setup)
                        self.assertEqual(ws.page_margins, base.page_margins)
                        self.assertEqual(ws.print_options, base.print_options)
                        self.assertEqual(ws.print_area.split("!", 1)[1], base.print_area.split("!", 1)[1])
                        self.assertEqual(ws.print_title_rows, base.print_title_rows)
                        self.assertEqual(ws.print_title_cols, base.print_title_cols)
                        self.assertEqual(ws.row_breaks, base.row_breaks)
                        self.assertEqual({k: v.width for k, v in ws.column_dimensions.items()}, {k: v.width for k, v in base.column_dimensions.items()})
                        self.assertEqual({k: v.height for k, v in ws.row_dimensions.items()}, {k: v.height for k, v in base.row_dimensions.items()})
                        for row in ws:
                            for cell in row:
                                self.assertEqual(cell.border, copy(base[cell.coordinate].border))
                                self.assertEqual(cell.fill, copy(base[cell.coordinate].fill))
                                if isinstance(cell, MergedCell):
                                    continue
                                self.assertEqual(cell.font.name, "游ゴシック")
                                if 7 <= cell.row <= 18 or cell.row == 20:
                                    self.assertEqual(cell.font.sz, 11)
                                    self.assertFalse(cell.font.bold)
                                if 7 <= cell.row <= 18:
                                    self.assertEqual(cell.alignment.vertical, "center")
                                    cols = (1, 2, 4, 5, 6) if is_tokuyou else (1, 2, 4)
                                    if cell.column in cols:
                                        self.assertEqual(cell.alignment.horizontal, "center")
                        for address in ("A1", "A2"):
                            self.assertEqual(ws[address].font.sz, 14)
                            self.assertTrue(ws[address].font.bold)
                        date = hokubu.TOKUYOU_DELIVERY_CELL if is_tokuyou else hokubu.YUHOUSE_DELIVERY_CELL
                        self.assertEqual(ws[date].value, delivery_date)
                        for address in ("A4", "E4", date):
                            self.assertEqual(ws[address].font.sz, 12)
                            self.assertTrue(ws[address].font.bold)
                        for address in ("A6", "B6", "D6", "E6", "F6", "H6"):
                            self.assertEqual(ws[address].font.sz, 11)
                            self.assertTrue(ws[address].font.bold)
                            self.assertEqual(ws[address].alignment.horizontal, "center")
                        for r in range(7, 19):
                            if ws.cell(r, 2).value:
                                self.assertEqual(ws.cell(r, 1).value, use_date)
                                records[ws.cell(r, 2).value] = (ws, r)
                    qty_col = 6 if is_tokuyou else 4
                    for food, expected in {"誤差数量": "4.6kg", "同一食品集約": "0.6kg" if is_tokuyou else "0.3kg", "パック": "5パック", "本数": "50本", "精度保持": "0.00000123456789kg"}.items():
                        ws, r = records[food]
                        self.assertEqual(ws.cell(r, qty_col).value, expected)
                    if is_tokuyou:
                        ws, r = records["誤差数量"]
                        self.assertEqual(ws.cell(r, 4).value, 3.2)
                        self.assertEqual(ws.cell(r, 5).value, 1.4)
                    self.assertNotIn("数量ゼロ", records)
                    self.assertNotIn("他業者の食品", records)
                    ws, r = records[LONG_FOOD]
                    self.assertEqual(ws.cell(r, 2).value, LONG_FOOD)
                    self.assertTrue(ws.cell(r, 2).alignment.shrink_to_fit)
                    self.assertFalse(ws.cell(r, 2).alignment.wrap_text)
                    with ZipFile(output) as archive:
                        self.assertEqual(archive.read("xl/vbaProject.bin"), VBA_BYTES)
                    wb.close()
            reference.close()
            reference.vba_archive.close()


if __name__ == "__main__":
    unittest.main()
