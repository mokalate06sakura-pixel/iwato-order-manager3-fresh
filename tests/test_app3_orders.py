"""注文書の共通日付区切り・印刷書式・既存の数量を実出力で確認する。"""
import io
import unittest

import openpyxl
import pandas as pd

from app3_test_support import APP_PATH, load_logic


def inspection(counts, supplier="仕入先A"):
    rows = []
    for day_index, (tokuyou, yuhouse) in enumerate(counts):
        for index in range(max(tokuyou, yuhouse)):
            rows.append({"使用日": f"8/{17 + day_index}", "仕入先": supplier,
                         "食品名": f"食品{index:03d}", "単位": "kg",
                         "特養入所者": 1.25 if index < tokuyou and index % 2 == 0 else 0,
                         "特養職員": 0.5 if index < tokuyou and index % 2 else 0,
                         "ユーハウス": 2.75 if index < yuhouse else 0})
    return pd.DataFrame(rows)


def upload(frame):
    buffer = io.BytesIO()
    frame.to_excel(buffer, index=False)
    buffer.seek(0)
    return buffer


def printed_pages(ws):
    last = 5
    pages = []
    for end in [item.id for item in ws.row_breaks.brk] + [ws.max_row]:
        pages.append((last + 1, end))
        last = end
    return pages


def page_dates(ws):
    return [[ws.cell(row, 1).value for row in range(first, last + 1)
             if ws.cell(row, 1).value] for first, last in printed_pages(ws)]


class OrderTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.logic = load_logic()

    def books(self, counts):
        frame = inspection(counts)
        return [openpyxl.load_workbook(io.BytesIO(self.logic.create_order_workbook(
            upload(frame), facility)[0])) for facility in ("特養", "ユーハウス")]

    def test_all_seven_dates_fit_one_page_with_six_writing_rows(self):
        for book in self.books([(1, 1)] * 7):
            ws = book.active
            self.assertEqual(ws.row_breaks.count, 0)
            self.assertEqual(ws.max_row, 18)
            self.assertEqual(ws["A5"].value, "使用日")
            for row in range(7, 18, 2):
                self.assertTrue(all(c.value is None for c in ws[row]))
                self.assertEqual(ws.row_dimensions[row].height, 30)
            self.assertEqual(ws["B6"].font.sz, 18)
            book.close()

    def test_pairs_then_last_three_and_shared_dates(self):
        books = self.books([(5, 3)] * 7)
        expected = [["8/17", "8/18"], ["8/19", "8/20"], ["8/21", "8/22", "8/23"]]
        for book in books:
            self.assertEqual(page_dates(book.active), expected)
            self.assertTrue(all(book.active.row_dimensions[row].height == 30
                                for row in range(6, book.active.max_row + 1)))
            book.close()

    def test_large_two_days_use_single_days(self):
        for book in self.books([(13, 9)] * 4):
            self.assertEqual(page_dates(book.active), [[f"8/{day}"] for day in range(17, 21)])
            self.assertEqual(book.active["B6"].font.sz, 18)
            book.close()

    def test_font_stages_and_row_height_only_as_last_stage(self):
        capacities = [self.logic._order_page_capacity(*stage)
                      for stage in self.logic.ORDER_LAYOUT_STAGES]
        self.assertEqual(sorted(set(capacities)), capacities)
        previous = 0
        for stage, capacity in zip(self.logic.ORDER_LAYOUT_STAGES, capacities):
            with self.subTest(stage=stage):
                size = capacity if not previous else previous + 1
                pages = self.logic._plan_order_pages(inspection([(size, size)]))
                self.assertEqual(len(pages), 1)
                self.assertEqual((pages[0]["font_size"], pages[0]["row_height"]), stage)
                for book in self.books([(size, size)]):
                    self.assertEqual(book.active["B6"].font.sz, stage[0])
                    self.assertEqual(book.active.row_dimensions[6].height, stage[1])
                    book.close()
            previous = capacity

    def test_opposite_dense_dates_use_actual_facility_totals(self):
        # 日ごとの最大値を足すと誤って2ページになる。実際は両方1ページ。
        for book in self.books([(12, 1), (1, 12)]):
            self.assertEqual(len(printed_pages(book.active)), 1)
            book.close()

    def test_missing_facility_dates_keep_identical_boundaries(self):
        for book in self.books([(5, 0), (5, 3), (0, 5), (5, 3), (5, 0), (5, 3), (0, 5)]):
            ws = book.active
            self.assertEqual(page_dates(ws), [["8/17", "8/18"], ["8/19", "8/20"],
                                              ["8/21", "8/22", "8/23"]])
            book.close()

    def test_oversized_day_splits_without_losing_products_or_dates(self):
        books = self.books([(57, 3), (2, 1)])
        self.assertEqual(page_dates(books[0].active), page_dates(books[1].active))
        for book, expected in zip(books, (59, 4)):
            ws = book.active
            self.assertEqual(sum(ws.cell(row, 2).value is not None
                                 for row in range(6, ws.max_row + 1)), expected)
            for first, last in printed_pages(ws):
                self.assertIsNotNone(ws.cell(first, 1).value)
                self.assertLessEqual(last - first + 1, self.logic._order_page_capacity(16, 28))
            book.close()

    def test_oversized_day_reserves_boundary_blank_only_on_last_segment(self):
        capacity = self.logic._order_page_capacity(16, 28)
        pages = self.logic._plan_order_pages(inspection([(capacity * 3 - 1, 1), (1, 1)]))
        self.assertEqual(len(pages), 4)
        self.assertEqual([page["trailing_blank"] for page in pages], [False, False, True, False])
        self.assertEqual((pages[-1]["font_size"], pages[-1]["row_height"]), (18, 30))

    def test_writing_row_at_every_date_boundary_including_page_end(self):
        for book in self.books([(8, 5)] * 7):
            ws = book.active
            date_rows = [row for row in range(6, ws.max_row + 1) if ws.cell(row, 1).value]
            for row in date_rows[1:]:
                self.assertTrue(all(c.value is None for c in ws[row - 1]))
                self.assertTrue(any(c.value is not None for c in ws[row - 2]))
                for col in range(1, 14):
                    self.assertNotIn(ws.cell(row - 1, col).border.bottom.style, ("thick", "medium"))
                    self.assertNotIn(ws.cell(row, col).border.top.style, ("thick", "medium"))
            for _, last in printed_pages(ws)[:-1]:
                self.assertTrue(all(c.value is None for c in ws[last]))
            book.close()

    def test_print_settings_and_original_facility_columns(self):
        for book, quantity, company_title in zip(self.books([(5, 3)] * 7),
                ("入所者", "ユーハウス入居者"), ("介護老人福祉施設いわと", "ユーハウスいわと")):
            ws = book.active
            self.assertEqual(ws["C5"].value, quantity)
            self.assertEqual(ws["A3"].value, "仕入先A 御中")
            self.assertIn(company_title, ws["B1"].value)
            self.assertEqual(ws["K3"].value, "(有) ハートミール")
            self.assertEqual(str(ws.page_setup.paperSize), ws.PAPERSIZE_A4)
            self.assertEqual(ws.page_setup.orientation, "landscape")
            self.assertEqual(ws.page_setup.fitToWidth, 1)
            self.assertEqual(ws.page_setup.fitToHeight, 0)
            self.assertTrue(ws.sheet_properties.pageSetUpPr.fitToPage)
            self.assertEqual(ws.print_title_rows, "$1:$5")
            self.assertEqual(ws.oddFooter.center.text, f"&P / {len(printed_pages(ws))}")
            self.assertEqual(ws.page_setup.firstPageNumber, 1)
            self.assertFalse(ws.sheet_view.showGridLines)
            self.assertFalse(ws.print_options.gridLines)
            self.assertTrue(ws["B6"].alignment.shrink_to_fit)
            self.assertFalse(ws["B6"].alignment.wrap_text)
            if quantity == "入所者":
                self.assertEqual(ws["E5"].value, "職員")
                self.assertIsNone(ws["C7"].value)
                self.assertEqual(ws["E7"].value, 0.5)
            book.close()

    def test_no_quantities_and_missing_columns_still_raise(self):
        for frame, facility in ((inspection([(1, 0)]), "ユーハウス"),
                                (inspection([(0, 1)]), "特養"),
                                (inspection([(1, 1)]).drop(columns="特養職員"), "特養"),
                                (inspection([(1, 1)]).drop(columns="食品名"), "特養")):
            with self.assertRaises(ValueError):
                self.logic.create_order_workbook(upload(frame), facility)

    def test_filename_supplier_sheets_and_independent_vendor_plans(self):
        frame = pd.concat([inspection([(5, 3)] * 7, "業者/A"),
                           inspection([(1, 1)], "業者?A"),
                           inspection([(0, 0)], "数量なし")], ignore_index=True)
        for facility, prefix in (("特養", "注文書_いわと"), ("ユーハウス", "注文書_ユーハウス")):
            payload, filename = self.logic.create_order_workbook(upload(frame), facility)
            self.assertEqual(filename, prefix + "_0817.xlsx")
            book = openpyxl.load_workbook(io.BytesIO(payload))
            self.assertEqual(book.sheetnames, ["業者＿A", "業者＿A_2"])
            self.assertEqual(len(printed_pages(book.worksheets[0])), 3)
            self.assertEqual(len(printed_pages(book.worksheets[1])), 1)
            book.close()


if __name__ == "__main__":
    unittest.main()
