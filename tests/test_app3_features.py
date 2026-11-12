import ast
from collections import Counter
from copy import copy
from decimal import Decimal
import io
import json
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

import openpyxl
import pandas as pd

from app3_test_support import APP_PATH, Upload, load_logic


def workbook_upload(sheets, name="master.xlsm"):
    wb = openpyxl.Workbook()
    wb.active.title = sheets[0]
    for name_ in sheets[1:]:
        wb.create_sheet(name_)
    buffer = io.BytesIO()
    wb.save(buffer)
    wb.close()
    return Upload(buffer.getvalue(), name)


def inventory_upload(mark="○", food_header="食品名"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["B2"], ws["C2"], ws["F2"] = "在", "発注種類", food_header
    # A列にコードがなくても食品行は数える。重複名は件数と集合を区別する。
    for row in (4, 5):
        ws.cell(row, 2, mark)
        ws.cell(row, 6, "発注しない食材")
    ws["F6"] = "発注する食材"
    buffer = io.BytesIO()
    wb.save(buffer)
    wb.close()
    return Upload(buffer.getvalue(), "在庫マスタ.xlsx")


class MenuCheckTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.logic = load_logic()

    def test_inventory_validation_counts_and_fresh_read(self):
        names, stats = self.logic.load_no_order_names(inventory_upload())
        self.assertEqual(names, {"発注しない食材"})
        self.assertEqual(stats["マスタ件数"], 3)
        self.assertEqual(stats["発注しない（○）件数"], 2)
        with self.assertRaises(ValueError):
            self.logic.load_no_order_names(inventory_upload(food_header="食品名コード"))
        with self.assertRaises(ValueError):
            self.logic.load_no_order_names(inventory_upload(mark=""))

    def test_all_blocks_both_dish_markers_and_three_error_types(self):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws["A1"] = "年月日：2026年11月2日（月）"
        ws["A2"] = "朝食"
        for _, col, gram in self.logic.MENU_CHECK_BLOCKS:
            ws.cell(3, col, "■料理A")
            ws.cell(4, col, "　食材")
            ws.cell(4, gram, "1 g")
            ws.cell(5, col, "◆料理B")
            ws.cell(6, col, "　食材")
            ws.cell(6, gram, "1 g")
        ws["A7"], ws["F7"] = "　ゼロ食材", "0 g"
        ws["A8"] = "　空欄食材"
        ws["A9"], ws["F9"] = "　食材", "1 g"
        ws["A10"], ws["F10"] = "　＊＊＊", "0 g"
        ws["A11"], ws["F11"] = "　発注しない食材", "0 g"
        ws["A12"] = "３時"
        ws["A13"], ws["F13"] = "　食材", "1 g"
        buffer = io.BytesIO()
        wb.save(buffer)
        wb.close()
        items = self.logic.extract_menu_ingredients(Upload(buffer.getvalue(), "献立.xlsx"))
        reds, excluded = self.logic.run_menu_check(items, {"発注しない食材"})
        self.assertEqual(Counter(r["種別"] for r in reds), {"A1": 1, "A2": 1, "A3": 1})
        self.assertEqual(excluded["＊＊＊（使用しない仮置き）"], 1)
        self.assertEqual(excluded["発注しない（在庫マスタ B列○）"], 1)
        self.assertEqual(items[-1]["食事区分"], "３時")
        self.assertEqual(items[-1]["料理名"], "")

    def test_reference_menu_has_one_real_error(self):
        root = Path.home() / "Downloads"
        menu = root / "献立予定（実施）表 (93).xlsx"
        master = root / "在庫マスタ一覧表 (1).xlsx"
        if not menu.exists() or not master.exists():
            self.skipTest("実サンプルがないため合成データのみ検証")
        names, stats = self.logic.load_no_order_names(Upload.from_path(master))
        items = self.logic.extract_menu_ingredients(Upload.from_path(menu))
        reds, excluded = self.logic.run_menu_check(items, names)
        self.assertEqual(len(items), 785)
        self.assertEqual(len(reds), 1)
        self.assertEqual(reds[0]["種別"], "A1")
        self.assertEqual(reds[0]["行"], 30)
        self.assertEqual(excluded["発注しない（在庫マスタ B列○）"], 386)
        self.assertEqual(excluded["＊＊＊（使用しない仮置き）"], 10)
        data = self.logic.build_menu_check_workbook(reds, excluded, len(items), sorted({i["日付"] for i in items}), stats)
        result = openpyxl.load_workbook(io.BytesIO(data))
        self.assertEqual(result.sheetnames, ["サマリ", "要修正（赤）"])
        self.assertEqual(result.worksheets[1].max_row, 2)
        self.assertEqual(result["サマリ"]["B6"].value, master.name)
        result.close()


class CardTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.logic = load_logic()

    def test_cleaning_order_filters_counts_and_decimal_precision(self):
        rows = [
            ["11/3", " B\t\n", " 豆乳\n", "本", 0.72, 0.18, 0, 0.9],
            ["11/2", "A", "野菜", "kg", 0.00012345, 0, 0, 0.00012345],
            ["11/2", "B", "別品", "本", 1, 0, 0, 99],
            ["11/3", "B", "豆乳", "本", 0.1, 0, 0, 0.1],
            ["11/3", "B", "", "本", 1, 0, 0, 1],
            ["11/3", "B", "ゼロ", "本", 0, 0, 0, 0],
        ]
        frame = pd.DataFrame(rows, columns=self.logic.CARD_REQUIRED_COLS + ["総合計"])
        buffer = io.BytesIO()
        frame.to_excel(buffer, index=False)
        cards, stats = self.logic.aggregate_card_rows(io.BytesIO(buffer.getvalue()))
        self.assertEqual([(c["仕入先"], c["使用日"], c["食品名"]) for c in cards],
                         [("B", "11/3", "豆乳"), ("B", "11/2", "別品"), ("A", "11/2", "野菜")])
        self.assertEqual(cards[0]["数量"], 1)
        self.assertEqual(cards[-1]["数量"], 0.00012345)
        self.assertEqual(stats["total_mismatch_rows"], 1)
        self.assertEqual(stats["skipped_blank_name"], 1)
        self.assertEqual(stats["skipped_zero_qty"], 1)
        chosen, _ = self.logic.aggregate_card_rows(io.BytesIO(buffer.getvalue()), {"11/2"})
        self.assertEqual(len(chosen), 2)
        empty, _ = self.logic.aggregate_card_rows(io.BytesIO(buffer.getvalue()), set())
        self.assertEqual(empty, [])

    def test_page_geometry_card_positions_and_styles(self):
        cards = [{"仕入先": "業者", "使用日": "11/2", "食品名": "食品" + str(i), "数量": i + 1, "単位": "本"} for i in range(5)]
        wb = openpyxl.load_workbook(io.BytesIO(self.logic.build_cards_workbook(cards)))
        ws = wb["仕入れカード"]
        width_cm = sum((ws.column_dimensions[c].width * 7 + 5) / 96 * 2.54 for c in "ABCDEFGHI")
        self.assertAlmostEqual(width_cm, 28.7, places=10)
        for page in range(2):
            height_cm = sum(ws.row_dimensions[r].height for r in range(page * 21 + 1, page * 21 + 22)) / 72 * 2.54
            self.assertAlmostEqual(height_cm, 20.0, places=10)
            self.assertEqual(ws.row_dimensions[page * 21 + 11].height, 66.6)
        self.assertEqual([ws[a].value for a in ("A3", "F3", "A14", "F14", "A24")], [f"食品{i}" for i in range(5)])
        self.assertEqual(ws["C5"].font.sz, 25)
        self.assertEqual(ws["D5"].font.sz, 18)
        self.assertEqual(ws["A1"].font.name, "Meiryo UI")
        self.assertEqual(ws["D1"].font.sz, 18)
        self.assertEqual(ws["A3"].font.sz, 30)
        self.assertEqual(ws["A1"].border.top.style, "medium")
        self.assertTrue(ws["B2"].border.left is None or ws["B2"].border.left.style is None)
        self.assertEqual([b.id for b in ws.row_breaks.brk], [21])
        self.assertEqual(str(ws.page_setup.paperSize), "9")
        self.assertEqual(ws.page_setup.orientation, "landscape")
        self.assertEqual(ws.page_setup.fitToWidth, 1)
        self.assertEqual(ws.page_setup.fitToHeight, 0)
        self.assertTrue(ws.sheet_properties.pageSetUpPr.fitToPage)
        self.assertTrue(ws.print_area.endswith("$A$1:$I$42"))
        for margin in ("left", "right", "top", "bottom"):
            self.assertAlmostEqual(getattr(ws.page_margins, margin), 0.5 / 2.54)
        self.assertEqual(self.logic._card_food_font_size("字" * 16), 18)
        self.assertEqual(self.logic._card_food_font_size("字" * 12), 23)
        wb.close()

    def test_reference_cards_independent_aggregation_and_155_bottles(self):
        source = Path.home() / "Desktop" / "検収簿_加工済_1102.xlsx"
        if not source.exists():
            self.skipTest("実サンプルがないため合成データのみ検証")
        cards, stats = self.logic.aggregate_card_rows(source)
        frame = pd.read_excel(source)
        self.assertEqual(len(frame), 294)
        self.assertEqual(len(cards), 287)
        expected = {}
        for _, row in frame.iterrows():
            key = tuple(str(row[c]).strip() for c in ("仕入先", "使用日", "食品名", "単位"))
            qty = sum(Decimal(str(row[c])) for c in ("特養入所者", "特養職員", "ユーハウス"))
            expected[key] = expected.get(key, Decimal(0)) + qty
        self.assertEqual({(c["仕入先"], c["使用日"], c["食品名"], c["単位"]): Decimal(str(c["数量"])) for c in cards}, expected)
        spot = [c for c in cards if c["使用日"] == "11/2月" and "三給" in c["仕入先"] and c["食品名"] == "まめぴよココア125ml"]
        self.assertEqual(len(spot), 1)
        self.assertEqual((spot[0]["数量"], spot[0]["単位"]), (155, "本"))
        self.assertEqual(stats["total_mismatch_rows"], 0)


class MasterTests(unittest.TestCase):
    def setUp(self):
        self.logic = load_logic()
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.logic.MASTER_DIR = Path(self.directory.name) / "masters"

    def test_registration_metadata_overwrite_and_private_temporary_copy(self):
        identifier = "hokubu_template"
        self.assertFalse(self.logic.get_master_status(identifier)["registered"])
        upload = workbook_upload(self.logic.MASTER_SPECS[identifier]["sheets"], "元テンプレ.xlsm")
        status = self.logic.register_master_file(identifier, upload)
        self.assertEqual(status["original_name"], "元テンプレ.xlsm")
        self.assertTrue(status["saved_at"].endswith("+09:00"))
        self.assertEqual(status["path"].name, "hokubu_template.xlsm")
        metadata = json.loads((self.logic.MASTER_DIR / "meta.json").read_text(encoding="utf-8"))
        self.assertEqual(set(metadata), {identifier})
        destination = Path(self.directory.name) / "template.xlsm"
        self.logic.copy_master_to_temp(identifier, destination)
        self.assertEqual(destination.read_bytes(), upload.getvalue())
        destination.write_bytes(b"generator changed its temporary copy")
        self.assertEqual(status["path"].read_bytes(), upload.getvalue())
        override = workbook_upload(self.logic.MASTER_SPECS[identifier]["sheets"], "今回だけ.xlsm")
        self.logic.copy_master_to_temp(identifier, destination, override)
        self.assertEqual(self.logic.get_master_status(identifier)["original_name"], "元テンプレ.xlsm")
        self.logic.register_master_file(identifier, override)
        self.assertEqual(self.logic.get_master_status(identifier)["original_name"], "今回だけ.xlsm")

    def test_missing_bad_and_failed_registration_do_not_destroy_master(self):
        with self.assertRaisesRegex(ValueError, "⑦で登録"):
            self.logic.copy_master_to_temp("maruhachi_code", Path(self.directory.name) / "copy.xlsm")
        good = workbook_upload(["タグ"])
        status = self.logic.register_master_file("maruhachi_code", good)
        before = status["path"].read_bytes()
        with self.assertRaises(ValueError):
            self.logic.register_master_file("maruhachi_code", workbook_upload(["別シート"]))
        self.assertEqual(status["path"].read_bytes(), before)
        with self.assertRaises(ValueError):
            self.logic.register_master_file("maruhachi_code", Upload(b"bad data", "wrong.xlsm"))
        original_write = self.logic._master_atomic_write
        def fail_metadata(path, payload):
            if path.name == "meta.json":
                raise OSError("metadata save failed")
            original_write(path, payload)
        with patch.object(self.logic, "_master_atomic_write", side_effect=fail_metadata):
            with self.assertRaises(OSError):
                self.logic.register_master_file("maruhachi_code", good)
        self.assertEqual(status["path"].read_bytes(), before)
        with self.assertRaises(ValueError):
            self.logic.get_master_status("../outside")

    def test_real_registered_masters_generate_both_supplier_outputs_without_changes(self):
        root = APP_PATH.parent
        sources = {
            "maruhachi_template": root / "丸八発注システム用/丸八ヒロタ発注書テンプレ改.xlsm",
            "maruhachi_code": root / "丸八発注システム用/丸八ヒロタ発注書コード一覧.xlsm",
            "hokubu_template": root / "北部市場発注システム用/北部市場発注書テンプレ1.xlsm",
        }
        inspection = Path.home() / "Desktop/検収簿_加工済_1102.xlsx"
        if not inspection.exists() or not all(p.exists() for p in sources.values()):
            self.skipTest("実テンプレートがありません")
        before = {key: path.read_bytes() for key, path in sources.items()}
        for key, path in sources.items():
            self.logic.register_master_file(key, Upload.from_path(path))
        output_root = Path(self.directory.name)
        copies = {key: self.logic.copy_master_to_temp(key, output_root / (key + ".xlsm")) for key in sources}
        from create_order_form_hokubu import generate_hokubu_order_forms_both_facilities
        from create_order_form_maruhachi import generate_maruhachi_order_forms_both_facilities
        outputs = generate_hokubu_order_forms_both_facilities(
            kenshu_xlsx_path=inspection, template_xlsm_path=copies["hokubu_template"],
            out_dir=output_root / "hokubu", out_prefix="北部市場発注書")
        outputs += generate_maruhachi_order_forms_both_facilities(
            kenshu_xlsx_path=inspection, template_xlsm_path=copies["maruhachi_template"],
            tag_xlsm_path=copies["maruhachi_code"], out_dir=output_root / "maruhachi", out_prefix="丸八発注書")
        self.assertEqual(len(outputs), 4)
        for path in outputs:
            book = openpyxl.load_workbook(path)
            self.assertTrue(book.sheetnames)
            book.close()
        for key, source in sources.items():
            self.assertEqual(source.read_bytes(), before[key])
            self.assertEqual(self.logic.get_master_status(key)["path"].read_bytes(), before[key])


class PreservationTests(unittest.TestCase):
    def test_original_functions_screens_and_supplier_arguments_unchanged(self):
        baseline = APP_PATH.parent / "backups/app3_features_20261112/app3_before.py"
        if not baseline.exists():
            self.skipTest("修正前バックアップがないため比較を省略")
        old, new = [ast.parse(p.read_text(encoding="utf-8")) for p in (baseline, APP_PATH)]
        old_functions = {n.name: n for n in old.body if isinstance(n, ast.FunctionDef) and n.lineno < 1597}
        new_functions = {n.name: n for n in new.body if isinstance(n, ast.FunctionDef)}
        for name, node in old_functions.items():
            self.assertEqual(ast.dump(node), ast.dump(new_functions[name]), name)
        def screens(tree):
            return {n.test.comparators[0].value: n for n in ast.walk(tree)
                    if isinstance(n, ast.If) and isinstance(n.test, ast.Compare)
                    and isinstance(n.test.left, ast.Name) and n.test.left.id == "page"
                    and isinstance(n.test.comparators[0], ast.Constant)}
        before, after = screens(old), screens(new)
        for page in ("① 検収簿整形", "② 業者別仕訳表", "③ 注文書作成"):
            self.assertEqual([ast.dump(n) for n in before[page].body], [ast.dump(n) for n in after[page].body])
        for name in ("generate_maruhachi_order_forms_both_facilities", "generate_hokubu_order_forms_both_facilities"):
            calls = lambda tree: [ast.dump(n) for n in ast.walk(tree) if isinstance(n, ast.Call) and isinstance(n.func, ast.Name) and n.func.id == name]
            self.assertEqual(calls(old), calls(new))

    def test_input_changes_clear_only_their_own_results(self):
        logic = load_logic()
        state = {"card_data": b"old", "card_dates": ["11/2"], "menu_check_data": b"other", "ins_data": b"existing"}
        with patch.object(logic.st, "session_state", state):
            logic._card_source_changed()
        self.assertEqual(state, {"menu_check_data": b"other", "ins_data": b"existing"})


if __name__ == "__main__":
    unittest.main()
