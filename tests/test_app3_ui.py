"""実際のStreamlit画面、登録済みマスタと一時選択の導線を検証する。"""
import io
import json
from pathlib import Path
import tempfile
import unittest

from streamlit.testing.v1 import AppTest

from app3_test_support import APP_PATH
from test_app3_features import workbook_upload


def injected_app(directory, uploads=None, fake_supplier=False):
    """ファイル選択だけを差し替え、アプリのUIと生成ボタンはそのまま実行。"""
    files = {}
    for key, path in (uploads or {}).items():
        files[key] = str(path)
    source = f'''
import io
import json
from pathlib import Path
from unittest.mock import patch
import streamlit as st

class Upload(io.BytesIO):
    def __init__(self, path):
        super().__init__(Path(path).read_bytes())
        self.name = Path(path).name
        self.size = len(self.getvalue())

files = {files!r}
def uploader(*args, **kwargs):
    path = files.get(kwargs.get("key"))
    return Upload(path) if path else None

def generate(**kwargs):
    output = Path(kwargs["out_dir"])
    output.mkdir(exist_ok=True)
    template = Path(kwargs["template_xlsm_path"])
    observed = {{"template_contents": template.read_bytes().hex(), "template_path": str(template)}}
    template.write_bytes(b"changed by generator")
    Path({str(Path(directory) / 'called.json')!r}).write_text(json.dumps(observed), encoding="utf-8")
    paths = (output / "tokuyou.xlsm", output / "yuhouse.xlsm")
    for path in paths:
        path.write_bytes(b"generated workbook")
    return paths

source = Path({str(APP_PATH)!r}).read_text(encoding="utf-8")
source = source.replace('MASTER_DIR = Path(__file__).resolve().parent / "masters"',
                        {('MASTER_DIR = Path(' + repr(str(Path(directory) / 'masters')) + ')')!r})
with patch("streamlit.file_uploader", side_effect=uploader):
    if {fake_supplier!r}:
        with patch("create_order_form_hokubu.generate_hokubu_order_forms_both_facilities", side_effect=generate):
            exec(compile(source, {str(APP_PATH)!r}, "exec"), {{"__file__": {str(APP_PATH)!r}, "__name__": "__main__"}})
    else:
        exec(compile(source, {str(APP_PATH)!r}, "exec"), {{"__file__": {str(APP_PATH)!r}, "__name__": "__main__"}})
'''
    return AppTest.from_string(source, default_timeout=30)


class AppScreensTests(unittest.TestCase):
    def test_all_eight_screens_and_missing_master_message(self):
        with tempfile.TemporaryDirectory() as directory:
            app = injected_app(directory).run()
            options = app.radio(key="main_page_selector").options
            self.assertEqual([label[0] for label in options], list("⓪①②③④⑤⑥⑦"))
            for page in options:
                app.radio(key="main_page_selector").set_value(page).run()
                self.assertEqual(list(app.exception), [], page)
            app.radio(key="main_page_selector").set_value("⑤ 北部市場発注書作成").run()
            app.button(key="btn_hokubu").click().run()
            self.assertTrue(any("検収簿" in item.value for item in app.warning))

    def test_registered_master_and_override_are_copied_without_modification(self):
        for override in (False, True):
            with self.subTest(override=override), tempfile.TemporaryDirectory() as directory:
                root = Path(directory)
                master_dir = root / "masters"
                master_dir.mkdir()
                master = workbook_upload(["特養 (北部市場)", "ユーハウス(北部市場)"])
                original = master.getvalue()
                (master_dir / "hokubu_template.xlsm").write_bytes(original)
                (master_dir / "meta.json").write_text("{}", encoding="utf-8")
                inspection = root / "inspection.xlsx"
                inspection.write_bytes(b"inspection is consumed by fake generator")
                upload_paths = {"hokubu_kenshu": inspection}
                expected = original
                if override:
                    alternate = workbook_upload(["特養 (北部市場)", "ユーハウス(北部市場)", "別の版"])
                    alternate_path = root / "alternate.xlsm"
                    alternate_path.write_bytes(alternate.getvalue())
                    upload_paths["hokubu_tpl"] = alternate_path
                    expected = alternate.getvalue()
                app = injected_app(root, upload_paths, fake_supplier=True).run()
                app.radio(key="main_page_selector").set_value("⑤ 北部市場発注書作成").run()
                app.button(key="btn_hokubu").click().run()
                self.assertEqual(list(app.exception), [])
                self.assertTrue(any("作成しました" in item.value for item in app.success))
                observed = json.loads((root / "called.json").read_text(encoding="utf-8"))
                self.assertEqual(bytes.fromhex(observed["template_contents"]), expected)
                self.assertNotEqual(Path(observed["template_path"]), master_dir / "hokubu_template.xlsm")
                self.assertEqual((master_dir / "hokubu_template.xlsm").read_bytes(), original)
                self.assertEqual((master_dir / "meta.json").read_text(encoding="utf-8"), "{}")
                self.assertIn("hokubu_tokuyou_data", app.session_state)

    def test_master_registration_screen(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "source.xlsm"
            path.write_bytes(workbook_upload(["タグ"]).getvalue())
            app = injected_app(directory, {"master_upload_maruhachi_code": path}).run()
            app.radio(key="main_page_selector").set_value("⑦ マスタ管理").run()
            app.button(key="master_register_maruhachi_code").click().run()
            self.assertEqual(list(app.exception), [])
            self.assertEqual((Path(directory) / "masters/maruhachi_code.xlsm").read_bytes(), path.read_bytes())
            self.assertTrue(any("丸八コード一覧を登録しました" in item.value for item in app.success))

    def test_real_sample_menu_and_card_generation_ui(self):
        home = Path.home()
        uploads = {
            "menu_check_menu": home / "Downloads/献立予定（実施）表 (93).xlsx",
            "menu_check_master": home / "Downloads/在庫マスタ一覧表 (1).xlsx",
            "card_kenshu": home / "Desktop/検収簿_加工済_1102.xlsx",
        }
        if not all(p.exists() for p in uploads.values()):
            self.skipTest("実サンプルがありません")
        with tempfile.TemporaryDirectory() as directory:
            app = injected_app(directory, uploads).run()
            app.button(key="menu_check_run").click().run()
            self.assertEqual(list(app.exception), [])
            self.assertEqual(len(app.session_state["menu_check_reds"]), 1)
            app.radio(key="main_page_selector").set_value("⑥ 仕入れカード作成").run()
            selected = app.multiselect(key="card_dates")
            self.assertEqual(selected.value, selected.options)
            app.button(key="card_run").click().run()
            self.assertEqual(list(app.exception), [])
            self.assertEqual(app.session_state["card_stats"]["card_count"], 287)
            self.assertEqual(len(app.session_state["card_preview"]), 100)
            app.multiselect(key="card_dates").set_value([]).run()
            app.button(key="card_run").click().run()
            self.assertTrue(any("1日以上" in item.value for item in app.warning))
            self.assertNotIn("card_data", app.session_state)


if __name__ == "__main__":
    unittest.main()
