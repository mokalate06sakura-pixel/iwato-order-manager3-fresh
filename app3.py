import io
import re
import tempfile
from pathlib import Path
from datetime import datetime

import pandas as pd
import streamlit as st
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.worksheet.page import PageMargins
from openpyxl.utils import get_column_letter
# 補助機能は、ファイル不足や内部エラーでアプリ全体が停止しないよう安全に読み込む
MARUHACHI_IMPORT_ERROR = None
HOKUBU_IMPORT_ERROR = None

try:
    from create_order_form_maruhachi import generate_maruhachi_order_forms_both_facilities
except Exception as exc:
    generate_maruhachi_order_forms_both_facilities = None
    MARUHACHI_IMPORT_ERROR = exc

try:
    from create_order_form_hokubu import generate_hokubu_order_forms_both_facilities
except Exception as exc:
    generate_hokubu_order_forms_both_facilities = None
    HOKUBU_IMPORT_ERROR = exc
# ------------------------------------------------------------
# Streamlit 基本設定
# ------------------------------------------------------------
st.set_page_config(page_title="発注・検収サポートシステム", layout="wide")

def apply_cute_theme():
    st.markdown("""
    <style>
    /* 全体背景 */
    .stApp {
        background: linear-gradient(180deg, #FFF7FB 0%, #F7FAFF 100%);
    }

    /* ページ横幅を少し締める（読みやすい） */
    .block-container {
        padding-top: 2rem;
        max-width: 1000px;
    }

    /* タイトルをかわいく */
    h1, h2, h3 {
        letter-spacing: 0.02em;
    }
    h1 {
        font-weight: 800;
        color: #6B4E71;
    }

    /* カード風コンテナ */
    .cute-card {
        background: rgba(255,255,255,0.85);
        border: 1px solid rgba(255, 192, 203, 0.35);
        border-radius: 18px;
        padding: 18px 18px 10px 18px;
        box-shadow: 0 10px 30px rgba(107, 78, 113, 0.08);
        backdrop-filter: blur(6px);
        margin-bottom: 14px;
    }
    .cute-label {
        font-weight: 700;
        color: #6B4E71;
        margin-bottom: 6px;
        display: flex;
        align-items: center;
        gap: 8px;
    }
    .cute-hint {
        color: rgba(107,78,113,0.7);
        font-size: 0.92rem;
        margin-bottom: 10px;
    }

    /* file_uploader をカードっぽく */
    [data-testid="stFileUploader"] {
        background: rgba(255,255,255,0.6);
        border: 1px dashed rgba(107,78,113,0.25);
        border-radius: 14px;
        padding: 12px;
    }

    /* ボタンをぷっくり可愛く */
    .stButton > button {
        background: linear-gradient(90deg, #FFB6C1 0%, #C7B3FF 100%);
        color: white;
        border: 0;
        border-radius: 999px;
        padding: 0.70rem 1.2rem;
        font-weight: 800;
        box-shadow: 0 10px 20px rgba(199, 179, 255, 0.25);
        transition: transform .08s ease-in-out, box-shadow .08s ease-in-out;
        width: 100%;
    }
    .stButton > button:hover {
        transform: translateY(-1px);
        box-shadow: 0 14px 24px rgba(199, 179, 255, 0.30);
    }
    .stButton > button:active {
        transform: translateY(1px);
        box-shadow: 0 8px 14px rgba(199, 179, 255, 0.20);
    }

    /* 成功・エラー表示をやさしく */
    [data-testid="stAlert"] {
        border-radius: 14px;
        border: 1px solid rgba(107,78,113,0.15);
    }

    /* 区切り線 */
    hr {
        border: none;
        height: 1px;
        background: rgba(107,78,113,0.15);
        margin: 18px 0;
    }
    </style>
    """, unsafe_allow_html=True)

apply_cute_theme()
# ------------------------------------------------------------
# 献ダテマン風 ゆるかわスタイル
# ------------------------------------------------------------
CUSTOM_CSS = """
<style>
body {
    background-color: #fffdf8;
}
.main {
    background-image: linear-gradient(90deg, rgba(0,0,0,0.03) 1px, transparent 1px),
                      linear-gradient(180deg, rgba(0,0,0,0.03) 1px, transparent 1px);
    background-size: 24px 24px;
}
.app-title {
    font-size: 34px;
    font-weight: bold;
    color: #ff7f50;
    padding: 0.3rem 1.4rem;
    display: inline-block;
    border-radius: 999px;
    background: #fff0e6;
    border: 2px solid #ffa76b;
}
.subtitle-pill {
    display: inline-block;
    padding: 0.25rem 1rem;
    border-radius: 999px;
    font-size: 13px;
    font-weight: 600;
    margin-right: 0.5rem;
    color: white;
}
.sub-orange { background: #ff9b50; }
.sub-green  { background: #5cb85c; }
.sub-blue   { background: #5bc0de; }

.feature-card {
    background: white;
    border-radius: 18px;
    padding: 1.2rem 1.5rem;
    margin-bottom: 1.4rem;
    box-shadow: 0 3px 6px rgba(0,0,0,0.06);
    border: 1px solid #f2e4d5;
}
.feature-title {
    font-weight: bold;
    font-size: 18px;
    margin-bottom: 0.3rem;
    color: #444;
}
.feature-sub {
    font-size: 12px;
    color: #777;
    margin-bottom: 0.7rem;
}
.small-note {
    font-size: 11px;
    color: #777;
    margin-top: 0.4rem;
}
.btn-cute {
    background: #ffb27a !important;
    color: white !important;
    font-weight: bold !important;
    border-radius: 10px !important;
}
hr.soft {
    border: none;
    border-top: 1px dashed #e0cbb0;
    margin: 0.4rem 0 0.8rem 0;
}

/* ---------------------------
   サイドバー: 献ダテマン風
---------------------------- */
[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #f7f5f2 0%, #efece6 100%);
    border-right: 1px solid rgba(120, 120, 120, 0.15);
}

[data-testid="stSidebar"] .block-container {
    padding-top: 1.2rem;
    padding-left: 1rem;
    padding-right: 1rem;
}

/* 自動ページメニュー上の「app3」をシステム名に置き換える */
[data-testid="stSidebarNav"] ul li:first-child a span {
    font-size: 0 !important;
}

[data-testid="stSidebarNav"] ul li:first-child a span::after {
    content: "発注・検収サポートシステム";
    font-size: 14px;
    color: #5f6570;
    white-space: normal;
}

/* サイドバー見出し */
[data-testid="stSidebar"] h2,
[data-testid="stSidebar"] h3 {
    color: #4a4a4a;
    font-weight: 800;
    margin-bottom: 0.6rem;
}

/* radio 全体 */
[data-testid="stSidebar"] [role="radiogroup"] {
    display: flex;
    flex-direction: column;
    gap: 2px;
    margin-top: 0.35rem;
}

/* 各メニュー項目：献立マン風の細い一覧 */
[data-testid="stSidebar"] [role="radiogroup"] label {
    background: #fff7fe;
    border: 1px solid #f28ae8;
    border-radius: 1px;
    padding: 7px 9px;
    box-shadow: none;
    transition: all 0.15s ease-in-out;
    cursor: pointer;
}

/* 検収簿整形は「検収関連」の水色 */
[data-testid="stSidebar"] [role="radiogroup"] label:first-of-type {
    background: #f0fcff;
    border-color: #52cce8;
}

/* 2番目の項目の直前に「発注関連」の見出しを表示 */
[data-testid="stSidebar"] [role="radiogroup"] label:nth-of-type(2) {
    position: relative;
    margin-top: 42px;
}

[data-testid="stSidebar"] [role="radiogroup"] label:nth-of-type(2)::before {
    content: "▼ 発注関連";
    position: absolute;
    top: -40px;
    left: -1px;
    width: calc(100% + 2px);
    box-sizing: border-box;
    padding: 7px 9px;
    color: #b600ad;
    font-size: 16px;
    font-weight: 900;
    background: linear-gradient(180deg, #fff0fd 0%, #f47ce9 100%);
    border: 1px solid #e55ada;
}

/* ホバー */
[data-testid="stSidebar"] [role="radiogroup"] label:hover {
    transform: none;
    box-shadow: inset 3px 0 0 rgba(219, 43, 202, 0.55);
    filter: brightness(0.98);
}

/* ラベル文字 */
[data-testid="stSidebar"] [role="radiogroup"] label p {
    font-size: 14px;
    font-weight: 700;
    color: #4a3548;
    margin: 0;
}

/* ラジオ丸 */
[data-testid="stSidebar"] input[type="radio"] {
    transform: scale(0.95);
    accent-color: #ef63df;
}

/* 選択中の項目 */
[data-testid="stSidebar"] label:has(input[type="radio"]:checked) {
    background: #ffd9fa;
    border-color: #df39d1;
    box-shadow: inset 4px 0 0 #df39d1;
}

[data-testid="stSidebar"] label:has(input[type="radio"]:checked) p {
    color: #8d087f;
}

/* 検収関連が選択中のときは水色 */
[data-testid="stSidebar"] label:first-of-type:has(input[type="radio"]:checked) {
    background: #c9f5ff;
    border-color: #24b9db;
    box-shadow: inset 4px 0 0 #24b9db;
}

[data-testid="stSidebar"] label:first-of-type:has(input[type="radio"]:checked) p {
    color: #087b96;
}

/* サイドバー上部 */
.sidebar-section-title {
    display: block;
    background: linear-gradient(180deg, #ffffff 0%, #e9e9e9 100%);
    border: 1px solid #bdbdbd;
    color: #333333;
    font-weight: 900;
    padding: 7px 9px;
    margin-bottom: 0.5rem;
    font-size: 17px;
}

.menu-group-title-blue {
    display: block;
    color: #087b96;
    font-size: 16px;
    font-weight: 900;
    padding: 7px 9px;
    margin-top: 0.45rem;
    margin-bottom: 0;
    background: linear-gradient(180deg, #edfcff 0%, #69d7ed 100%);
    border: 1px solid #42c5e0;
}
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

# ------------------------------------------------------------
# 共通ユーティリティ
# ------------------------------------------------------------
def parse_mmdd(value: str):
    """文字列 '12/8月' などから月日だけ抜き出して datetime に変換"""
    if value is None:
        return None
    s = str(value)
    m = re.search(r"\d+/\d+", s)
    if not m:
        return None
    try:
        return datetime.strptime(m.group(), "%m/%d").replace(year=2000)
    except Exception:
        return None

def detect_min_usage_date_token(df, col="使用日"):
    """使用日の最も古い日付を MMDD 形式 '1208' のように返す"""
    if col not in df.columns:
        return ""
    dt_list = [parse_mmdd(v) for v in df[col]]
    dt_list = [d for d in dt_list if d is not None]
    if not dt_list:
        return ""
    return min(dt_list).strftime("%m%d")


# ------------------------------------------------------------
# ① 検収簿整形ロジック（修正版：不要列を削除）
# ------------------------------------------------------------
def _is_blank(value):
    """Excelの空白セル（NaN、None、空文字、空白だけの文字列）を判定する。"""
    return pd.isna(value) or (isinstance(value, str) and value.strip() == "")


def apply_ek_blank_rows_and_f_zero(df):
    """VBA「EK空行削除_後にF空白へ0」と同じデータ整形を行う。

    元ファイルの見出し行は pandas が読み取る際に除外されているため、
    ここではデータ行だけを対象にする。
    """
    if df.shape[1] >= 11:
        # ExcelのE～K列（0始まりでは4～10）がすべて空白の行を削除
        ek_all_blank = df.iloc[:, 4:11].apply(
            lambda row: all(_is_blank(value) for value in row), axis=1
        )
        df = df.loc[~ek_all_blank].copy()

    if df.shape[1] >= 6:
        # ExcelのF列（0始まりでは5）の空白を0で埋める
        f_col = df.columns[5]
        f_blank = df[f_col].map(_is_blank)
        df.loc[f_blank, f_col] = 0

    return df


def format_inspection_workbook(uploaded_file):
    df = pd.read_excel(uploaded_file, header=[6, 7])

    # VBA「EK空行削除_後にF空白へ0」を自動適用
    df = apply_ek_blank_rows_and_f_zero(df)

    # ---- MultiIndex → フラット化 ----
    flat_cols = []
    for top, sub in df.columns:
        top = "" if str(top).startswith("Unnamed") else str(top)
        sub = "" if str(sub).startswith("Unnamed") else str(sub)

        if top == "":
            flat_cols.append(sub)
        elif sub == "":
            flat_cols.append(top)
        else:
            flat_cols.append(f"{top}_{sub}")

    df.columns = flat_cols

    # F列に相当する「換算値」の空白を確実に0で補完する。
    # 元Excelの列位置だけでなく、見出し名でも処理することで、
    # 結合見出しや列構成の違いがあっても出力へ反映させる。
    if "換算値" in df.columns:
        conversion_blank = df["換算値"].map(_is_blank)
        df.loc[conversion_blank, "換算値"] = 0

    # ---- 欠損補完 ----
    for col in ["納品日", "使用日", "朝昼夕", "仕入先"]:
        if col in df.columns:
            df[col] = df[col].ffill()

    # ---- 朝昼夕用の並び順 ----
    order_map = {"朝食": 1, "昼食": 2, "夕食": 3}
    df["食事順"] = df["朝昼夕"].map(order_map).fillna(0)

    # ---- ソート ----
    df = df.sort_values(["使用日", "食事順", "食品名"])

    # ------------------------------------------------------------
    # 🔥 ここが今回の重要修正ポイント
    # ------------------------------------------------------------

    # ❶ いわと列名（確実に拾えるように）
    iwato_in = [c for c in df.columns if "いわと" in c and "入所" in c]
    iwato_staff = [c for c in df.columns if "いわと" in c and "職員" in c]

    # ❷ ユーハウス列名（部分一致で拾う）
    yuhouse_in = [
        c for c in df.columns 
        if ("ユーハウス" in c or "ユー" in c or "ケアハウス" in c)
        and "入" in c
    ]

    # デバッグ表示（必要なら） print(df.columns)

    # ---- 最終的に残す列 ----
    needed_cols = [
        "納品日",
        "使用日",
        "朝昼夕",
        "仕入先",
        "食品名",
        "換算値",
        "総合計",
        "単位",
    ]

    # 自動で見つけた列を追加
    needed_cols += iwato_in[:1]           # I列：いわと入所者
    needed_cols += iwato_staff[:1]        # J列：いわと職員
    needed_cols += yuhouse_in[:1]         # L列：ユーハウス入所者

    # 存在するものだけ残す
    needed_cols = [c for c in needed_cols if c in df.columns]

    df_out = df[needed_cols].copy()

    # Excel出力直前にも再確認し、換算値の空白を残さない
    if "換算値" in df_out.columns:
        conversion_blank = df_out["換算値"].map(_is_blank)
        df_out.loc[conversion_blank, "換算値"] = 0

    # ---- 出力 ----
    buffer = io.BytesIO()
    df_out.to_excel(buffer, index=False)
    buffer.seek(0)

    token = detect_min_usage_date_token(df_out, "使用日")
    fname = f"検収簿_加工済_{token}.xlsx" if token else "検収簿_加工済.xlsx"

    return buffer.read(), fname

# ------------------------------------------------------------
# 業者別仕訳表 作成ロジック
# ------------------------------------------------------------
def _safe_sheet_name(value, used_names):
    """仕入先名をExcelで使用可能な一意のシート名へ変換する。"""
    name = str(value).strip() if not pd.isna(value) else "仕入先未設定"
    name = re.sub(r'[\\/*?:\[\]]', '＿', name)
    name = name[:31] or "仕入先未設定"

    base = name
    index = 2
    while name in used_names:
        suffix = f"_{index}"
        name = f"{base[:31 - len(suffix)]}{suffix}"
        index += 1

    used_names.add(name)
    return name


def create_vendor_journal_workbook(uploaded_file):
    """加工済み検収簿から、仕入先ごとのテーブル付き仕訳表を作成する。"""
    df = pd.read_excel(uploaded_file)

    if "仕入先" not in df.columns:
        raise ValueError("『仕入先』列が見つかりません。検収簿（加工済）を選択してください。")

    # A列（通常は納品日）は削除せず、Excel出力後に非表示にする。
    # 列番号ではなく元の見出し名を判定して名称を変更する。
    # これによりH列「単位」を保持し、I・J・K列だけを変更する。
    rename_map = {}

    for col in df.columns:
        col_text = str(col)

        if "介護老人福祉施設いわと" in col_text and "入所者" in col_text:
            rename_map[col] = "特養入所者"

        elif "介護老人福祉施設いわと" in col_text and "職員" in col_text:
            rename_map[col] = "特養職員"

        elif (
            ("ケアハウス" in col_text or "ユーハウス" in col_text or "ユー" in col_text)
            and ("入所者" in col_text or "入居者" in col_text)
            and "職員" not in col_text
        ):
            rename_map[col] = "ユーハウス"

    df = df.rename(columns=rename_map)

    required_headers = ["単位", "特養入所者", "特養職員", "ユーハウス"]
    missing_headers = [name for name in required_headers if name not in df.columns]
    if missing_headers:
        raise ValueError(
            "必要な列が見つかりません: "
            + "、".join(missing_headers)
            + "。検収簿（加工済）の見出しを確認してください。"
        )

    # 各業者シートの最終列に、手入力用の「コメント」欄を追加
    # 既にコメント列が存在する場合も、最終列へ移動する。
    if "コメント" in df.columns:
        comment_values = df.pop("コメント")
        df["コメント"] = comment_values
    else:
        df["コメント"] = ""

    # 仕入先が空白の行も独立シートにまとめる
    df["仕入先"] = df["仕入先"].fillna("仕入先未設定").astype(str).str.strip()
    df.loc[df["仕入先"] == "", "仕入先"] = "仕入先未設定"

    buffer = io.BytesIO()
    used_names = set()

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for table_index, (vendor, vendor_df) in enumerate(df.groupby("仕入先", sort=True), start=1):
            sheet_name = _safe_sheet_name(vendor, used_names)
            vendor_df = vendor_df.reset_index(drop=True)
            vendor_df.to_excel(writer, sheet_name=sheet_name, index=False)

            ws = writer.book[sheet_name]

            # A3・横向き印刷設定
            ws.page_setup.paperSize = ws.PAPERSIZE_A3
            ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
            ws.sheet_properties.pageSetUpPr.fitToPage = True
            ws.page_setup.fitToWidth = 1
            ws.page_setup.fitToHeight = 0
            ws.print_options.horizontalCentered = True
            ws.freeze_panes = "B2"

            # A列はデータを保持したまま非表示にする
            ws.column_dimensions["A"].hidden = True

            # 印刷範囲
            ws.print_area = ws.dimensions
            ws.page_margins = PageMargins(
                left=0.25, right=0.25, top=0.4, bottom=0.4,
                header=0.2, footer=0.2
            )

            # 見出しと列幅
            for cell in ws[1]:
                cell.font = Font(name="ＭＳ ゴシック", size=11, bold=True)
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

            for column_cells in ws.columns:
                letter = get_column_letter(column_cells[0].column)
                max_length = max(
                    (len(str(cell.value)) if cell.value is not None else 0)
                    for cell in column_cells
                )
                ws.column_dimensions[letter].width = min(max(max_length + 2, 10), 32)

            ws.row_dimensions[1].height = 32

    buffer.seek(0)
    token = detect_min_usage_date_token(df, "使用日")
    fname = f"業者別仕訳表_{token}.xlsx" if token else "業者別仕訳表.xlsx"
    return buffer.read(), fname


# ------------------------------------------------------------
# 注文書 書式設定（いわと／ユーハウス共通）
# ------------------------------------------------------------
def apply_order_style(ws, is_tokuyou=False):
    font_body = Font(name="ＭＳ ゴシック", size=18)
    border = Border(
        left=Side("thin"),
        right=Side("thin"),
        top=Side("thin"),
        bottom=Side("thin")
    )

    header_row = 6

    # --- 6行目：ヘッダー行 ---
    for cell in ws[header_row]:
        cell.font = Font(name="ＭＳ ゴシック", size=12, bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border

    # --- 7行目以降：データ行 ---
    for row in ws.iter_rows(min_row=header_row + 1):
        for c in row:
            c.font = font_body
            c.border = border
            c.alignment = Alignment(
                vertical="center",
                wrap_text=False,      # 折り返しなし
            )

    # --- 行高 ---
    for i in range(1, ws.max_row + 1):
        ws.row_dimensions[i].height = 30

    # ------------------------------------------------------------
    # 列幅設定（注文書仕様）
    # ------------------------------------------------------------

    # A列：使用日
    ws.column_dimensions["A"].width = 15.18

    # B列：食品名（広く）
    ws.column_dimensions["B"].width = 60.09

    # D〜H列：7.73 に変更（数量・単位・確認欄）
    for col in ["D", "E", "F", "G", "H"]:
        ws.column_dimensions[col].width = 7.73

    # C・I・J・K・L・M は 15.18
    for col in ["C", "I", "J", "K", "L", "M"]:
        ws.column_dimensions[col].width = 15.18

    # 特養用マクロの指定
    if is_tokuyou:
        for col in ["I", "L", "M"]:
            ws.column_dimensions[col].width = 7

        # C～E列の外枠を太線にする（内側の罫線は維持）
        thick = Side(style="thick")
        start_row = 6
        end_row = ws.max_row
        if end_row >= start_row:
            for row in range(start_row, end_row + 1):
                ws.cell(row, 3).border = Border(
                    left=thick,
                    right=ws.cell(row, 3).border.right,
                    top=thick if row == start_row else ws.cell(row, 3).border.top,
                    bottom=thick if row == end_row else ws.cell(row, 3).border.bottom,
                )
                ws.cell(row, 5).border = Border(
                    left=ws.cell(row, 5).border.left,
                    right=thick,
                    top=thick if row == start_row else ws.cell(row, 5).border.top,
                    bottom=thick if row == end_row else ws.cell(row, 5).border.bottom,
                )
            for col in range(3, 6):
                ws.cell(start_row, col).border = Border(
                    left=ws.cell(start_row, col).border.left,
                    right=ws.cell(start_row, col).border.right,
                    top=thick,
                    bottom=ws.cell(start_row, col).border.bottom,
                )
                ws.cell(end_row, col).border = Border(
                    left=ws.cell(end_row, col).border.left,
                    right=ws.cell(end_row, col).border.right,
                    top=ws.cell(end_row, col).border.top,
                    bottom=thick,
                )

    # ------------------------------------------------------------
    # B列（食品名）を縮小して全体表示
    # ------------------------------------------------------------
    for row in ws.iter_rows(min_row=7, max_row=ws.max_row, min_col=2, max_col=2):
        for cell in row:
            cell.alignment = Alignment(
                horizontal="left",
                vertical="center",
                wrap_text=False,        # 折り返しなし
                shrink_to_fit=True      # 縮小して全体を表示
            )

    # ------------------------------------------------------------
    # 印刷設定
    # ------------------------------------------------------------
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_margins = PageMargins(left=0.3, right=0.3, top=0.5, bottom=0.5)

    # 印刷範囲（A〜M列）
    ws.print_area = f"A1:M{ws.max_row}"



# ------------------------------------------------------------
# ヘッダー（いわと）
# ------------------------------------------------------------
def create_header_iwato(ws, supplier):
    ws.merge_cells("A3:B3")
    ws["A3"] = f"{supplier} 御中"
    ws["A3"].font = Font(name="ＭＳ ゴシック", size=28, bold=True)

    ws["B1"] = "注文書（介護老人福祉施設いわと）"
    ws["B1"].alignment = Alignment(horizontal="center")
    ws["B1"].font = Font(name="ＭＳ ゴシック", size=26, bold=True)

    ws["K3"] = "(有) ハートミール"
    ws["K3"].alignment = Alignment(horizontal="right")
    ws["K3"].font = Font(name="ＭＳ ゴシック", size=24, bold=True)



# ------------------------------------------------------------
# ヘッダー（ユーハウス）
# ------------------------------------------------------------
def create_header_yuhouse(ws, supplier):
    ws.merge_cells("A3:B3")
    ws["A3"] = f"{supplier} 御中"
    ws["A3"].font = Font(name="ＭＳ ゴシック", size=28, bold=True)

    ws["B1"] = "注文書（ユーハウスいわと）"
    ws["B1"].alignment = Alignment(horizontal="center")
    ws["B1"].font = Font(name="ＭＳ ゴシック", size=26, bold=True)

    ws["K3"] = "(有) ハートミール"
    ws["K3"].alignment = Alignment(horizontal="right")
    ws["K3"].font = Font(name="ＭＳ ゴシック", size=24, bold=True)


# ------------------------------------------------------------
# ③ 注文書作成（特養 / ユーハウス 共通・並び順修正版）
# ------------------------------------------------------------
def create_order_workbook(uploaded_file, order_type):
    df = pd.read_excel(uploaded_file)

    # 欠損補完
    for c in ["使用日", "仕入先", "食品名", "単位"]:
        if c in df.columns:
            df[c] = df[c].ffill()

    df["使用日"] = df["使用日"].astype(str)

    # ------------------------------------------------------------
    # 🔶 特養（いわと）
    # ------------------------------------------------------------
    if "特養" in order_type:
        raw_qty = "介護老人福祉施設いわと_入所者"
        raw_staff = "介護老人福祉施設いわと_職員"

        if raw_qty not in df.columns:
            df[raw_qty] = 0
        if raw_staff not in df.columns:
            df[raw_staff] = 0

        df[raw_qty] = pd.to_numeric(df[raw_qty], errors="coerce").fillna(0)
        df[raw_staff] = pd.to_numeric(df[raw_staff], errors="coerce").fillna(0)

    # ------------------------------------------------------------
    # 🔷 ユーハウス（ケアハウス）
    # ------------------------------------------------------------
    else:
        # ゆるマッチで入居者列を探す
        cand_cols = [
            c for c in df.columns
            if ("ケアハウス" in c or "ユー" in c or "ユ" in c)
            and ("入" in c or "居" in c)
            and ("職" not in c)
        ]

        if len(cand_cols) == 0:
            raw_qty = "ケアハウス入居者"
            df[raw_qty] = 0
        else:
            raw_qty = cand_cols[0]  # 例：ケアハウスユー…_入所者

        df[raw_qty] = pd.to_numeric(df.get(raw_qty, 0), errors="coerce").fillna(0)
        raw_staff = None  # ユーハウスは職員欄なし

    # ------------------------------------------------------------
    # 評価項目の空列作成
    # ------------------------------------------------------------
    for c in ["鮮度", "品温", "異物", "包装", "期限", "備考欄", "検収者"]:
        if c not in df.columns:
            df[c] = ""

    df["納品日"] = ""  # 納品日は常に空欄

    suppliers = df["仕入先"].dropna().unique()

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:

        for supplier in suppliers:
            sub = df[df["仕入先"] == supplier].copy()

            sub["使用日_dt"] = sub["使用日"].apply(parse_mmdd)
            sub = sub.sort_values(["使用日_dt", "食品名"])

            # 表示名に変換
            if "特養" in order_type:
                sub = sub.rename(columns={
                    raw_qty: "入所者",
                    raw_staff: "職員"
                })
                qty_label = "入所者"
                staff_label = "職員"
            else:
                sub = sub.rename(columns={raw_qty: "ユーハウス入居者"})
                qty_label = "ユーハウス入居者"
                staff_label = None

            # 並べる列順
            col_order = [
                "使用日",
                "食品名",
                qty_label,
                "単位",
            ]

            if staff_label:
                col_order.append(staff_label)

            col_order += [
                "鮮度", "品温", "異物", "包装", "期限",
                "備考欄", "納品日", "検収者"
            ]

            # 不足列を補完
            for c in col_order:
                if c not in sub.columns:
                    sub[c] = ""

            sub = sub[col_order]

            # VBAと同様に、主数量（C列）が空白または0の行を削除
            qty_values = pd.to_numeric(sub[qty_label], errors="coerce").fillna(0)
            sub = sub.loc[qty_values != 0].copy()

            # 特養用：職員数（E列）が0なら空欄にする
            if staff_label:
                staff_values = pd.to_numeric(sub[staff_label], errors="coerce").fillna(0)
                sub[staff_label] = sub[staff_label].astype(object)
                sub.loc[staff_values == 0, staff_label] = ""

            # 同じ使用日は2行目以降空欄に
            sub["使用日"] = sub["使用日"].mask(sub["使用日"].duplicated(), "")

            sheet = str(supplier)[:30]
            sub.to_excel(writer, sheet_name=sheet, index=False, startrow=5)

        # 書式 & ヘッダー
        wb = writer.book
        for supplier in suppliers:
            ws = wb[str(supplier)[:30]]
            apply_order_style(ws, is_tokuyou="特養" in order_type)

            if "特養" in order_type:
                create_header_iwato(ws, supplier)
            else:
                create_header_yuhouse(ws, supplier)
                ws["C6"] = "ユーハウス入居者"

    # ファイル名（使用日の最古日）
    token = detect_min_usage_date_token(df, "使用日")

    if "特養" in order_type:
        base_name = "注文書_いわと"
    else:
        base_name = "注文書_ユーハウス"

    fname = f"{base_name}_{token}.xlsx" if token else f"{base_name}.xlsx"

    buffer.seek(0)
    return buffer.read(), fname



# ------------------------------------------------------------
# 🖥️ UI構築（左メニューでページ切替）
# ------------------------------------------------------------
st.markdown(
    """
<div style="margin-bottom: 1.5rem;">
  <span class="app-title">発注・検収サポートシステム</span>
</div>

<div style="margin-bottom: 2.0rem;">
  <span class="subtitle-pill sub-orange">毎日の業務をかんたんに</span>
  <span class="subtitle-pill sub-green">発注書を自動作成</span>
  <span class="subtitle-pill sub-blue">検収簿を整形</span>
</div>
""",
    unsafe_allow_html=True,
)

with st.sidebar:
    st.markdown(
        '<div class="sidebar-section-title">▦ ユーザーメニュー</div>',
        unsafe_allow_html=True,
    )
    st.markdown(
        '<div class="menu-group-title-blue">▼ 検収関連</div>',
        unsafe_allow_html=True,
    )

    page = st.radio(
        "画面を選択してください",
               [
            "① 検収簿整形",
            "② 業者別仕訳表",
            "③ 注文書作成",
            "④ 丸八発注書作成",
            "⑤ 北部市場発注書作成",
        ],
        label_visibility="collapsed",
    )

# ------------------------------------------------------------
# ① 検収簿整形
# ------------------------------------------------------------
if page == "① 検収簿整形":
    st.markdown(
        """
<div class="feature-card">
  <div class="feature-title">① 検収簿を整える"</div>
  <div class="feature-sub">
    MultiIndex の検収記録簿を<br>
    A〜H列だけの加工済みファイルに整形します。<br>
    ※ 献ダテマンから出力したファイルを<br>
      「<b>検収記録簿_原本.xlsx</b>」の名前で保存して下さい。
  </div>
  <hr class="soft"/>
</div>
        """,
        unsafe_allow_html=True,
    )

    ins_file = st.file_uploader(
        "検収簿（原本 Excel）をアップロード",
        type=["xlsx"],
        key="ins"
    )

    if ins_file:
        if st.button("📘 検収簿を整形する", key="btn_ins"):
            st.session_state["ins_data"], st.session_state["ins_fname"] = \
                format_inspection_workbook(ins_file)
            st.success("検収簿の整形が完了しました！")

        if "ins_data" in st.session_state:
            st.download_button(
                "📥 検収簿（加工済）をダウンロード",
                st.session_state["ins_data"],
                st.session_state["ins_fname"]
            )

# ------------------------------------------------------------
# ② 業者別仕訳表
# ------------------------------------------------------------
elif page == "② 業者別仕訳表":
    st.markdown(
        """
<div class="feature-card">
  <div class="feature-title">② 業者別仕訳表を作成</div>
  <div class="feature-sub">
    加工済み検収簿を仕入先ごとのシートに分割し、<br>
    通常のセル形式・A3横向き印刷設定で出力します。
  </div>
  <hr class="soft"/>
</div>
        """,
        unsafe_allow_html=True,
    )

    vendor_file = st.file_uploader(
        "検収簿（加工済 Excel）をアップロード",
        type=["xlsx"],
        key="vendor_journal_src",
    )

    if vendor_file:
        try:
            if st.button("📊 業者別仕訳表を作成する", key="btn_vendor_journal"):
                (
                    st.session_state["vendor_journal_data"],
                    st.session_state["vendor_journal_fname"],
                ) = create_vendor_journal_workbook(vendor_file)
                st.success("業者別仕訳表の作成が完了しました！")

            if "vendor_journal_data" in st.session_state:
                st.download_button(
                    "📥 業者別仕訳表をダウンロード",
                    data=st.session_state["vendor_journal_data"],
                    file_name=st.session_state["vendor_journal_fname"],
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

        except Exception as e:
            st.error("業者別仕訳表の作成中にエラーが発生しました。")
            st.exception(e)


# ------------------------------------------------------------
# ③ 注文書作成
# ------------------------------------------------------------
elif page == "③ 注文書作成":
    st.markdown(
        """
<div class="feature-card">
  <div class="feature-title">③ 注文書を作成</div>
  <div class="feature-sub">
    特養（介護老人福祉施設いわと）<br>
    かユーハウスいわと を選んで、<br>
    仕入先別にシート作成された注文書を作成します。
  </div>
  <hr class="soft"/>
</div>
        """,
        unsafe_allow_html=True,
    )

    order_type = st.radio(
        "作成する注文書の種類を選んでください",
        ("特養（介護老人福祉施設いわと）", "ユーハウスいわと"),
        horizontal=True,
        key="order_type",
    )

    order_file = st.file_uploader(
        "注文書のもとになる検収簿 Excel をアップロード",
        type=["xlsx"],
        key="order_src",
    )

    st.markdown(
        '<p class="small-note">※ 検収簿整形で加工したもの、または同じ形式の検収簿ファイルを想定しています。</p>',
        unsafe_allow_html=True,
    )

    if order_file:
        try:
            if st.button("📗 注文書を作成する", key="btn_order"):
                st.session_state["order_data"], st.session_state["order_fname"] = \
                    create_order_workbook(order_file, order_type)
                st.success(f"{order_type} の注文書が作成されました！")

            if "order_data" in st.session_state:
                st.download_button(
                    "📥 注文書ファイルをダウンロード",
                    st.session_state["order_data"],
                    st.session_state["order_fname"],
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

        except Exception as e:
            st.error("注文書作成中にエラーが発生しました。アップロードファイルを確認してください。")
            st.exception(e)

# ------------------------------------------------------------
# ④ 丸八発注書作成
# ------------------------------------------------------------
elif page == "④ 丸八発注書作成":
    st.markdown(
        """
<div class="feature-card">
  <div class="feature-title">④ 丸八発注書を作成</div>
  <div class="feature-sub">
    検収簿_加工済、丸八テンプレ、丸八コード一覧を使って<br>
    特養用・ユーハウス用の丸八発注書を自動作成します。
  </div>
  <hr class="soft"/>
</div>
        """,
        unsafe_allow_html=True,
    )

    mcol1, mcol2, mcol3 = st.columns(3)

    with mcol1:
        st.markdown(
            """
            <div class="feature-card">
              <div class="feature-title">📄 検収簿_加工済</div>
              <div class="feature-sub">丸八ヒロタの行を含む加工済ファイル</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        kenshu_file = st.file_uploader(
            "検収簿_加工済（xlsx）",
            type=["xlsx"],
            key="kenshu_maruhachi"
        )

    with mcol2:
        st.markdown(
            """
            <div class="feature-card">
              <div class="feature-title">🧾 丸八発注書テンプレ</div>
              <div class="feature-sub">丸八ヒロタ専用の発注書テンプレート</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        template_file = st.file_uploader(
            "丸八発注書テンプレ（xlsm）",
            type=["xlsm"],
            key="tpl_maruhachi"
        )

    with mcol3:
        st.markdown(
            """
            <div class="feature-card">
              <div class="feature-title">🏷️ 丸八コード一覧</div>
              <div class="feature-sub">タグシート付きのコード対応表</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        tag_file = st.file_uploader(
            "丸八コード一覧（xlsm）",
            type=["xlsm"],
            key="tag_maruhachi"
        )

    st.markdown("### 出力")
    bcol1, bcol2 = st.columns(2)

    with bcol1:
        btn = st.button("📦 丸八発注書を作成", key="btn_maruhachi")

    with bcol2:
        st.markdown(
            '<p class="small-note">※ 特養用とユーハウス用を同時に作成します。</p>',
            unsafe_allow_html=True,
        )

    if btn:
        if generate_maruhachi_order_forms_both_facilities is None:
            st.error("丸八発注書機能を読み込めませんでした。")
            st.exception(MARUHACHI_IMPORT_ERROR)
        elif not (kenshu_file and template_file and tag_file):
            st.error("⚠ 3つのファイル（検収簿・テンプレ・コード一覧）をすべて選択してください。")
        else:
            st.success("丸八発注書を作成中です…")

            with tempfile.TemporaryDirectory() as td:
                td = Path(td)

                k_path = td / "kenshu.xlsx"
                t_path = td / "template.xlsm"
                m_path = td / "tag.xlsm"

                k_path.write_bytes(kenshu_file.getbuffer())
                t_path.write_bytes(template_file.getbuffer())
                m_path.write_bytes(tag_file.getbuffer())

                out_dir = td / "out"

                tokuyou_xlsm, yuhouse_xlsm = generate_maruhachi_order_forms_both_facilities(
                    kenshu_xlsx_path=k_path,
                    template_xlsm_path=t_path,
                    tag_xlsm_path=m_path,
                    out_dir=out_dir,
                    out_prefix="丸八発注書",
                )

                st.success("作成完了 ✅ ダウンロードできます")

                dcol1, dcol2 = st.columns(2)
                with dcol1:
                    st.download_button(
                        "📥 特養：丸八発注書をダウンロード",
                        data=tokuyou_xlsm.read_bytes(),
                        file_name=tokuyou_xlsm.name,
                        mime="application/vnd.ms-excel.sheet.macroEnabled.12",
                    )
                with dcol2:
                    st.download_button(
                        "📥 ユーハウス：丸八発注書をダウンロード",
                        data=yuhouse_xlsm.read_bytes(),
                        file_name=yuhouse_xlsm.name,
                        mime="application/vnd.ms-excel.sheet.macroEnabled.12",
                    )

# ------------------------------------------------------------
# ⑤ 北部市場発注書作成
# ------------------------------------------------------------
elif page == "⑤ 北部市場発注書作成":
    st.markdown(
        """
<div class="feature-card">
  <div class="feature-title">⑤ 北部市場発注書を作成</div>
  <div class="feature-sub">
    検収簿_加工済と北部市場発注書テンプレを使って<br>
    特養用・ユーハウス用の発注書を自動作成します。
  </div>
  <hr class="soft"/>
</div>
        """,
        unsafe_allow_html=True,
    )

    hokubu_kenshu = st.file_uploader(
        "検収簿_加工済（xlsx）",
        type=["xlsx"],
        key="hokubu_kenshu"
    )

    hokubu_template = st.file_uploader(
        "北部市場発注書テンプレ（xlsm）",
        type=["xlsm"],
        key="hokubu_tpl"
    )

    btn_hokubu = st.button("📦 北部市場発注書を作成", key="btn_hokubu")

    if btn_hokubu:
        if generate_hokubu_order_forms_both_facilities is None:
            st.error("北部市場発注書機能を読み込めませんでした。")
            st.exception(HOKUBU_IMPORT_ERROR)
        elif not (hokubu_kenshu and hokubu_template):
            st.error("⚠ 検収簿_加工済 と 北部市場テンプレを選択してください。")
        else:
            st.success("北部市場発注書を作成中です…")

            with tempfile.TemporaryDirectory() as td:
                td = Path(td)
                k_path = td / "kenshu.xlsx"
                t_path = td / "template.xlsm"

                k_path.write_bytes(hokubu_kenshu.getbuffer())
                t_path.write_bytes(hokubu_template.getbuffer())

                out_dir = td / "out"

                tokuyou_xlsm, yuhouse_xlsm = generate_hokubu_order_forms_both_facilities(
                    kenshu_xlsx_path=k_path,
                    template_xlsm_path=t_path,
                    out_dir=out_dir,
                    out_prefix="北部市場発注書",
                )

                st.success("北部市場発注書を作成しました。")

                c1, c2 = st.columns(2)
                with c1:
                    st.download_button(
                        "📥 特養：北部市場発注書をダウンロード",
                        data=tokuyou_xlsm.read_bytes(),
                        file_name=tokuyou_xlsm.name,
                        mime="application/vnd.ms-excel.sheet.macroEnabled.12",
                    )
                with c2:
                    st.download_button(
                        "📥 ユーハウス：北部市場発注書をダウンロード",
                        data=yuhouse_xlsm.read_bytes(),
                        file_name=yuhouse_xlsm.name,
                        mime="application/vnd.ms-excel.sheet.macroEnabled.12",
                    )
