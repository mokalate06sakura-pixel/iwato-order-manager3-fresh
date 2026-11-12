import io
import re
import tempfile
from pathlib import Path
from datetime import datetime

import pandas as pd
import streamlit as st
import openpyxl
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

# ------------------------------------------------------------
# 🌸 Streamlit かわいい安定版テーマ
# ------------------------------------------------------------
def apply_cute_theme():
    st.html(
        """
        <style>

        /* ====================================================
           全体
        ==================================================== */
        .stApp {
            background:
                linear-gradient(
                    180deg,
                    #FFF8FC 0%,
                    #F7FAFF 55%,
                    #FFFDF8 100%
                );
        }

        .block-container {
            padding-top: 2.4rem;
            padding-bottom: 3rem;
            max-width: 1080px;
        }

        /* ====================================================
           アプリ上部タイトル
        ==================================================== */
        .app-hero {
            background:
                linear-gradient(
                    135deg,
                    #FFF0F7 0%,
                    #F1F7FF 55%,
                    #FFF9EC 100%
                );
            border: 1px solid #F0DCE9;
            border-radius: 24px;
            padding: 24px 28px;
            margin-bottom: 24px;
            box-shadow:
                0 7px 22px rgba(110, 85, 120, 0.09);
        }

        .app-hero-title {
            font-size: 30px;
            font-weight: 800;
            color: #6B4E71;
            margin-bottom: 8px;
        }

        .app-hero-sub {
            font-size: 16px;
            color: #665B68;
            line-height: 1.8;
            margin-bottom: 14px;
        }

        .cute-pill {
            display: inline-block;
            padding: 7px 12px;
            margin-right: 7px;
            margin-bottom: 5px;
            border-radius: 999px;
            font-size: 14px;
            font-weight: 700;
        }

        .pill-pink {
            background: #FFEAF3;
            color: #A55278;
        }

        .pill-green {
            background: #E9F8EE;
            color: #4B8060;
        }

        .pill-blue {
            background: #EAF3FF;
            color: #4C7199;
        }

        /* ====================================================
           各ページのメインカード
        ==================================================== */
        .main-feature-card {
            background: rgba(255,255,255,0.92);
            border: 1px solid #F0DCE9;
            border-radius: 22px;
            padding: 24px 28px;
            margin-top: 4px;
            margin-bottom: 22px;
            box-shadow:
                0 6px 18px rgba(120,90,120,0.08);
        }

        .main-feature-title {
            font-size: 24px;
            font-weight: 800;
            color: #6B4E71;
            margin-bottom: 10px;
        }

        .main-feature-sub {
            font-size: 16px;
            color: #594F5C;
            line-height: 1.85;
        }

        /* ====================================================
           注意書き
        ==================================================== */
        .main-note-box {
            background: #FFF8E8;
            border-left: 5px solid #F1B66B;
            border-radius: 12px;
            padding: 12px 16px;
            margin-top: 15px;
            color: #6A5643;
            line-height: 1.7;
        }

        .main-note-blue {
            background: #EEF6FF;
            border-left: 5px solid #8BB9E8;
            border-radius: 12px;
            padding: 12px 16px;
            margin-top: 15px;
            color: #4B6075;
            line-height: 1.7;
        }

        .main-note-green {
            background: #EFF9F2;
            border-left: 5px solid #8BC69D;
            border-radius: 12px;
            padding: 12px 16px;
            margin-top: 15px;
            color: #496453;
            line-height: 1.7;
        }

        /* ====================================================
           見出し
        ==================================================== */
        h1 {
            color: #6B4E71;
            font-weight: 800;
        }

        h2, h3 {
            color: #59445E;
            font-weight: 750;
        }

        /* ====================================================
           ボタン
        ==================================================== */
        .stButton > button {
            border-radius: 16px;
            font-weight: 700;
            min-height: 48px;
            border: 1px solid #E7D7E5;
            box-shadow:
                0 3px 9px rgba(100,80,110,0.08);
        }

        .stButton > button:hover {
            border-color: #D9B8D2;
            box-shadow:
                0 5px 12px rgba(100,80,110,0.12);
        }

        /* ====================================================
           ダウンロードボタン
        ==================================================== */
        [data-testid="stDownloadButton"] button {
            border-radius: 16px;
            font-weight: 700;
            min-height: 48px;
            border: 1px solid #D8E5D9;
            box-shadow:
                0 3px 9px rgba(90,120,100,0.08);
        }

        /* ====================================================
           ファイルアップロード
        ==================================================== */
        [data-testid="stFileUploader"] {
            background: #FBFCFF;
            border: 1px solid #E1E8F2;
            border-radius: 18px;
            padding: 8px;
        }

        [data-testid="stFileUploaderDropzone"] {
            border-radius: 14px;
            background: #F5F8FC;
        }

        /* ====================================================
           通知
        ==================================================== */
        [data-testid="stAlert"] {
            border-radius: 14px;
        }

        /* ====================================================
           サイドバー
        ==================================================== */
        [data-testid="stSidebar"] {
            background:
                linear-gradient(
                    180deg,
                    #FFF8FB 0%,
                    #F7F4F1 100%
                );
        }

        .sidebar-section-title {
            margin-bottom: 0.8rem;
            padding: 11px 12px;
            border-radius: 13px;
            background: #F6EAF4;
            color: #6B4E71;
            font-weight: 800;
            font-size: 18px;
        }

        .menu-group-title-blue {
            margin-top: 0.8rem;
            margin-bottom: 0.5rem;
            padding: 8px 10px;
            border-radius: 10px;
            background: #EAF4FF;
            color: #3F6F9F;
            font-weight: 700;
        }

        .menu-group-title-orange {
            margin-top: 1.2rem;
            margin-bottom: 0.5rem;
            padding: 8px 10px;
            border-radius: 10px;
            background: #FFF1E8;
            color: #B85C38;
            font-weight: 700;
        }

        /* ====================================================
           小さいカード
        ==================================================== */
        .mini-card {
            background: rgba(255,255,255,0.90);
            border: 1px solid #E7E2EB;
            border-radius: 17px;
            padding: 15px 17px;
            margin-bottom: 10px;
        }

        .mini-card-title {
            color: #6B4E71;
            font-size: 17px;
            font-weight: 800;
            margin-bottom: 5px;
        }

        .mini-card-sub {
            color: #69616C;
            font-size: 14px;
            line-height: 1.6;
        }

        </style>
        """
    )


apply_cute_theme()

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


# ------------------------------------------------------------
# ①・② 共通 Excel印刷書式
# A3縦 / 罫線 / 納品日区切り線 / 行高26 / 文字16
# ------------------------------------------------------------
def apply_inspection_print_style(ws):
    # 基本フォント
    body_font = Font(
        name="ＭＳ ゴシック",
        size=16
    )

    header_font = Font(
        name="ＭＳ ゴシック",
        size=16,
        bold=True
    )

    # 罫線
    thin = Side(
        style="thin",
        color="000000"
    )

    medium = Side(
        style="medium",
        color="000000"
    )

    thin_border = Border(
        left=thin,
        right=thin,
        top=thin,
        bottom=thin
    )

    # A3縦
    ws.page_setup.paperSize = ws.PAPERSIZE_A3
    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT

    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0

    ws.print_options.horizontalCentered = True

    ws.page_margins = PageMargins(
        left=0.25,
        right=0.25,
        top=0.35,
        bottom=0.35,
        header=0.15,
        footer=0.15
    )

    # 全セル
    for row in range(1, ws.max_row + 1):
        ws.row_dimensions[row].height = 26

        for col in range(1, ws.max_column + 1):
            cell = ws.cell(
                row=row,
                column=col
            )

            if row == 1:
                cell.font = header_font
                cell.alignment = Alignment(
                    horizontal="center",
                    vertical="center",
                    wrap_text=True
                )
            else:
                cell.font = body_font
                cell.alignment = Alignment(
                    vertical="center",
                    wrap_text=False
                )

            cell.border = thin_border

    # 見出し名と列番号
    header_map = {}

    for col in range(1, ws.max_column + 1):
        value = ws.cell(
            row=1,
            column=col
        ).value

        if value is not None:
            header_map[str(value).strip()] = col

    # 納品日が変わったら上に太線
    delivery_col = header_map.get("納品日")

    if delivery_col is not None:
        previous_value = None

        for row in range(2, ws.max_row + 1):
            current_value = ws.cell(
                row=row,
                column=delivery_col
            ).value

            current_text = (
                ""
                if current_value is None
                else str(current_value).strip()
            )

            if row == 2:
                previous_value = current_text
                continue

            if current_text != previous_value:
                for col in range(1, ws.max_column + 1):
                    cell = ws.cell(
                        row=row,
                        column=col
                    )

                    cell.border = Border(
                        left=thin,
                        right=thin,
                        top=medium,
                        bottom=thin
                    )

            previous_value = current_text

    # 列幅
    width_map = {
        "納品日": 15,
        "使用日": 15,
        "朝昼夕": 11,
        "仕入先": 24,
        "食品名": 38,
        "換算値": 13,
        "総合計": 13,
        "単位": 11,
        "特養入所者": 15,
        "特養職員": 15,
        "ユーハウス": 15,
        "コメント": 28,
    }

    for header, width in width_map.items():
        col_num = header_map.get(header)

        if col_num is not None:
            letter = get_column_letter(col_num)
            ws.column_dimensions[letter].width = width

    # 食品名は左揃え
    food_col = header_map.get("食品名")

    if food_col is not None:
        for row in range(2, ws.max_row + 1):
            ws.cell(
                row=row,
                column=food_col
            ).alignment = Alignment(
                horizontal="left",
                vertical="center",
                wrap_text=False,
                shrink_to_fit=True
            )

    # 印刷範囲
    ws.print_area = (
        f"A1:"
        f"{get_column_letter(ws.max_column)}"
        f"{ws.max_row}"
    )

    # 各ページに見出し行を表示
    ws.print_title_rows = "1:1"

def detect_min_usage_date_token(df, col="使用日"):
    """使用日の最も古い日付を MMDD 形式 '1208' のように返す"""
    if col not in df.columns:
        return ""

    dt_list = [
        parse_mmdd(v)
        for v in df[col]
    ]

    dt_list = [
        d
        for d in dt_list
        if d is not None
    ]

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

    # ------------------------------------------------------------
    # MultiIndex → フラット化
    # ------------------------------------------------------------
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

    # ------------------------------------------------------------
    # 換算値の空白を0
    # ------------------------------------------------------------
    if "換算値" in df.columns:
        conversion_blank = df["換算値"].map(_is_blank)
        df.loc[conversion_blank, "換算値"] = 0

    # ------------------------------------------------------------
    # 欠損補完
    # ------------------------------------------------------------
    for col in ["納品日", "使用日", "朝昼夕", "仕入先"]:
        if col in df.columns:
            df[col] = df[col].ffill()

    # ------------------------------------------------------------
    # 朝昼夕の並び順
    # ------------------------------------------------------------
    order_map = {
        "朝食": 1,
        "昼食": 2,
        "夕食": 3,
    }

    df["食事順"] = (
        df["朝昼夕"]
        .map(order_map)
        .fillna(0)
    )

    # ------------------------------------------------------------
    # ソート
    # ------------------------------------------------------------
    df = df.sort_values(
        ["使用日", "食事順", "食品名"]
    )

    # ------------------------------------------------------------
    # 特養・ユーハウスの元列を検索
    # ------------------------------------------------------------

    # 特養入所者
    iwato_in = [
        c for c in df.columns
        if (
            "いわと" in str(c)
            and "入所" in str(c)
            and "職員" not in str(c)
        )
    ]

    # 特養職員
    iwato_staff = [
        c for c in df.columns
        if (
            "いわと" in str(c)
            and "職員" in str(c)
        )
    ]

    # ユーハウス
    yuhouse_in = [
        c for c in df.columns
        if (
            (
                "ケアハウス" in str(c)
                or "ユーハウス" in str(c)
                or "ユー" in str(c)
            )
            and (
                "入所者" in str(c)
                or "入居者" in str(c)
                or "入" in str(c)
            )
            and "職員" not in str(c)
        )
    ]

    # ------------------------------------------------------------
    # 必要列
    # ------------------------------------------------------------
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

    if iwato_in:
        needed_cols.append(iwato_in[0])

    if iwato_staff:
        needed_cols.append(iwato_staff[0])

    if yuhouse_in:
        needed_cols.append(yuhouse_in[0])

    # 実際に存在する列だけ残す
    needed_cols = [
        c for c in needed_cols
        if c in df.columns
    ]

    df_out = df[needed_cols].copy()

    # ------------------------------------------------------------
    # ★列名を統一
    # ------------------------------------------------------------
    rename_map = {}

    if iwato_in:
        rename_map[iwato_in[0]] = "特養入所者"

    if iwato_staff:
        rename_map[iwato_staff[0]] = "特養職員"

    if yuhouse_in:
        rename_map[yuhouse_in[0]] = "ユーハウス"

    df_out = df_out.rename(
        columns=rename_map
    )

    # ------------------------------------------------------------
    # 必須の数量列が取得できたか確認
    # ------------------------------------------------------------
    missing_facility_cols = []

    if "特養入所者" not in df_out.columns:
        missing_facility_cols.append("特養入所者")

    if "特養職員" not in df_out.columns:
        missing_facility_cols.append("特養職員")

    if "ユーハウス" not in df_out.columns:
        missing_facility_cols.append("ユーハウス")

    if missing_facility_cols:
        raise ValueError(
            "検収簿から次の数量列を取得できませんでした："
            + "、".join(missing_facility_cols)
            + "。元の検収記録簿の見出しを確認してください。"
        )

    # ------------------------------------------------------------
    # 換算値を再確認
    # ------------------------------------------------------------
    if "換算値" in df_out.columns:
        conversion_blank = (
            df_out["換算値"]
            .map(_is_blank)
        )

        df_out.loc[
            conversion_blank,
            "換算値"
        ] = 0

    # ------------------------------------------------------------
    # Excel出力
    # ------------------------------------------------------------
    buffer = io.BytesIO()

    with pd.ExcelWriter(
        buffer,
        engine="openpyxl"
    ) as writer:

        df_out.to_excel(
            writer,
            index=False,
            sheet_name="検収簿"
        )

        ws = writer.book["検収簿"]

        apply_inspection_print_style(ws)

    buffer.seek(0)

    token = detect_min_usage_date_token(
        df_out,
        "使用日"
    )

    fname = (
        f"検収簿_加工済_{token}.xlsx"
        if token
        else "検収簿_加工済.xlsx"
    )

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
    """加工済み検収簿から、仕入先ごとの仕訳表を作成する。"""

    # ------------------------------------------------------------
    # Excel読み込み
    # ------------------------------------------------------------
    df = pd.read_excel(uploaded_file)

    if "仕入先" not in df.columns:
        raise ValueError(
            "『仕入先』列が見つかりません。"
            "検収簿（加工済）を選択してください。"
        )

    # ------------------------------------------------------------
    # 列名を統一
    # ------------------------------------------------------------
    rename_map = {}

    for col in df.columns:
        col_text = str(col)

        if (
            "介護老人福祉施設いわと" in col_text
            and "入所者" in col_text
        ):
            rename_map[col] = "特養入所者"

        elif (
            "介護老人福祉施設いわと" in col_text
            and "職員" in col_text
        ):
            rename_map[col] = "特養職員"

        elif (
            (
                "ケアハウス" in col_text
                or "ユーハウス" in col_text
                or "ユー" in col_text
            )
            and (
                "入所者" in col_text
                or "入居者" in col_text
            )
            and "職員" not in col_text
        ):
            rename_map[col] = "ユーハウス"

    df = df.rename(
        columns=rename_map
    )

    # ------------------------------------------------------------
    # 必須列チェック
    # ------------------------------------------------------------
    required_headers = [
        "単位",
        "特養入所者",
        "特養職員",
        "ユーハウス",
    ]

    missing_headers = [
        name
        for name in required_headers
        if name not in df.columns
    ]

    if missing_headers:
        raise ValueError(
            "必要な列が見つかりません: "
            + "、".join(missing_headers)
            + "。検収簿（加工済）の見出しを確認してください。"
        )

    # ------------------------------------------------------------
    # コメント列を最後へ
    # ------------------------------------------------------------
    if "コメント" in df.columns:
        comment_values = df.pop(
            "コメント"
        )
        df["コメント"] = (
            comment_values
        )
    else:
        df["コメント"] = ""

    # ------------------------------------------------------------
    # 仕入先の空欄処理
    # ------------------------------------------------------------
    df["仕入先"] = (
        df["仕入先"]
        .fillna("仕入先未設定")
        .astype(str)
        .str.strip()
    )

    df.loc[
        df["仕入先"] == "",
        "仕入先"
    ] = "仕入先未設定"

    # ------------------------------------------------------------
    # Excel出力
    # ------------------------------------------------------------
    buffer = io.BytesIO()

    used_names = set()

    with pd.ExcelWriter(
        buffer,
        engine="openpyxl"
    ) as writer:

        for (
            table_index,
            (vendor, vendor_df)
        ) in enumerate(
            df.groupby(
                "仕入先",
                sort=True
            ),
            start=1
        ):

            # ----------------------------------------------------
            # シート名
            # ----------------------------------------------------
            sheet_name = (
                _safe_sheet_name(
                    vendor,
                    used_names
                )
            )

            vendor_df = (
                vendor_df
                .reset_index(
                    drop=True
                )
            )

            # ----------------------------------------------------
            # Excelへ書き込み
            # ----------------------------------------------------
            vendor_df.to_excel(
                writer,
                sheet_name=sheet_name,
                index=False
            )

            ws = writer.book[
                sheet_name
            ]

            # ----------------------------------------------------
            # ①・② 共通の印刷書式を適用
            # ----------------------------------------------------
            apply_inspection_print_style(
                ws
            )

    # ------------------------------------------------------------
    # 出力ファイル名
    # ------------------------------------------------------------
    buffer.seek(0)

    token = (
        detect_min_usage_date_token(
            df,
            "使用日"
        )
    )

    fname = (
        f"業者別仕訳表_{token}.xlsx"
        if token
        else "業者別仕訳表.xlsx"
    )

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
# ③ 注文書作成
# 特養 / ユーハウス
# ------------------------------------------------------------
def create_order_workbook(uploaded_file, order_type):

    df = pd.read_excel(uploaded_file)

    # ------------------------------------------------------------
    # 基本必須列チェック
    # ------------------------------------------------------------
    required_cols = [
        "使用日",
        "仕入先",
        "食品名",
        "単位",
    ]

    missing_cols = [
        c for c in required_cols
        if c not in df.columns
    ]

    if missing_cols:
        raise ValueError(
            "必要な列が見つかりません："
            + "、".join(missing_cols)
            + "。①検収簿整形で作成した"
              "加工済み検収簿を使用してください。"
        )

    # ------------------------------------------------------------
    # 欠損補完
    # ------------------------------------------------------------
    for c in [
        "使用日",
        "仕入先",
        "食品名",
        "単位",
    ]:
        df[c] = df[c].ffill()

    df["使用日"] = (
        df["使用日"]
        .astype(str)
    )

    # ------------------------------------------------------------
    # 特養
    # ------------------------------------------------------------
    if "特養" in order_type:

        raw_qty = "特養入所者"
        raw_staff = "特養職員"

        missing_tokuyou = []

        if raw_qty not in df.columns:
            missing_tokuyou.append(
                "特養入所者"
            )

        if raw_staff not in df.columns:
            missing_tokuyou.append(
                "特養職員"
            )

        if missing_tokuyou:
            raise ValueError(
                "特養の数量列が見つかりません："
                + "、".join(missing_tokuyou)
                + "。①検収簿整形で作成した"
                  "最新の加工済み検収簿を使用してください。"
            )

        df[raw_qty] = pd.to_numeric(
            df[raw_qty],
            errors="coerce"
        ).fillna(0)

        df[raw_staff] = pd.to_numeric(
            df[raw_staff],
            errors="coerce"
        ).fillna(0)

    # ------------------------------------------------------------
    # ユーハウス
    # ------------------------------------------------------------
    else:

        raw_qty = "ユーハウス"
        raw_staff = None

        if raw_qty not in df.columns:
            raise ValueError(
                "『ユーハウス』列が見つかりません。"
                "①検収簿整形で作成した"
                "最新の加工済み検収簿を使用してください。"
            )

        df[raw_qty] = pd.to_numeric(
            df[raw_qty],
            errors="coerce"
        ).fillna(0)

    # ------------------------------------------------------------
    # 検収用空欄
    # ------------------------------------------------------------
    inspection_cols = [
        "鮮度",
        "品温",
        "異物",
        "包装",
        "期限",
        "備考欄",
        "検収者",
    ]

    for c in inspection_cols:
        if c not in df.columns:
            df[c] = ""

    # 納品日は空欄
    df["納品日"] = ""

    # ------------------------------------------------------------
    # 仕入先整理
    # ------------------------------------------------------------
    df["仕入先"] = (
        df["仕入先"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    # 仕入先空欄は除外
    df = df[
        df["仕入先"] != ""
    ].copy()

    suppliers = (
        df["仕入先"]
        .drop_duplicates()
        .tolist()
    )

    if not suppliers:
        raise ValueError(
            "仕入先が見つかりません。"
        )

    # ------------------------------------------------------------
    # Excel出力
    # ------------------------------------------------------------
    buffer = io.BytesIO()

    used_sheet_names = set()
    created_sheet_count = 0

    with pd.ExcelWriter(
        buffer,
        engine="openpyxl"
    ) as writer:

        for supplier in suppliers:

            sub = df[
                df["仕入先"] == supplier
            ].copy()

            # ----------------------------------------------------
            # 使用日で並び替え
            # ----------------------------------------------------
            sub["使用日_dt"] = (
                sub["使用日"]
                .apply(parse_mmdd)
            )

            sub = sub.sort_values(
                [
                    "使用日_dt",
                    "食品名",
                ],
                na_position="last",
            )

            # ----------------------------------------------------
            # 特養
            # ----------------------------------------------------
            if "特養" in order_type:

                # 加工済み検収簿では
                # 「特養入所者」「特養職員」を使用
                sub = sub.rename(
                    columns={
                        "特養入所者": "入所者",
                        "特養職員": "職員",
                    }
                )

                qty_label = "入所者"
                staff_label = "職員"

                qty_values = pd.to_numeric(
                    sub[qty_label],
                    errors="coerce"
                ).fillna(0)

                staff_values = pd.to_numeric(
                    sub[staff_label],
                    errors="coerce"
                ).fillna(0)

                # ------------------------------------------------
                # 入所者または職員の
                # どちらかに数量があれば残す
                # ------------------------------------------------
                sub = sub.loc[
                    (qty_values != 0)
                    |
                    (staff_values != 0)
                ].copy()

                if sub.empty:
                    continue

                # 再計算
                qty_values = pd.to_numeric(
                    sub[qty_label],
                    errors="coerce"
                ).fillna(0)

                staff_values = pd.to_numeric(
                    sub[staff_label],
                    errors="coerce"
                ).fillna(0)

                # 0は注文書では空欄
                sub[qty_label] = (
                    sub[qty_label]
                    .astype(object)
                )

                sub[staff_label] = (
                    sub[staff_label]
                    .astype(object)
                )

                sub.loc[
                    qty_values == 0,
                    qty_label
                ] = ""

                sub.loc[
                    staff_values == 0,
                    staff_label
                ] = ""

            # ----------------------------------------------------
            # ユーハウス
            # ----------------------------------------------------
            else:

                # 加工済み検収簿の
                # 「ユーハウス」を使用
                sub = sub.rename(
                    columns={
                        "ユーハウス":
                        "ユーハウス入居者"
                    }
                )

                qty_label = (
                    "ユーハウス入居者"
                )

                staff_label = None

                qty_values = pd.to_numeric(
                    sub[qty_label],
                    errors="coerce"
                ).fillna(0)

                # 数量0は除外
                sub = sub.loc[
                    qty_values != 0
                ].copy()

                if sub.empty:
                    continue

            # ----------------------------------------------------
            # 出力列
            # ----------------------------------------------------
            col_order = [
                "使用日",
                "食品名",
                qty_label,
                "単位",
            ]

            if staff_label:
                col_order.append(
                    staff_label
                )

            col_order += [
                "鮮度",
                "品温",
                "異物",
                "包装",
                "期限",
                "備考欄",
                "納品日",
                "検収者",
            ]

            for c in col_order:
                if c not in sub.columns:
                    sub[c] = ""

            sub = sub[
                col_order
            ].copy()

            # ----------------------------------------------------
            # 同じ使用日は最初だけ表示
            # ----------------------------------------------------
            sub["使用日"] = (
                sub["使用日"].mask(
                    sub["使用日"]
                    .duplicated(),
                    ""
                )
            )

            # ----------------------------------------------------
            # Excelシート名
            # ----------------------------------------------------
            sheet_name = re.sub(
                r'[\\/*?:\[\]]',
                '＿',
                str(supplier)
            )

            sheet_name = (
                sheet_name[:31]
                or "仕入先"
            )

            base_sheet_name = (
                sheet_name
            )

            index = 2

            while (
                sheet_name
                in used_sheet_names
            ):
                suffix = f"_{index}"

                sheet_name = (
                    base_sheet_name[
                        :31 - len(suffix)
                    ]
                    + suffix
                )

                index += 1

            used_sheet_names.add(
                sheet_name
            )

            # ----------------------------------------------------
            # Excelへ書き込み
            # ----------------------------------------------------
            sub.to_excel(
                writer,
                sheet_name=sheet_name,
                index=False,
                startrow=5,
            )

            created_sheet_count += 1

            ws = writer.book[
                sheet_name
            ]

            apply_order_style(
                ws,
                is_tokuyou=(
                    "特養"
                    in order_type
                ),
            )

            # ----------------------------------------------------
            # ヘッダー
            # ----------------------------------------------------
            if "特養" in order_type:

                create_header_iwato(
                    ws,
                    supplier
                )

            else:

                create_header_yuhouse(
                    ws,
                    supplier
                )

                ws["C6"] = (
                    "ユーハウス入居者"
                )

        # --------------------------------------------------------
        # 数量ありのシートが0件
        # --------------------------------------------------------
        if created_sheet_count == 0:
            raise ValueError(
                "注文数量が入力されている"
                "データが見つかりませんでした。"
            )

    # ------------------------------------------------------------
    # ファイル名
    # ------------------------------------------------------------
    token = detect_min_usage_date_token(
        df,
        "使用日"
    )

    if "特養" in order_type:
        base_name = (
            "注文書_いわと"
        )
    else:
        base_name = (
            "注文書_ユーハウス"
        )

    if token:
        fname = (
            f"{base_name}_{token}.xlsx"
        )
    else:
        fname = (
            f"{base_name}.xlsx"
        )

    buffer.seek(0)

    return buffer.read(), fname



# ============================================================
# ⓪ 献立チェック ロジック
# ------------------------------------------------------------
# 献立表（下処理室・調理室用）の食材行を読み取り、
# グラム0（A1）／グラム空欄（A2）／完全重複（A3）を検出する。
# 対象外：食材名「＊＊＊」（使用しない仮置き行）と、
# 在庫マスタで B列に「○」（発注しない）がある食品。
# ============================================================

# (ブロック表示名, 名前列, グラム列) ※列は1始まり
MENU_CHECK_BLOCKS = [
    ("ユーハウス（いわと常食）", 1, 6),
    ("特養・常食", 10, 15),
    ("特養・軟菜食きざみ", 19, 24),
    ("特養・ブレンダー食", 28, 33),
]
MENU_CHECK_MEALS = ("朝食", "昼食", "夕食")


def _mc_normalize_name(name):
    """照合用に名前を整える：先頭の「・」除去、空白（全角含む）をすべて除去。"""
    if name is None:
        return ""
    s = str(name).strip().lstrip("・").strip()
    return re.sub(r"[\s　]+", "", s)


def _mc_parse_gram(value):
    """'36   g ' のような文字列から数値を取り出す。取れなければ None。"""
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    m = re.search(r"[\d.]+", str(value))
    return float(m.group(0)) if m else None


def load_no_order_names(uploaded_file):
    """在庫マスタ（アップロード）から B列「○」（発注しない）の食品名集合と統計を返す。

    列構成が想定と違う場合は、黙って続行せず例外にして止める。
    """
    wb = openpyxl.load_workbook(
        io.BytesIO(uploaded_file.getbuffer()), data_only=True
    )
    ws = wb.active
    header = [ws.cell(2, c).value for c in (2, 3, 6)]
    if (
        header[0] != "在"
        or header[1] != "発注種類"
        or not (header[2] and "食品名" in str(header[2]))
    ):
        raise ValueError(
            "在庫マスタの列構成が想定と違います"
            "（2行目の見出しが B:在 / C:発注種類 / F:食品名 ではありません）。"
            "マスタの形式が変更された可能性があるため、処理を止めました。"
        )
    names = set()
    total = 0
    for row in ws.iter_rows(min_row=4, values_only=True):
        if row[0] is None:
            continue
        total += 1
        if row[1] == "○" and row[5]:
            names.add(_mc_normalize_name(row[5]))
    if not names:
        raise ValueError(
            "在庫マスタの B列「○」（発注しない）が0件でした。"
            "マスタの形式が変更された可能性があるため、処理を止めました。"
        )
    stats = {
        "ファイル名": uploaded_file.name,
        "マスタ件数": total,
        "発注しない（○）件数": len(names),
    }
    return names, stats


def extract_menu_ingredients(uploaded_file):
    """献立表（アップロード）から食材行を取り出す。"""
    wb = openpyxl.load_workbook(
        io.BytesIO(uploaded_file.getbuffer()), data_only=True
    )
    ws = wb.active
    items = []
    day = None
    meal = None
    dish_by_block = {}
    for r in range(1, ws.max_row + 1):
        a = ws.cell(r, 1).value
        if isinstance(a, str) and a.startswith("年月日"):
            day = a.replace("年月日：", "").strip()
        if isinstance(a, str) and a.strip() in MENU_CHECK_MEALS:
            meal = a.strip()
        for block_name, name_col, gram_col in MENU_CHECK_BLOCKS:
            v = ws.cell(r, name_col).value
            if not isinstance(v, str):
                continue
            if v.strip().startswith(("■", "◆")):
                dish_by_block[block_name] = v.strip()
                continue
            if v.startswith("　") and v.strip():
                name = v.strip()
                if name.startswith("〃"):
                    continue
                gram_raw = ws.cell(r, gram_col).value
                items.append(
                    {
                        "日付": day,
                        "食事区分": meal,
                        "ブロック": block_name,
                        "料理名": dish_by_block.get(block_name, ""),
                        "食材名": name,
                        "グラム_raw": gram_raw,
                        "グラム": _mc_parse_gram(gram_raw),
                        "行": r,
                    }
                )
    return items


def run_menu_check(items, no_order_names):
    """チェックを実行し、(赤リスト, 対象外内訳) を返す。"""
    from collections import Counter

    reds = []
    excluded = Counter()
    seen = {}
    for it in items:
        if it["食材名"] == "＊＊＊":
            excluded["＊＊＊（使用しない仮置き）"] += 1
            continue
        if _mc_normalize_name(it["食材名"]) in no_order_names:
            excluded["発注しない（在庫マスタ B列○）"] += 1
            continue
        base = {
            "日付": it["日付"],
            "食事区分": it["食事区分"],
            "ブロック": it["ブロック"],
            "料理名": it["料理名"],
            "食材名": it["食材名"],
            "グラム表示": it["グラム_raw"],
            "行": it["行"],
        }
        if it["グラム"] is None:
            reds.append(
                {
                    **base,
                    "種別": "A2",
                    "内容": "グラムが空欄（または数値でない）ため発注書に上がりません",
                }
            )
        elif it["グラム"] == 0:
            reds.append(
                {
                    **base,
                    "種別": "A1",
                    "内容": "グラムが0のため発注書に上がりません",
                }
            )
        key = (
            it["日付"],
            it["ブロック"],
            it["食事区分"],
            it["料理名"],
            it["食材名"],
            it["グラム"],
        )
        if key in seen:
            reds.append(
                {
                    **base,
                    "種別": "A3",
                    "内容": f"まったく同じ行が {seen[key]} 行目にもあります（二重発注の恐れ）",
                }
            )
        else:
            seen[key] = it["行"]
    return reds, excluded


def build_menu_check_workbook(reds, excluded, items_count, days, master_stats):
    """チェック結果Excel（サマリ＋要修正シート）を bytes で返す。"""
    from openpyxl.styles import PatternFill

    wb = openpyxl.Workbook()
    red_fill = PatternFill("solid", fgColor="FCE4E4")
    head_fill = PatternFill("solid", fgColor="4472C4")
    head_font = Font(color="FFFFFF", bold=True)

    ws = wb.active
    ws.title = "サマリ"
    summary = [
        ["献立チェック結果"],
        [],
        [
            "対象期間",
            f"{days[0]} 〜 {days[-1]}（{len(days)}日分）" if days else "不明",
        ],
        ["読み取った食材行数", items_count],
        ["要修正（赤）の件数", len(reds)],
        ["使用した在庫マスタ", master_stats["ファイル名"]],
        ["マスタ件数", master_stats["マスタ件数"]],
        ["発注しない（○）件数", master_stats["発注しない（○）件数"]],
        [],
        ["種別ごとの件数"],
    ]
    from collections import Counter

    for kind, cnt in sorted(Counter(r["種別"] for r in reds).items()):
        summary.append([kind, cnt])
    summary.append([])
    summary.append(["チェック対象外の内訳"])
    for k, cnt in excluded.items():
        summary.append([k, cnt])
    for row in summary:
        ws.append(row)
    ws["A1"].font = Font(bold=True, size=14)
    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 60

    ws2 = wb.create_sheet("要修正（赤）")
    headers = [
        "日付",
        "食事区分",
        "ブロック",
        "料理名",
        "食材名",
        "グラム",
        "種別",
        "内容",
        "献立表の行",
        "対応済",
        "対応メモ",
    ]
    ws2.append(headers)
    for c in range(1, len(headers) + 1):
        cell = ws2.cell(1, c)
        cell.fill = head_fill
        cell.font = head_font
    for r in reds:
        ws2.append(
            [
                r["日付"],
                r["食事区分"],
                r["ブロック"],
                r["料理名"],
                r["食材名"],
                r["グラム表示"],
                r["種別"],
                r["内容"],
                r["行"],
                "",
                "",
            ]
        )
        for c in range(1, len(headers) + 1):
            ws2.cell(ws2.max_row, c).fill = red_fill
    widths = [20, 10, 22, 26, 30, 10, 8, 52, 10, 8, 20]
    for i, w in enumerate(widths, 1):
        ws2.column_dimensions[get_column_letter(i)].width = w
    for row in ws2.iter_rows(min_row=2):
        row[7].alignment = Alignment(wrap_text=True, vertical="top")

    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


# ============================================================
# ⑥ 仕入れカード作成 ロジック
# 検収簿_加工済 → 仕入先×使用日×食品名×単位で集計し、
# A4横・1ページ4枚（2列×2段）のカードExcelを作る。
# 余白は上下左右とも0.5cmで対称、印刷可能領域（28.7×20.0cm）を
# 2列×2段に分割し、上下カード間にはスペーサー行（0.6cm）を入れる。
# 折って切る位置がスペーサー中央を通るため下段の日付が切れない。
# ============================================================
CARD_REQUIRED_COLS = ["使用日", "仕入先", "食品名", "単位",
                      "特養入所者", "特養職員", "ユーハウス"]


def _card_clean_text(value):
    """前後の空白・タブ・改行（全角スペース含む）を取り除く"""
    if value is None:
        return ""
    if isinstance(value, float) and pd.isna(value):
        return ""
    text = str(value)
    return re.sub(r"^[\s　]+|[\s　]+$", "", text)


def _card_date_text(value):
    """使用日を表示用文字列にする（日付型なら m/d、それ以外は文字列そのまま）"""
    if isinstance(value, (datetime, pd.Timestamp)):
        return f"{value.month}/{value.day}"
    return _card_clean_text(value)


def _card_load_df(source):
    df = pd.read_excel(source, sheet_name=0)
    df.columns = [_card_clean_text(c) for c in df.columns]
    return df


def list_card_dates(source):
    """検収簿_加工済に含まれる使用日を、出現順に返す"""
    df = _card_load_df(source)
    if "使用日" not in df.columns:
        return []
    dates = []
    for v in df["使用日"]:
        t = _card_date_text(v)
        if t and t not in dates:
            dates.append(t)
    return dates


def aggregate_card_rows(source, date_filter=None):
    """検収簿_加工済 → カード用に集計する。

    戻り値: (cards, stats)
      cards: [{"仕入先","使用日","食品名","数量","単位"}, ...]
             並びは 仕入先（出現順）→ 使用日（出現順）→ 品目（出現順）
      stats: 件数などの集計情報
    """
    df = _card_load_df(source)
    missing = [c for c in CARD_REQUIRED_COLS if c not in df.columns]
    if missing:
        raise ValueError(
            "検収簿_加工済に必要な列が見つかりません: " + "、".join(missing)
        )

    has_total = "総合計" in df.columns
    stats = {
        "total_rows": int(len(df)),
        "skipped_blank_name": 0,
        "skipped_zero_qty": 0,
        "skipped_date": 0,
        "total_mismatch_rows": 0,
    }

    vendor_order = {}
    date_order = {}
    cards_map = {}
    card_first_idx = {}

    for idx, row in df.iterrows():
        food = _card_clean_text(row["食品名"])
        if not food:
            stats["skipped_blank_name"] += 1
            continue
        date_text = _card_date_text(row["使用日"])
        if date_filter is not None and date_text not in date_filter:
            stats["skipped_date"] += 1
            continue

        def _num(col):
            v = pd.to_numeric(pd.Series([row[col]]), errors="coerce").iloc[0]
            return 0.0 if pd.isna(v) else float(v)

        qty = _num("特養入所者") + _num("特養職員") + _num("ユーハウス")
        if has_total:
            total = _num("総合計")
            if abs(qty - total) > 0.001:
                stats["total_mismatch_rows"] += 1
        if qty == 0:
            stats["skipped_zero_qty"] += 1
            continue

        vendor = _card_clean_text(row["仕入先"])
        unit = _card_clean_text(row["単位"])
        if vendor not in vendor_order:
            vendor_order[vendor] = len(vendor_order)
        dkey = (vendor, date_text)
        if dkey not in date_order:
            date_order[dkey] = len(date_order)
        key = (vendor, date_text, food, unit)
        if key not in cards_map:
            cards_map[key] = {
                "仕入先": vendor,
                "使用日": date_text,
                "食品名": food,
                "数量": 0.0,
                "単位": unit,
            }
            card_first_idx[key] = idx
        cards_map[key]["数量"] += qty

    def _sort_key(item):
        key, _card = item
        vendor, date_text, _food, _unit = key
        return (
            vendor_order[vendor],
            date_order[(vendor, date_text)],
            card_first_idx[key],
        )

    cards = []
    for _key, card in sorted(cards_map.items(), key=_sort_key):
        q = card["数量"]
        card["数量"] = int(q) if float(q).is_integer() else round(q, 3)
        cards.append(card)

    stats["card_count"] = len(cards)
    stats["vendor_count"] = len(vendor_order)
    return cards, stats


# ---------------- レイアウト定数（cm / pt） ----------------
_CARD_COL_CM = [4.28, 4.28, 3.05, 2.44]   # カード内4列（合計14.05cm）
_CARD_SPACER_CM = 0.60
_CARD_ROW_HEIGHTS_PT = [40, 6, 128, 8, 36, 8, 8, 8, 8]  # 9行 + 残り1行で 1枚分
_CARD_BASE_HEIGHT_PT = 10.0 / 2.54 * 72   # 縮小前のカード高さ 283.4646pt
_CARD_HEIGHT_PT = 9.7 / 2.54 * 72         # 実際のカード高さ（上下間スペーサー分を差引）
_CARD_SPACER_ROW_PT = round(0.6 / 2.54 * 72, 2)  # 上下カード間のスペーサー行 17.01pt
_PAGE_ROWS = 21                            # カード10行＋スペーサー1行＋カード10行


def _cm_to_col_width(cm):
    """cm → Excel列幅単位（標準フォント Calibri 11 前提）"""
    px = cm / 2.54 * 96
    return round((px - 5) / 7, 2)


def _card_food_font_size(food):
    L = len(food)
    if L >= 16:
        return 18
    if L >= 12:
        return 23
    return 30


def _draw_one_card(ws, top_row, left_col, card, thick):
    """1枚のカードを描く（枠は10行×4列の外周のみ太線）"""
    for r in range(top_row, top_row + 10):
        for c in range(left_col, left_col + 4):
            cell = ws.cell(row=r, column=c)
            cell.border = Border(
                left=thick if c == left_col else None,
                right=thick if c == left_col + 3 else None,
                top=thick if r == top_row else None,
                bottom=thick if r == top_row + 9 else None,
            )

    # 左上：仕入先
    cell = ws.cell(row=top_row, column=left_col)
    cell.value = card["仕入先"]
    cell.font = Font(name="Meiryo UI", size=10)
    cell.alignment = Alignment(horizontal="left", vertical="top")

    # 右上：使用日
    cell = ws.cell(row=top_row, column=left_col + 3)
    cell.value = card["使用日"]
    cell.font = Font(name="Meiryo UI", size=18, bold=True)
    cell.alignment = Alignment(horizontal="right", vertical="top")

    # 食品名（4列結合）
    ws.merge_cells(
        start_row=top_row + 2, start_column=left_col,
        end_row=top_row + 2, end_column=left_col + 3,
    )
    cell = ws.cell(row=top_row + 2, column=left_col)
    cell.value = card["食品名"]
    cell.font = Font(
        name="Meiryo UI", size=_card_food_font_size(card["食品名"]), bold=True
    )
    cell.alignment = Alignment(
        horizontal="left", vertical="center", wrap_text=True
    )

    # 数量
    cell = ws.cell(row=top_row + 4, column=left_col + 2)
    cell.value = card["数量"]
    cell.font = Font(name="Meiryo UI", size=25, bold=True)
    cell.alignment = Alignment(horizontal="right", vertical="center")

    # 単位
    cell = ws.cell(row=top_row + 4, column=left_col + 3)
    cell.value = card["単位"]
    cell.font = Font(name="Meiryo UI", size=18, bold=True)
    cell.alignment = Alignment(horizontal="left", vertical="center")


def build_cards_workbook(cards):
    """カード一覧 → 仕入れカードExcel（bytes）"""
    from openpyxl.worksheet.pagebreak import Break
    from openpyxl.worksheet.properties import PageSetupProperties

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "仕入れカード"

    # 列幅（A4横の印刷可能幅 28.7cm に一致させる）
    widths_cm = _CARD_COL_CM + [_CARD_SPACER_CM] + _CARD_COL_CM
    for i, cm in enumerate(widths_cm, start=1):
        ws.column_dimensions[get_column_letter(i)].width = _cm_to_col_width(cm)

    pages = (len(cards) + 3) // 4
    total_rows = pages * _PAGE_ROWS

    # 行高（1ページ合計 = 20.0cm になるよう設定）
    scale = _CARD_HEIGHT_PT / _CARD_BASE_HEIGHT_PT
    heights = [round(h * scale, 2) for h in _CARD_ROW_HEIGHTS_PT]
    heights.append(round(_CARD_HEIGHT_PT - sum(heights), 2))
    page_heights = heights + [_CARD_SPACER_ROW_PT] + heights
    for page in range(pages):
        for j, h in enumerate(page_heights):
            ws.row_dimensions[page * _PAGE_ROWS + j + 1].height = h

    thick = Side(style="medium")
    for idx, card in enumerate(cards):
        page_no = idx // 4
        pos = idx % 4
        top_row = page_no * _PAGE_ROWS + (12 if pos >= 2 else 1)
        left_col = 6 if pos % 2 == 1 else 1
        _draw_one_card(ws, top_row, left_col, card, thick)

    # 改ページ（1ページ＝21行ごと）
    for p in range(1, pages):
        ws.row_breaks.append(Break(id=p * _PAGE_ROWS))

    # 印刷設定：A4横、余白は上下左右とも0.5cmで対称、横は1ページ幅に合わせる
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    margin_in = 0.5 / 2.54
    ws.page_margins = PageMargins(
        left=margin_in, right=margin_in, top=margin_in, bottom=margin_in,
        header=0, footer=0,
    )
    ws.print_area = f"A1:I{total_rows}"

    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


# ------------------------------------------------------------
# ------------------------------------------------------------
# 🖥️ UI構築
# ------------------------------------------------------------

# ============================================================
# 🌸 アプリ上部
# ============================================================
st.html(
    '<div class="app-hero">'
    '<div class="app-hero-title">🌷 発注・検収サポートシステム</div>'
    '<div class="app-hero-sub">毎日の発注・検収業務を、やさしく・かんたんに。</div>'
    '<span class="cute-pill pill-pink">🧾 検収簿を整形</span>'
    '<span class="cute-pill pill-green">📦 発注書を自動作成</span>'
    '<span class="cute-pill pill-blue">✨ 業務をかんたんに</span>'
    '</div>',
)

# ============================================================
# サイドバー
# ============================================================
with st.sidebar:
    st.html(
        '<div class="sidebar-section-title">🌸 ユーザーメニュー</div>'
    )

    st.html(
        '<div class="menu-group-title-blue">▼ 📘 検収・注文関連</div>'
    )

    page = st.radio(
        "画面を選択してください",
        [
            "⓪ 献立チェック",
            "① 検収簿整形",
            "② 業者別仕訳表",
            "③ 注文書作成",
            "④ 丸八発注書作成",
            "⑤ 北部市場発注書作成",
            "⑥ 仕入れカード作成",
        ],
        key="main_page_selector",
        label_visibility="collapsed",
    )

    st.html(
        '<div class="menu-group-title-orange">'
        '📦 ④・⑤は業者別発注書メニューです'
        '</div>'
    )



# ============================================================
# ⓪ 献立チェック
# ============================================================
if page == "⓪ 献立チェック":

    st.html(
        '<div class="main-feature-card">'
        '<div class="main-feature-title">🔍 ⓪ 献立チェック</div>'
        '<div class="main-feature-sub">'
        '献立を印刷する前に、グラム0や重複で<br>'
        '発注書に上がらない食材がないか確認します。'
        '</div>'
        '<div class="main-note-box">'
        '📌 献立表と在庫マスタ一覧表の<br>'
        '<b>2つのファイル</b> を選択してください。'
        '</div>'
        '</div>',
    )

    mc_col1, mc_col2 = st.columns(2)

    with mc_col1:
        st.html(
            '<div class="mini-card">'
            '<div class="mini-card-title">📄 献立表</div>'
            '<div class="mini-card-sub">下処理室・調理室用の献立表（1週間〜1か月分）</div>'
            '</div>',
        )
        menu_file = st.file_uploader(
            "献立表（xlsx）",
            type=["xlsx"],
            key="menu_check_menu",
        )

    with mc_col2:
        st.html(
            '<div class="mini-card">'
            '<div class="mini-card-title">📦 在庫マスタ一覧表</div>'
            '<div class="mini-card-sub">B列○（発注しない）の判定に使います</div>'
            '</div>',
        )
        master_file = st.file_uploader(
            "在庫マスタ一覧表（xlsx）",
            type=["xlsx"],
            key="menu_check_master",
        )

    if st.button(
        "🔍 献立チェックを実行する",
        key="btn_menu_check",
        use_container_width=True,
    ):
        if not (menu_file and master_file):
            st.warning("⚠ 献立表と在庫マスタ一覧表の両方を選択してください。")
        else:
            try:
                with st.spinner("献立をチェックしています…"):
                    no_order_names, master_stats = load_no_order_names(master_file)
                    menu_items = extract_menu_ingredients(menu_file)
                    mc_reds, mc_excluded = run_menu_check(menu_items, no_order_names)
                    mc_days = sorted(
                        {it["日付"] for it in menu_items if it["日付"]}
                    )
                    st.session_state["menu_check_data"] = build_menu_check_workbook(
                        mc_reds,
                        mc_excluded,
                        len(menu_items),
                        mc_days,
                        master_stats,
                    )
                    st.session_state["menu_check_fname"] = "献立チェック結果.xlsx"
                    st.session_state["menu_check_reds"] = mc_reds
                    st.session_state["menu_check_summary"] = {
                        "食材行数": len(menu_items),
                        "日数": len(mc_days),
                        "対象外": dict(mc_excluded),
                        "マスタ○件数": master_stats["発注しない（○）件数"],
                    }
                st.success("🌸 献立チェックが完了しました！")
            except Exception as e:
                st.session_state.pop("menu_check_data", None)
                st.session_state.pop("menu_check_reds", None)
                st.session_state.pop("menu_check_summary", None)
                st.error("献立チェック中にエラーが発生しました。")
                st.exception(e)

    if (
        "menu_check_data" in st.session_state
        and "menu_check_reds" in st.session_state
    ):
        mc_reds = st.session_state["menu_check_reds"]
        mc_summary = st.session_state.get("menu_check_summary", {})

        st.markdown("### 📋 チェック結果")
        st.write(
            f"対象：{mc_summary.get('日数', 0)}日分／食材 {mc_summary.get('食材行数', 0)} 行"
            f"（発注しない○は {mc_summary.get('マスタ○件数', 0)} 件で判定）"
        )
        for k, cnt in mc_summary.get("対象外", {}).items():
            st.write(f"・対象外：{k} {cnt} 行")

        if mc_reds:
            st.warning(f"⚠ 要修正（赤）が {len(mc_reds)} 件あります。献立を直してから印刷してください。")
            st.dataframe(
                pd.DataFrame(mc_reds)[
                    ["日付", "食事区分", "ブロック", "料理名", "食材名", "グラム表示", "種別", "内容"]
                ].rename(columns={"グラム表示": "グラム"}),
                use_container_width=True,
                hide_index=True,
            )
        else:
            st.success("🌸 要修正（赤）はありませんでした。")

        st.download_button(
            label="📥 献立チェック結果をダウンロード",
            data=st.session_state["menu_check_data"],
            file_name=st.session_state["menu_check_fname"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_menu_check",
            use_container_width=True,
        )

# ============================================================
# ① 検収簿整形
# ============================================================
elif page == "① 検収簿整形":

    st.html(
        '<div class="main-feature-card">'
        '<div class="main-feature-title">🌷 ① 検収簿を整える</div>'
        '<div class="main-feature-sub">'
        '献ダテマンから出力した検収記録簿を、<br>'
        '発注・検収アプリで使いやすい形に整えます。'
        '</div>'
        '<div class="main-note-box">'
        '📌 献ダテマンから出力したファイルを<br>'
        '<b>「検収記録簿_原本.xlsx」</b> の名前で保存して下さい。'
        '</div>'
        '</div>',
    )

    ins_file = st.file_uploader(
        "📄 検収簿（原本 Excel）をアップロード",
        type=["xlsx"],
        key="ins",
    )

    if ins_file is not None:
        if st.button(
            "📘 検収簿を整形する",
            key="btn_ins",
            use_container_width=True,
        ):
            try:
                (
                    st.session_state["ins_data"],
                    st.session_state["ins_fname"],
                ) = format_inspection_workbook(ins_file)
                st.success("🌸 検収簿の整形が完了しました！")
            except Exception as e:
                st.error("検収簿の整形中にエラーが発生しました。")
                st.exception(e)

        if (
            "ins_data" in st.session_state
            and "ins_fname" in st.session_state
        ):
            st.download_button(
                label="📥 検収簿（加工済）をダウンロード",
                data=st.session_state["ins_data"],
                file_name=st.session_state["ins_fname"],
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="download_inspection",
                use_container_width=True,
            )


# ============================================================
# ② 業者別仕訳表
# ============================================================
elif page == "② 業者別仕訳表":

    st.html(
        '<div class="main-feature-card">'
        '<div class="main-feature-title">📊 ② 業者別仕訳表を作成</div>'
        '<div class="main-feature-sub">'
        '加工済み検収簿を、仕入先ごとのシートに分けて自動作成します。'
        '</div>'
        '<div class="main-note-blue">'
        '💡 「① 検収簿整形」で作成した <b>検収簿_加工済</b> を使用してください。<br>'
        'A3縦・文字16・行高26・罫線付きで出力します。'
        '</div>'
        '</div>',
    )

    vendor_file = st.file_uploader(
        "📄 検収簿（加工済 Excel）をアップロード",
        type=["xlsx"],
        key="vendor_journal_src",
    )

    if vendor_file is not None:
        try:
            if st.button(
                "📊 業者別仕訳表を作成する",
                key="btn_vendor_journal",
                use_container_width=True,
            ):
                (
                    st.session_state["vendor_journal_data"],
                    st.session_state["vendor_journal_fname"],
                ) = create_vendor_journal_workbook(vendor_file)
                st.success("🌸 業者別仕訳表の作成が完了しました！")

            if (
                "vendor_journal_data" in st.session_state
                and "vendor_journal_fname" in st.session_state
            ):
                st.download_button(
                    label="📥 業者別仕訳表をダウンロード",
                    data=st.session_state["vendor_journal_data"],
                    file_name=st.session_state["vendor_journal_fname"],
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="download_vendor_journal",
                    use_container_width=True,
                )
        except Exception as e:
            st.error("業者別仕訳表の作成中にエラーが発生しました。")
            st.exception(e)


# ============================================================
# ③ 注文書作成
# ============================================================
elif page == "③ 注文書作成":

    st.html(
        '<div class="main-feature-card">'
        '<div class="main-feature-title">📗 ③ 注文書を作成</div>'
        '<div class="main-feature-sub">'
        '特養またはユーハウスを選択して、<br>'
        '仕入先別の注文書を自動作成します。'
        '</div>'
        '<div class="main-note-green">'
        '🌿 特養では <b>「入所者」と「職員」</b> を同じ注文書にまとめて出力します。'
        '</div>'
        '</div>',
    )

    order_type = st.radio(
        "作成する注文書の種類を選んでください",
        (
            "特養（介護老人福祉施設いわと）",
            "ユーハウスいわと",
        ),
        horizontal=True,
        key="order_type",
    )

    order_file = st.file_uploader(
        "📄 注文書のもとになる検収簿 Excel",
        type=["xlsx"],
        key="order_src",
    )

    st.info(
        "① 検収簿整形で作成した「検収簿_加工済」ファイルを使用してください。"
    )

    if order_file is not None:
        current_file_id = (
            order_file.name,
            order_file.size,
            order_type,
        )

        previous_file_id = st.session_state.get("order_current_file_id")

        if previous_file_id != current_file_id:
            st.session_state["order_current_file_id"] = current_file_id
            st.session_state.pop("order_data", None)
            st.session_state.pop("order_fname", None)

        if st.button(
            "📗 注文書を作成する",
            key="btn_order",
            use_container_width=True,
        ):
            try:
                order_bytes, order_fname = create_order_workbook(
                    order_file,
                    order_type,
                )
                st.session_state["order_data"] = order_bytes
                st.session_state["order_fname"] = order_fname
                st.success(f"🌸 {order_type} の注文書を作成しました！")
            except Exception as e:
                st.session_state.pop("order_data", None)
                st.session_state.pop("order_fname", None)
                st.error("注文書作成中にエラーが発生しました。")
                st.exception(e)

        if (
            "order_data" in st.session_state
            and "order_fname" in st.session_state
        ):
            st.download_button(
                label="📥 注文書ファイルをダウンロード",
                data=st.session_state["order_data"],
                file_name=st.session_state["order_fname"],
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="download_order",
                use_container_width=True,
            )


# ============================================================
# ④ 丸八発注書作成
# ============================================================
elif page == "④ 丸八発注書作成":

    st.html(
        '<div class="main-feature-card">'
        '<div class="main-feature-title">📦 ④ 丸八発注書を作成</div>'
        '<div class="main-feature-sub">'
        '検収簿_加工済・丸八発注書テンプレート・<br>'
        '丸八コード一覧から、特養用とユーハウス用の発注書を同時に作成します。'
        '</div>'
        '<div class="main-note-box">📌 下の3つのファイルをすべて選択してください。</div>'
        '</div>',
    )

    mcol1, mcol2, mcol3 = st.columns(3)

    with mcol1:
        st.html(
            '<div class="mini-card">'
            '<div class="mini-card-title">📄 検収簿_加工済</div>'
            '<div class="mini-card-sub">丸八ヒロタのデータを含む加工済ファイル</div>'
            '</div>',
        )
        kenshu_file = st.file_uploader(
            "検収簿_加工済（xlsx）",
            type=["xlsx"],
            key="kenshu_maruhachi",
        )

    with mcol2:
        st.html(
            '<div class="mini-card">'
            '<div class="mini-card-title">🧾 丸八発注書テンプレ</div>'
            '<div class="mini-card-sub">丸八ヒロタ専用の発注書テンプレート</div>'
            '</div>',
        )
        template_file = st.file_uploader(
            "丸八発注書テンプレ（xlsm）",
            type=["xlsm"],
            key="tpl_maruhachi",
        )

    with mcol3:
        st.html(
            '<div class="mini-card">'
            '<div class="mini-card-title">🏷️ 丸八コード一覧</div>'
            '<div class="mini-card-sub">タグシート付きのコード対応表</div>'
            '</div>',
        )
        tag_file = st.file_uploader(
            "丸八コード一覧（xlsm）",
            type=["xlsm"],
            key="tag_maruhachi",
        )

    btn = st.button(
        "📦 丸八発注書を作成",
        key="btn_maruhachi",
        use_container_width=True,
    )

    if btn:
        if generate_maruhachi_order_forms_both_facilities is None:
            st.error("丸八発注書機能を読み込めませんでした。")
            if MARUHACHI_IMPORT_ERROR is not None:
                st.exception(MARUHACHI_IMPORT_ERROR)

        elif not (kenshu_file and template_file and tag_file):
            st.warning("⚠ 3つのファイルをすべて選択してください。")

        else:
            try:
                with st.spinner("丸八発注書を作成しています…"):
                    with tempfile.TemporaryDirectory() as td:
                        td = Path(td)

                        k_path = td / "kenshu.xlsx"
                        t_path = td / "template.xlsm"
                        m_path = td / "tag.xlsm"

                        k_path.write_bytes(kenshu_file.getbuffer())
                        t_path.write_bytes(template_file.getbuffer())
                        m_path.write_bytes(tag_file.getbuffer())

                        out_dir = td / "out"

                        tokuyou_xlsm, yuhouse_xlsm = (
                            generate_maruhachi_order_forms_both_facilities(
                                kenshu_xlsx_path=k_path,
                                template_xlsm_path=t_path,
                                tag_xlsm_path=m_path,
                                out_dir=out_dir,
                                out_prefix="丸八発注書",
                            )
                        )

                        st.session_state["maruhachi_tokuyou_data"] = tokuyou_xlsm.read_bytes()
                        st.session_state["maruhachi_tokuyou_fname"] = tokuyou_xlsm.name
                        st.session_state["maruhachi_yuhouse_data"] = yuhouse_xlsm.read_bytes()
                        st.session_state["maruhachi_yuhouse_fname"] = yuhouse_xlsm.name

                st.success("🌸 丸八発注書を作成しました！")

            except Exception as e:
                st.session_state.pop("maruhachi_tokuyou_data", None)
                st.session_state.pop("maruhachi_tokuyou_fname", None)
                st.session_state.pop("maruhachi_yuhouse_data", None)
                st.session_state.pop("maruhachi_yuhouse_fname", None)
                st.error("丸八発注書の作成中にエラーが発生しました。")
                st.exception(e)

    if (
        "maruhachi_tokuyou_data" in st.session_state
        and "maruhachi_yuhouse_data" in st.session_state
    ):
        st.markdown("### 📥 作成済みファイル")

        dcol1, dcol2 = st.columns(2)

        with dcol1:
            st.download_button(
                "📥 特養：丸八発注書",
                data=st.session_state["maruhachi_tokuyou_data"],
                file_name=st.session_state["maruhachi_tokuyou_fname"],
                mime="application/vnd.ms-excel.sheet.macroEnabled.12",
                key="download_maruhachi_tokuyou",
                use_container_width=True,
            )

        with dcol2:
            st.download_button(
                "📥 ユーハウス：丸八発注書",
                data=st.session_state["maruhachi_yuhouse_data"],
                file_name=st.session_state["maruhachi_yuhouse_fname"],
                mime="application/vnd.ms-excel.sheet.macroEnabled.12",
                key="download_maruhachi_yuhouse",
                use_container_width=True,
            )


# ============================================================
# ⑤ 北部市場発注書作成
# ============================================================
elif page == "⑤ 北部市場発注書作成":

    st.html(
        '<div class="main-feature-card">'
        '<div class="main-feature-title">🥕 ⑤ 北部市場発注書を作成</div>'
        '<div class="main-feature-sub">'
        '検収簿_加工済と北部市場発注書テンプレートを使って、<br>'
        '特養用・ユーハウス用の発注書を同時に作成します。'
        '</div>'
        '<div class="main-note-green">'
        '🌿 2つのファイルを選択して、「北部市場発注書を作成」を押してください。'
        '</div>'
        '</div>',
    )

    hcol1, hcol2 = st.columns(2)

    with hcol1:
        st.html(
            '<div class="mini-card">'
            '<div class="mini-card-title">📄 検収簿_加工済</div>'
            '<div class="mini-card-sub">北部市場販売のデータを含む加工済ファイル</div>'
            '</div>',
        )
        hokubu_kenshu = st.file_uploader(
            "検収簿_加工済（xlsx）",
            type=["xlsx"],
            key="hokubu_kenshu",
        )

    with hcol2:
        st.html(
            '<div class="mini-card">'
            '<div class="mini-card-title">🧾 北部市場テンプレート</div>'
            '<div class="mini-card-sub">北部市場専用の発注書テンプレート</div>'
            '</div>',
        )
        hokubu_template = st.file_uploader(
            "北部市場発注書テンプレート（xlsm）",
            type=["xlsm"],
            key="hokubu_tpl",
        )

    btn_hokubu = st.button(
        "🥕 北部市場発注書を作成",
        key="btn_hokubu",
        use_container_width=True,
    )

    if btn_hokubu:
        if generate_hokubu_order_forms_both_facilities is None:
            st.error("北部市場発注書機能を読み込めませんでした。")
            if HOKUBU_IMPORT_ERROR is not None:
                st.exception(HOKUBU_IMPORT_ERROR)

        elif not (hokubu_kenshu and hokubu_template):
            st.warning(
                "⚠ 検収簿_加工済 と 北部市場テンプレートを両方選択してください。"
            )

        else:
            try:
                with st.spinner("北部市場発注書を作成しています…"):
                    with tempfile.TemporaryDirectory() as td:
                        td = Path(td)

                        k_path = td / "kenshu.xlsx"
                        t_path = td / "template.xlsm"

                        k_path.write_bytes(hokubu_kenshu.getbuffer())
                        t_path.write_bytes(hokubu_template.getbuffer())

                        out_dir = td / "out"

                        tokuyou_xlsm, yuhouse_xlsm = (
                            generate_hokubu_order_forms_both_facilities(
                                kenshu_xlsx_path=k_path,
                                template_xlsm_path=t_path,
                                out_dir=out_dir,
                                out_prefix="北部市場発注書",
                            )
                        )

                        st.session_state["hokubu_tokuyou_data"] = tokuyou_xlsm.read_bytes()
                        st.session_state["hokubu_tokuyou_fname"] = tokuyou_xlsm.name
                        st.session_state["hokubu_yuhouse_data"] = yuhouse_xlsm.read_bytes()
                        st.session_state["hokubu_yuhouse_fname"] = yuhouse_xlsm.name

                st.success("🌸 北部市場発注書を作成しました！")

            except Exception as e:
                st.session_state.pop("hokubu_tokuyou_data", None)
                st.session_state.pop("hokubu_tokuyou_fname", None)
                st.session_state.pop("hokubu_yuhouse_data", None)
                st.session_state.pop("hokubu_yuhouse_fname", None)
                st.error("北部市場発注書の作成中にエラーが発生しました。")
                st.exception(e)

    if (
        "hokubu_tokuyou_data" in st.session_state
        and "hokubu_yuhouse_data" in st.session_state
    ):
        st.markdown("### 📥 作成済みファイル")

        c1, c2 = st.columns(2)

        with c1:
            st.download_button(
                label="📥 特養：北部市場発注書",
                data=st.session_state["hokubu_tokuyou_data"],
                file_name=st.session_state["hokubu_tokuyou_fname"],
                mime="application/vnd.ms-excel.sheet.macroEnabled.12",
                key="download_hokubu_tokuyou",
                use_container_width=True,
            )

        with c2:
            st.download_button(
                label="📥 ユーハウス：北部市場発注書",
                data=st.session_state["hokubu_yuhouse_data"],
                file_name=st.session_state["hokubu_yuhouse_fname"],
                mime="application/vnd.ms-excel.sheet.macroEnabled.12",
                key="download_hokubu_yuhouse",
                use_container_width=True,
            )

# ============================================================
# ⑥ 仕入れカード作成
# ============================================================
elif page == "⑥ 仕入れカード作成":

    st.html(
        '<div class="main-feature-card">'
        '<div class="main-feature-title">🎴 ⑥ 仕入れカードを作成</div>'
        '<div class="main-feature-sub">'
        '検収簿_加工済から、仕入先ごとの仕入れカード（A4横・1ページ4枚）を作成します。<br>'
        'カードの数量は、特養入所者＋特養職員＋ユーハウスの合計です。'
        '</div>'
        '<div class="main-note-green">'
        '🌿 カードは4枚とも同じ大きさになるよう寸法を統一しています。'
        '印刷した紙を半分・半分に切ると、4枚とも約14.9cm × 10.5cm になります。'
        '</div>'
        '</div>',
    )

    st.html(
        '<div class="mini-card">'
        '<div class="mini-card-title">📄 検収簿_加工済</div>'
        '<div class="mini-card-sub">「① 検収簿整形」で作成したファイルを使用してください</div>'
        '</div>',
    )
    card_kenshu = st.file_uploader(
        "検収簿_加工済（xlsx）",
        type=["xlsx"],
        key="card_kenshu",
    )

    card_dates_selected = None
    if card_kenshu is not None:
        try:
            all_dates = list_card_dates(card_kenshu)
        except Exception:
            all_dates = []
        if all_dates:
            card_dates_selected = st.multiselect(
                "カードを作る使用日（未選択なら全日）",
                all_dates,
                default=all_dates,
                key="card_dates",
            )

    btn_card = st.button(
        "🎴 仕入れカードを作成する",
        key="btn_card",
        use_container_width=True,
    )

    if btn_card:
        if not card_kenshu:
            st.warning("⚠ 検収簿_加工済 を選択してください。")
        else:
            try:
                with st.spinner("仕入れカードを作成しています…"):
                    date_filter = (
                        set(card_dates_selected)
                        if card_dates_selected
                        else None
                    )
                    cards, card_stats = aggregate_card_rows(
                        card_kenshu, date_filter=date_filter
                    )
                    if not cards:
                        st.session_state.pop("card_data", None)
                        st.warning(
                            "⚠ カードになる品目がありませんでした"
                            "（数量0の品目はカードにしません）。"
                        )
                    else:
                        st.session_state["card_data"] = build_cards_workbook(cards)
                        st.session_state["card_fname"] = "仕入れカード.xlsx"
                        st.session_state["card_stats"] = card_stats
                        st.session_state["card_preview"] = cards[:100]
                        st.success(
                            f"🌸 仕入れカードを作成しました！"
                            f"（{card_stats['card_count']}枚 / "
                            f"{(card_stats['card_count'] + 3) // 4}ページ、"
                            f"仕入先 {card_stats['vendor_count']}社）"
                        )
                        if card_stats["skipped_zero_qty"]:
                            st.info(
                                f"数量0のためカードにしなかった品目: "
                                f"{card_stats['skipped_zero_qty']}件"
                            )
                        if card_stats["total_mismatch_rows"]:
                            st.warning(
                                f"「総合計」と3施設の合計が合わない行: "
                                f"{card_stats['total_mismatch_rows']}件"
                                "（カードの数量は3施設の合計で計算しています）"
                            )
            except Exception as e:
                st.session_state.pop("card_data", None)
                st.error("仕入れカードの作成中にエラーが発生しました。")
                st.exception(e)

    if "card_data" in st.session_state:
        st.markdown("### 📥 作成済みファイル")
        st.download_button(
            label="📥 仕入れカードをダウンロード",
            data=st.session_state["card_data"],
            file_name=st.session_state.get("card_fname", "仕入れカード.xlsx"),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_card",
            use_container_width=True,
        )
        if st.session_state.get("card_preview"):
            st.markdown("### 🎴 カード内容（先頭100件まで）")
            st.dataframe(
                pd.DataFrame(st.session_state["card_preview"]),
                use_container_width=True,
            )

