import io
import re
from datetime import datetime

import pandas as pd
import streamlit as st
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.worksheet.page import PageMargins

# ------------------------------------------------------------
# Streamlit 基本設定
# ------------------------------------------------------------
st.set_page_config(page_title="発注・検収サポートシステム", layout="wide")

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
def format_inspection_workbook(uploaded_file):
    df = pd.read_excel(uploaded_file, header=[6, 7])

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

    df_out = df[needed_cols]

    # ---- 出力 ----
    buffer = io.BytesIO()
    df_out.to_excel(buffer, index=False)
    buffer.seek(0)

    token = detect_min_usage_date_token(df_out, "使用日")
    fname = f"検収簿_加工済_{token}.xlsx" if token else "検収簿_加工済.xlsx"

    return buffer.read(), fname

# ------------------------------------------------------------
# 注文書 書式設定（いわと／ユーハウス共通）
# ------------------------------------------------------------
def apply_order_style(ws):
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
        cell.font = Font(name="ＭＳ ゴシック", size=18, bold=True)
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

            # 同じ使用日は2行目以降空欄に
            sub["使用日"] = sub["使用日"].mask(sub["使用日"].duplicated(), "")

            sheet = str(supplier)[:30]
            sub.to_excel(writer, sheet_name=sheet, index=False, startrow=5)

        # 書式 & ヘッダー
        wb = writer.book
        for supplier in suppliers:
            ws = wb[str(supplier)[:30]]
            apply_order_style(ws)

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
# 🖥️ UI構築（かわいい献ダテマン風）
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


col_left, col_right = st.columns([1, 1])


# ------------------------------------------------------------
# ① 検収簿整形
# ------------------------------------------------------------
with col_left:
    st.markdown(
        """
<div class="feature-card">
  <div class="feature-title">① 検収簿を整える</div>
  <div class="feature-sub">
    MultiIndex の検収記録簿を<br>
    A〜H列だけの加工済みファイルに整形します。<br>
    ※ 献ダテマンから出力したファイルを<br>
      「<b>検収簿_原本.xlsx</b>」の名前で保存して下さい。
  </div>
  <hr class="soft"/>
</div>
        """,
        unsafe_allow_html=True,
    )

    ins_file = st.file_uploader("検収簿（原本 Excel）をアップロード", type=["xlsx"], key="ins")

if ins_file:
    if st.button("📘 検収簿を整形する", key="btn_ins"):
        st.session_state["ins_data"], st.session_state["ins_fname"] = \
            format_inspection_workbook(ins_file)
        st.success("検収簿の整形が完了しました！")

    # 整形が完了したらダウンロードボタンを出す
    if "ins_data" in st.session_state:
        st.download_button(
            "📥 検収簿（加工済）をダウンロード",
            st.session_state["ins_data"],
            st.session_state["ins_fname"]
        )



# ------------------------------------------------------------
# ② 注文書（特養 / ユーハウス）選択式
# ------------------------------------------------------------
with col_right:
    st.markdown(
        """
<div class="feature-card">
  <div class="feature-title">② 注文書を作成</div>
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

    # 種別選択
    order_type = st.radio(
        "作成する注文書の種類を選んでください",
        ("特養（介護老人福祉施設いわと）", "ユーハウスいわと"),
        horizontal=True,
        key="order_type",
    )

    # ファイルアップロード
    order_file = st.file_uploader(
        "注文書のもとになる検収簿 Excel をアップロード",
        type=["xlsx"],
        key="order_src",
    )

    st.markdown(
        '<p class="small-note">※ 検収簿整形で加工したもの、または同じ形式の検収簿ファイルを想定しています。</p>',
        unsafe_allow_html=True,
    )

    # 🔥 注文書作成ボタン（正しい位置）
    if order_file:
        try:
            if st.button("📗 注文書を作成する", key="btn_order"):
                st.session_state["order_data"], st.session_state["order_fname"] = \
                    create_order_workbook(order_file, order_type)
                st.success(f"{order_type} の注文書が作成されました！")

            # 作成後にダウンロードボタンを出す
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




