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
# ① 検収簿整形ロジック（ログ付き）
# ------------------------------------------------------------
def format_inspection_workbook(uploaded_file):
    print("\n=== 📘 検収簿 整形処理 開始 =====================")

    # --------------------------------------------------------
    # ① MultiIndex → 読み込み
    # --------------------------------------------------------
    df = pd.read_excel(uploaded_file, header=[6, 7])
    print("✔ MultiIndex ヘッダー読み込み完了")

    # --------------------------------------------------------
    # ② 列名フラット化
    # --------------------------------------------------------
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
    print("✔ 列名フラット化完了")

    # Unnamed の削除
    for i in range(6):
        df.columns = [c.replace(f"Unnamed: {i}_level_0_", "") for c in df.columns]

    print("✔ Unnamed カラム削除完了")

    # --------------------------------------------------------
    # ③ 欠損補完
    # --------------------------------------------------------
    for col in ["納品日", "使用日", "朝昼夕", "仕入先"]:
        if col in df.columns:
            df[col] = df[col].ffill()

    print("✔ 欠損補完完了")

    # --------------------------------------------------------
    # ④ 朝昼夕並び替え用番号
    # --------------------------------------------------------
    order_map = {"朝食": 1, "昼食": 2, "夕食": 3}
    df["食事順"] = df["朝昼夕"].map(order_map)

    print("✔ 朝昼夕 並び順マッピング完了")

    # --------------------------------------------------------
    # ⑤ 並び替え
    # --------------------------------------------------------
    df = df.sort_values(["使用日", "食事順", "食品名"])
    print("✔ 並び替え完了（使用日 → 朝昼夕 → 食品名）")

    # --------------------------------------------------------
    # ⑥ 必要列だけ抽出（A〜K）
    # --------------------------------------------------------
    extract_cols = [
        "納品日",
        "使用日",
        "朝昼夕",
        "仕入先",
        "食品名",
        "換算値",
        "総合計",
        "単位",
        "介護老人福祉施設いわと_入所者",
        "介護老人福祉施設いわと_職員",
        "ケアハウスユー…_入所者",
    ]

    extract_cols = [c for c in extract_cols if c in df.columns]
    df_out = df[extract_cols]

    print("✔ 列抽出完了（A〜K 列）")

    # --------------------------------------------------------
    # ⑦ 出力
    # --------------------------------------------------------
    buffer = io.BytesIO()
    df_out.to_excel(buffer, index=False)
    buffer.seek(0)

    print("🎉 完了：検収簿の整形が正常終了しました")
    print("=========================================\n")

    return buffer.read(), "検収記録簿_加工済.xlsx"


# ------------------------------------------------------------
# ② 注文書（特養・ユーハウス）の共通処理
# ------------------------------------------------------------
def apply_order_style(ws):
    font_body = Font(name="ＭＳ ゴシック", size=18)
    border = Border(
        left=Side("thin"), right=Side("thin"),
        top=Side("thin"), bottom=Side("thin")
    )

    header_row = 6

    # ヘッダー部分
    for cell in ws[header_row]:
        cell.font = Font(name="ＭＳ ゴシック", size=18, bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border

    # データ部
    for row in ws.iter_rows(min_row=header_row + 1):
        for c in row:
            c.font = font_body
            c.border = border
            c.alignment = Alignment(vertical="center")

    # 行高
    for i in range(1, ws.max_row + 1):
        ws.row_dimensions[i].height = 30

    # 列幅
    ws.column_dimensions["A"].width = 15.18
    ws.column_dimensions["B"].width = 60.09
    for col in ["D", "E", "F", "G", "H"]:
        ws.column_dimensions[col].width = 7.73
    for col in ["C", "I", "J", "K", "L", "M"]:
        ws.column_dimensions[col].width = 15.18

    # B列のみ縮小表示
    for row in ws.iter_rows(min_row=7, max_row=ws.max_row, min_col=2, max_col=2):
        for cell in row:
            cell.alignment = Alignment(
                horizontal="left", vertical="center", shrink_to_fit=True
            )

    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_margins = PageMargins(left=0.3, right=0.3, top=0.5, bottom=0.5)
    ws.print_area = f"A1:M{ws.max_row}"


# ------------------------------------------------------------
# 注文書：特養ヘッダー
# ------------------------------------------------------------
def create_header_iwato(ws, supplier):
    ws.merge_cells("A3:B3")
    ws["A3"] = f"{supplier}　御中"
    ws["A3"].font = Font(name="ＭＳ ゴシック", size=28, bold=True)

    ws["B1"] = "注文書（介護老人福祉施設いわと）"
    ws["B1"].font = Font(name="ＭＳ ゴシック", size=26, bold=True)
    ws["B1"].alignment = Alignment(horizontal="center")

    ws["K3"] = "(有) ハートミール"
    ws["K3"].font = Font(name="ＭＳ ゴシック", size=24, bold=True)
    ws["K3"].alignment = Alignment(horizontal="right")


# ------------------------------------------------------------
# 注文書：ユーハウスヘッダー
# ------------------------------------------------------------
def create_header_yuhouse(ws, supplier):
    ws.merge_cells("A3:B3")
    ws["A3"] = f"{supplier}　御中"
    ws["A3"].font = Font(name="ＭＳ ゴシック", size=28, bold=True)

    ws["B1"] = "注文書（ユーハウスいわと）"
    ws["B1"].font = Font(name="ＭＳ ゴシック", size=26, bold=True)
    ws["B1"].alignment = Alignment(horizontal="center")

    ws["K3"] = "(有) ハートミール"
    ws["K3"].font = Font(name="ＭＳ ゴシック", size=24, bold=True)
    ws["K3"].alignment = Alignment(horizontal="right")
# ------------------------------------------------------------
# ③ 注文書作成（特養 / ユーハウスを選択式で統合）
# ------------------------------------------------------------
def create_order_workbook(uploaded_file, order_type):
    df = pd.read_excel(uploaded_file)

    # 欠損補完
    for c in ["使用日", "仕入先", "食品名"]:
        if c in df.columns:
            df[c] = df[c].ffill()

    # 使用日文字化
    df["使用日"] = df["使用日"].astype(str)

    # 数値列
    if order_type == "特養（いわと）":
        qty_col = "入所者"
        extra_cols = ["職員"]
    else:
        qty_col = "ユーハウス入所者"
        extra_cols = []

    df[qty_col] = pd.to_numeric(df[qty_col], errors="coerce").fillna(0)
    for c in extra_cols:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

    # 出力列
    keep_cols = [
        "使用日", "食品名", qty_col, "単位",
        "鮮度", "品温", "異物", "包装", "期限",
        "備考欄", "納品時間", "検収者"
    ]
    for c in keep_cols:
        if c not in df.columns:
            df[c] = ""

    suppliers = df["仕入先"].dropna().unique()

    # ファイル名の接頭辞
    token = detect_min_usage_date_token(df, "使用日")

    if order_type == "特養（いわと）":
        base_name = "注文書_いわと"
    else:
        base_name = "注文書_ユーハウス"

    out_name = f"{base_name}{token}.xlsx" if token else f"{base_name}.xlsx"

    # Excel 出力
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        pd.DataFrame({"dummy": [1]}).to_excel(writer, sheet_name="_dummy", index=False)

        for supplier in suppliers:
            sub = df[df["仕入先"] == supplier].copy()

            # 集計
            group_cols = ["使用日", "食品名", "単位"]
            sum_cols = [qty_col] + extra_cols
            sub = sub.groupby(group_cols, as_index=False)[sum_cols].sum()

            # 列追加
            for c in keep_cols:
                if c not in sub.columns:
                    sub[c] = ""

            # 日付で並べ替え
            sub["使用日_dt"] = sub["使用日"].apply(parse_mmdd)
            sub = sub.sort_values(["使用日_dt", "食品名"], na_position="last")

            # 出力列順に揃える
            sub = sub[keep_cols]
            sub["使用日"] = sub["使用日"].mask(sub["使用日"].duplicated(), "")

            sheet_name = str(supplier)[:30]
            sub.to_excel(writer, sheet_name=sheet_name, index=False, startrow=5)

        # スタイル適用
        wb = writer.book
        for supplier in suppliers:
            sheet = str(supplier)[:30]
            ws = wb[sheet]

            apply_order_style(ws)

            # ヘッダー
            if order_type == "特養（いわと）":
                create_header_iwato(ws, supplier)
            else:
                create_header_yuhouse(ws, supplier)
                ws["C6"].value = "入居者"

            # 「納品時間 → 納品日」
            for cell in ws[6]:
                if cell.value == "納品時間":
                    cell.value = "納品日"

    buffer.seek(0)
    return buffer.read(), out_name


# ------------------------------------------------------------
# 🖥️ UI構築（かわいい献ダテマン風）
# ------------------------------------------------------------
st.markdown(
    """
<div>
  <span class="app-title">発注・検収サポートシステム</span><br/>
  <span class="subtitle-pill sub-orange">毎日の業務をかんたんに</span>
  <span class="subtitle-pill sub-green">発注書を自動作成</span>
  <span class="subtitle-pill sub-blue">検収簿を整形</span>
</div>
<br/>
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

    if ins_file and st.button("📘 検収簿を整形する", key="btn_ins"):
        data, fname = format_inspection_workbook(ins_file)
        st.success("検収簿の整形が完了しました！")
        st.download_button("📥 ダウンロード（検収簿 加工済）", data, fname)


# ------------------------------------------------------------
# ② 注文書（特養 / ユーハウス）選択式
# ------------------------------------------------------------
with col_right:
    st.markdown(
        """
<div class="feature-card">
  <div class="feature-title">② 注文書を作成する</div>
  <div class="feature-sub">
      特養（いわと）・ユーハウスを選択できます。<br>
      1つのファイルからどちらの注文書も自動生成！
  </div>
  <hr class="soft"/>
</div>
""",
        unsafe_allow_html=True,
    )

    # 🟢 選択式
    order_type = st.radio(
        "作成する注文書を選んでください",
        ["特養（いわと）", "ユーハウス"],
        horizontal=True,
        key="ordertype"
    )

    order_file = st.file_uploader(
        "検収簿（整形済み Excel）をアップロード", type=["xlsx"], key="orderfile"
    )

    if order_file and st.button("📗 注文書を作成する", key="btn_order"):
        data, fname = create_order_workbook(order_file, order_type)
        st.success(f"{order_type} の注文書を作成しました！")
        st.download_button("📥 ダウンロード（注文書）", data, fname)

