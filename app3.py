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

    # ★★★ ここをあなたの仕様に合わせて修正 ★★★
    needed_cols = [
        "納品日",
        "使用日",
        "朝昼夕",
        "仕入先",
        "食品名",
        "換算値",
        "総合計",
        "単位",
        "介護老人福祉施設いわと_入所者",  # I列
        "介護老人福祉施設いわと_職員",    # J列
        "ケアハウスユーハウス_入所者",     # L列
    ]

    # 存在する列だけ残す
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

    # 種別選択（ラジオボタン）
    order_type = st.radio(
        "作成する注文書の種類を選んでください",
        ("特養（介護老人福祉施設いわと）", "ユーハウスいわと"),
        horizontal=True,
        key="order_type",
    )

    # ファイルアップロード（共通）
    order_file = st.file_uploader(
        "注文書のもとになる検収簿 Excel をアップロード",
        type=["xlsx"],
        key="order_src",
    )

    st.markdown(
        '<p class="small-note">※ inspection_formatter / 検収簿整形で加工したもの、<br>　または同じ形式の検収簿ファイルを想定しています。</p>',
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
        st.error("注文書作成中にエラーが発生しました。アップロードしたファイルの形式を確認してください。")
        st.exception(e)



