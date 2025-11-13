# NOTE: Streamlit is required to run this app.
# Ensure you install it via: pip install streamlit

import io, zipfile, datetime
import pandas as pd
from io import BytesIO
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.worksheet.page import PageMargins

try:
    import streamlit as st
except ModuleNotFoundError:
    raise ModuleNotFoundError("Streamlit is not installed. Please run `pip install streamlit` in your environment.")

st.set_page_config(page_title="いわと発注管理", layout="centered")
TITLE = "いわと発注管理"
st.title(TITLE)
st.caption("ブラウザだけで『検収簿の加工 → 仕入先別注文書（いわと／ユーハウス）』を作成します。")

# STEP 1：検収簿の加工
def to_excel_bytes(df, startrow=0):
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, startrow=startrow)
    bio.seek(0)
    return bio.getvalue()

with st.expander("STEP 1：検収簿の加工（空欄補完付き）", expanded=True):
    uploaded_raw = st.file_uploader("検収記録簿（原本 .xlsx）をアップロード", type=["xlsx"], key="raw")

    if st.button("加工する ▶", use_container_width=True, disabled=(uploaded_raw is None)):
        try:
            raw_bytes = uploaded_raw.read()
            bio = BytesIO(raw_bytes)
            df = pd.read_excel(bio, header=[6, 7])
            df.columns = ['_'.join([str(i) for i in col if str(i) != 'nan']).strip() for col in df.columns]

            fill_cols = [
                'Unnamed: 0_level_0_納品日',
                'Unnamed: 1_level_0_使用日',
                'Unnamed: 2_level_0_朝昼夕',
                'Unnamed: 3_level_0_仕入先'
            ]
            df[fill_cols] = df[fill_cols].ffill()

            order_map = {'朝食': 1, '昼食': 2, '夕食': 3, '3時': 4}
            df['朝昼夕_order'] = df['Unnamed: 2_level_0_朝昼夕'].map(order_map).fillna(5)

            df_sorted = df.sort_values(by=[
                'Unnamed: 1_level_0_使用日',
                '朝昼夕_order',
                'Unnamed: 5_level_0_食品名'
            ])

            cols = [
                'Unnamed: 0_level_0_納品日',
                'Unnamed: 1_level_0_使用日',
                'Unnamed: 2_level_0_朝昼夕',
                'Unnamed: 3_level_0_仕入先',
                'Unnamed: 5_level_0_食品名',
                'Unnamed: 6_level_0_換算値',
                'Unnamed: 7_level_0_総合計',
                'Unnamed: 8_level_0_単位',
                '介護老人福祉施設いわと_入所者',
                '介護老人福祉施設いわと_職員',
                'ケアハウスユー…_入所者'
            ]
            df_out = df_sorted[cols].reset_index(drop=True)

            ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            st.success("✅ 加工が完了しました。ダウンロードしてください。")
            st.download_button(
                label="加工済ファイルをダウンロード",
                data=to_excel_bytes(df_out, startrow=0),
                file_name=f"検収簿_加工済_空欄補完済み_{ts}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        except Exception as e:
            st.error(f"❌ 加工中にエラー：{e}")

st.markdown("---")

# STEP 2：発注書作成（いわと／ユーハウス）
from order_form_iwato_with_checker2 import generate_zip as iwato_zip
from order_form_yuhouse2 import generate_zip as yuhouse_zip

with st.expander("STEP 2：仕入先別 注文書を作成（ZIP）", expanded=True):
    uploaded_proc = st.file_uploader("加工済ファイル（.xlsx）をアップロード", type=["xlsx"], key="proc")
    facility = st.radio("施設を選択", options=["いわと", "ユーハウス"], horizontal=True)
    st.caption("出力仕様：A4横／MSゴシック22pt／行高30／細罫線／検収者列あり")

    if st.button("注文書を作成 ▶", use_container_width=True, disabled=(uploaded_proc is None)):
        try:
            proc_bytes = uploaded_proc.read()
            df = pd.read_excel(BytesIO(proc_bytes), header=0)

            if facility == "いわと":
                zip_data = iwato_zip(df)
            else:
                zip_data = yuhouse_zip(df)

            ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            st.success("✅ 仕入先別の注文書をZIPで用意しました。ダウンロードしてください。")
            st.download_button(
                label="注文書ZIPをダウンロード",
                data=zip_data,
                file_name=f"注文書_{facility}_{ts}.zip",
                mime="application/zip",
                use_container_width=True
            )
        except Exception as e:
            st.error(f"❌ 作成中にエラー：{e}")
