import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="実績データ集計（申込＋発行）", layout="wide")

# =====================
# ファイルパス
# =====================
AF_MASTER_PATH = "AFマスター.xlsx"
AFF_ONLY_PATH = "AFF_AFコード.xlsx"

# =====================
# 共通関数
# =====================
def normalize_assign(val):
    if val is None:
        return ""
    return (
        str(val)
        .replace(" ", "")
        .replace("　", "")
        .replace("\n", "")
        .replace("\r", "")
        .strip()
    )

def convert_date(val):
    try:
        s = str(int(val))
        return pd.Timestamp(
            year=int(s[:4]),
            month=int(s[4:6]),
            day=int(s[6:8]),
        )
    except:
        return pd.NaT

# =====================
# AFマスター（安全版）
# =====================
def read_af_master(path):
    df = pd.read_excel(path, header=None, engine="openpyxl")

    header_row = None
    for i in range(len(df)):
        row = df.iloc[i].astype(str).apply(normalize_assign)
        if row.str.contains("AFコード|AFｺｰﾄﾞ|ＡＦコード", regex=True).any():
            header_row = i
            break

    if header_row is None:
        st.error("AFマスターに『AFコード』列が見つかりません。")
        st.stop()

    df.columns = df.iloc[header_row]
    df = df.iloc[header_row + 1 :].reset_index(drop=True)
    df.columns = [normalize_assign(c) for c in df.columns]

    required = ["AFコード", "割り振り", "領域"]
    for col in required:
        if col not in df.columns:
            st.error(f"AFマスターに必要な列『{col}』がありません。")
            st.stop()

    df["AFコード"] = df["AFコード"].apply(normalize_assign)
    df["割り振り"] = df["割り振り"].apply(normalize_assign)

    return df[["AFコード", "割り振り", "領域"]]

def read_affiliate_master(path):
    df = pd.read_excel(path, engine="openpyxl")
    df.columns = ["AFコード", "領域"]
    df["割り振り"] = df["AFコード"]

    df["AFコード"] = df["AFコード"].apply(normalize_assign)
    df["割り振り"] = df["割り振り"].apply(normalize_assign)

    return df[["AFコード", "割り振り", "領域"]]

# =====================
# 実績データ処理
# =====================
def process_raw(df_raw, af_master, start, end, kind):
    af_map = (
        af_master.set_index("AFコード")[["割り振り", "領域"]]
        .to_dict("index")
    )

    records = []

    for _, row in df_raw.iterrows():
        dt = row["日付"]
        if pd.isna(dt) or not (start <= dt <= end):
            continue

        for col in df_raw.columns[1:]:
            val = row[col]
            if pd.isna(val) or val == 0:
                continue

            info = af_map.get(normalize_assign(col))
            if info is None:
                continue

            records.append([
                kind,
                dt,
                info["割り振り"],
                info["領域"],
                val,
            ])

    return pd.DataFrame(
        records,
        columns=["種別", "日付", "割り振り", "領域", "実績"],
    )

# =====================
# サマリ
# =====================
def create_area_summary(df):
    pvt = df.pivot_table(
        index="領域",
        columns="日付",
        values="実績",
        aggfunc="sum",
        fill_value=0,
    )

    pvt["total"] = pvt.sum(axis=1)
    total_row = pvt.sum().to_frame().T
    total_row.index = ["total"]
    pvt = pd.concat([pvt, total_row])

    pvt.columns = [
        c.strftime("%Y/%m/%d") if isinstance(c, pd.Timestamp) else c
        for c in pvt.columns
    ]
    cols = ["total"] + [c for c in pvt.columns if c != "total"]
    return pvt[cols]

# =====================
# Streamlit UI
# =====================
st.title("📊 実績データ集計（申込＋発行）")

apply_file = st.file_uploader("📤 申込データ", type=["xlsx"])
issue_file = st.file_uploader("📤 発行データ", type=["xlsx"])

if not apply_file or not issue_file:
    st.stop()

df_apply = pd.read_excel(apply_file, engine="openpyxl")
df_issue = pd.read_excel(issue_file, engine="openpyxl")

df_apply.rename(columns={df_apply.columns[0]: "日付"}, inplace=True)
df_issue.rename(columns={df_issue.columns[0]: "日付"}, inplace=True)

df_apply["日付"] = df_apply["日付"].apply(convert_date)
df_issue["日付"] = df_issue["日付"].apply(convert_date)

min_date = min(df_apply["日付"].min(), df_issue["日付"].min())
max_date = max(df_apply["日付"].max(), df_issue["日付"].max())

start, end = map(
    pd.to_datetime,
    st.date_input("📅 期間選択", value=[min_date, max_date]),
)

af_master = pd.concat(
    [
        read_af_master(AF_MASTER_PATH),
        read_affiliate_master(AFF_ONLY_PATH),
    ],
    ignore_index=True,
)

df_apply_p = process_raw(df_apply, af_master, start, end, "申込")
df_issue_p = process_raw(df_issue, af_master, start, end, "発行")

st.subheader("✅ 領域別サマリ（申込）")
st.dataframe(create_area_summary(df_apply_p), use_container_width=True)

st.subheader("✅ 領域別サマリ（発行）")
st.dataframe(create_area_summary(df_issue_p), use_container_width=True)
