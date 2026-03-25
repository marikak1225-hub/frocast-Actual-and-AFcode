import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

st.set_page_config(page_title="集計アプリ（申込・発行対応）", layout="wide")

#  ファイルパス（app.py と同じフォルダ）
AF_MASTER_PATH = "AFマスター.xlsx"
AFF_ONLY_PATH = "AFF_AFコード.xlsx"
TARGET申込_PATH = "目標申込件数マスター.xlsx"
TARGET発行_PATH = "目標発行件数マスター.xlsx"

# normalize
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

# AFマスター（動的ヘッダー）
def read_af_master(path):
    df = pd.read_excel(path, header=None, engine="openpyxl")

    header_row = None
    for i in range(len(df)):
        row = df.iloc[i].astype(str).str.replace(" ", "")
        if row.str.contains("AFコード|AFｺｰﾄﾞ|ＡＦコード", regex=True).any():
            header_row = i
            break

    if header_row is None:
        st.error("AFマスターに『AFコード』列が見つかりません。")
        st.stop()

    df.columns = df.iloc[header_row]
    df = df.iloc[header_row + 1:].reset_index(drop=True)

    df.columns = [normalize_assign(c) for c in df.columns]

    required_cols = ["AFコード", "割り振り", "領域"]
    for col in required_cols:
        if col not in df.columns:
            st.error(f"AFマスターに必要な列『{col}』がありません。")
            st.stop()

    df["割り振り"] = df["割り振り"].apply(normalize_assign)
    return df

# Affiliate 専用 AFコードマスター読込
def read_affiliate_master(path):
    df = pd.read_excel(path, engine="openpyxl", header=0)

    # このファイルは A列＝AFコード、B列＝領域 の二列構造
    df.columns = ["AFコード", "領域"]

    # 割り振りは AFコードと同じ（任意だが Listing/Display 形式に合わせる）
    df["割り振り"] = df["AFコード"]

    # normalize
    df["AFコード"] = df["AFコード"].apply(normalize_assign)
    df["割り振り"] = df["割り振り"].apply(normalize_assign)

    # 領域はそのまま（= Affiliate）
    return df[["AFコード", "割り振り", "領域"]]

# 日付変換
def convert_date(val):
    try:
        s = str(int(val))
        return pd.Timestamp(
            year=int(s[:4]),
            month=int(s[4:6]),
            day=int(s[6:8])
        )
    except:
        return pd.NaT

# 割り振り・領域付与
def attach_assign_area(df_raw, af_master, start, end):

    af_map = af_master.set_index("AFコード")[["割り振り", "領域"]].to_dict("index")

    records = []

    for _, row in df_raw.iterrows():
        dt = row["日付"]
        if pd.isna(dt) or not (start <= dt <= end):
            continue

        for col in df_raw.columns[1:]:
            val = row[col]
            if pd.isna(val) or val == 0:
                continue

            info = af_map.get(col)
            if info is None:
                continue

            assign = normalize_assign(info["割り振り"])
            area = info["領域"]

            records.append([dt, assign, area, val])

    df = pd.DataFrame(records, columns=["日付", "割り振り", "領域", "実績"])
    df = df.groupby(["日付", "割り振り", "領域"], as_index=False).sum()
    return df

# 目標マスター
def read_target_master(path):
    df = pd.read_excel(path, header=4, engine="openpyxl")
    df.columns = [normalize_assign(c) for c in df.columns]

    date_col = df.columns[1]
    df = df.rename(columns={date_col: "日付"})
    df["日付"] = pd.to_datetime(df["日付"], errors="coerce")
    return df

def get_target_value(date, assign, target_master):
    if pd.isna(date):
        return 0

    assign = normalize_assign(assign)

    row = target_master[target_master["日付"] == date]
    if row.empty:
        return 0
    if assign not in target_master.columns:
        return 0

    val = row.iloc[0][assign]
    return 0 if pd.isna(val) else val

# サマリ
def create_summary(df_data, af_master):

    af_master_sorted = (
        af_master.groupby("領域")["割り振り"].apply(list).to_dict()
    )

    ordered_pairs = []
    for area, assigns in af_master_sorted.items():
        for assign in assigns:
            ordered_pairs.append((area, assign))

    pt_act = df_data.pivot_table(
        index="日付",
        columns=["領域", "割り振り"],
        values="実績",
        aggfunc="sum"
    ).fillna(0)

    pt_tar = df_data.pivot_table(
        index="日付",
        columns=["領域", "割り振り"],
        values="目標",
        aggfunc="sum"
    ).fillna(0)

    out = pd.DataFrame(index=pt_act.index)

    for (area, assign) in ordered_pairs:
        out[(area, assign, "実績")] = pt_act.get((area, assign), 0)
        out[(area, assign, "目標")] = pt_tar.get((area, assign), 0)

    out.columns = pd.MultiIndex.from_tuples(out.columns)
    out.index = out.index.map(lambda x: f"{x.year}/{x.month}/{x.day}")
    return out

# Excel 出力
def to_excel(summary_df, detail_df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:

        summary_df.to_excel(writer, sheet_name="サマリ")
        workbook = writer.book
        fmt = workbook.add_format({"align": "center"})
        ws1 = writer.sheets["サマリ"]
        ws1.set_column(0, len(summary_df.columns) + 2, 15, fmt)

        detail_df.to_excel(writer, sheet_name="明細", index=False)
        ws2 = writer.sheets["明細"]
        ws2.set_column(0, len(detail_df.columns) + 2, 15)

    return output.getvalue()

# Streamlit UI
st.title("📊 申込 / 発行 ・AFコード　集計")

mode = st.radio("集計対象", ["申込データ集計", "発行データ集計"], horizontal=True)

uploaded = st.file_uploader("📤 実績データ", type=["xlsx"])
if uploaded is None:
    st.stop()

df_raw = pd.read_excel(uploaded, engine="openpyxl")
df_raw.rename(columns={df_raw.columns[0]: "日付"}, inplace=True)
df_raw["日付"] = df_raw["日付"].apply(convert_date)

min_date, max_date = df_raw["日付"].min(), df_raw["日付"].max()
date_range = st.date_input("📅 期間選択", value=[min_date, max_date])
start, end = pd.to_datetime(date_range[0]), pd.to_datetime(date_range[1])

# ★ AFマスター + Affiliateマスター を結合
af_master_base = read_af_master(AF_MASTER_PATH)
af_master_aff  = read_affiliate_master(AFF_ONLY_PATH)

# ★ 縦方向に結合
af_master = pd.concat([af_master_base, af_master_aff], ignore_index=True)

# 目標マスター
if mode == "申込データ集計":
    target_master = read_target_master(TARGET申込_PATH)
else:
    target_master = read_target_master(TARGET発行_PATH)

df_data = attach_assign_area(df_raw, af_master, start, end)

df_data["目標"] = df_data.apply(
    lambda r: get_target_value(r["日付"], r["割り振り"], target_master),
    axis=1
)

summary_df = create_summary(df_data, af_master)

st.subheader("✅ サマリ（AFマスター順）")
st.dataframe(summary_df, use_container_width=True)

df_data["日付"] = df_data["日付"].map(lambda x: f"{x.year}/{x.month}/{x.day}")

excel_bytes = to_excel(summary_df, df_data)

st.download_button(
    label="📤 集計結果をダウンロード（Excel）",
    data=excel_bytes,
    file_name="集計結果.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True
)
