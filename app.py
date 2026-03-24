# ============================================
# app.py（完全版：申込/発行・AF順・目標一致・Excel出力）
# ============================================

import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import os

# ============================================
# 基本設定
# ============================================

st.set_page_config(page_title="集計アプリ（申込・発行対応）", layout="wide")

BASE = r"C:\work\shukei_app"
AF_MASTER_PATH = os.path.join(BASE, "AFマスター.xlsx")
TARGET申込_PATH = os.path.join(BASE, "目標申込件数マスター.xlsx")
TARGET発行_PATH = os.path.join(BASE, "目標発行件数マスター.xlsx")

# ============================================
# normalize 関数（揺れ吸収）
# ============================================

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

# ============================================
# AFマスター読込（動的ヘッダー）
# ============================================

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

    if "割り振り" not in df.columns or "領域" not in df.columns or "AFコード" not in df.columns:
        st.error("AFマスターに必要な列（AFコード, 割り振り, 領域）がありません。")
        st.stop()

    df["割り振り"] = df["割り振り"].apply(normalize_assign)

    return df

# ============================================
# 日付（YYYYMMDD → Timestamp）
# ============================================

def convert_date(val):
    try:
        s = str(int(val))
        return pd.Timestamp(year=int(s[:4]), month=int(s[4:6]), day=int(s[6:8]))
    except:
        return pd.NaT

# ============================================
# CVデータ → 割り振り/領域付与 → 縦持ち
# ============================================

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

    # ✅ 重複行を合算（同日付 × 同割り振り × 同領域）
    df = df.groupby(["日付", "割り振り", "領域"], as_index=False).sum()

    return df

# ============================================
# 目標マスター読込（5 行目ヘッダー / B列＝日）
# ============================================

def read_target_master(path):
    df = pd.read_excel(path, header=4, engine="openpyxl")

    df.columns = [normalize_assign(c) for c in df.columns]

    date_col = df.columns[1]  # ✅ B列を日付列とみなす
    df = df.rename(columns={date_col: "日付"})
    df["日付"] = pd.to_datetime(df["日付"], errors="coerce")

    return df

# ============================================
# INDEX/MATCH（目標値取得）
# ============================================

def get_target_value(date, assign, target_master):
    if pd.isna(date):
        return 0

    assign = normalize_assign(assign)

    row = target_master[target_master["日付"] == date]
    if len(row) == 0:
        return 0

    if assign not in target_master.columns:
        return 0

    val = row.iloc[0][assign]
    if pd.isna(val):
        return 0

    return val

# ============================================
# サマリ（AF順・3段マルチヘッダー）
# ============================================

def create_summary(df_data, af_master):

    af_master_sorted = (
        af_master.groupby("領域")["割り振り"].apply(list).to_dict()
    )

    ordered_pairs = []
    for area in af_master_sorted:
        for assign in af_master_sorted[area]:
            ordered_pairs.append((area, assign))

    pt_act = df_data.pivot_table(
        index="日付", columns=["領域", "割り振り"], values="実績", aggfunc="sum"
    ).fillna(0)

    pt_tar = df_data.pivot_table(
        index="日付", columns=["領域", "割り振り"], values="目標", aggfunc="sum"
    ).fillna(0)

    out = pd.DataFrame(index=pt_act.index)

    for (area, assign) in ordered_pairs:
        out[(area, assign, "実績")] = pt_act.get((area, assign), 0)
        out[(area, assign, "目標")] = pt_tar.get((area, assign), 0)

    out.columns = pd.MultiIndex.from_tuples(out.columns)

    # ✅ 日付を yyyy/m/d に統一
    out.index = out.index.map(lambda x: f"{x.year}/{x.month}/{x.day}")

    return out

# ============================================
# Excel 出力（2シート）
# ============================================

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

# ============================================
# Streamlit UI
# ============================================

st.title("📊 集計アプリ（申込 / 発行 対応）")

mode = st.radio("集計対象", ["申込データ集計", "発行データ集計"], horizontal=True)

uploaded = st.file_uploader("📤 CVデータ（横持ちの実績データ）", type=["xlsx"])
if uploaded is None:
    st.stop()

df_raw = pd.read_excel(uploaded, engine="openpyxl")
df_raw.rename(columns={df_raw.columns[0]: "日付"}, inplace=True)
df_raw["日付"] = df_raw["日付"].apply(convert_date)

min_date, max_date = df_raw["日付"].min(), df_raw["日付"].max()

date_range = st.date_input("📅 期間選択", value=[min_date, max_date])
start, end = pd.to_datetime(date_range[0]), pd.to_datetime(date_range[1])

af_master = read_af_master(AF_MASTER_PATH)

if mode == "申込データ集計":
    target_master = read_target_master(TARGET申込_PATH)
else:
    target_master = read_target_master(TARGET発行_PATH)

df_data = attach_assign_area(df_raw, af_master, start, end)

df_data["目標"] = df_data.apply(
    lambda r: get_target_value(r["日付"], r["割り振り"], target_master), axis=1
)

summary_df = create_summary(df_data, af_master)

st.subheader("✅ サマリ")
st.dataframe(summary_df, use_container_width=True)

st.subheader("📋 明細")
st.dataframe(df_data, use_container_width=True)

excel_bytes = to_excel(summary_df, df_data)

st.download_button(
    label="📥 集計結果をダウンロード（Excel）",
    data=excel_bytes,
    file_name="集計結果.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True
)