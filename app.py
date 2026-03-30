import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="実績データ集計（申込＋発行）", layout="wide")

# =====================
# ファイルパス
# =====================
AF_MASTER_PATH = "AFマスター.xlsx"
AFF_ONLY_PATH = "AFF_AFコード.xlsx"
TARGET申込_PATH = "目標申込件数マスター.xlsx"
TARGET発行_PATH = "目標発行件数マスター.xlsx"

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
            day=int(s[6:8])
        )
    except:
        return pd.NaT

# =====================
# マスター読込
# =====================
def read_af_master(path):
    df = pd.read_excel(path, header=None, engine="openpyxl")
    header_row = df.index[df.iloc[:, 0].astype(str).str.contains("AF")][0]
    df.columns = df.iloc[header_row]
    df = df.iloc[header_row + 1:]
    df.columns = [normalize_assign(c) for c in df.columns]
    return df[["AFコード", "割り振り", "領域"]]

def read_affiliate_master(path):
    df = pd.read_excel(path, engine="openpyxl")
    df.columns = ["AFコード", "領域"]
    df["割り振り"] = df["AFコード"]
    return df[["AFコード", "割り振り", "領域"]]

def read_target_master(path):
    df = pd.read_excel(path, header=4, engine="openpyxl")
    df.columns = [normalize_assign(c) for c in df.columns]
    date_col = df.columns[1]
    df = df.rename(columns={date_col: "日付"})
    df["日付"] = pd.to_datetime(df["日付"], errors="coerce")
    return df

# =====================
# 実績データ整形
# =====================
def process_raw(df_raw, af_master, start, end, kind):
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

            records.append([
                kind,
                dt,
                normalize_assign(info["割り振り"]),
                info["領域"],
                val
            ])

    return pd.DataFrame(
        records,
        columns=["種別", "日付", "割り振り", "領域", "実績"]
    )

# =====================
# サマリ生成
# =====================
def create_area_summary(df):
    pvt = df.pivot_table(
        index="領域",
        columns="日付",
        values="実績",
        aggfunc="sum",
        fill_value=0
    )
    pvt["total"] = pvt.sum(axis=1)
    total_row = pvt.sum().to_frame().T
    total_row.index = ["total"]
    pvt = pd.concat([pvt, total_row])

    cols = ["total"] + [c for c in pvt.columns if c != "total"]
    pvt = pvt[cols]
    pvt.columns = [c.strftime("%Y/%m/%d") if not isinstance(c, str) else c for c in pvt.columns]
    return pvt

def create_assign_summary(df):
    pvt = df.pivot_table(
        index="日付",
        columns="割り振り",
        values="実績",
        aggfunc="sum",
        fill_value=0
    )
    pvt["total"] = pvt.sum(axis=1)
    pvt.index = pvt.index.strftime("%Y/%m/%d")
    return pvt

# =====================
# Excel 出力
# =====================
def to_excel(area_apply, area_issue, assign_apply, assign_issue, raw_df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:

        area_apply.to_excel(writer, sheet_name="領域別_申込")
        area_issue.to_excel(writer, sheet_name="領域別_発行")

        assign_apply.to_excel(writer, sheet_name="割り振り別_申込")
        assign_issue.to_excel(writer, sheet_name="割り振り別_発行")

        raw_df.to_excel(writer, sheet_name="日別_集計ローデータ", index=False)

    return output.getvalue()

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

start, end = map(pd.to_datetime, st.date_input(
    "📅 期間選択",
    value=[min_date, max_date]
))

af_master = pd.concat([
    read_af_master(AF_MASTER_PATH),
    read_affiliate_master(AFF_ONLY_PATH)
], ignore_index=True)

df_apply_p = process_raw(df_apply, af_master, start, end, "申込")
df_issue_p = process_raw(df_issue, af_master, start, end, "発行")

df_all = pd.concat([df_apply_p, df_issue_p], ignore_index=True)

# ===== サマリ =====
area_apply = create_area_summary(df_apply_p)
area_issue = create_area_summary(df_issue_p)

assign_apply = create_assign_summary(df_apply_p)
assign_issue = create_assign_summary(df_issue_p)

st.subheader("✅ 領域別サマリ（申込）")
st.dataframe(area_apply, use_container_width=True)

st.subheader("✅ 領域別サマリ（発行）")
st.dataframe(area_issue, use_container_width=True)

excel_bytes = to_excel(
    area_apply,
    area_issue,
    assign_apply,
    assign_issue,
    df_all
)

filename = f"実績集計結果_{start:%Y%m%d}_{end:%Y%m%d}.xlsx"

st.download_button(
    "📥 集計結果をダウンロード（Excel）",
    data=excel_bytes,
    file_name=filename,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True
)
``
