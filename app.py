import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="実績データ集計", layout="wide")

AF_MASTER_PATH = "AFマスター.xlsx"
AFF_ONLY_PATH = "AFF_AFコード.xlsx"
TARGET_APPLY_PATH = "目標申込件数マスター.xlsx"
TARGET_ISSUE_PATH = "目標発行件数マスター.xlsx"

# =========================
# 共通関数
# =========================
def normalize(val):
    if pd.isna(val):
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
        return pd.Timestamp(int(s[:4]), int(s[4:6]), int(s[6:8]))
    except:
        return pd.NaT

# =========================
# マスター読込
# =========================
def read_af_master(path):
    df = pd.read_excel(path, header=None, engine="openpyxl")
    header = None
    for i in range(len(df)):
        row = df.iloc[i].astype(str).apply(normalize)
        if row.str.contains("AFコード|AFｺｰﾄﾞ|ＡＦコード").any():
            header = i
            break
    if header is None:
        st.error("AFコード列が見つかりません")
        st.stop()

    df.columns = df.iloc[header]
    df = df.iloc[header + 1:].reset_index(drop=True)
    df.columns = [normalize(c) for c in df.columns]
    df["AFコード"] = df["AFコード"].apply(normalize)
    df["割り振り"] = df["割り振り"].apply(normalize)
    return df[["AFコード", "割り振り", "領域"]]

def read_aff_master(path):
    df = pd.read_excel(path, engine="openpyxl")
    df.columns = ["AFコード", "領域"]
    df["割り振り"] = df["AFコード"]
    df = df.applymap(normalize)
    return df[["AFコード", "割り振り", "領域"]]

def read_target(path):
    df = pd.read_excel(path, header=4, engine="openpyxl")
    df.columns = [normalize(c) for c in df.columns]
    date_col = df.columns[1]
    df = df.rename(columns={date_col: "日付"})
    df["日付"] = pd.to_datetime(df["日付"], errors="coerce")
    return df

def get_target(df_target, date, assign):
    row = df_target[df_target["日付"] == date]
    if row.empty or assign not in df_target.columns:
        return 0
    val = row.iloc[0][assign]
    return 0 if pd.isna(val) else val

# =========================
# 実績ローデータ整形（★ダブルカウント修正済）
# =========================
def process_raw(df_raw, af_master, start, end, kind):
    af_map = af_master.set_index("AFコード")[["割り振り", "領域"]].to_dict("index")
    records = []

    for _, r in df_raw.iterrows():
        dt = r["日付"]
        if pd.isna(dt) or not (start <= dt <= end):
            continue

        for col in df_raw.columns[1:]:
            val = r[col]
            if pd.isna(val) or val == 0:
                continue

            info = af_map.get(normalize(col))
            if info is None:
                continue

            records.append(
                [kind, dt, info["割り振り"], info["領域"], val]
            )

    df = pd.DataFrame(
        records,
        columns=["種別", "日付", "割り振り", "領域", "実績"],
    )

    if df.empty:
        return df

    # ✅ 日付・割り振り・領域単位で統合（ダブルカウント防止）
    df = (
        df
        .groupby(["種別", "日付", "割り振り", "領域"], as_index=False)
        .agg({"実績": "sum"})
    )

    return df

# =========================
# 割り振り別ブロック作成
# =========================
def create_blocks(df, target_master):
    act = df.pivot_table(
        index="日付",
        columns="割り振り",
        values="実績",
        aggfunc="sum",
        fill_value=0
    )

    tar = act.copy()
    for d in act.index:
        for c in act.columns:
            tar.loc[d, c] = get_target(target_master, d, c)

    for t in [act, tar]:
        t["total"] = t.sum(axis=1)

    gap = act - tar
    ratio = act.divide(tar).replace([float("inf"), -float("inf")], pd.NA)

    for df_ in [act, tar, gap, ratio]:
        df_.index = df_.index.strftime("%Y/%m/%d")

    return act, tar, gap, ratio

# =========================
# Excel出力
# =========================
def to_excel(area_apply, area_issue, blocks_apply, blocks_issue, raw_df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:

        # 領域別
        ws_area = writer.book.add_worksheet("領域別")
        writer.sheets["領域別"] = ws_area
        row = 0
        for title, df in [("申込", area_apply), ("発行", area_issue)]:
            ws_area.write(row, 0, title)
            df.to_excel(writer, sheet_name="領域別", startrow=row + 1)
            row += len(df) + 4

        # 割り振り別
        ws_assign = writer.book.add_worksheet("割り振り別")
        writer.sheets["割り振り別"] = ws_assign
        row = 0
        for label, blocks in [("申込", blocks_apply), ("発行", blocks_issue)]:
            ws_assign.write(row, 0, label)
            row += 1
            for title, df in zip(
                ["■実績", "■目標", "■GAP", "■Target vs Actual"], blocks
            ):
                ws_assign.write(row, 0, title)
                df.to_excel(writer, sheet_name="割り振り別", startrow=row + 1)
                row += len(df) + 3
            row += 2

        # ローデータ
        raw_df.to_excel(writer, sheet_name="日別_集計ローデータ", index=False)

    return output.getvalue()

# =========================
# UI
# =========================
st.title("📊 実績データ集計")

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

min_d = min(df_apply["日付"].min(), df_issue["日付"].min())
max_d = max(df_apply["日付"].max(), df_issue["日付"].max())

start, end = map(
    pd.to_datetime,
    st.date_input("📅 期間選択", value=[min_d, max_d])
)

af_master = pd.concat(
    [read_af_master(AF_MASTER_PATH), read_aff_master(AFF_ONLY_PATH)],
    ignore_index=True
)

target_apply = read_target(TARGET_APPLY_PATH)
target_issue = read_target(TARGET_ISSUE_PATH)

df_a = process_raw(df_apply, af_master, start, end, "申込")
df_i = process_raw(df_issue, af_master, start, end, "発行")

# =========================
# 領域別
# =========================
area_apply = df_a.pivot_table(
    index="領域", columns="日付", values="実績", aggfunc="sum", fill_value=0
)
area_issue = df_i.pivot_table(
    index="領域", columns="日付", values="実績", aggfunc="sum", fill_value=0
)

for df in [area_apply, area_issue]:
    df["total"] = df.sum(axis=1)
    df.loc["total"] = df.sum()
    df.columns = ["total"] + [
        c.strftime("%Y/%m/%d") for c in df.columns if c != "total"
    ]

st.subheader("✅ 領域別サマリ（申込）")
st.dataframe(area_apply, use_container_width=True)

st.subheader("✅ 領域別サマリ（発行）")
st.dataframe(area_issue, use_container_width=True)

# =========================
# 割り振り別（Excel用）
# =========================
blocks_apply = create_blocks(df_a, target_apply)
blocks_issue = create_blocks(df_i, target_issue)

raw = pd.concat([df_a, df_i], ignore_index=True)
raw["日付"] = raw["日付"].dt.strftime("%Y/%m/%d")

excel = to_excel(area_apply, area_issue, blocks_apply, blocks_issue, raw)

st.download_button(
    "📥 集計結果をダウンロード（Excel）",
    excel,
    f"実績集計結果_{start:%Y%m%d}_{end:%Y%m%d}.xlsx",
    use_container_width=True,
)
