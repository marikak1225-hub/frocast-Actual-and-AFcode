import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="実績データ集計", layout="wide")

AF_MASTER_PATH = "AFマスター.xlsx"
AFF_ONLY_PATH = "AFF_AFコード.xlsx"
TARGET_APPLY_PATH = "目標申込件数マスター.xlsx"
TARGET_ISSUE_PATH = "目標発行件数マスター.xlsx"

# =========================
# 共通
# =========================
def normalize(val):
    if pd.isna(val):
        return ""
    return str(val).replace(" ", "").replace("　", "").strip()

def convert_date(val):
    try:
        s = str(int(val))
        return pd.Timestamp(int(s[:4]), int(s[4:6]), int(s[6:8]))
    except:
        return pd.NaT

# =========================
# マスター
# =========================
def read_af_master(path):
    df = pd.read_excel(path, header=None)
    header = df.apply(lambda r: r.astype(str).str.contains("AFコード")).any(axis=1).idxmax()
    df.columns = df.iloc[header]
    df = df.iloc[header+1:].reset_index(drop=True)
    df.columns = df.columns.map(normalize)
    df["AFコード"] = df["AFコード"].map(normalize)
    df["割り振り"] = df["割り振り"].map(normalize)
    return df[["AFコード", "割り振り", "領域"]]

def read_aff_master(path):
    df = pd.read_excel(path)
    df.columns = ["AFコード", "領域"]
    df["割り振り"] = df["AFコード"]
    return df.applymap(normalize)

def read_target(path):
    df = pd.read_excel(path, header=4)
    df.columns = df.columns.map(normalize)
    df.rename(columns={df.columns[1]: "日付"}, inplace=True)
    df["日付"] = pd.to_datetime(df["日付"])
    return df

def get_target(df, date, assign):
    row = df[df["日付"] == date]
    return 0 if row.empty or assign not in df.columns else row.iloc[0][assign]

# =========================
# ローデータ（統合済）
# =========================
def process_raw(df_raw, af_master, start, end, kind):
    af_map = af_master.set_index("AFコード")[["割り振り", "領域"]].to_dict("index")
    rows = []

    for _, r in df_raw.iterrows():
        if not (start <= r["日付"] <= end):
            continue
        for col in df_raw.columns[1:]:
            if pd.isna(r[col]) or r[col] == 0:
                continue
            info = af_map.get(normalize(col))
            if info:
                rows.append([kind, r["日付"], info["割り振り"], info["領域"], r[col]])

    df = pd.DataFrame(rows, columns=["種別", "日付", "割り振り", "領域", "実績"])
    return df.groupby(["種別", "日付", "割り振り", "領域"], as_index=False).sum()

# =========================
# 割り振り別
# =========================
def create_blocks(df, target):
    act = df.pivot_table(index="日付", columns="割り振り", values="実績", aggfunc="sum").fillna(0)
    tar = act.copy()
    for d in act.index:
        for c in act.columns:
            tar.loc[d, c] = get_target(target, d, c)

    gap = act - tar
    ratio = act / tar

    for t in [act, tar, gap, ratio]:
        t["total"] = t.sum(axis=1)
        t.index = t.index.strftime("%Y/%m/%d")

    return act, tar, gap, ratio

# =========================
# Excel出力
# =========================
def to_excel(area_a, area_i, blocks_a, blocks_i, raw):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        wb = writer.book

        base = wb.add_format({"font_name": "Meiryo UI", "border": 1})
        bold = wb.add_format({"font_name": "Meiryo UI", "border": 1, "bold": True})
        red = wb.add_format({"font_name": "Meiryo UI", "border": 1, "font_color": "red"})
        pct = wb.add_format({"font_name": "Meiryo UI", "border": 1, "num_format": "0.0%"})
        pct_red = wb.add_format({"font_name": "Meiryo UI", "border": 1, "num_format": "0.0%", "font_color": "red"})

        # ==== 領域別 ====
        ws = wb.add_worksheet("領域別")
        writer.sheets["領域別"] = ws
        row = 0
        for title, df in [("申込", area_a), ("発行", area_i)]:
            ws.write(row, 0, title, bold)
            row += 1
            df.insert(0, "total", df.sum(axis=1))
            df.loc["total"] = df.sum()
            df.to_excel(writer, "領域別", startrow=row)
            start = row + 1
            end = start + len(df)
            ws.conditional_format(start, 1, end, 1, {"type": "no_errors", "format": bold})
            ws.conditional_format(start, 0, start, len(df.columns), {"type": "no_errors", "format": bold})
            row = end + 2

        # ==== 割り振り別 ====
        ws = wb.add_worksheet("割り振り別")
        writer.sheets["割り振り別"] = ws
        row = 0
        for label, blocks in [("申込", blocks_a), ("発行", blocks_i)]:
            ws.write(row, 0, label, bold)
            row += 1
            for name, df in zip(["実績", "目標", "GAP", "Target vs Actual"], blocks):
                ws.write(row, 0, name, bold)
                df.to_excel(writer, "割り振り別", startrow=row+1)
                start = row + 2
                end = start + len(df)
                if name == "GAP":
                    ws.conditional_format(start, 1, end, len(df.columns), {"type": "<", "value": 0, "format": red})
                if name == "Target vs Actual":
                    ws.set_column(1, len(df.columns), None, pct)
                    ws.conditional_format(start, 1, end, len(df.columns), {"type": "<", "value": 0, "format": pct_red})
                row = end + 2

        # ==== ローデータ ====
        raw.to_excel(writer, "日別_集計ローデータ", index=False)

    return output.getvalue()

# =========================
# UI
# =========================
st.title("📊 実績データ集計")

apply = st.file_uploader("申込", type="xlsx")
issue = st.file_uploader("発行", type="xlsx")
if not apply or not issue:
    st.stop()

dfa = pd.read_excel(apply)
dfi = pd.read_excel(issue)
dfa.rename(columns={dfa.columns[0]: "日付"}, inplace=True)
dfi.rename(columns={dfi.columns[0]: "日付"}, inplace=True)

dfa["日付"] = dfa["日付"].map(convert_date)
dfi["日付"] = dfi["日付"].map(convert_date)

start, end = map(pd.to_datetime, st.date_input("期間", [dfa["日付"].min(), dfa["日付"].max()]))

af = pd.concat([read_af_master(AF_MASTER_PATH), read_aff_master(AFF_ONLY_PATH)])
ta = read_target(TARGET_APPLY_PATH)
ti = read_target(TARGET_ISSUE_PATH)

ra = process_raw(dfa, af, start, end, "申込")
ri = process_raw(dfi, af, start, end, "発行")

raw = pd.concat([ra, ri])
raw["目標"] = raw.apply(lambda r: get_target(ta if r["種別"]=="申込" else ti, r["日付"], r["割り振り"]), axis=1)
raw["日付"] = raw["日付"].dt.strftime("%Y/%m/%d")

area_a = ra.pivot_table(index="領域", columns="日付", values="実績", aggfunc="sum").fillna(0)
area_i = ri.pivot_table(index="領域", columns="日付", values="実績", aggfunc="sum").fillna(0)

excel = to_excel(area_a, area_i, create_blocks(ra, ta), create_blocks(ri, ti), raw)

st.download_button("📥 Excelダウンロード", excel, "実績集計.xlsx")
