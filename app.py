import streamlit as st
import pandas as pd
from io import BytesIO
import os
from datetime import datetime

st.set_page_config(page_title="実績データ集計", layout="wide")

# =========================
# SharePoint（OneDrive同期）上のマスターパス
# ※ 環境に合わせて1回だけ修正してください
# =========================
AF_MASTER_PATH = (
    "C:/Users/xxx/OneDrive - 会社名/実績集計マスター/AFマスター.xlsx"
)
TARGET_APPLY_PATH = (
    "C:/Users/xxx/OneDrive - 会社名/実績集計マスター/目標申込件数マスター.xlsx"
)
TARGET_ISSUE_PATH = (
    "C:/Users/xxx/OneDrive - 会社名/実績集計マスター/目標発行件数マスター.xlsx"
)

# =========================
# 共通関数
# =========================
def normalize(val):
    if pd.isna(val):
        return ""
    return str(val).replace(" ", "").replace("　", "").strip()

def convert_date(val):
    try:
        s = str(int(val))
        return pd.Timestamp(f"{s[:4]}/{int(s[4:6])}/{int(s[6:8])}")
    except:
        return pd.NaT

def file_info(path):
    if not os.path.exists(path):
        return "❌ 見つかりません"
    t = os.path.getmtime(path)
    return datetime.fromtimestamp(t).strftime("%Y/%m/%d %H:%M")

# =========================
# マスター読込
# =========================
def read_af_master(path):
    df = pd.read_excel(path, header=None)
    header = df.apply(
        lambda r: r.astype(str).str.contains("AFコード"), axis=1
    ).idxmax()
    df.columns = df.iloc[header]
    df = df.iloc[header + 1:].reset_index(drop=True)
    df.columns = df.columns.map(normalize)
    df["AFコード"] = df["AFコード"].map(normalize)
    df["割り振り"] = df["割り振り"].map(normalize)
    return df[["AFコード", "割り振り", "領域"]]

def read_target(path):
    df = pd.read_excel(path, header=4)
    df.columns = df.columns.map(normalize)
    df.rename(columns={df.columns[1]: "日付"}, inplace=True)
    df["日付"] = pd.to_datetime(df["日付"], errors="coerce")
    return df

def get_target(df, date, assign):
    row = df[df["日付"] == date]
    if row.empty or assign not in df.columns:
        return 0
    v = row.iloc[0][assign]
    return 0 if pd.isna(v) else v

# =========================
# 実績ローデータ（ダブルカウント防止）
# =========================
def process_raw(df_raw, af_master, start, end, kind):
    af_map = af_master.set_index("AFコード")[["割り振り", "領域"]].to_dict("index")
    rows = []

    for _, r in df_raw.iterrows():
        if pd.isna(r["日付"]) or not (start <= r["日付"] <= end):
            continue

        for col in df_raw.columns[1:]:
            if pd.isna(r[col]) or r[col] == 0:
                continue

            info = af_map.get(normalize(col))
            if info:
                rows.append([kind, r["日付"], info["割り振り"], info["領域"], r[col]])

    df = pd.DataFrame(
        rows, columns=["種別", "日付", "割り振り", "領域", "実績"]
    )

    return (
        df.groupby(["種別", "日付", "割り振り", "領域"], as_index=False)
          .agg({"実績": "sum"})
    )

# =========================
# 割り振り別ブロック（元仕様）
# =========================
def create_blocks(df, target):
    act = df.pivot_table(
        index="日付",
        columns="割り振り",
        values="実績",
        aggfunc="sum",
        fill_value=0,
    )

    tar = act.copy()
    for d in act.index:
        for c in act.columns:
            tar.loc[d, c] = get_target(target, d, c)

    for t in [act, tar]:
        t["total"] = t.sum(axis=1)
        t.index = t.index.strftime("%Y/%m/%d")

    gap = act - tar
    ratio = act.divide(tar)

    return act, tar, gap, ratio

# =========================
# Excel出力（完成版）
# =========================
def to_excel(area_a, area_i, blocks_a, blocks_i, raw):
    output = BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        wb = writer.book
        bold = wb.add_format({"font_name": "Meiryo UI", "bold": True})
        title_fmt = wb.add_format(
            {"font_name": "Meiryo UI", "bold": True, "border": 1}
        )

        # ===== 領域別 =====
        ws = wb.add_worksheet("領域別")
        writer.sheets["領域別"] = ws
        row = 0

        for title, df in [("申込", area_a), ("発行", area_i)]:
            ws.write(row, 0, title, title_fmt)
            row += 1

            df2 = df.copy()
            df2.insert(0, "total", df2.sum(axis=1))
            df2.loc["total"] = df2.sum()

            df2.columns = ["total"] + [
                c.strftime("%Y/%m/%d") for c in df2.columns[1:]
            ]

            df2.to_excel(writer, sheet_name="領域別", startrow=row)

            ws.set_column(1, 1, None, bold)
            ws.set_row(row + len(df2), None, bold)

            row += len(df2) + 3

        # ===== 割り振り別 =====
        ws = wb.add_worksheet("割り振り別")
        writer.sheets["割り振り別"] = ws
        row = 0

        for label, blocks in [("申込", blocks_a), ("発行", blocks_i)]:
            ws.write(row, 0, label, bold)
            row += 1

            for name, df in zip(
                ["■実績", "■目標", "■GAP", "■Target vs Actual"],
                blocks,
            ):
                ws.write(row, 0, name, bold)
                df.to_excel(writer, sheet_name="割り振り別", startrow=row + 1)
                row += len(df) + 3

        # ===== ローデータ =====
        raw.to_excel(writer, sheet_name="日別_集計ローデータ", index=False)

    return output.getvalue()

# =========================
# UI
# =========================
st.title("📊 実績データ集計")

# ---- 注意書き（★要件どおり表示）----
st.info(
    "【マスター管理について】\n\n"
    "マスターは SharePoint で管理しています。\n"
    "ツールは起動時にそのファイルを直接参照します。\n"
    "更新時は SharePoint 上のファイルを上書きしてください。\n"
    "UI や GitHub からマスターを書き換えることは想定していません。"
)

# ---- 参照中マスター情報 ----
st.sidebar.markdown("### 📂 参照中マスター")
st.sidebar.write("AFマスター：", file_info(AF_MASTER_PATH))
st.sidebar.write("目標申込：", file_info(TARGET_APPLY_PATH))
st.sidebar.write("目標発行：", file_info(TARGET_ISSUE_PATH))

apply = st.file_uploader("📤 申込データ", type="xlsx")
issue = st.file_uploader("📤 発行データ", type="xlsx")
if not apply or not issue:
    st.stop()

dfa = pd.read_excel(apply)
dfi = pd.read_excel(issue)
dfa.rename(columns={dfa.columns[0]: "日付"}, inplace=True)
dfi.rename(columns={dfi.columns[0]: "日付"}, inplace=True)

dfa["日付"] = dfa["日付"].apply(convert_date)
dfi["日付"] = dfi["日付"].apply(convert_date)

start, end = map(
    pd.to_datetime,
    st.date_input("📅 期間選択", [dfa["日付"].min(), dfa["日付"].max()]),
)

af = read_af_master(AF_MASTER_PATH)
ta = read_target(TARGET_APPLY_PATH)
ti = read_target(TARGET_ISSUE_PATH)

ra = process_raw(dfa, af, start, end, "申込")
ri = process_raw(dfi, af, start, end, "発行")

# ---- UI サマリ（領域別）----
def make_summary_df(area_df):
    df = area_df.copy()
    df.insert(0, "total", df.sum(axis=1))
    df.loc["total"] = df.sum()
    df.columns = ["total"] + [
        pd.to_datetime(c).strftime("%Y/%m/%d") for c in df.columns[1:]
    ]
    return df

st.subheader("📌 サマリ（領域別）")

st.markdown("### 申込")
area_apply = ra.pivot_table(
    index="領域", columns="日付", values="実績", aggfunc="sum", fill_value=0
)
st.dataframe(make_summary_df(area_apply), use_container_width=True)

st.markdown("### 発行")
area_issue = ri.pivot_table(
    index="領域", columns="日付", values="実績", aggfunc="sum", fill_value=0
)
st.dataframe(make_summary_df(area_issue), use_container_width=True)

# ---- ローデータ ----
raw = pd.concat([ra, ri], ignore_index=True)
raw["目標"] = raw.apply(
    lambda r: get_target(
        ta if r["種別"] == "申込" else ti,
        r["日付"],
        r["割り振り"],
    ),
    axis=1,
)
raw["日付"] = raw["日付"].dt.strftime("%Y/%m/%d")

# ---- Excel出力 ----
excel = to_excel(
    area_apply,
    area_issue,
    create_blocks(ra, ta),
    create_blocks(ri, ti),
    raw,
)

st.download_button(
    "📥 集計結果をダウンロード（Excel）",
    excel,
    f"実績集計結果_{start:%Y%m%d}_{end:%Y%m%d}.xlsx",
    use_container_width=True,
)
