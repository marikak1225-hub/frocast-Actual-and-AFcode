import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

st.set_page_config(page_title="集計アプリ（申込・発行対応）", layout="wide")

# クラウド環境では app.py と同じフォルダに置かれたファイルを参照
AF_MASTER_PATH = "AFマスター.xlsx"
TARGET申込_PATH = "目標申込件数マスター.xlsx"
TARGET発行_PATH = "目標発行件数マスター.xlsx"

# normalize 関数（揺れ吸収）

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

# AFマスター読込（動的ヘッダー検出）
def read_af_master(path):
    df = pd.read_excel(path, header=None, engine="openpyxl")

    # ヘッダー行を探す
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

# 日付（YYYYMMDD → Timestamp）
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

# CVデータ → 割り振り/領域付与
def attach_assign_area(df_raw, af_master, start, end):
    # AFコード → 割り振り/領域 辞書
    af_map = af_master.set_index("AFコード")[["割り振り", "領域"]].to_dict("index")

    records = []

    for _, row in df_raw.iterrows():
        dt = row["日付"]
        if pd.isna(dt) or not (start <= dt <= end):
            continue

        # A列（日付）以外のAFコード列
        for col in df_raw.columns[1:]:
            val = row[col]
            if pd.isna(val) or val == 0:
                continue

            info = af_map.get(col)
            if info is None:
                # AFマスターに存在しないコードは無視
                continue

            assign = normalize_assign(info["割り振り"])
            area = info["領域"]

            records.append([dt, assign, area, val])

    # DataFrame 化
    df = pd.DataFrame(records, columns=["日付", "割り振り", "領域", "実績"])

    #  同じ日付 × 割り振り × 領域 は合算
    df = df.groupby(["日付", "割り振り", "領域"], as_index=False).sum()

    return df

# 目標マスター読込（5行目ヘッダー、B列＝日）
def read_target_master(path):
    # ヘッダーは5行目（header=4）
    df = pd.read_excel(path, header=4, engine="openpyxl")

    # 列名 normalize
    df.columns = [normalize_assign(c) for c in df.columns]

    # B列（日付列）を強制的に「日付」に統一
    date_col = df.columns[1]
    df = df.rename(columns={date_col: "日付"})

    # Timestamp へ変換
    df["日付"] = pd.to_datetime(df["日付"], errors="coerce")

    return df

# 目標値取得（Excel INDEX/MATCH 相当）
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

# サマリ作成（AFマスター順 → 3段ヘッダー）
def create_summary(df_data, af_master):

    # AFマスター順の「領域→割り振り」順
    af_master_sorted = (
        af_master.groupby("領域")["割り振り"].apply(list).to_dict()
    )

    ordered_pairs = []
    for area, assigns in af_master_sorted.items():
        for assign in assigns:
            ordered_pairs.append((area, assign))

    # Pivot 実績
    pt_act = df_data.pivot_table(
        index="日付",
        columns=["領域", "割り振り"],
        values="実績",
        aggfunc="sum"
    ).fillna(0)

    # Pivot 目標
    pt_tar = df_data.pivot_table(
        index="日付",
        columns=["領域", "割り振り"],
        values="目標",
        aggfunc="sum"
    ).fillna(0)

    # 出力 DataFrame
    out = pd.DataFrame(index=pt_act.index)

    for (area, assign) in ordered_pairs:
        out[(area, assign, "実績")] = pt_act.get((area, assign), 0)
        out[(area, assign, "目標")] = pt_tar.get((area, assign), 0)

    out.columns = pd.MultiIndex.from_tuples(out.columns)

    # 日付フォーマットを yyyy/m/d に統一
    out.index = out.index.map(lambda x: f"{x.year}/{x.month}/{x.day}")

    return out

# Excel 出力（2シート）
def to_excel(summary_df, detail_df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:

        # --- サマリ ---
        summary_df.to_excel(writer, sheet_name="サマリ")

        workbook = writer.book
        fmt = workbook.add_format({"align": "center"})
        ws1 = writer.sheets["サマリ"]
        ws1.set_column(0, len(summary_df.columns) + 2, 15, fmt)

        # --- 明細 ---
        detail_df.to_excel(writer, sheet_name="明細", index=False)
        ws2 = writer.sheets["明細"]
        ws2.set_column(0, len(detail_df.columns) + 2, 15)

    return output.getvalue()

# Streamlit UI
st.title("📊 申込 / 発行 ・AFコード　集計")

# --- 集計モード選択 ---
mode = st.radio("集計対象", ["申込データ集計", "発行データ集計"], horizontal=True)

# --- ファイルアップロード ---
uploaded = st.file_uploader("📤 実績データ（セキュリティラベルなしにしてアップロード）", type=["xlsx"])
if uploaded is None:
    st.stop()

df_raw = pd.read_excel(uploaded, engine="openpyxl")

# --- 日付変換 ---
df_raw.rename(columns={df_raw.columns[0]: "日付"}, inplace=True)
df_raw["日付"] = df_raw["日付"].apply(convert_date)

min_date, max_date = df_raw["日付"].min(), df_raw["日付"].max()

# --- 集計期間 ---
date_range = st.date_input("📅 期間選択", value=[min_date, max_date])
start, end = pd.to_datetime(date_range[0]), pd.to_datetime(date_range[1])

# --- AFマスター読込 ---
af_master = read_af_master(AF_MASTER_PATH)

# --- 目標マスター切替 ---
if mode == "申込データ集計":
    target_master = read_target_master(TARGET申込_PATH)
else:
    target_master = read_target_master(TARGET発行_PATH)

# --- CVデータに割り振り付与 ---
df_data = attach_assign_area(df_raw, af_master, start, end)

# --- 目標値付与（INDEX/MATCH） ---
df_data["目標"] = df_data.apply(
    lambda r: get_target_value(r["日付"], r["割り振り"], target_master),
    axis=1
)

# --- サマリ作成 ---
summary_df = create_summary(df_data, af_master)

# --- サマリ表示 ---
st.subheader("✅ サマリ（AFマスター順）")
st.dataframe(summary_df, use_container_width=True)

# UIには明細を表示しない → Excel だけに出力
df_data["日付"] = df_data["日付"].map(lambda x: f"{x.year}/{x.month}/{x.day}")

# --- Excel ダウンロード ---
excel_bytes = to_excel(summary_df, df_data)

st.download_button(
    label="📤 集計結果をダウンロード（Excel）",
    data=excel_bytes,
    file_name="集計結果.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True
)
