import os
from datetime import datetime
import pandas as pd
import gspread

# ── 설정 ──────────────────────────────────────────────────────────────
INPUT_FOLDER     = "/Users/gilbert/Desktop/Sentbiz Data"         # xlsx 파일들이 있는 로컬 폴더
OUTPUT_FOLDER_ID = "1mldXezakLMwLD-NZ_bz06zV_i3aaeVjb"          # Google Drive 폴더 ID
# ──────────────────────────────────────────────────────────────────────

HEADERS = [
    "Product", "Client Segment", "Currency Classification", "Platform",
    "Destination Country", "Destination Currency", "Partner Code",
    "Client", "Date", "Count", "Volume (USD)",
    "Rev - Txn Fee", "Rev - FX Fee", "Cost - Txn Fee", "Cost - FX Fee",
]

CONV_RATES = {"IDR": 16990, "USD": 1, "SGD": 1.3}


# ── xlsx 스마트 읽기 (헤더 위치 자동 감지) ────────────────────────────

def smart_read(path):
    df = pd.read_excel(path)
    first_col = str(df.columns[0])
    if first_col.startswith("Unnamed"):   # OSB: 상단 2개 빈 행 → header=2
        df = pd.read_excel(path, header=2)
    elif first_col.isdigit():             # SenDA: 숫자 인덱스 행 → header=1
        df = pd.read_excel(path, header=1)
    return df


# ── 소스 자동 감지 ────────────────────────────────────────────────────

def detect_source(columns):
    cols = set(str(c) for c in columns)
    if "Baokim's Transaction ID" in cols:
        return "VND"
    if "Funding Method" in cols:
        return "ERP"
    return "PIVOT"  # 이미 피벗된 출력물 (OSB / NSB / SenDA)


# ── 이미 피벗된 파일: 값만 추출 ──────────────────────────────────────

def pass_through(df):
    # 15컬럼 순서 그대로 반환 (컬럼명 무시, 위치 기준)
    return df.iloc[:, :15].values.tolist()


# ── VND Collection: 원본 → 피벗 ──────────────────────────────────────

def pivot_vnd(df):
    num_cols = ["Volume (USD)", "Sum of Rev - Txn Fee", "Sum of Rev - FX Fee",
                "Sum of Cost - Txn Fee", "Sum of Cost - FX Fee"]
    for c in num_cols:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

    keys = ["Product", "Type", "Target Currency", "Collection/Payout",
            "Merchant's Transaction ID", "Client Name", "Date"]
    pivot = df.groupby(keys, dropna=False).agg(
        Count=("Volume (USD)", "count"),
        Volume=("Volume (USD)", "sum"),
        RevTxn=("Sum of Rev - Txn Fee", "sum"),
        RevFX=("Sum of Rev - FX Fee", "sum"),
        CostTxn=("Sum of Cost - Txn Fee", "sum"),
        CostFX=("Sum of Cost - FX Fee", "sum"),
    ).reset_index()

    return [r[:5] + ["", ""] + r[5:] for r in pivot.values.tolist()]


# ── ERP Data: 원본 → 가공 → 피벗 ────────────────────────────────────

def pivot_erp(df):
    df = df[df["Conversion Method"] != "Transaction base"].copy()

    df["Funding Amount"] = pd.to_numeric(df["Funding Amount"], errors="coerce").fillna(0)
    df["Request Amount"] = pd.to_numeric(df["Request Amount"], errors="coerce").fillna(0)
    df["FX Spread"]      = pd.to_numeric(df["FX Spread"], errors="coerce").fillna(0)

    base = df["Funding Amount"].where(df["Funding Amount"] > 0, df["Request Amount"])
    df["Funding Amount (USD)"] = base / df["Funding Currency"].map(CONV_RATES).fillna(1)
    df["Rev - FX Fee"]  = df["FX Spread"] / 100 * df["Funding Amount (USD)"]
    df["Rev - Txn Fee"] = 0
    df["Cost - Txn Fee"] = 0
    df["Cost - FX Fee"]  = 0

    df["Product"]                = "Bulk Conversion Deal"
    df["Client Segment"]         = "Non-KR FI"
    df["Currency Classification"] = df["Top-up Currency"]
    df["Platform"]               = "Manual"
    df["Date"] = pd.to_datetime(df["Received Date"]).dt.strftime("%d/%m/%Y")

    keys = ["Product", "Client Segment", "Currency Classification", "Platform",
            "Funding Currency", "Customer", "Date"]
    pivot = df.groupby(keys, dropna=False).agg(
        Count=("Customer", "count"),
        Volume=("Funding Amount (USD)", "sum"),
        RevTxn=("Rev - Txn Fee", "sum"),
        RevFX=("Rev - FX Fee", "sum"),
        CostTxn=("Cost - Txn Fee", "sum"),
        CostFX=("Cost - FX Fee", "sum"),
    ).reset_index()

    return [r[:5] + ["", ""] + r[5:] for r in pivot.values.tolist()]


# ── 메인 ─────────────────────────────────────────────────────────────

def main():
    xlsx_files = sorted(f for f in os.listdir(INPUT_FOLDER) if f.endswith(".xlsx"))
    print(f"파일 {len(xlsx_files)}개 발견")

    all_rows = []
    for fname in xlsx_files:
        path = os.path.join(INPUT_FOLDER, fname)
        df = smart_read(path)
        source = detect_source(df.columns)
        print(f"  {fname} → {source}", end="")

        if source == "PIVOT":
            rows = pass_through(df)
        elif source == "VND":
            rows = pivot_vnd(df)
        elif source == "ERP":
            rows = pivot_erp(df)

        all_rows.extend(rows)
        print(f" ({len(rows)} rows)")

    print(f"총 {len(all_rows)} rows → Google Sheet 생성 중...")

    # NaN / inf / Timestamp → 직렬화 가능한 값으로 정리
    import math
    def clean(v):
        if v is None:
            return ""
        if isinstance(v, float) and (math.isnan(v) or math.isinf(v)):
            return ""
        if hasattr(v, 'isoformat'):  # Timestamp, datetime
            return str(v)
        return v
    all_rows = [[clean(v) for v in row] for row in all_rows]

    gc = gspread.oauth()
    title = datetime.today().strftime("%Y%m%d") + "_GBD_Data"
    sh = gc.create(title, folder_id=OUTPUT_FOLDER_ID)
    sh.sheet1.update([HEADERS] + all_rows, "A1")

    print(f"완료: {sh.url}")


if __name__ == "__main__":
    main()
