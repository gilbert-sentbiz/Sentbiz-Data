import os
from datetime import datetime
import pandas as pd
import gspread

# ── 설정 ──────────────────────────────────────────────────────────────
INPUT_FOLDER    = "/path/to/xlsx/folder"  # xlsx 파일들이 있는 로컬 폴더
OUTPUT_FOLDER_ID = None                    # Google Drive 폴더 ID (None = 내 드라이브 루트)
# ──────────────────────────────────────────────────────────────────────

HEADERS = [
    "Product", "Client Segment", "Currency Classification", "Platform",
    "Destination Country", "Destination Currency", "Partner Code",
    "Client", "Date", "Count", "Volume (USD)",
    "Rev - Txn Fee", "Rev - FX Fee", "Cost - Txn Fee", "Cost - FX Fee",
]


# ── 소스 자동 감지 ────────────────────────────────────────────────────

def detect_source(columns):
    cols = set(columns)
    if "transfer_id" in cols:
        return "OSB"
    if "Baokim's Transaction ID" in cols:
        return "VND Collection"
    if "Funding Method" in cols:
        return "ERP Data"
    if "Amount (USD)" in cols and "Client" in cols:
        return "NSB Data"
    if "Volume (USD)" in cols:
        return "SenDA"
    raise ValueError(f"소스 감지 실패 — 첫 5개 컬럼: {list(columns)[:5]}")


# ── 피벗 함수 (각 소스별) ─────────────────────────────────────────────

def _to_num(df, cols):
    for c in cols:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)


def pivot_osb(df):
    df = df[df["transfer_status"] == "COMPLETE"].copy()
    _to_num(df, ["usd_base_amount", "Rev - Txn Fee", "Rev - FX Fee", "Cost - Txn Fee", "Cost - FX Fee"])
    keys = ["Product", "Client Segment", "Currency Clasification", "Platform",
            "receive_country", "receive_currency",
            "payment_gateway_wing_money_account_number", "corp_name", "Date"]
    return _agg(df, keys, count_col="transfer_id", volume_col="usd_base_amount",
                rev_txn="Rev - Txn Fee", rev_fx="Rev - FX Fee",
                cost_txn="Cost - Txn Fee", cost_fx="Cost - FX Fee")


def pivot_nsb(df):
    _to_num(df, ["Amount (USD)", "Sum of Rev - Txn Fee", "Sum of Rev - FX Fee",
                 "Sum of Cost - Txn Fee", "Sum of Cost - FX Fee"])
    keys = ["Product", "Type", "Currency Clasification", "Platform",
            "destinationCountry", "destinationCurrency", "masterCodeCategory", "Client", "Date"]
    return _agg(df, keys, count_col="id", volume_col="Amount (USD)",
                rev_txn="Sum of Rev - Txn Fee", rev_fx="Sum of Rev - FX Fee",
                cost_txn="Sum of Cost - Txn Fee", cost_fx="Sum of Cost - FX Fee")


def pivot_senda(df):
    _to_num(df, ["Volume (USD)", "Rev - Txn Fee", "Rev - FX Fee (USD)", "Cost - Txn Fee", "Cost - FX Fee"])
    keys = ["Product", "Client Segment", "Currency Clasification", "Platform",
            "destinationCountry", "destinationCurrency", "masterCodeCategory",
            "franchisePartnerName", "Date"]
    return _agg(df, keys, count_col="id", volume_col="Volume (USD)",
                rev_txn="Rev - Txn Fee", rev_fx="Rev - FX Fee (USD)",
                cost_txn="Cost - Txn Fee", cost_fx="Cost - FX Fee")


def pivot_vnd(df):
    _to_num(df, ["Volume (USD)", "Sum of Rev - Txn Fee", "Sum of Rev - FX Fee",
                 "Sum of Cost - Txn Fee", "Sum of Cost - FX Fee"])
    keys = ["Product", "Type", "Target Currency", "Collection/Payout",
            "Merchant's Transaction ID", "Client Name", "Date"]
    rows = _agg(df, keys, count_col="Volume (USD)", volume_col="Volume (USD)",
                rev_txn="Sum of Rev - Txn Fee", rev_fx="Sum of Rev - FX Fee",
                cost_txn="Sum of Cost - Txn Fee", cost_fx="Sum of Cost - FX Fee")
    return [r[:5] + ["", ""] + r[5:] for r in rows]  # 6-7번 빈칸 삽입


def pivot_erp(df):
    _to_num(df, ["Funding Amount (USD)", "Sum of Rev - Txn Fee", "Sum of Rev - FX Fee",
                 "Sum of Cost - Txn Fee", "Sum of Cost - FX Fee"])
    keys = ["Product", "Client Segment", "Currency Classification", "Platform",
            "Funding Currency", "Customer", "Date"]
    rows = _agg(df, keys, count_col="Customer", volume_col="Funding Amount (USD)",
                rev_txn="Sum of Rev - Txn Fee", rev_fx="Sum of Rev - FX Fee",
                cost_txn="Sum of Cost - Txn Fee", cost_fx="Sum of Cost - FX Fee")
    return [r[:5] + ["", ""] + r[5:] for r in rows]  # 6-7번 빈칸 삽입


def _agg(df, keys, count_col, volume_col, rev_txn, rev_fx, cost_txn, cost_fx):
    pivot = df.groupby(keys, dropna=False).agg(
        Count=(count_col, "count"),
        Volume=(volume_col, "sum"),
        RevTxn=(rev_txn, "sum"),
        RevFX=(rev_fx, "sum"),
        CostTxn=(cost_txn, "sum"),
        CostFX=(cost_fx, "sum"),
    ).reset_index()
    return pivot.values.tolist()


PIVOT_FUNCS = {
    "OSB":            pivot_osb,
    "NSB Data":       pivot_nsb,
    "SenDA":          pivot_senda,
    "VND Collection": pivot_vnd,
    "ERP Data":       pivot_erp,
}


# ── 메인 ─────────────────────────────────────────────────────────────

def main():
    xlsx_files = sorted(f for f in os.listdir(INPUT_FOLDER) if f.endswith(".xlsx"))
    print(f"파일 {len(xlsx_files)}개 발견")

    all_rows = []
    for fname in xlsx_files:
        path = os.path.join(INPUT_FOLDER, fname)
        df = pd.read_excel(path)
        source = detect_source(df.columns)
        print(f"  {fname} → {source}", end="")
        rows = PIVOT_FUNCS[source](df)
        all_rows.extend(rows)
        print(f" ({len(rows)} rows)")

    print(f"총 {len(all_rows)} rows → Google Sheet 생성 중...")

    gc = gspread.oauth()
    title = datetime.today().strftime("%Y%m%d") + "_GBD_Data"
    sh = gc.create(title, folder_id=OUTPUT_FOLDER_ID)
    sh.sheet1.update([HEADERS] + all_rows, "A1")

    print(f"완료: {sh.url}")


if __name__ == "__main__":
    main()
