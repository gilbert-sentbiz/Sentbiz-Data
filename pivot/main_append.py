import os
import math
from datetime import datetime, date as _date
import pandas as pd
import gspread

SHEETS_EPOCH = _date(1899, 12, 30)

def to_serial_date(val):
    """DD/MM/YYYY 문자열 → Google Sheets 날짜 시리얼 숫자"""
    try:
        dt = datetime.strptime(str(val), "%d/%m/%Y").date()
        return (dt - SHEETS_EPOCH).days
    except (ValueError, TypeError):
        return val


def load_week_mapping(sh):
    """Week Mapping 탭에서 {날짜 시리얼: 주차 레이블} 딕셔너리 반환"""
    wm = sh.worksheet("Week Mapping")
    mapping = {}
    for row in wm.get_all_values()[1:]:  # 헤더 스킵
        if len(row) >= 5 and row[4]:
            try:
                mapping[int(float(row[4]))] = row[3]
            except (ValueError, IndexError):
                pass
    return mapping

# ── 설정 ──────────────────────────────────────────────────────────────
INPUT_FOLDER    = "/Users/gilbert/Desktop/Sentbiz Data"
MASTER_SHEET_ID = "1m1scU5eLl-WmOVB7FgzR_RN7avwxIdOGx-90kusDcqI"
# ──────────────────────────────────────────────────────────────────────

HEADERS = [
    "Product", "Client Segment", "Currency Classification", "Platform",
    "Destination Country", "Destination Currency", "Partner Code",
    "Client", "Date", "Count", "Volume (USD)",
    "Rev - Txn Fee", "Rev - FX Fee", "Cost - Txn Fee", "Cost - FX Fee",
]

CONV_RATES = {"IDR": 16990, "USD": 1, "SGD": 1.3}

# 원본 탭 이름 우선순위 (먼저 발견되는 탭을 읽음)
RAW_TABS = ["ERP SenDa Raw Data", "Raw Data - Transactions", "Original Data", "Raw Data"]


# ── xlsx 스마트 읽기 ──────────────────────────────────────────────────

def smart_read(path):
    xl = pd.ExcelFile(path)
    for tab in RAW_TABS:
        if tab in xl.sheet_names:
            df = xl.parse(tab)
            if not df.empty and len(df.columns) > 0:
                return df
    raise ValueError(f"원본 탭 없음 — 발견된 탭: {xl.sheet_names}")


# ── 소스 자동 감지 ────────────────────────────────────────────────────

def detect_source(columns):
    cols = set(str(c) for c in columns)
    if "transfer_id" in cols:
        return "OSB"
    if "Baokim's Transaction ID" in cols:
        return "VND"
    if "Funding Method" in cols:
        return "ERP"
    if "Amount (USD)" in cols and "Client" in cols:
        return "NSB"
    if "Volume (USD)" in cols and "id" in cols:
        return "SenDA"
    return "PIVOT"  # 이미 피벗된 출력물


# ── 공통 유틸 ─────────────────────────────────────────────────────────

def _to_num(df, cols):
    for c in cols:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)


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


# ── 피벗 함수 ─────────────────────────────────────────────────────────

def pivot_osb(df):
    df = df[df["transfer_status"] == "COMPLETE"].copy()
    _to_num(df, ["usd_base_amount", "Rev - Txn Fee", "Rev - FX Fee", "Cost - Txn Fee", "Cost - FX Fee"])
    keys = ["Product", "Client Segment", "Currency Clasification", "Platform",
            "receive_country", "receive_currency",
            "payment_gateway_wing_money_account_number", "corp_name", "Date"]
    return _agg(df, keys, "transfer_id", "usd_base_amount",
                "Rev - Txn Fee", "Rev - FX Fee", "Cost - Txn Fee", "Cost - FX Fee")


def pivot_nsb(df):
    _to_num(df, ["Amount (USD)", "Sum of Rev - Txn Fee", "Sum of Rev - FX Fee",
                 "Sum of Cost - Txn Fee", "Sum of Cost - FX Fee"])
    keys = ["Product", "Type", "Currency Clasification", "Platform",
            "destinationCountry", "destinationCurrency", "masterCodeCategory", "Client", "Date"]
    return _agg(df, keys, "id", "Amount (USD)",
                "Sum of Rev - Txn Fee", "Sum of Rev - FX Fee",
                "Sum of Cost - Txn Fee", "Sum of Cost - FX Fee")


def pivot_senda(df):
    _to_num(df, ["Volume (USD)", "Rev - Txn Fee", "Rev - FX Fee (USD)", "Cost - Txn Fee", "Cost - FX Fee"])
    keys = ["Product", "Client Segment", "Currency Clasification", "Platform",
            "destinationCountry", "destinationCurrency", "masterCodeCategory",
            "franchisePartnerName", "Date"]
    return _agg(df, keys, "id", "Volume (USD)",
                "Rev - Txn Fee", "Rev - FX Fee (USD)", "Cost - Txn Fee", "Cost - FX Fee")


def pivot_vnd(df):
    _to_num(df, ["Volume (USD)", "Sum of Rev - Txn Fee", "Sum of Rev - FX Fee",
                 "Sum of Cost - Txn Fee", "Sum of Cost - FX Fee"])
    keys = ["Product", "Type", "Target Currency", "Collection/Payout",
            "Merchant's Transaction ID", "Client Name", "Date"]
    rows = _agg(df, keys, "Volume (USD)", "Volume (USD)",
                "Sum of Rev - Txn Fee", "Sum of Rev - FX Fee",
                "Sum of Cost - Txn Fee", "Sum of Cost - FX Fee")
    return [r[:5] + ["", ""] + r[5:] for r in rows]


def pivot_erp(df):
    df = df[df["Conversion Method"] != "Transaction base"].copy()
    df["Funding Amount"] = pd.to_numeric(df["Funding Amount"], errors="coerce").fillna(0)
    df["Request Amount"] = pd.to_numeric(df["Request Amount"], errors="coerce").fillna(0)
    df["FX Spread"]      = pd.to_numeric(df["FX Spread"], errors="coerce").fillna(0)

    base = df["Funding Amount"].where(df["Funding Amount"] > 0, df["Request Amount"])
    df["Funding Amount (USD)"] = base / df["Funding Currency"].map(CONV_RATES).fillna(1)
    df["Rev - FX Fee"]   = df["FX Spread"] / 100 * df["Funding Amount (USD)"]
    df["Rev - Txn Fee"]  = 0
    df["Cost - Txn Fee"] = 0
    df["Cost - FX Fee"]  = 0
    df["Product"]                 = "Bulk Conversion Deal"
    df["Client Segment"]          = "Non-KR FI"
    df["Currency Classification"] = df["Top-up Currency"]
    df["Platform"]                = "Manual"
    df["Date"] = pd.to_datetime(df["Received Date"]).dt.strftime("%d/%m/%Y")

    keys = ["Product", "Client Segment", "Currency Classification", "Platform",
            "Funding Currency", "Customer", "Date"]
    rows = _agg(df, keys, "Customer", "Funding Amount (USD)",
                "Rev - Txn Fee", "Rev - FX Fee", "Cost - Txn Fee", "Cost - FX Fee")
    return [r[:5] + ["", ""] + r[5:] for r in rows]


def pass_through(df):
    return df.iloc[:, :15].values.tolist()


PIVOT_FUNCS = {
    "OSB":   pivot_osb,
    "NSB":   pivot_nsb,
    "SenDA": pivot_senda,
    "VND":   pivot_vnd,
    "ERP":   pivot_erp,
    "PIVOT": pass_through,
}


# ── 유틸 ─────────────────────────────────────────────────────────────

def clean(v):
    if v is None:
        return ""
    if isinstance(v, float) and (math.isnan(v) or math.isinf(v)):
        return ""
    if hasattr(v, "isoformat"):
        return str(v)
    return v


def rows_to_weeks(rows, week_mapping):
    """시리얼 날짜(index 8)에서 시트 기준 주차 레이블 추출"""
    weeks = set()
    for row in rows:
        serial = row[8]
        if isinstance(serial, (int, float)):
            label = week_mapping.get(int(serial))
            if label:
                weeks.add(label)
    return weeks


def delete_weeks(ws, target_weeks):
    """Data 탭에서 target_weeks에 해당하는 행을 삭제 (Week 컬럼 = index 15)"""
    all_vals = ws.get_all_values()
    WEEK_COL = 15  # 0-based

    to_delete = [
        i + 2
        for i, row in enumerate(all_vals[1:])
        if len(row) > WEEK_COL and row[WEEK_COL] in target_weeks
    ]

    if not to_delete:
        return 0, len(all_vals)

    ranges = []
    start = end = to_delete[0]
    for r in to_delete[1:]:
        if r == end + 1:
            end = r
        else:
            ranges.append((start, end))
            start = end = r
    ranges.append((start, end))

    for s, e in reversed(ranges):
        ws.delete_rows(s, e)

    return len(to_delete), len(all_vals) - len(to_delete)


def write_formulas(ws, start_row, num_rows):
    """새로 추가된 행(start_row~)에 P~V 수식을 직접 생성해서 쓴다."""
    formula_rows = []
    for r in range(start_row, start_row + num_rows):
        formula_rows.append([
            f"=XLOOKUP(I{r},'Week Mapping'!$A$2:$A$733,'Week Mapping'!$D$2:$D$733)",
            f"=L{r}+M{r}",
            f"=L{r}+M{r}-N{r}-O{r}",
            f"=month(I{r})",
            f"=year(I{r})",
            f'=if(or(B{r}="KR MERCHANT",B{r}="KR FI"),"Glenn","Valex")',
            f'=if(or(B{r}="KR MERCHANT",B{r}="KR FI"),"Joyce","Bernice")',
        ])
    ws.update(formula_rows, f"P{start_row}:V{start_row + num_rows - 1}",
              value_input_option="USER_ENTERED")


# ── 메인 ─────────────────────────────────────────────────────────────

def main():
    xlsx_files = sorted(f for f in os.listdir(INPUT_FOLDER) if f.endswith(".xlsx"))
    print(f"파일 {len(xlsx_files)}개 발견")

    all_rows = []
    for fname in xlsx_files:
        path = os.path.join(INPUT_FOLDER, fname)
        if fname.startswith("~$"):
            print(f"  {fname} → 건너뜀 (Excel 임시 잠금 파일)")
            continue
        try:
            df = smart_read(path)
        except ValueError as e:
            print(f"  {fname} → 건너뜀 ({e})")
            continue
        except Exception as e:
            print(f"  {fname} → 건너뜀 (파일 읽기 실패: {e})")
            continue
        source = detect_source(df.columns)
        rows = PIVOT_FUNCS[source](df)
        all_rows.extend(rows)
        print(f"  {fname} → {source} ({len(rows)} rows)")

    all_rows = [[clean(v) for v in row] for row in all_rows]

    # Client Segment(index 1) 대문자 정규화, 날짜(index 8) 시리얼 변환
    for row in all_rows:
        if isinstance(row[1], str):
            row[1] = row[1].upper().strip()
        row[8] = to_serial_date(row[8])

    gc = gspread.oauth()
    sh = gc.open_by_key(MASTER_SHEET_ID)
    ws = sh.worksheet("Data")

    week_mapping = load_week_mapping(sh)
    target_weeks = rows_to_weeks(all_rows, week_mapping)
    print(f"대상 주차: {sorted(target_weeks)}")

    deleted, remaining = delete_weeks(ws, target_weeks)
    if deleted:
        print(f"기존 {deleted} rows 삭제 완료")

    ws.append_rows(all_rows, value_input_option="USER_ENTERED")
    start_row = remaining + 1  # 헤더 포함 remaining행 다음부터
    write_formulas(ws, start_row, len(all_rows))
    print(f"완료: {len(all_rows)} rows → Data 탭에 추가됨 (수식 작성 완료)")


if __name__ == "__main__":
    main()
