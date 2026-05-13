import pandas as pd


GROUP_KEYS = [
    "Product",
    "Client Segment",
    "Currency Clasification",
    "Platform",
    "receive_country",
    "receive_currency",
    "payment_gateway_wing_money_account_number",
    "corp_name",
    "Date",
]

HEADERS = GROUP_KEYS + [
    "Count of transfer_id",
    "Sum of usd_base_amount",
    "Sum of Rev - Txn Fee",
    "Sum of Rev - FX Fee",
    "Sum of Cost - Txn Fee",
    "Sum of Cost - FX Fee",
]


def run(spreadsheet):
    raw = spreadsheet.worksheet("Original Data").get_all_values()
    df = pd.DataFrame(raw[1:], columns=raw[0])
    df = df[df["transfer_status"] == "COMPLETE"].copy()

    num_cols = ["usd_base_amount", "Rev - Txn Fee", "Rev - FX Fee", "Cost - Txn Fee", "Cost - FX Fee"]
    for col in num_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    pivot = df.groupby(GROUP_KEYS, dropna=False).agg(
        **{
            "Count of transfer_id": ("transfer_id", "count"),
            "Sum of usd_base_amount": ("usd_base_amount", "sum"),
            "Sum of Rev - Txn Fee": ("Rev - Txn Fee", "sum"),
            "Sum of Rev - FX Fee": ("Rev - FX Fee", "sum"),
            "Sum of Cost - Txn Fee": ("Cost - Txn Fee", "sum"),
            "Sum of Cost - FX Fee": ("Cost - FX Fee", "sum"),
        }
    ).reset_index()

    rows = pivot.values.tolist()

    ws = spreadsheet.worksheet("Pivoted Data")
    ws.clear()
    ws.update([[""] * len(HEADERS), [""] * len(HEADERS), HEADERS] + rows, "A1")
    print(f"  sheet1: {len(rows)} rows written to 'Pivoted Data'")
