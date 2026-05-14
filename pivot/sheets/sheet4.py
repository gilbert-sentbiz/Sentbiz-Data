import pandas as pd


GROUP_KEYS = [
    "Product",
    "Type",
    "Target Currency",
    "Collection/Payout",
    "Merchant's Transaction ID",
    "Client Name",
    "Date",
]

HEADERS = GROUP_KEYS + [
    "COUNT of Volume (USD)",
    "SUM of Volume (USD)",
    "SUM of Sum of Rev - Txn Fee",
    "SUM of Sum of Rev - FX Fee",
    "SUM of Sum of Cost - Txn Fee",
    "SUM of Sum of Cost - FX Fee",
]


def run(spreadsheet):
    raw = spreadsheet.worksheet("Raw Data").get_all_values()
    df = pd.DataFrame(raw[1:], columns=raw[0])

    num_cols = ["Volume (USD)", "Sum of Rev - Txn Fee", "Sum of Rev - FX Fee", "Sum of Cost - Txn Fee", "Sum of Cost - FX Fee"]
    for col in num_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    pivot = df.groupby(GROUP_KEYS, dropna=False).agg(
        **{
            "COUNT of Volume (USD)": ("Volume (USD)", "count"),
            "SUM of Volume (USD)": ("Volume (USD)", "sum"),
            "SUM of Sum of Rev - Txn Fee": ("Sum of Rev - Txn Fee", "sum"),
            "SUM of Sum of Rev - FX Fee": ("Sum of Rev - FX Fee", "sum"),
            "SUM of Sum of Cost - Txn Fee": ("Sum of Cost - Txn Fee", "sum"),
            "SUM of Sum of Cost - FX Fee": ("Sum of Cost - FX Fee", "sum"),
        }
    ).reset_index()

    rows = pivot.values.tolist()

    ws = spreadsheet.worksheet("Pivot Table 1")
    ws.clear()
    ws.update([HEADERS] + rows, "A1")
    print(f"  sheet4: {len(rows)} rows written to 'Pivot Table 1'")
