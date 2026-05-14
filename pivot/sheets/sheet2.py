import pandas as pd


GROUP_KEYS = [
    "Product",
    "Type",
    "Currency Clasification",
    "Platform",
    "destinationCountry",
    "destinationCurrency",
    "masterCodeCategory",
    "Client",
    "Date",
]

HEADERS = GROUP_KEYS + [
    "COUNTA of id",
    "SUM of Amount (USD)",
    "SUM of Sum of Rev - Txn Fee",
    "SUM of Sum of Rev - FX Fee",
    "SUM of Sum of Cost - Txn Fee",
    "SUM of Sum of Cost - FX Fee",
]


def run(spreadsheet):
    raw = spreadsheet.worksheet("Raw Data - Transactions").get_all_values()
    df = pd.DataFrame(raw[1:], columns=raw[0])

    num_cols = ["Amount (USD)", "Sum of Rev - Txn Fee", "Sum of Rev - FX Fee", "Sum of Cost - Txn Fee", "Sum of Cost - FX Fee"]
    for col in num_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    pivot = df.groupby(GROUP_KEYS, dropna=False).agg(
        **{
            "COUNTA of id": ("id", "count"),
            "SUM of Amount (USD)": ("Amount (USD)", "sum"),
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
    print(f"  sheet2: {len(rows)} rows written to 'Pivot Table 1'")
