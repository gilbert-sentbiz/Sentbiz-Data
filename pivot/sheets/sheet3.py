import pandas as pd


GROUP_KEYS = [
    "Product",
    "Client Segment",
    "Currency Clasification",
    "Platform",
    "destinationCountry",
    "destinationCurrency",
    "masterCodeCategory",
    "franchisePartnerName",
    "Date",
]

HEADERS = GROUP_KEYS + [
    "COUNTA of id",
    "SUM of Volume (USD)",
    "SUM of Rev - Txn Fee",
    "SUM of Rev - FX Fee (USD)",
    "SUM of Cost - Txn Fee",
    "SUM of Cost - FX Fee",
]

INDEX_ROW = [str(i) for i in range(1, len(HEADERS))] + [""]


def run(spreadsheet):
    raw = spreadsheet.worksheet("Original Data").get_all_values()
    df = pd.DataFrame(raw[1:], columns=raw[0])

    num_cols = ["Volume (USD)", "Rev - Txn Fee", "Rev - FX Fee (USD)", "Cost - Txn Fee", "Cost - FX Fee"]
    for col in num_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    pivot = df.groupby(GROUP_KEYS, dropna=False).agg(
        **{
            "COUNTA of id": ("id", "count"),
            "SUM of Volume (USD)": ("Volume (USD)", "sum"),
            "SUM of Rev - Txn Fee": ("Rev - Txn Fee", "sum"),
            "SUM of Rev - FX Fee (USD)": ("Rev - FX Fee (USD)", "sum"),
            "SUM of Cost - Txn Fee": ("Cost - Txn Fee", "sum"),
            "SUM of Cost - FX Fee": ("Cost - FX Fee", "sum"),
        }
    ).reset_index()

    rows = pivot.values.tolist()

    ws = spreadsheet.worksheet("Pivot Table 1")
    ws.clear()
    ws.update([INDEX_ROW, HEADERS] + rows, "A1")
    print(f"  sheet3: {len(rows)} rows written to 'Pivot Table 1'")
