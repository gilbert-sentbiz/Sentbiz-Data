import pandas as pd


GROUP_KEYS = [
    "Product",
    "Client Segment",
    "Currency Classification",
    "Platform",
    "Funding Currency",
    "Customer",
    "Date",
]

HEADERS = GROUP_KEYS + [
    "COUNTA of Customer",
    "SUM of Funding Amount (USD)",
    "SUM of Sum of Rev - Txn Fee",
    "SUM of Sum of Rev - FX Fee",
    "SUM of Sum of Cost - Txn Fee",
    "SUM of Sum of Cost - FX Fee",
]


def run(spreadsheet):
    raw = spreadsheet.worksheet("Filtered Data").get_all_values()

    # deduplicate column names (e.g. 'Conversion Rate' appears twice)
    seen = {}
    headers = []
    for h in raw[0]:
        if h in seen:
            seen[h] += 1
            headers.append(f"{h}_{seen[h]}")
        else:
            seen[h] = 0
            headers.append(h)

    df = pd.DataFrame(raw[1:], columns=headers)

    num_cols = ["Funding Amount (USD)", "Sum of Rev - Txn Fee", "Sum of Rev - FX Fee", "Sum of Cost - Txn Fee", "Sum of Cost - FX Fee"]
    for col in num_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    pivot = df.groupby(GROUP_KEYS, dropna=False).agg(
        **{
            "COUNTA of Customer": ("Customer", "count"),
            "SUM of Funding Amount (USD)": ("Funding Amount (USD)", "sum"),
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
    print(f"  sheet5: {len(rows)} rows written to 'Pivot Table 1'")
