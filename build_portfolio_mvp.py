import pandas as pd
from pathlib import Path
from datetime import datetime
import time
import yfinance as yf

SOURCE = Path(r"C:/Users/bgand/.openclaw/media/inbound/e5dc6078-2b7f-4d35-9ed3-8b4f2a89c10e.xlsx")
OUTPUT = Path("portfolio-mvp.html")
SUMMARY = Path("portfolio-updates.txt")


def parse_money(value):
    if isinstance(value, str):
        clean = value.replace("$", "").replace(",", "").strip()
        if not clean:
            return None
        try:
            return float(clean)
        except ValueError:
            return None
    if value in ("", None):
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def pct_format(value):
    if value is None or pd.isna(value):
        return ""
    return f"{value:+.1f}%"


def money_format(value):
    if value is None or pd.isna(value):
        return ""
    return f"${value:,.2f}"


def fetch_quote(symbol: str):
    try:
        data = yf.download(symbol, period="2d", interval="1d", progress=False, auto_adjust=False)
        data = data.dropna(subset=["Close"])
        if data.empty:
            return None
        last = data.iloc[-1]
        prev = data.iloc[-2] if len(data) > 1 else None
        price = float(last["Close"])
        prev_close = float(prev["Close"]) if prev is not None else None
        day_change = price - prev_close if prev_close else None
        day_pct = (day_change / prev_close * 100) if prev_close else None
        return {
            "price": price,
            "prev_close": prev_close,
            "day_change": day_change,
            "day_pct": day_pct,
        }
    except Exception:
        return None


def fetch_quotes(symbols):
    quotes = {}
    for symbol in symbols:
        clean = symbol.strip().upper()
        if not clean or clean in quotes:
            continue
        quote = fetch_quote(clean)
        if quote:
            quotes[clean] = quote
        time.sleep(0.2)
    return quotes


def build():
    df = pd.read_excel(SOURCE)

    df["Symbol_clean"] = df["Symbol"].astype(str).str.strip().str.upper()
    symbols = [s for s in df["Symbol_clean"].dropna().unique() if s and s != "NAN"]
    quotes = fetch_quotes(symbols)

    df["Buy"] = df["Buy price"].apply(parse_money)
    df["Units"] = pd.to_numeric(df["Units"], errors="coerce")
    df["Dec19"] = pd.to_numeric(df["Price As of Dec 19 2025"], errors="coerce")

    def current_from_quote(row):
        sym = row["Symbol_clean"]
        quote = quotes.get(sym)
        if quote:
            return quote["price"]
        return parse_money(row["Current Price"])

    def day_change_from_quote(row, field):
        sym = row["Symbol_clean"]
        quote = quotes.get(sym)
        if quote:
            return quote[field]
        return None

    df["Current"] = df.apply(current_from_quote, axis=1)
    df["DayChange"] = df.apply(lambda r: day_change_from_quote(r, "day_change"), axis=1)
    df["DayPct"] = df.apply(lambda r: day_change_from_quote(r, "day_pct"), axis=1)

    df["Pct_vs_buy"] = ((df["Current"] - df["Buy"]) / df["Buy"]) * 100
    df["Pct_vs_dec"] = ((df["Current"] - df["Dec19"]) / df["Dec19"]) * 100
    df["CostBasis"] = df["Units"] * df["Buy"]
    df["CurrentValue"] = df["Units"] * df["Current"]

    holdings = df[df["Units"].fillna(0) > 0].copy()
    holdings = holdings.sort_values("Pct_vs_buy", ascending=True)

    total_cost = holdings["CostBasis"].sum(min_count=1)
    total_value = holdings["CurrentValue"].sum(min_count=1)
    total_gain = total_value - total_cost
    pct_gain = (total_gain / total_cost * 100) if total_cost else 0

    display_cols = [
        "Stock Name",
        "Symbol",
        "Symbol_clean",
        "Units",
        "Buy",
        "Current",
        "CostBasis",
        "CurrentValue",
        "Pct_vs_buy",
        "Pct_vs_dec",
        "DayChange",
        "DayPct",
    ]

    table_df = holdings[display_cols].copy()

    summary_lines = []
    for _, row in table_df.iterrows():
        name = row["Stock Name"]
        sym = row["Symbol_clean"] or row["Symbol"]
        units = row["Units"] or 0
        now_val = money_format(row["Current"])
        delta_buy = pct_format(row["Pct_vs_buy"])
        delta_dec = pct_format(row["Pct_vs_dec"])
        day_move = money_format(row["DayChange"])
        day_pct = pct_format(row["DayPct"])
        summary_lines.append(
            f"- {name} ({sym}) — {units:.2f} sh · Now {now_val or 'n/a'} · Δ vs buy {delta_buy or 'n/a'} · Δ vs Dec19 {delta_dec or 'n/a'} · Today {day_move or 'n/a'} ({day_pct or 'n/a'})"
        )

    table_df = table_df.drop(columns=["Symbol_clean"])
    table_df["Units"] = table_df["Units"].map(lambda x: f"{x:.2f}")
    table_df["Buy"] = table_df["Buy"].map(money_format)
    table_df["Current"] = table_df["Current"].map(money_format)
    table_df["CostBasis"] = table_df["CostBasis"].map(money_format)
    table_df["CurrentValue"] = table_df["CurrentValue"].map(money_format)
    table_df["Pct_vs_buy"] = table_df["Pct_vs_buy"].map(pct_format)
    table_df["Pct_vs_dec"] = table_df["Pct_vs_dec"].map(pct_format)
    table_df["DayChange"] = table_df["DayChange"].map(lambda v: money_format(v) if v is not None else "")
    table_df["DayPct"] = table_df["DayPct"].map(pct_format)
    table_df = table_df.rename(
        columns={
            "Stock Name": "Name",
            "Pct_vs_buy": "Δ vs Buy",
            "Pct_vs_dec": "Δ vs Dec 19",
            "DayChange": "Δ Today",
            "DayPct": "Δ Today %",
        }
    )

    table_html = table_df.to_html(index=False, border=0, escape=False, na_rep="", classes="data-table")

    generated_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S ET")

    html = f"""<!doctype html>
<html lang=\"en\">
<head>
<meta charset=\"utf-8\" />
<title>Portfolio MVP</title>
<style>
body {{ font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; margin: 32px; background: #0b111b; color: #f1f5f9; }}
h1 {{ font-size: 28px; margin-bottom: 4px; }}
.summary {{ margin-bottom: 24px; font-size: 15px; color: #cbd5f5; }}
strong {{ color: #fef08a; }}
.data-table {{ width: 100%; border-collapse: collapse; font-size: 14px; }}
.data-table thead tr {{ background: #111b2f; text-align: left; }}
.data-table th, .data-table td {{ padding: 8px 10px; border-bottom: 1px solid rgba(148, 163, 184, 0.2); }}
.data-table tr:nth-child(even) {{ background: rgba(15, 23, 42, 0.65); }}
.badge-loss {{ color: #f87171; }}
.badge-gain {{ color: #34d399; }}
footer {{ margin-top: 24px; font-size: 12px; color: #94a3b8; }}
</style>
</head>
<body>
  <h1>Portfolio MVP</h1>
  <div class=\"summary\">
    Snapshot generated {generated_at}. Holdings shown: <strong>{len(table_df)}</strong>.
    Total cost basis: <strong>{money_format(total_cost)}</strong> · Current value: <strong>{money_format(total_value)}</strong>
    · Net: <strong>{money_format(total_gain)} ({pct_gain:+.1f}%)</strong>
  </div>
  {table_html}
  <footer>Data source: holdings workbook (sheet export) · Robinhood Jan statement loaded separately for reconciliation · Intraday quotes via Yahoo Finance.</footer>
</body>
</html>"""

    SUMMARY.write_text("\n".join(summary_lines), encoding="utf-8")
    OUTPUT.write_text(html, encoding="utf-8")


if __name__ == "__main__":
    build()
