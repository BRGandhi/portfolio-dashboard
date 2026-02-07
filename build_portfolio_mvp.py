import json
import math
import time
from datetime import datetime
from pathlib import Path
from textwrap import dedent

import pandas as pd
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

    view_cols = [
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
    view_df = holdings[view_cols].copy()

    interactive_records = []
    for _, row in view_df.iterrows():
        sym_clean = row["Symbol_clean"]
        interactive_records.append(
            {
                "name": row["Stock Name"],
                "symbol": sym_clean,
                "units": round(float(row["Units"]), 4) if not math.isnan(row["Units"]) else 0,
                "buy": row["Buy"],
                "current": row["Current"],
                "costBasis": row["CostBasis"],
                "currentValue": row["CurrentValue"],
                "pctVsBuy": row["Pct_vs_buy"],
                "pctVsDec": row["Pct_vs_dec"],
                "dayChange": row["DayChange"],
                "dayPct": row["DayPct"],
            }
        )

    summary_lines = []
    for rec in interactive_records:
        summary_lines.append(
            f"- {rec['name']} ({rec['symbol']}) — {rec['units']:.2f} sh · Now {money_format(rec['current']) or 'n/a'} · "
            f"Δ vs buy {pct_format(rec['pctVsBuy']) or 'n/a'} · Δ vs Dec19 {pct_format(rec['pctVsDec']) or 'n/a'} · "
            f"Today {money_format(rec['dayChange']) or 'n/a'} ({pct_format(rec['dayPct']) or 'n/a'})"
        )

    SUMMARY.write_text("\n".join(summary_lines), encoding="utf-8")

    html = dedent(
        """<!doctype html>
<html lang=\"en\">
<head>
<meta charset=\"utf-8\" />
<title>Portfolio Control Room</title>
<style>
body {{ font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; margin: 24px; background: #050b16; color: #f8fafc; }}
h1 {{ font-size: 28px; margin-bottom: 6px; }}
.summary {{ margin-bottom: 16px; font-size: 15px; color: #cbd5f5; }}
strong {{ color: #fef08a; }}
#controls {{ display: flex; flex-wrap: wrap; gap: 16px; margin-bottom: 16px; }}
#controls input {{ flex: 1; padding: 8px 12px; border-radius: 6px; border: 1px solid rgba(148,163,184,.3); background: #0f172a; color: #f8fafc; }}
button.sort-btn {{ padding: 8px 12px; border-radius: 6px; border: 1px solid rgba(148,163,184,.3); background: #111b2f; color: #f8fafc; cursor: pointer; }}
button.sort-btn.active {{ border-color: #38bdf8; color: #38bdf8; }}
table {{ width: 100%; border-collapse: collapse; font-size: 13px; }}
thead {{ background: #0f172a; }}
th, td {{ padding: 8px 10px; border-bottom: 1px solid rgba(148, 163, 184, 0.2); text-align: left; }}
tr:nth-child(even) {{ background: rgba(15, 23, 42, 0.65); }}
tr:hover {{ background: rgba(59,130,246,.15); cursor: pointer; }}
.badge-profit {{ color: #34d399; font-weight: 600; }}
.badge-loss {{ color: #f87171; font-weight: 600; }}
#detail-modal {{ position: fixed; inset: 0; background: rgba(5,11,22,.9); display: none; align-items: center; justify-content: center; padding: 32px; z-index: 1000; }}
#detail-card {{ width: min(900px, 90vw); background: #0b111f; border-radius: 16px; padding: 24px; border: 1px solid rgba(56,189,248,.2); }}
#detail-card header {{ display: flex; justify-content: space-between; align-items: center; margin-bottom: 16px; }}
#detail-card button {{ background: transparent; border: 0; color: #94a3b8; font-size: 24px; cursor: pointer; }}
#alert {{ margin-top: 12px; font-size: 13px; color: #fbbf24; }}
footer {{ margin-top: 24px; font-size: 12px; color: #94a3b8; }}
</style>
</head>
<body>
  <h1>Portfolio Control Room</h1>
  <div class=\"summary\">
    Snapshot generated {generated_at}. Holdings shown: <strong>{holding_count}</strong>. Cost basis <strong>{total_cost}</strong> · Market value <strong>{total_value}</strong> · Net <strong>{total_gain} ({pct_gain:+.1f}%)</strong>
  </div>
  <section id=\"controls\">
    <input id=\"search\" type=\"search\" placeholder=\"Filter by ticker, name, or notes...\" />
    <button class=\"sort-btn\" data-field=\"pctVsBuy\">Sort Δ vs Buy</button>
    <button class=\"sort-btn\" data-field=\"pctVsDec\">Sort Δ vs Dec</button>
    <button class=\"sort-btn\" data-field=\"dayPct\">Sort Today %</button>
  </section>
  <table>
    <thead>
      <tr>
        <th>Ticker</th>
        <th>Name</th>
        <th>Units</th>
        <th>Buy</th>
        <th>Now</th>
        <th>Δ vs Buy</th>
        <th>Δ vs Dec19</th>
        <th>Today</th>
        <th>Today %</th>
        <th>Value</th>
      </tr>
    </thead>
    <tbody id=\"positions-body\"></tbody>
  </table>
  <div id=\"alert\"></div>
  <div id=\"detail-modal\">
    <div id=\"detail-card\">
      <header>
        <div>
          <div id=\"detail-title\" style=\"font-size: 18px; font-weight: 600;\"></div>
          <div id=\"detail-meta\" style=\"font-size: 13px; color: #94a3b8;\"></div>
        </div>
        <button id=\"detail-close\">&times;</button>
      </header>
      <div id=\"detail-body\"></div>
    </div>
  </div>
  <footer>Data source: holdings workbook (sheet export) · Robinhood Jan statement · Intraday quotes via Yahoo Finance API. Dashboard refreshes quotes automatically every 60 seconds.</footer>
  <script id=\"portfolio-data\" type=\"application/json\">{data_json}</script>
  <script>
  const tableBody = document.getElementById('positions-body');
  const data = JSON.parse(document.getElementById('portfolio-data').textContent);
  let filtered = [...data];
  let currentSort = null;

  const fmtMoney = (v) => v == null ? '—' : '$' + v.toLocaleString(undefined, {{ minimumFractionDigits: 2, maximumFractionDigits: 2 }});
  const fmtPct = (v) => v == null ? '—' : (v > 0 ? '+' : '') + v.toFixed(1) + '%';

  function renderTable() {{
    tableBody.innerHTML = '';
    filtered.forEach(row => {{
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td>${{row.symbol}}</td>
        <td>${{row.name}}</td>
        <td>${{row.units.toFixed(2)}}</td>
        <td>${{fmtMoney(row.buy)}}</td>
        <td>${{fmtMoney(row.current)}}</td>
        <td class="${{row.pctVsBuy >= 0 ? 'badge-profit' : 'badge-loss'}}">${{fmtPct(row.pctVsBuy)}}</td>
        <td class="${{row.pctVsDec >= 0 ? 'badge-profit' : 'badge-loss'}}">${{fmtPct(row.pctVsDec)}}</td>
        <td>${{fmtMoney(row.dayChange)}}</td>
        <td class="${{row.dayPct >= 0 ? 'badge-profit' : 'badge-loss'}}">${{fmtPct(row.dayPct)}}</td>
        <td>${{fmtMoney(row.currentValue)}}</td>`;
      tr.addEventListener('click', () => openDetail(row));
      tableBody.appendChild(tr);
    }});
  }}

  function applyFilter(term) {{
    const needle = term.toLowerCase();
    filtered = data.filter(row => row.symbol.toLowerCase().includes(needle) || row.name.toLowerCase().includes(needle));
    applySort(currentSort?.field, true);
  }}

  function applySort(field, skipToggle) {{
    if (!field) {{ renderTable(); return; }}
    if (!skipToggle) {{
      if (currentSort?.field === field) {{
        currentSort.dir = currentSort.dir === 'desc' ? 'asc' : 'desc';
      }} else {{
        currentSort = {{ field, dir: 'desc' }};
      }}
    }} else if (!currentSort) {{
      currentSort = {{ field, dir: 'desc' }};
    }}
    const dir = currentSort.dir === 'desc' ? -1 : 1;
    filtered.sort((a, b) => {{
      const av = a[field] ?? -Infinity;
      const bv = b[field] ?? -Infinity;
      return av > bv ? dir : av < bv ? -dir : 0;
    }});
    renderTable();
    document.querySelectorAll('.sort-btn').forEach(btn => btn.classList.toggle('active', btn.dataset.field === field));
  }}

  function openDetail(row) {{
    document.getElementById('detail-title').textContent = `${{row.symbol}} · ${{row.name}}`;
    document.getElementById('detail-meta').textContent = `${{row.units.toFixed(2)}} units · ${{fmtMoney(row.currentValue)}} current value`;
    const tvSymbol = encodeURIComponent(row.symbol);
    document.getElementById('detail-body').innerHTML = `
      <div style="margin-bottom:12px;font-size:13px;color:#94a3b8;">
        Δ vs buy <span class="${{row.pctVsBuy >=0 ? 'badge-profit' : 'badge-loss'}}">${{fmtPct(row.pctVsBuy)}}</span> ·
        Today <span class="${{row.dayPct >=0 ? 'badge-profit' : 'badge-loss'}}">${{fmtMoney(row.dayChange)}} (${{fmtPct(row.dayPct)}})</span>
      </div>
      <iframe style="width:100%;height:420px;border:0;border-radius:12px;background:#0b111f;"
        src="https://s.tradingview.com/widgetembed/?frameElementId=tradingview_${{tvSymbol}}&symbol=${{tvSymbol}}&interval=60&hidetoptoolbar=1&symboledit=1&saveimage=1&toolbarbg=f1f3f6&studies=[]&theme=dark"
        loading="lazy"></iframe>`;
    document.getElementById('detail-modal').style.display = 'flex';
  }}

  document.getElementById('detail-close').addEventListener('click', () => {{
    document.getElementById('detail-modal').style.display = 'none';
  }});
  document.getElementById('detail-modal').addEventListener('click', (e) => {{
    if (e.target.id === 'detail-modal') {{
      e.currentTarget.style.display = 'none';
    }}
  }});

  document.getElementById('search').addEventListener('input', (e) => applyFilter(e.target.value));
  document.querySelectorAll('.sort-btn').forEach(btn => {{
    btn.addEventListener('click', () => applySort(btn.dataset.field));
  }});

  function chunk(array, size) {{
    const chunks = [];
    for (let i = 0; i < array.length; i += size) chunks.push(array.slice(i, i + size));
    return chunks;
  }}

  async function refreshLiveQuotes() {{
    const alertBar = document.getElementById('alert');
    try {{
      const symbols = data.map(d => d.symbol).filter(Boolean);
      const groups = chunk(symbols, 20);
      for (const group of groups) {{
        const query = group.join(',');
        const resp = await fetch(`https://query1.finance.yahoo.com/v7/finance/quote?symbols=${{query}}`);
        if (!resp.ok) continue;
        const payload = await resp.json();
        payload.quoteResponse.result.forEach(item => {{
          const sym = item.symbol?.toUpperCase();
          const record = data.find(d => d.symbol === sym);
          if (!record) return;
          const price = item.regularMarketPrice;
          if (price != null) {{
            record.current = price;
            record.currentValue = price * record.units;
          }}
          if (item.previousClose) {{
            record.dayChange = price - item.previousClose;
            record.dayPct = ((record.dayChange) / item.previousClose) * 100;
          }}
          if (record.buy) {{
            record.pctVsBuy = ((record.current - record.buy) / record.buy) * 100;
          }}
        }});
      }}
      applySort(currentSort?.field || 'pctVsBuy', true);
      alertBar.textContent = `Quotes refreshed ${{new Date().toLocaleTimeString()}}`;
    }} catch (err) {{
      alertBar.textContent = 'Live quote refresh failed; using last known values.';
    }}
  }}

  renderTable();
  applySort('pctVsBuy', true);
  refreshLiveQuotes();
  setInterval(refreshLiveQuotes, 60000);
  </script>
</body>
</html>"""
    ).format(
        generated_at=datetime.now().strftime("%Y-%m-%d %H:%M:%S ET"),
        holding_count=len(view_df),
        total_cost=money_format(total_cost),
        total_value=money_format(total_value),
        total_gain=money_format(total_gain),
        pct_gain=pct_gain,
        data_json=json.dumps(interactive_records)
    )

    OUTPUT.write_text(html, encoding="utf-8")


if __name__ == "__main__":
    build()
