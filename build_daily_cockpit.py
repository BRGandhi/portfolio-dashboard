import json
import math
import numbers
from datetime import datetime
from pathlib import Path
from textwrap import dedent

import pandas as pd
import yfinance as yf

SOURCE = Path(r"C:/Users/bgand/.openclaw/media/inbound/e5dc6078-2b7f-4d35-9ed3-8b4f2a89c10e.xlsx")
OUTPUT = Path("portfolio-control-room.html")
DIGEST = Path("portfolio-daily-digest.txt")
NOTES_PATH = Path("portfolio-notes.json")


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


def pct_format(value, decimals=1):
    if value is None or pd.isna(value):
        return "—"
    return f"{value:+.{decimals}f}%"


def money_format(value, decimals=2):
    if value is None or pd.isna(value):
        return "—"
    return f"${value:,.{decimals}f}"


def fetch_quote(symbol: str):
    try:
        data = yf.download(symbol, period="5d", interval="1d", progress=False, auto_adjust=False)
        data = data.dropna(subset=["Close"])
        if data.empty:
            return None
        last = data.iloc[-1]
        prev = data.iloc[-2] if len(data) > 1 else None
        price = float(last["Close"])
        prev_close = float(prev["Close"]) if prev is not None else None
        day_change = price - prev_close if prev_close else None
        day_pct = (day_change / prev_close * 100) if prev_close else None
        week_ago = data.iloc[0]["Close"] if len(data) >= 5 else None
        week_pct = ((price - week_ago) / week_ago * 100) if week_ago else None
        return {
            "price": price,
            "prev_close": prev_close,
            "day_change": day_change,
            "day_pct": day_pct,
            "week_pct": week_pct,
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
    return quotes


def classify(row):
    triggers = []
    day_pct = row.get("DayPct") or 0
    pct_buy = row.get("Pct_vs_buy") or 0
    pct_dec = row.get("Pct_vs_dec") or 0
    week_pct = row.get("WeekPct") or 0
    target = row.get("Target")
    current = row.get("Current")

    if day_pct <= -3:
        triggers.append("▼ Day drop >3%")
    if week_pct <= -10:
        triggers.append("▼ Week drop >10%")
    if pct_buy <= -15:
        triggers.append("▼ 15% under buy")
    if pct_dec <= -20:
        triggers.append("▼ 20% under Dec ref")
    if target and current and current >= target:
        triggers.append("🎯 Target reached")

    if not triggers:
        status = "stable"
    elif any(t.startswith("▼") for t in triggers):
        status = "attention"
    else:
        status = "action"

    return status, triggers


def load_notes():
    if NOTES_PATH.exists():
        try:
            return json.loads(NOTES_PATH.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            return {}
    return {}


def clean_for_json(value):
    if isinstance(value, dict):
        return {k: clean_for_json(v) for k, v in value.items()}
    if isinstance(value, list):
        return [clean_for_json(v) for v in value]
    if isinstance(value, (pd.Series, pd.Index)):
        return clean_for_json(value.tolist())
    if isinstance(value, pd.Timestamp):
        return value.isoformat()
    if isinstance(value, numbers.Number):
        if pd.isna(value) or math.isinf(value):
            return None
        return float(value)
    return value


def build():
    df = pd.read_excel(SOURCE)
    df["Symbol_clean"] = df["Symbol"].astype(str).str.strip().str.upper()
    symbols = [s for s in df["Symbol_clean"].dropna().unique() if s and s != "NAN"]
    quotes = fetch_quotes(symbols)

    df["Units"] = pd.to_numeric(df["Units"], errors="coerce")
    df["Buy"] = df["Buy price"].apply(parse_money)
    df["Sell"] = df["Sell price"].apply(parse_money)
    df["CurrentExcel"] = df["Current Price"].apply(parse_money)
    df["DecRef"] = pd.to_numeric(df["Price As of Dec 19 2025"], errors="coerce")
    df["Target"] = pd.to_numeric(df["Target (20%)"], errors="coerce")

    def current_from_quote(row):
        sym = row["Symbol_clean"]
        quote = quotes.get(sym)
        if quote:
            return quote["price"]
        return row["CurrentExcel"]

    def metric_from_quote(row, field):
        sym = row["Symbol_clean"]
        quote = quotes.get(sym)
        if quote:
            return quote.get(field)
        return None

    df["Current"] = df.apply(current_from_quote, axis=1)
    df["DayChange"] = df.apply(lambda r: metric_from_quote(r, "day_change"), axis=1)
    df["DayPct"] = df.apply(lambda r: metric_from_quote(r, "day_pct"), axis=1)
    df["WeekPct"] = df.apply(lambda r: metric_from_quote(r, "week_pct"), axis=1)

    df["Pct_vs_buy"] = ((df["Current"] - df["Buy"]) / df["Buy"]) * 100
    df["Pct_vs_dec"] = ((df["Current"] - df["DecRef"]) / df["DecRef"]) * 100
    df["CostBasis"] = df["Units"] * df["Buy"]
    df["CurrentValue"] = df["Units"] * df["Current"]
    df["DayDollar"] = df["Units"] * df["DayChange"]

    holdings = df[df["Units"].fillna(0) > 0].copy()
    holdings = holdings.sort_values("Pct_vs_buy", ascending=True)

    total_cost = holdings["CostBasis"].sum(min_count=1) or 0
    total_value = holdings["CurrentValue"].sum(min_count=1) or 0
    day_move = holdings["DayDollar"].sum(min_count=1)
    if pd.isna(day_move):
        day_move = 0
    pct_gain = ((total_value - total_cost) / total_cost * 100) if total_cost else 0

    notes = load_notes()
    records = []
    attention = []

    for _, row in holdings.iterrows():
        status, triggers = classify(row)
        symbol = row["Symbol_clean"]
        record = {
            "name": row["Stock Name"],
            "symbol": symbol,
            "units": float(row["Units"] or 0),
            "buy": row["Buy"],
            "current": row["Current"],
            "costBasis": row["CostBasis"],
            "currentValue": row["CurrentValue"],
            "pctVsBuy": row["Pct_vs_buy"],
            "pctVsDec": row["Pct_vs_dec"],
            "dayChange": row["DayChange"],
            "dayPct": row["DayPct"],
            "weekPct": row["WeekPct"],
            "target": row["Target"],
            "status": status,
            "triggers": triggers,
            "note": notes.get(symbol, ""),
            "dayDollar": row["DayDollar"],
        }
        if triggers:
            attention.append(record)
        records.append(record)

    attention_sorted = sorted(attention, key=lambda r: (r["dayPct"] or -999))
    top_losers = sorted(records, key=lambda r: (r["dayPct"] or 0))[:5]
    top_winners = sorted(records, key=lambda r: (r["dayPct"] or 0), reverse=True)[:5]

    digest_lines = [
        f"Snapshot {datetime.now().strftime('%b %d %H:%M')} — Net {money_format(total_value - total_cost)} ({pct_format(pct_gain)})",
        f"Market value {money_format(total_value)} · Cost basis {money_format(total_cost)} · Today {money_format(day_move)}",
        "--- Movers ---",
    ]
    for rec in top_losers + top_winners:
        digest_lines.append(
            f"{rec['symbol']}: {pct_format(rec['dayPct'])} today · vs buy {pct_format(rec['pctVsBuy'])} · note {rec['note'] or '—'}"
        )
    DIGEST.write_text("\n".join(digest_lines), encoding="utf-8")

    generated_at = datetime.now().strftime("%Y-%m-%d %H:%M")
    top_symbol = records[-1]["symbol"] if records else "—"

    records_json = json.dumps(clean_for_json(records))
    attention_json = json.dumps(clean_for_json(attention_sorted[:10]))
    winners_json = json.dumps(clean_for_json(top_winners))
    losers_json = json.dumps(clean_for_json(top_losers))

    template = dedent(
        """<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8" />
<title>Portfolio Control Room</title>
<meta name="viewport" content="width=device-width, initial-scale=1" />
<style>
:root {
  font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
  color-scheme: dark;
  --bg: #030712;
  --card: #111c36;
  --text: #e2e8f0;
  --muted: #94a3b8;
  --accent: #38bdf8;
}
body { margin: 0; padding: 32px; background: radial-gradient(circle at top, rgba(56,189,248,.08), transparent 60%), var(--bg); color: var(--text); min-height: 100vh; }
main { max-width: 1200px; margin: 0 auto; }
header { margin-bottom: 24px; }
h1 { font-size: 28px; margin: 0; }
summary { font-size: 15px; color: var(--muted); }
.grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 16px; margin-bottom: 24px; }
.card { background: var(--card); border: 1px solid rgba(148,163,184,.2); border-radius: 16px; padding: 16px; box-shadow: 0 20px 60px rgba(15,23,42,.45); }
.card h2 { margin: 0; font-size: 13px; text-transform: uppercase; letter-spacing: 0.08em; color: var(--muted); }
.card strong { display: block; font-size: 24px; margin-top: 4px; }
.card small { color: var(--muted); }
section { margin-bottom: 32px; }
section h3 { margin-bottom: 12px; font-size: 18px; }
.list { display: grid; gap: 8px; }
.list-item { padding: 12px; border-radius: 12px; background: linear-gradient(120deg, rgba(248, 113, 113, .2), rgba(30, 41, 59, .9)); border: 1px solid rgba(248,113,113,.4); cursor: pointer; }
.list-item.ok { background: rgba(15,23,42,.8); border-color: rgba(148,163,184,.2); cursor: default; }
.list-item h4 { margin: 0 0 4px; font-size: 14px; }
.list-item span { font-size: 13px; color: var(--muted); }
.table-wrap { overflow-x: auto; border-radius: 18px; border: 1px solid rgba(148,163,184,.2); }
table { width: 100%; border-collapse: collapse; font-size: 13px; min-width: 780px; }
thead { background: rgba(15,23,42,.7); }
th, td { text-align: left; padding: 12px 14px; border-bottom: 1px solid rgba(148,163,184,.15); }
tr:hover { background: rgba(56,189,248,.08); }
.tag { font-size: 11px; padding: 2px 8px; border-radius: 999px; text-transform: uppercase; letter-spacing: .06em; }
.tag.attention { background: rgba(251,113,133,.2); color: #fecdd3; border: 1px solid rgba(251,113,133,.35); }
.tag.stable { background: rgba(52,211,153,.15); color: #bbf7d0; border: 1px solid rgba(52,211,153,.3); }
.tag.action { background: rgba(59,130,246,.2); color: #bfdbfe; border: 1px solid rgba(59,130,246,.35); }
.controls { display: flex; flex-wrap: wrap; gap: 12px; margin-bottom: 12px; }
.controls input, .controls select { flex: 1; min-width: 220px; padding: 10px 14px; border-radius: 999px; border: 1px solid rgba(148,163,184,.3); background: rgba(15,23,42,.6); color: var(--text); }
.detail-modal { position: fixed; inset: 0; background: rgba(2,6,23,.85); display: none; align-items: center; justify-content: center; z-index: 50; }
.detail-card { background: #050e1f; border-radius: 20px; padding: 24px; border: 1px solid rgba(148,163,184,.3); width: min(720px, 90vw); max-height: 90vh; overflow-y: auto; }
@media (max-width: 640px) { body { padding: 20px; } }
</style>
</head>
<body>
<main>
  <header>
    <h1>Portfolio Control Room</h1>
    <summary>Generated __GENERATED_AT__ · Holdings __HOLDING_COUNT__ · Cost basis __COST_BASIS__ · Market value __MARKET_VALUE__ · Today __DAY_MOVE__ · Net __NET_PCT__</summary>
  </header>

  <div class="grid">
    <div class="card"><h2>Market Value</h2><strong>__MARKET_VALUE__</strong><small>Δ vs buy __NET_PCT__</small></div>
    <div class="card"><h2>Cost Basis</h2><strong>__COST_BASIS__</strong><small>Net __NET_VALUE__</small></div>
    <div class="card"><h2>Today (est)</h2><strong>__DAY_MOVE__</strong><small>Top mover __TOP_SYMBOL__</small></div>
    <div class="card"><h2>Flags Active</h2><strong>__FLAGS__</strong><small>Auto classified moves</small></div>
  </div>

  <section>
    <h3>Attention Queue</h3>
    <div class="list" id="attention-list"></div>
  </section>

  <section>
    <h3>Positions</h3>
    <div class="controls">
      <input id="search" placeholder="Filter by ticker or name..." />
      <select id="filter">
        <option value="all">All statuses</option>
        <option value="attention">Attention</option>
        <option value="action">Target hit</option>
        <option value="stable">Stable</option>
      </select>
    </div>
    <div class="table-wrap">
      <table>
        <thead>
          <tr>
            <th>Ticker</th>
            <th>Name</th>
            <th>Units</th>
            <th>Buy</th>
            <th>Now</th>
            <th>Day %</th>
            <th>Vs Buy</th>
            <th>Vs Dec19</th>
            <th>Target</th>
            <th>Value</th>
            <th>Status</th>
          </tr>
        </thead>
        <tbody id="table-body"></tbody>
      </table>
    </div>
  </section>

  <section>
    <h3>Movers</h3>
    <div class="grid">
      <div class="card">
        <h2>Winners</h2>
        <div id="winners"></div>
      </div>
      <div class="card">
        <h2>Losers</h2>
        <div id="losers"></div>
      </div>
    </div>
  </section>
</main>

<div class="detail-modal" id="detail-modal">
  <div class="detail-card">
    <header>
      <div>
        <div id="detail-title" style="font-size:20px;font-weight:600;"></div>
        <div id="detail-sub" style="font-size:13px;color:var(--muted);"></div>
      </div>
      <button id="detail-close">&times;</button>
    </header>
    <div id="detail-body"></div>
  </div>
</div>

<script id="records" type="application/json">__RECORDS_JSON__</script>
<script id="attention" type="application/json">__ATTENTION_JSON__</script>
<script id="winners-data" type="application/json">__WINNERS_JSON__</script>
<script id="losers-data" type="application/json">__LOSERS_JSON__</script>
<script>
const records = JSON.parse(document.getElementById('records').textContent);
const attention = JSON.parse(document.getElementById('attention').textContent);
const winners = JSON.parse(document.getElementById('winners-data').textContent);
const losers = JSON.parse(document.getElementById('losers-data').textContent);
const bodyEl = document.getElementById('table-body');
let filtered = [...records];

const fmtMoney = (v) => v == null ? '—' : '$' + v.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const fmtPct = (v) => v == null ? '—' : (v > 0 ? '+' : '') + v.toFixed(1) + '%';

function renderTable() {
  bodyEl.innerHTML = '';
  filtered.forEach(row => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${row.symbol}</td>
      <td>${row.name}</td>
      <td>${row.units.toFixed(2)}</td>
      <td>${fmtMoney(row.buy)}</td>
      <td>${fmtMoney(row.current)}</td>
      <td>${fmtPct(row.dayPct)}</td>
      <td>${fmtPct(row.pctVsBuy)}</td>
      <td>${fmtPct(row.pctVsDec)}</td>
      <td>${fmtMoney(row.target)}</td>
      <td>${fmtMoney(row.currentValue)}</td>
      <td><span class="tag ${row.status}">${row.status}</span></td>`;
    tr.addEventListener('click', () => openDetail(row));
    bodyEl.appendChild(tr);
  });
}

function renderLists() {
  const attentionList = document.getElementById('attention-list');
  attentionList.innerHTML = '';
  if (!attention.length) {
    attentionList.innerHTML = '<div class="list-item ok"><h4>All clear</h4><span>No triggers right now.</span></div>';
    return;
  }
  attention.forEach(item => {
    const div = document.createElement('div');
    div.className = 'list-item';
    div.innerHTML = `<h4>${item.symbol} · ${fmtPct(item.dayPct)} today</h4><span>${item.triggers.join(' · ')}</span>`;
    div.addEventListener('click', () => openDetail(item));
    attentionList.appendChild(div);
  });
}

function renderMovers(list, elId) {
  const wrap = document.getElementById(elId);
  wrap.innerHTML = '';
  if (!list.length) {
    wrap.innerHTML = '<p style="color:var(--muted);">No data yet.</p>';
    return;
  }
  list.forEach(item => {
    const p = document.createElement('p');
    p.style.margin = '6px 0';
    p.innerHTML = `<strong>${item.symbol}</strong> ${fmtPct(item.dayPct)} · vs buy ${fmtPct(item.pctVsBuy)}`;
    p.addEventListener('click', () => openDetail(item));
    wrap.appendChild(p);
  });
}

function openDetail(row) {
  const modal = document.getElementById('detail-modal');
  document.getElementById('detail-title').textContent = `${row.symbol} · ${row.name}`;
  document.getElementById('detail-sub').textContent = `${row.units.toFixed(2)} units · ${fmtMoney(row.currentValue)} current value`;
  const stored = localStorage.getItem(`note-${row.symbol}`) || row.note || '';
  document.getElementById('detail-body').innerHTML = `
    <div style="display:flex;gap:12px;flex-wrap:wrap;margin-bottom:16px;">
      <span class="tag ${row.status}">${row.status}</span>
      <span class="tag">Day ${fmtPct(row.dayPct)}</span>
      <span class="tag">Vs buy ${fmtPct(row.pctVsBuy)}</span>
      <span class="tag">Vs Dec ${fmtPct(row.pctVsDec)}</span>
      <span class="tag">Target ${fmtMoney(row.target)}</span>
    </div>
    <label style="font-size:13px;color:var(--muted);display:block;margin-bottom:6px;">Scratchpad (stored locally)</label>
    <textarea id="note-field" style="width:100%;min-height:120px;border-radius:12px;border:1px solid rgba(148,163,184,.3);background:#020617;color:var(--text);padding:12px;">${stored}</textarea>
    <button id="save-note" style="margin-top:12px;padding:10px 16px;border-radius:999px;border:0;background:var(--accent);color:#04101f;font-weight:600;cursor:pointer;">Save note</button>
    <iframe style="width:100%;height:360px;border:0;border-radius:16px;margin-top:12px;" loading="lazy" src="https://s.tradingview.com/widgetembed/?symbol=${row.symbol}&interval=60&theme=dark"></iframe>`;
  document.getElementById('save-note').addEventListener('click', () => {
    const val = document.getElementById('note-field').value;
    localStorage.setItem(`note-${row.symbol}`, val);
  });
  modal.style.display = 'flex';
}

document.getElementById('detail-close').addEventListener('click', () => {
  document.getElementById('detail-modal').style.display = 'none';
});
document.getElementById('detail-modal').addEventListener('click', (e) => {
  if (e.target.id === 'detail-modal') {
    e.currentTarget.style.display = 'none';
  }
});

document.getElementById('search').addEventListener('input', (e) => {
  const term = e.target.value.toLowerCase();
  filtered = records.filter(row => row.symbol.toLowerCase().includes(term) || row.name.toLowerCase().includes(term));
  applyFilter(document.getElementById('filter').value, true);
});

document.getElementById('filter').addEventListener('change', (e) => {
  applyFilter(e.target.value);
});

function applyFilter(status, skipSource) {
  let list = records;
  if (!skipSource) {
    const term = document.getElementById('search').value.toLowerCase();
    list = records.filter(row => row.symbol.toLowerCase().includes(term) || row.name.toLowerCase().includes(term));
  } else {
    list = filtered;
  }
  if (status !== 'all') list = list.filter(row => row.status === status);
  filtered = list;
  renderTable();
}

renderTable();
renderLists();
renderMovers(winners, 'winners');
renderMovers(losers, 'losers');
</script>
</body>
</html>"""
    )

    html = (
        template
        .replace("__GENERATED_AT__", generated_at)
        .replace("__HOLDING_COUNT__", str(len(records)))
        .replace("__COST_BASIS__", money_format(total_cost))
        .replace("__MARKET_VALUE__", money_format(total_value))
        .replace("__DAY_MOVE__", money_format(day_move))
        .replace("__NET_PCT__", pct_format(pct_gain))
        .replace("__NET_VALUE__", money_format(total_value - total_cost))
        .replace("__TOP_SYMBOL__", top_symbol)
        .replace("__FLAGS__", str(len(attention_sorted)))
        .replace("__RECORDS_JSON__", records_json)
        .replace("__ATTENTION_JSON__", attention_json)
        .replace("__WINNERS_JSON__", winners_json)
        .replace("__LOSERS_JSON__", losers_json)
    )

    OUTPUT.write_text(html, encoding="utf-8")


if __name__ == "__main__":
    build()
