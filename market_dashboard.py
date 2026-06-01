"""
Market volume dashboard.

Pulls settled and open market data, maps to categories via series lookup,
and exports an Excel dashboard showing volume by category over time.
Also includes a trade P&L sheet built from trades.csv.
"""

import sys
import os
import csv
from collections import defaultdict
from datetime import datetime, timezone

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, LineChart, Reference

from pykalshi import KalshiClient, MarketStatus

TRADES_CSV = os.path.join(os.path.dirname(__file__), "trades.csv")

EXCEL_OUT = os.path.join(os.path.dirname(__file__), "market_dashboard.xlsx")
SETTLED_PAGES = 10        # pages of 200 = up to 2,000 settled markets
MIN_VOLUME    = 1.0       # drop truly zero-volume markets


# ---------------------------------------------------------------------------
# Data collection
# ---------------------------------------------------------------------------

def build_series_category_map(client) -> dict:
    """Return a lazy-populating cache: look up series only as needed."""
    # Seed with common ones so most markets need no API call
    seed = {
        "KXNBA": "Sports", "KXNFL": "Sports", "KXMLB": "Sports",
        "KXNHL": "Sports", "KXNHLOVERTIME": "Sports", "KXMMA": "Sports",
        "KXNASCAR": "Sports", "KXPGA": "Sports", "KXTENNIS": "Sports",
        "KXBTC": "Crypto", "KXETH": "Crypto", "KXXRPD": "Crypto",
        "KXSOLD": "Crypto", "KXAVAXD": "Crypto", "KXSUID": "Crypto",
        "KXCPI": "Economics", "KXFOMC": "Economics", "KXPCE": "Economics",
        "KXJOBS": "Economics", "KXGDP": "Economics", "KXUNRATE": "Economics",
        "KXSPX": "Financials", "KXNASDAQ": "Financials", "KXDOW": "Financials",
        "KXPOTUS": "Politics", "KXSENATE": "Politics", "KXHOUSE": "Politics",
        "KXOSCARS": "Entertainment", "KXEMMYS": "Entertainment",
    }

    class LazyMap(dict):
        def __missing__(self, key):
            try:
                s = client.get_series(key)
                val = s.category or "Unknown"
            except Exception:
                val = "Unknown"
            self[key] = val
            return val

    m = LazyMap(seed)
    print(f"Series map initialised ({len(seed)} pre-seeded, rest looked up on demand).")
    return m


def infer_series(ticker: str) -> str:
    """Extract series prefix from a market ticker (e.g. KXBTC from KXBTC-26MAY15-T88799)."""
    return ticker.split("-")[0]


def pull_settled(client, series_map: dict, pages: int = SETTLED_PAGES) -> list[dict]:
    """Pull settled markets across multiple pages and return as list of dicts."""
    print(f"Fetching settled markets ({pages} pages x 200)...")
    rows = []
    cursor = None

    for page in range(pages):
        kwargs = {"status": MarketStatus.SETTLED, "limit": 200}
        if cursor:
            kwargs["cursor"] = cursor

        batch = client.get_markets(**kwargs)
        if not batch:
            break

        for m in batch:
            vol = float(m.volume_fp or 0)
            if vol < MIN_VOLUME:
                continue
            if m.ticker.startswith("KXMVE"):
                continue

            series = m.series_ticker or infer_series(m.ticker)
            category = series_map.get(series, "Unknown")
            close_dt = _parse_dt(m.close_time)

            rows.append({
                "ticker":   m.ticker,
                "series":   series,
                "category": category,
                "volume":   vol,
                "result":   (m.result or "").lower(),
                "close_month": close_dt.strftime("%Y-%m") if close_dt else "Unknown",
                "close_date":  close_dt.date() if close_dt else None,
            })

        print(f"  Page {page + 1}/{pages}: {len(batch)} markets, {len(rows)} kept so far")

        if len(batch) < 200:
            break  # last page

    print(f"  Total settled markets with volume: {len(rows)}")
    return rows


def pull_open_24h(client, series_map: dict) -> list[dict]:
    """Pull open and recently closed markets, return 24h volume by category."""
    rows = []
    for status, label in [(MarketStatus.OPEN, "open"), (MarketStatus.CLOSED, "closed")]:
        print(f"Fetching {label} markets for 24h volume snapshot...")
        try:
            markets = client.get_markets(status=status, limit=200, fetch_all=True)
        except Exception as e:
            print(f"  Warning: could not fetch {label} markets: {e}")
            continue
        count = 0
        for m in markets:
            try:
                if m.ticker.startswith("KXMVE"):
                    continue
                vol_24h = float(m.volume_24h_fp or 0)
                if vol_24h < MIN_VOLUME:
                    continue
                series = m.series_ticker or infer_series(m.ticker)
                category = series_map.get(series, "Unknown")
                rows.append({
                    "ticker":     m.ticker,
                    "series":     series,
                    "category":   category,
                    "volume_24h": vol_24h,
                    "total_vol":  float(m.volume_fp or 0),
                })
                count += 1
            except Exception:
                continue
        print(f"  {count} {label} markets with 24h volume > 0 (excl. MVE)")
    return rows


def _parse_dt(s: str | None) -> datetime | None:
    if not s:
        return None
    try:
        return datetime.fromisoformat(s.replace("Z", "+00:00"))
    except Exception:
        return None


# ---------------------------------------------------------------------------
# Analysis
# ---------------------------------------------------------------------------

def category_totals(settled_rows: list[dict]) -> pd.DataFrame:
    df = pd.DataFrame(settled_rows)
    summary = (
        df.groupby("category")
          .agg(markets=("ticker", "count"), total_volume=("volume", "sum"))
          .reset_index()
          .sort_values("total_volume", ascending=False)
          .reset_index(drop=True)
    )
    summary.columns = ["Category", "Markets", "Total Volume"]
    summary["Total Volume"] = summary["Total Volume"].round(0).astype(int)
    return summary


def volume_over_time(settled_rows: list[dict]) -> pd.DataFrame:
    df = pd.DataFrame(settled_rows)
    pivot = (
        df.groupby(["close_month", "category"])["volume"]
          .sum()
          .reset_index()
          .pivot(index="close_month", columns="category", values="volume")
          .fillna(0)
          .astype(int)
          .reset_index()
          .rename(columns={"close_month": "Month"})
    )
    pivot = pivot.sort_values("Month").reset_index(drop=True)
    return pivot


def active_now(open_rows: list[dict]) -> pd.DataFrame:
    df = pd.DataFrame(open_rows)
    summary = (
        df.groupby("category")
          .agg(
              open_markets=("ticker", "count"),
              volume_24h=("volume_24h", "sum"),
              total_volume=("total_vol", "sum"),
          )
          .reset_index()
          .sort_values("volume_24h", ascending=False)
          .reset_index(drop=True)
    )
    summary.columns = ["Category", "Open Markets", "24h Volume", "Total Volume"]
    for col in ["24h Volume", "Total Volume"]:
        summary[col] = summary[col].round(0).astype(int)
    return summary


# ---------------------------------------------------------------------------
# Trade P&L
# ---------------------------------------------------------------------------

def _infer_category(ticker: str) -> str:
    t = ticker.upper()
    if any(x in t for x in ["ATP", "ITF", "WTA", "CRICKET", "NBA", "NFL", "MLB", "NHL", "TENNIS", "SOCCER"]):
        return "Sports"
    if any(x in t for x in ["TRUMP", "VANCE", "HEARING", "POTUS"]):
        return "Politics"
    if any(x in t for x in ["EARNINGS", "CPI", "FOMC", "PCE", "GDP"]):
        return "Economics"
    if any(x in t for x in ["MRBEAST", "MUSIC", "STREAM"]):
        return "Entertainment"
    return "Other"


def build_trade_pnl(client) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Return (settled_df, open_df) from trades.csv with P&L computed."""
    if not os.path.exists(TRADES_CSV):
        return pd.DataFrame(), pd.DataFrame()

    fills = list(csv.DictReader(open(TRADES_CSV, newline="")))
    if not fills:
        return pd.DataFrame(), pd.DataFrame()

    # Group fills by ticker
    by_ticker = defaultdict(list)
    for r in fills:
        by_ticker[r["ticker"]].append(r)

    # Fetch market status for all unique tickers
    print(f"Fetching market status for {len(by_ticker)} tickers...")
    results = {}
    statuses = {}
    for ticker in by_ticker:
        try:
            m = client.get_market(ticker)
            statuses[ticker] = str(m.status)
            results[ticker] = (getattr(m, "result", None) or "").lower()
        except Exception:
            statuses[ticker] = "unknown"
            results[ticker] = ""

    def compute_position(ticker_fills):
        net_yes = 0.0
        cash_out = 0.0
        cash_in = 0.0
        fees = 0.0
        total_contracts = 0.0
        for r in ticker_fills:
            c = float(r["contracts"])
            p = float(r["yes_price"])
            fee = float(r["fee"]) if r["fee"] else 0.0
            side = r["side"]
            action = r["action"]
            fees += fee
            if side == "yes":
                if action == "buy":
                    net_yes += c; cash_out += c * p + fee; total_contracts += c
                else:
                    net_yes -= c; cash_in += c * p - fee
            else:
                if action == "buy":
                    net_yes -= c; cash_out += c * (1 - p) + fee; total_contracts += c
                else:
                    # "sell no" = Kalshi records exiting a YES position from the NO side
                    net_yes -= c; cash_in += c * p - fee
        avg_price = (cash_out - cash_in) / total_contracts if total_contracts else 0
        return net_yes, cash_out, cash_in, fees, avg_price

    settled_rows = []
    open_rows = []

    for ticker, ticker_fills in sorted(by_ticker.items()):
        net_yes, cash_out, cash_in, fees, avg_price = compute_position(ticker_fills)
        result = results.get(ticker, "")
        status = statuses.get(ticker, "")
        category = ticker_fills[0].get("category") or _infer_category(ticker)
        trade_type = ticker_fills[0].get("type", "")
        is_finalized = "FINALIZED" in status.upper()

        if is_finalized and result:
            if result == "yes":
                payout = max(net_yes, 0) * 1.0
                if net_yes < 0:
                    payout = (-net_yes) * 0.0  # held NO, lost
            else:
                payout = max(-net_yes, 0) * 1.0 if net_yes < 0 else 0.0
            pnl = payout + cash_in - cash_out
            settled_rows.append({
                "Ticker":    ticker,
                "Category":  category,
                "Type":      trade_type,
                "Result":    result.upper(),
                "Net YES":   int(net_yes),
                "Avg Price": round(avg_price, 3),
                "Cost":      round(cash_out - cash_in, 2),
                "Payout":    round(payout, 2),
                "P&L":       round(pnl, 2),
                "Fees":      round(fees, 2),
            })
        else:
            cost_basis = cash_out - cash_in
            open_rows.append({
                "Ticker":      ticker,
                "Category":    category,
                "Type":        trade_type,
                "Net YES":     int(net_yes),
                "Avg Price":   round(avg_price, 3),
                "Cost Basis":  round(cost_basis, 2),
                "Fills":       len(ticker_fills),
            })

    settled_df = pd.DataFrame(settled_rows) if settled_rows else pd.DataFrame(
        columns=["Ticker","Category","Type","Result","Net YES","Avg Price","Cost","Payout","P&L","Fees"]
    )
    open_df = pd.DataFrame(open_rows) if open_rows else pd.DataFrame(
        columns=["Ticker","Category","Type","Net YES","Avg Price","Cost Basis","Fills"]
    )

    if not settled_df.empty:
        settled_df = settled_df.sort_values("P&L", ascending=False).reset_index(drop=True)
    if not open_df.empty:
        open_df = open_df.sort_values("Cost Basis", ascending=False).reset_index(drop=True)

    return settled_df, open_df


# ---------------------------------------------------------------------------
# Excel export
# ---------------------------------------------------------------------------

COLORS = {
    "header_blue":  "1F4E79",
    "header_green": "375623",
    "header_teal":  "215868",
    "alt_row":      "EBF3FB",
}


def _style_sheet(ws, header_color: str):
    thin = Side(style="thin", color="D0D0D0")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF", size=11)
        cell.fill = PatternFill("solid", fgColor=header_color)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border
    ws.row_dimensions[1].height = 28

    alt_fill = PatternFill("solid", fgColor=COLORS["alt_row"])
    for i, row in enumerate(ws.iter_rows(min_row=2), start=2):
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(horizontal="right" if isinstance(cell.value, (int, float)) else "left")
            if i % 2 == 0:
                cell.fill = alt_fill

    for col in ws.columns:
        max_len = max((len(str(c.value or "")) for c in col), default=10)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 3, 35)

    ws.freeze_panes = "A2"


def _add_bar_chart(ws, title, data_col, n_rows, anchor, horizontal=True):
    chart = BarChart()
    chart.type = "bar" if horizontal else "col"
    chart.title = title
    chart.style = 10
    chart.width = 22
    chart.height = 14
    chart.legend = None

    data = Reference(ws, min_col=data_col, min_row=1, max_row=n_rows + 1)
    cats = Reference(ws, min_col=1, min_row=2, max_row=n_rows + 1)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    ws.add_chart(chart, anchor)


def _add_line_chart(ws, pivot_df, anchor):
    n_rows = len(pivot_df)
    n_cols = len(pivot_df.columns)

    chart = LineChart()
    chart.title = "Monthly Volume by Category"
    chart.style = 10
    chart.y_axis.title = "Contracts"
    chart.x_axis.title = "Month"
    chart.width = 30
    chart.height = 16

    cats = Reference(ws, min_col=1, min_row=2, max_row=n_rows + 1)
    for col_idx in range(2, n_cols + 1):
        data = Reference(ws, min_col=col_idx, min_row=1, max_row=n_rows + 1)
        chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    ws.add_chart(chart, anchor)


def _style_pnl_sheet(ws, header_color: str):
    thin = Side(style="thin", color="D0D0D0")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF", size=11)
        cell.fill = PatternFill("solid", fgColor=header_color)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border
    ws.row_dimensions[1].height = 28

    green_fill  = PatternFill("solid", fgColor="C6EFCE")
    red_fill    = PatternFill("solid", fgColor="FFC7CE")
    alt_fill    = PatternFill("solid", fgColor=COLORS["alt_row"])

    # Find P&L column index
    pnl_col = None
    for cell in ws[1]:
        if str(cell.value) == "P&L":
            pnl_col = cell.column
            break

    for i, row in enumerate(ws.iter_rows(min_row=2), start=2):
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(horizontal="right" if isinstance(cell.value, (int, float)) else "left")
            if pnl_col and cell.column == pnl_col:
                if isinstance(cell.value, (int, float)):
                    cell.fill = green_fill if cell.value >= 0 else red_fill
                    cell.font = Font(bold=True)
            elif i % 2 == 0:
                cell.fill = alt_fill

    for col in ws.columns:
        max_len = max((len(str(c.value or "")) for c in col), default=10)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 3, 50)

    ws.freeze_panes = "A2"


def export_excel(cat_df, time_df, active_df, settled_df=None, open_df=None):
    sheets = {"By Category": cat_df, "Over Time": time_df, "Active Now": active_df}
    if settled_df is not None and not settled_df.empty:
        sheets["Settled P&L"] = settled_df
    if open_df is not None and not open_df.empty:
        sheets["Open Positions"] = open_df

    with pd.ExcelWriter(EXCEL_OUT, engine="openpyxl") as writer:
        for name, df in sheets.items():
            df.to_excel(writer, sheet_name=name, index=False)

    wb = load_workbook(EXCEL_OUT)

    ws_cat = wb["By Category"]
    _style_sheet(ws_cat, COLORS["header_blue"])
    _add_bar_chart(ws_cat, "Total Volume by Category", 3, len(cat_df), "E2", horizontal=True)

    ws_time = wb["Over Time"]
    _style_sheet(ws_time, COLORS["header_green"])
    _add_line_chart(ws_time, time_df, "A" + str(len(time_df) + 4))

    ws_active = wb["Active Now"]
    _style_sheet(ws_active, COLORS["header_teal"])
    _add_bar_chart(ws_active, "24h Volume by Category", 3, len(active_df), "F2", horizontal=True)

    if "Settled P&L" in wb.sheetnames:
        _style_pnl_sheet(wb["Settled P&L"], COLORS["header_blue"])
        # Add totals row
        ws = wb["Settled P&L"]
        last_row = ws.max_row + 2
        ws.cell(last_row, 1, "TOTAL").font = Font(bold=True)
        pnl_col = next((c.column for c in ws[1] if c.value == "P&L"), None)
        if pnl_col:
            col_letter = get_column_letter(pnl_col)
            ws.cell(last_row, pnl_col).value = f"=SUM({col_letter}2:{col_letter}{ws.max_row - 1})"
            ws.cell(last_row, pnl_col).font = Font(bold=True)

    if "Open Positions" in wb.sheetnames:
        _style_pnl_sheet(wb["Open Positions"], COLORS["header_teal"])

    wb.save(EXCEL_OUT)
    print(f"\nDashboard saved: {EXCEL_OUT}")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    with KalshiClient.from_env() as client:
        series_map   = build_series_category_map(client)
        settled_rows = pull_settled(client, series_map)
        open_rows    = pull_open_24h(client, series_map)
        settled_pnl, open_positions = build_trade_pnl(client)

    if not settled_rows and not open_rows and settled_pnl.empty:
        print("No data found.")
        return

    cat_df    = category_totals(settled_rows) if settled_rows else pd.DataFrame()
    time_df   = volume_over_time(settled_rows) if settled_rows else pd.DataFrame()
    active_df = active_now(open_rows) if open_rows else pd.DataFrame()

    print("\n--- By Category (settled) ---")
    print(cat_df.to_string(index=False))
    print("\n--- Active Now (24h volume) ---")
    print(active_df.to_string(index=False))

    if not settled_pnl.empty:
        total_pnl = settled_pnl["P&L"].sum()
        wins = (settled_pnl["P&L"] > 0).sum()
        print(f"\n--- Trade P&L ({len(settled_pnl)} settled, {wins}W/{len(settled_pnl)-wins}L, total {total_pnl:+.2f}) ---")
        print(settled_pnl[["Ticker","Result","Net YES","P&L"]].to_string(index=False))
    if not open_positions.empty:
        print(f"\n--- Open Positions ({len(open_positions)}) ---")
        print(open_positions[["Ticker","Net YES","Cost Basis"]].to_string(index=False))

    export_excel(cat_df, time_df, active_df, settled_pnl, open_positions)


if __name__ == "__main__":
    main()
