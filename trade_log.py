"""Trade logger: tracks systematic vs discretionary Kalshi fills."""

import csv
import os
import time
from datetime import datetime, timezone

from pykalshi import KalshiClient
from pykalshi.enums import Action, Side, TimeInForce, SelfTradePrevention

LOG_FILE = os.path.join(os.path.dirname(__file__), "trades.csv")

_COLUMNS = [
    "fill_id", "trade_id", "order_id", "ticker",
    "side", "action", "contracts", "yes_price",
    "fee", "type", "strategy", "category", "bet", "logged_at", "ts",
]


# --- Internal helpers ---

def _load_logged_ids() -> set:
    if not os.path.exists(LOG_FILE):
        return set()
    with open(LOG_FILE, newline="") as f:
        return {row["fill_id"] for row in csv.DictReader(f)}


def _max_logged_ts() -> int | None:
    """Latest fill timestamp already in the log, or None if empty/untracked."""
    if not os.path.exists(LOG_FILE):
        return None
    max_ts = None
    with open(LOG_FILE, newline="") as f:
        for row in csv.DictReader(f):
            ts = row.get("ts")
            if ts:
                max_ts = max(max_ts or 0, int(float(ts)))
    return max_ts


def _append_rows(rows: list) -> None:
    write_header = not os.path.exists(LOG_FILE)
    with open(LOG_FILE, "a", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=_COLUMNS)
        if write_header:
            writer.writeheader()
        writer.writerows(rows)


def _fill_to_row(fill, trade_type: str, strategy: str | None, market=None) -> dict:
    category = ""
    bet = ""
    if market is not None:
        category = getattr(market, "category", "") or ""
        bet = getattr(market, "title", "") or ""
    return {
        "fill_id": fill.fill_id or fill.trade_id,
        "trade_id": fill.trade_id,
        "order_id": fill.order_id,
        "ticker": fill.ticker,
        "side": fill.side.value if hasattr(fill.side, "value") else fill.side,
        "action": fill.action.value if hasattr(fill.action, "value") else fill.action,
        "contracts": fill.count_fp,
        "yes_price": fill.yes_price_dollars,
        "fee": fill.fee_cost_dollars or "0",
        "type": trade_type,
        "strategy": strategy or "",
        "category": category,
        "bet": bet,
        "logged_at": datetime.now(timezone.utc).isoformat(),
        "ts": fill.ts,
    }


def _log_fills(fills, trade_type: str, strategy: str | None, markets: dict = None) -> int:
    logged = _load_logged_ids()
    new_rows = [
        _fill_to_row(f, trade_type, strategy, (markets or {}).get(f.ticker))
        for f in fills
        if (f.fill_id or f.trade_id) not in logged
    ]
    if new_rows:
        _append_rows(new_rows)
    return len(new_rows)


# --- Public API ---

def place_systematic_order(
    client: KalshiClient,
    ticker: str,
    action: Action,
    side: Side,
    count_fp: str,
    *,
    yes_price_dollars: str | None = None,
    no_price_dollars: str | None = None,
    strategy: str | None = None,
    **kwargs,
):
    """Place an order and automatically log its fills as systematic.

    All extra keyword args are forwarded to portfolio.place_order().
    Returns the Order object.
    """
    market = client.get_market(ticker if isinstance(ticker, str) else ticker.ticker)
    order = client.portfolio.place_order(
        ticker, action, side, count_fp,
        yes_price_dollars=yes_price_dollars,
        no_price_dollars=no_price_dollars,
        **kwargs,
    )
    # Brief pause so fills appear in the API before fetching
    time.sleep(0.5)
    fills = client.portfolio.get_fills(order_id=order.order_id)
    logged = _log_fills(fills, "systematic", strategy, {market.ticker: market})
    print(f"Order placed ({order.order_id}): logged {logged} fill(s) as systematic"
          + (f" [{strategy}]" if strategy else ""))
    return order


_SYNC_LOOKBACK_SEC = 86400  # re-check the last day of fills as a safety margin against sync gaps


def sync_discretionary(client: KalshiClient) -> int:
    """Fetch fills since the last logged one (with a lookback safety margin)
    and log any not already tracked as discretionary.

    Paginates through *all* matching fills (not just the most recent page),
    so a missed scheduled run can't cause fills to permanently fall outside
    the fetch window and go untracked.

    Call this periodically to capture trades placed via the Kalshi app.
    Returns the number of new discretionary fills logged.
    """
    last_ts = _max_logged_ts()
    min_ts = max(last_ts - _SYNC_LOOKBACK_SEC, 0) if last_ts else None
    fills = client.portfolio.get_fills(min_ts=min_ts, fetch_all=True)
    # Fetch market metadata for any tickers not yet logged
    logged_ids = _load_logged_ids()
    new_fills = [f for f in fills if (f.fill_id or f.trade_id) not in logged_ids]
    markets = {}
    for f in new_fills:
        if f.ticker not in markets:
            try:
                markets[f.ticker] = client.get_market(f.ticker)
            except Exception:
                pass
    logged = _log_fills(fills, "discretionary", None, markets)
    if logged:
        print(f"Synced {logged} new discretionary fill(s).")
    else:
        print("No new discretionary fills found.")
    return logged


def summary() -> None:
    """Print a breakdown of logged trades by type (systematic vs discretionary)."""
    if not os.path.exists(LOG_FILE):
        print("No trades logged yet.")
        return

    from collections import defaultdict

    stats = defaultdict(lambda: {
        "fills": 0, "tickers": set(),
        "buy_contracts": 0.0, "buy_cost": 0.0,
        "sell_contracts": 0.0, "sell_proceeds": 0.0,
    })
    strategies = defaultdict(set)
    categories = defaultdict(lambda: defaultdict(float))  # type -> category -> net_cost

    with open(LOG_FILE, newline="") as f:
        for row in csv.DictReader(f):
            t = row["type"]
            s = stats[t]
            s["fills"] += 1
            s["tickers"].add(row["ticker"])
            contracts = float(row["contracts"])
            price = float(row["yes_price"])
            cost = contracts * price
            if row["action"] == "buy":
                s["buy_contracts"] += contracts
                s["buy_cost"] += cost
                categories[t][row.get("category") or "Unknown"] += cost
            else:
                s["sell_contracts"] += contracts
                s["sell_proceeds"] += cost
                categories[t][row.get("category") or "Unknown"] -= cost
            if row.get("strategy"):
                strategies[t].add(row["strategy"])

    print(f"\n{'='*45}")
    print(f"{'TRADE LOG SUMMARY':^45}")
    print(f"{'='*45}")

    for trade_type in ["systematic", "discretionary"]:
        if trade_type not in stats:
            continue
        s = stats[trade_type]
        net = s["buy_cost"] - s["sell_proceeds"]
        print(f"\n  [{trade_type.upper()}]")
        print(f"  Fills:             {s['fills']}")
        print(f"  Unique tickers:    {len(s['tickers'])}")
        print(f"  Bought:            {s['buy_contracts']:.0f} contracts  (${s['buy_cost']:.2f})")
        print(f"  Sold:              {s['sell_contracts']:.0f} contracts  (${s['sell_proceeds']:.2f})")
        print(f"  Net deployed:      ${net:.2f}")
        if strategies[trade_type]:
            print(f"  Strategies:        {', '.join(sorted(strategies[trade_type]))}")
        if categories[trade_type]:
            print(f"  By category:")
            for cat, amt in sorted(categories[trade_type].items(), key=lambda x: -x[1]):
                print(f"    {cat:<25} ${amt:.2f}")

    print()
