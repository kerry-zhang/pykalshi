"""Standalone script for Windows Task Scheduler: syncs discretionary fills."""

import sys
import os
import logging
from datetime import datetime

# Ensure the project root is on the path regardless of working directory
sys.path.insert(0, os.path.dirname(__file__))

logging.basicConfig(
    filename=os.path.join(os.path.dirname(__file__), "sync_trades.log"),
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
)

try:
    from pykalshi import KalshiClient
    from trade_log import sync_discretionary

    logging.info("Starting discretionary sync")
    with KalshiClient.from_env() as client:
        count = sync_discretionary(client)
    logging.info("Sync complete: %d new fill(s)", count)

except Exception as e:
    logging.error("Sync failed: %s", e)
    sys.exit(1)
