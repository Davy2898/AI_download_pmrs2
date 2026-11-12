"""Entry point for running the PMRS claims scraper."""

import sys

from scrape_rejected_claims import main as run_scraper


if __name__ == "__main__":
    sys.exit(run_scraper())
