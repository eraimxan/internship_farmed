"""
cpi_scripts/run.py — Main runner: fetch → build Excel → build HTML

Steps:
  1. Discover API parameters for CPI index 703076
  2. Fetch data for all regions + both comparison types
  3. Save CSV to output/cpi_data.csv
  4. Build Excel to output/CPI_Kazakhstan.xlsx
  5. Build HTML to output/cpi_kazakhstan.html

Usage:
    python run.py                   # uses cookie from ../data/cookie.txt (resumes from checkpoint)
    python run.py --cookie "..."    # provide cookie string directly
    python run.py --no-resume       # ignore checkpoint, fetch everything from scratch
    python run.py --skip-fetch      # skip fetching, rebuild outputs from existing CSV
    python run.py --skip-excel      # skip Excel build
    python run.py --skip-html       # skip HTML build

Requirements:
    pip install requests pandas openpyxl
"""

import sys
import argparse
import logging
from pathlib import Path

# Allow running from any directory
sys.path.insert(0, str(Path(__file__).resolve().parent))

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler()],
)
log = logging.getLogger(__name__)


def main():
    parser = argparse.ArgumentParser(
        description="Full CPI pipeline: fetch → Excel → HTML",
        formatter_class=argparse.RawTextHelpFormatter,
    )
    parser.add_argument("--cookie",       default="", help="Cookie string for taldau.stat.gov.kz")
    parser.add_argument("--skip-fetch",   action="store_true", help="Skip data fetching (use existing CSV)")
    parser.add_argument("--skip-excel",   action="store_true", help="Skip Excel generation")
    parser.add_argument("--skip-html",    action="store_true", help="Skip HTML generation")
    parser.add_argument("--no-resume",    action="store_true", help="Ignore checkpoint; fetch everything from scratch")
    parser.add_argument("--workers",      type=int, default=0,  help="Parallel region workers (0 = use config default)")
    args = parser.parse_args()

    from cpi_config import DATA_CSV, EXCEL_OUT, HTML_OUT

    # ─── STEP 1: FETCH ───────────────────────────────────────────
    if not args.skip_fetch:
        log.info("=" * 60)
        log.info("STEP 1: Fetching CPI data from taldau.stat.gov.kz …")
        log.info("=" * 60)
        try:
            import fetch
            # Temporarily override sys.argv for argparse in fetch.main()
            orig_argv = sys.argv
            sys.argv = ["fetch.py"]
            if args.cookie:
                sys.argv += ["--cookie", args.cookie]
            if args.no_resume:
                sys.argv += ["--no-resume"]
            if args.workers:
                sys.argv += ["--workers", str(args.workers)]
            df, params = fetch.main()
            sys.argv = orig_argv
            log.info("Fetch complete. Rows: %d", len(df))
        except SystemExit as e:
            log.error("Fetch failed with exit code %s", e.code)
            sys.exit(e.code)
        except Exception as exc:
            log.error("Fetch error: %s", exc, exc_info=True)
            sys.exit(1)
    else:
        log.info("Skipping fetch (--skip-fetch). Using existing %s", DATA_CSV)
        if not DATA_CSV.exists():
            log.error("CSV not found: %s", DATA_CSV)
            sys.exit(1)

    # ─── STEP 2: EXCEL ───────────────────────────────────────────
    if not args.skip_excel:
        log.info("=" * 60)
        log.info("STEP 2: Building Excel file …")
        log.info("=" * 60)
        try:
            import pandas as pd
            import build_excel

            df = pd.read_csv(DATA_CSV, encoding="utf-8-sig", low_memory=False)
            params_path = DATA_CSV.parent / "cpi_params.json"
            build_excel.build_excel(
                df,
                params_path if params_path.exists() else None,
                EXCEL_OUT,
            )
            log.info("Excel ready: %s", EXCEL_OUT)
        except Exception as exc:
            log.error("Excel build error: %s", exc, exc_info=True)
    else:
        log.info("Skipping Excel (--skip-excel).")

    # ─── STEP 3: HTML ────────────────────────────────────────────
    if not args.skip_html:
        log.info("=" * 60)
        log.info("STEP 3: Building HTML visualization …")
        log.info("=" * 60)
        try:
            import pandas as pd
            import build_html

            df = pd.read_csv(DATA_CSV, encoding="utf-8-sig", low_memory=False)
            params_path = DATA_CSV.parent / "cpi_params.json"
            build_html.build_html(
                df,
                params_path if params_path.exists() else None,
                HTML_OUT,
            )
            log.info("HTML ready: %s", HTML_OUT)
        except Exception as exc:
            log.error("HTML build error: %s", exc, exc_info=True)
    else:
        log.info("Skipping HTML (--skip-html).")

    log.info("=" * 60)
    log.info("Done. Output files:")
    if DATA_CSV.exists():
        log.info("  CSV:   %s  (%.1f MB)", DATA_CSV, DATA_CSV.stat().st_size / 1e6)
    if EXCEL_OUT.exists():
        log.info("  Excel: %s  (%.1f MB)", EXCEL_OUT, EXCEL_OUT.stat().st_size / 1e6)
    if HTML_OUT.exists():
        log.info("  HTML:  %s  (%.1f MB)", HTML_OUT, HTML_OUT.stat().st_size / 1e6)
    log.info("=" * 60)


if __name__ == "__main__":
    main()
