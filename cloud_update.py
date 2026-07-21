#!/usr/bin/env python3
"""
HAP DATABASE - CLOUD UPDATE (runs inside GitHub Actions).

Runs the standard HAP update (hap_pipeline.fetch_and_build - the same
pipeline the local helper uses) and writes the results into docs/ so
GitHub Pages can serve them.

Outputs (in docs/):
  hap_state.json              the data payload the web page reads
  HAP_Database_latest.xlsx    master spreadsheet + "Rents & UAs Detail" tab
  HAP_Rents_UAs_latest.xlsx   standalone row-level Rents & UAs workbook
                              (both HUD sheets combined; stable link)

docs/index.html is a static page and is NOT touched by this script.
"""
import os, sys, json, datetime, tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import hap_pipeline as hp
import hap_writer as hw
import hap_html as hh

DOCS = os.path.join(HERE, "docs")
os.makedirs(DOCS, exist_ok=True)

DETAIL_NUM  = set(hp.DETAIL_COLS[2:])                       # bedroom/unit counts + $ cols
DETAIL_TEXT = {"property_id", "contract_number"}

def fetch_rents_ua_detail():
    """Re-download the (small) rents/UA file and return the typed
    row-level detail table - both HUD sheets combined."""
    tmp = tempfile.mkdtemp(prefix="hap_ru_")
    try:
        path = hp._download("RentsAndUAs", hp.SOURCE_URLS["RentsAndUAs"], tmp)
        raw = hp._read_excel_any(path)
        return hp.rents_ua_detail(raw)
    finally:
        try:
            for f in os.listdir(tmp): os.remove(os.path.join(tmp, f))
            os.rmdir(tmp)
        except Exception:
            pass

def build_rent_tiers(detail):
    """Contracts where ANY bedroom size has more than one row (multiple
    rent/UA tiers). Returns {contract: [[bed, units, rent, fmr, ua], ...]}
    with every row for that contract, sorted by bedroom then rent, so the
    site can show the full breakdown in a popup."""
    import pandas as pd
    def clean(x):
        if x is None or (isinstance(x, float) and pd.isna(x)): return None
        f = float(x)
        return int(f) if f.is_integer() else f
    tiers = {}
    cols = ["assistance_bedroom_count","assistance_unit_count",
            "contract_rent_amount","fair_market_rent_amount",
            "utility_allowance_amount"]
    for c, g in detail.groupby("contract_number"):
        sizes = g.dropna(subset=["assistance_bedroom_count"])                  .groupby("assistance_bedroom_count").size()
        if len(sizes) and sizes.max() > 1:
            # pure HUD source order — no sorting. The master's "first row per
            # size" logic follows source order (including 5BR and 6BR rows
            # collapsing into one 5BR+ bucket), so preserving it here keeps
            # tier 1 identical to the master everywhere. Display sorting is
            # done client-side in the popup.
            tiers[str(c)] = [[clean(x) for x in row] for row in g[cols].values.tolist()]
    return tiers

def main():
    print("Running the standard HAP update...", flush=True)
    master = hp.fetch_and_build(progress=lambda m: print("  " + m, flush=True))
    asof = datetime.date.today().isoformat()

    print("Fetching row-level Rents & UAs detail...", flush=True)
    detail = fetch_rents_ua_detail()
    print(f"  {len(detail):,} detail rows.", flush=True)

    with open(os.path.join(DOCS, "hap_state.json"), "w", encoding="utf-8") as f:
        f.write(hh.payload_json(master, asof))

    extra = [("Rents & UAs Detail", detail, DETAIL_NUM, set(), DETAIL_TEXT)]
    hw.write_master(master, os.path.join(DOCS, "HAP_Database_latest.xlsx"), extra_sheets=extra)

    # flag data for the site's multi-tier popup
    tiers = build_rent_tiers(detail)
    with open(os.path.join(DOCS, "rent_tiers.json"), "w", encoding="utf-8") as f:
        json.dump({"contracts": tiers}, f, separators=(",", ":"))
    print(f"  {len(tiers):,} contracts flagged with multiple rent/UA tiers.", flush=True)

    # standalone workbook at a stable link
    hw.write_master(detail, os.path.join(DOCS, "HAP_Rents_UAs_latest.xlsx"),
                    sheet_name="Rents & UAs",
                    num_cols=DETAIL_NUM, date_cols=set(), text_cols=DETAIL_TEXT)
    print(f"  {len(master):,} contracts written. Data as of {asof}. Done.", flush=True)

if __name__ == "__main__":
    main()
