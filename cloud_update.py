#!/usr/bin/env python3
"""
HAP DATABASE - CLOUD UPDATE (runs inside GitHub Actions).

Thin wrapper only. Runs the EXACT same update as the local helper
(hap_server.do_update) -- same downloads, same pipeline, same rent fix --
and writes the results into docs/ so GitHub Pages can serve them.
hap_pipeline.py / hap_writer.py / hap_html.py are NOT modified.

Outputs (in docs/):
  hap_state.json              the data payload the web page reads
  HAP_Database_latest.xlsx    finished spreadsheet, same format as always

docs/index.html is a static page and is NOT touched by this script.
"""
import os, sys, datetime

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import hap_pipeline as hp   # the updater -- untouched
import hap_writer as hw     # untouched
import hap_html as hh       # untouched (reused only for payload_json)

DOCS = os.path.join(HERE, "docs")
os.makedirs(DOCS, exist_ok=True)

def main():
    print("Running the standard HAP update (unchanged pipeline)...", flush=True)
    master = hp.fetch_and_build(progress=lambda m: print("  " + m, flush=True))
    asof = datetime.date.today().isoformat()

    with open(os.path.join(DOCS, "hap_state.json"), "w", encoding="utf-8") as f:
        f.write(hh.payload_json(master, asof))
    hw.write_master(master, os.path.join(DOCS, "HAP_Database_latest.xlsx"))
    print(f"  {len(master):,} contracts written. Data as of {asof}. Done.", flush=True)

if __name__ == "__main__":
    main()
