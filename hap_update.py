#!/usr/bin/env python3
"""
HAP DATABASE - ONE-CLICK UPDATE (no Excel required)
Downloads the six HUD source files, rebuilds the master table with the
CORRECT rent/UA logic (first row per contract+bedroom, not the buggy sum),
and writes a finished, formatted spreadsheet next to this script.

Just run it (double-click "Run HAP Update.bat"). No setup, no Power Query,
no macros. Output:  HAP_Database_<today>.xlsx
"""
import os, sys, tempfile, datetime, traceback
import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import hap_pipeline as hp
import hap_writer as hw
import hap_html as hh

HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}

def log(msg): print(msg, flush=True)

def download(name, url, folder):
    import requests
    ext = ".xls" if url.lower().endswith(".xls") else ".xlsx"
    dest = os.path.join(folder, name + ext)
    log(f"   downloading {name} ...")
    r = requests.get(url, headers=HEADERS, timeout=300)
    r.raise_for_status()
    with open(dest, "wb") as f:
        f.write(r.content)
    return dest

def read_excel_any(path):
    if path.lower().endswith(".xls"):
        return pd.read_excel(path, sheet_name=0, engine="xlrd", dtype=str)
    return pd.read_excel(path, sheet_name=0, engine="openpyxl", dtype=str)

def main():
    log("=" * 60)
    log("  HAP DATABASE UPDATE")
    log("=" * 60)
    tmp = tempfile.mkdtemp(prefix="hap_")
    try:
        log("\n[1/3] Downloading the six HUD files (a few minutes)...")
        files = {}
        for name, url in hp.SOURCE_URLS.items():
            try:
                files[name] = download(name, url, tmp)
            except Exception as e:
                log(f"\n   ! Could not download {name}")
                log(f"     URL: {url}")
                log(f"     Error: {e}")
                log("     If HUD changed this link, update it in hap_pipeline.py (SOURCE_URLS).")
                raise

        log("\n[2/3] Building the master table (applying the rent fix)...")
        dfs = {n: read_excel_any(p) for n, p in files.items()}
        master = hp.build_master(dfs["HAPRents"], dfs["HAPRenewal1"],
                                 dfs["HAPInfo"], dfs["REACScores"], dfs["RentsAndUAs"])
        log(f"     {len(master):,} contracts, {len(master.columns)} columns.")

        log("\n[3/3] Writing the finished spreadsheet...")
        today = datetime.date.today().isoformat()
        out = os.path.join(HERE, f"HAP_Database_{today}.xlsx")
        hw.write_master(master, out)
        log(f"\nDONE  ->  {out}")
        log(f"      {len(master):,} rows written.")

        # open it for the user (Windows)
        try:
            os.startfile(out)  # type: ignore[attr-defined]
        except Exception:
            pass
        return 0
    except Exception:
        log("\n" + "!" * 60)
        log("  Something went wrong. Details below:")
        log("!" * 60)
        traceback.print_exc()
        return 1
    finally:
        # tidy temp downloads
        try:
            for f in os.listdir(tmp): os.remove(os.path.join(tmp, f))
            os.rmdir(tmp)
        except Exception:
            pass

if __name__ == "__main__":
    rc = main()
    if os.name == "nt":
        input("\nPress Enter to close...")
    sys.exit(rc)
