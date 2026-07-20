"""
HAP Database pipeline (pure Python) - reproduces the Excel Power Query flow
EXACTLY, with the corrected rents/UA logic (first row per contract+bedroom,
NOT the buggy sum). Builds the MasterTable from the six HUD source files.
"""
import pandas as pd
import numpy as np

# ---- source URLs (from the workbook's SourceURLs sheet) ----
SOURCE_URLS = {
    "HAPInfo":     "https://www.hud.gov/sites/dfiles/Housing/documents/MF-Properties-with-Assistance-Sec8-Contracts1.xlsx",
    "HAPRents":    "https://www.hud.gov/sites/dfiles/Housing/documents/MF-Assistance-Sec8-Contracts1.xlsx",
    "HAPRenewal1": "https://www.hud.gov/sites/dfiles/Housing/documents/contractrenewaallcontracts.xls",
    "REACScores":  "https://www.hud.gov/sites/default/files/Housing/documents/MF-Inspection-Report.xls",
    "RentsAndUAs": "https://www.hud.gov/sites/dfiles/Housing/documents/contractsrentutilityamt.xlsx",
}

def _t(s):
    """Trim like Power Query Text.Trim (only when not null)."""
    if s is None: return None
    if isinstance(s, float) and pd.isna(s): return None
    return str(s).strip()

def _num(x):
    try:
        if x is None or (isinstance(x, float) and pd.isna(x)): return None
        return float(x)
    except (ValueError, TypeError):
        return None

def br_tag(b):
    v = _num(b)
    if v is None: return None
    bi = int(v)
    return "5plus" if bi >= 5 else str(bi)

# ---------- per-source shaping (mirrors each M query) ----------
def shape_hap_rents(df):
    df = df.copy()
    df["contract_number"] = df["contract_number"].map(_t)
    df["property_id"] = df["property_id"].map(_t)
    df = df[df["contract_number"].notna() & (df["contract_number"] != "")]
    return df

import datetime as _dt
_EPOCH = _dt.datetime(1899,12,30)
def parse_date(x):
    if x is None or (isinstance(x,float) and pd.isna(x)) or (isinstance(x,str) and x.strip()==""): return None
    if isinstance(x,(_dt.datetime,_dt.date)): return pd.Timestamp(x)
    if isinstance(x,pd.Timestamp): return x
    s=str(x).strip()
    try:
        f=float(s)
        if 1 < f < 200000: return pd.Timestamp(_EPOCH + _dt.timedelta(days=f))
    except ValueError: pass
    return pd.to_datetime(s, errors="coerce")

def shape_renewal1(df):
    df = df.copy()
    df["contract_number"] = df["contract_number"].map(_t)
    df["property_id"] = df["property_id"].map(_t)
    # robust date key for dedup (handles serials AND real dates); keep raw value intact
    df["_eff"] = df["tracs_effective_date"].map(parse_date)
    df = df.sort_values("_eff", ascending=False, kind="mergesort")
    df = df.drop_duplicates(subset=["contract_number"], keep="first")
    # normalise the two output date cols to real Timestamps so the writer formats them
    df["tracs_effective_date"] = df["tracs_effective_date"].map(parse_date)
    df["tracs_overall_expiration_date"] = df["tracs_overall_expiration_date"].map(parse_date)
    return df.drop(columns=["_eff"])

def shape_hapinfo(df):
    df = df.copy()
    for c in ["property_id","property_name_text","city_name_text","state_code","county_name_text"]:
        if c in df.columns: df[c] = df[c].map(_t)
    df = df.drop_duplicates(subset=["property_id"], keep="first")
    return df

def shape_reac(df):
    df = df.copy()
    df = df.rename(columns={"REMS Property Id":"property_id"})
    df["property_id"] = df["property_id"].map(_t)
    def first2(x):
        if x is None or (isinstance(x,float) and pd.isna(x)): return None
        return str(x)[:2]
    df["Last REAC Numeric 1"] = df["Inspection Score1"].map(first2)
    df["Last REAC Numeric 2"] = df["Inspection Score2"].map(first2)
    df["Last REAC Numeric 3"] = df["Inspection Score3"].map(first2)
    df = df.drop_duplicates(subset=["property_id"], keep="first")
    return df

def shape_rents_uas(df):
    """The FIXED pivot: first value per contract+bedroom (not sum)."""
    df = df.copy()
    df["contract_number"] = df["contract_number"].map(_t)
    df["BR_Tag"] = df["assistance_bedroom_count"].map(br_tag)
    df = df[df["contract_number"].notna() & (df["contract_number"]!="") & df["BR_Tag"].notna()]
    df["contract_rent_amount"] = df["contract_rent_amount"].map(_num)
    df["utility_allowance_amount"] = df["utility_allowance_amount"].map(_num)
    # FIRST row per (contract, BR_Tag), preserving source order  == List.First
    rent = (df.dropna(subset=["contract_rent_amount"])
              .drop_duplicates(["contract_number","BR_Tag"], keep="first")
              .pivot(index="contract_number", columns="BR_Tag", values="contract_rent_amount"))
    ua   = (df.dropna(subset=["utility_allowance_amount"])
              .drop_duplicates(["contract_number","BR_Tag"], keep="first")
              .pivot(index="contract_number", columns="BR_Tag", values="utility_allowance_amount"))
    rent = rent.rename(columns={"0":"0BR_Rent","1":"1BR_Rent","2":"2BR_Rent","3":"3BR_Rent","4":"4BR_Rent","5plus":"5plusBR_Rent"})
    ua   = ua.rename(columns={"0":"0BR_UA","1":"1BR_UA","2":"2BR_UA","3":"3BR_UA","4":"4BR_UA","5plus":"5plusBR_UA"})
    # unit mix: SUM of assistance_unit_count per (contract, BR) — a contract
    # can have several rows for the same bedroom size (assistance tiers), so
    # unlike rents (first value) the counts must be summed.
    df["assistance_unit_count"] = df["assistance_unit_count"].map(_num)
    units = (df.dropna(subset=["assistance_unit_count"])
               .groupby(["contract_number","BR_Tag"])["assistance_unit_count"].sum()
               .unstack())
    units = units.rename(columns={"0":"0BR_Units","1":"1BR_Units","2":"2BR_Units","3":"3BR_Units","4":"4BR_Units","5plus":"5plusBR_Units"})
    out = rent.join(ua, how="outer").join(units, how="outer").reset_index()
    return out

# ---------- final master assembly (mirrors MasterTable_NEW) ----------
FINAL_ORDER = [
    "Property Name","Total Units","City","State","County","Renewal Type","Program Type","Doc Type",
    "HAP Effective Date","HAP Expiration Date","Rent-to-FMR Ratio",
    "Last REAC 1","Last REAC 2","Last REAC 3","Last REAC Date 1","Last REAC Date 2","Last REAC Date 3",
    "1BR Rent","2BR Rent","3BR Rent","4BR Rent","0BR Rent","5BR+ Rent",
    "1BR UA","2BR UA","3BR UA","4BR UA","0BR UA","5BR+ UA",
    "Contract Units","Property ID","Contract Number","Property ID2",
    "Last REAC 1 Numeric","Last REAC 2 Numeric","Last REAC 3 Numeric",
    "property_phone_number","address_line1_text","address_line2_text","city_name_text","state_code",
    "zip_code","zip4_code","county_code","county_name_text","msa_code","msa_name_text",
    "congressional_district_code","placed_base_city_name_text","property_total_unit_count","property_category_name",
    "primary_financing_type","associated_financing_Number","is_insured_ind","is_202_811_ind","is_hud_held_ind",
    "is_hud_owned_ind","is_hospital_ind","is_nursing_home_ind","is_board_and_care_ind","is_assisted_living_ind",
    "is_refinanced_ind","is_221d3_ind","is_221d4_ind","is_236_ind","is_non_insured_ind","is_bmir_ind",
    "is_risk_sharing_ind","is_mip_ind","is_co_insured_ind","is_opportunity_zone_ind","ownership_effective_date",
    "owner_participant_id","owner_company_type","owner_individual_first_name","owner_individual_middle_name",
    "owner_individual_last_name","owner_individual_full_name","owner_individual_title_text","owner_organization_name",
    "owner_address_line1","owner_address_line2","owner_city_name","owner_state_code","owner_zip_code","owner_zip4_code",
    "owner_main_phone_number_text","owner_main_fax_number_text","owner_contact_email_text","mgmt_agent_participant_id",
    "mgmt_agent_company_type","mgmt_agent_indv_first_name","mgmt_agent_indv_last_name","mgmt_agent_indv_middle_name",
    "mgmt_agent_full_name","mgmt_agent_indv_title_text","mgmt_agent_org_name","mgmt_agent_address_line1",
    "mgmt_agent_address_line2","mgmt_agent_city_name","mgmt_agent_state_code","mgmt_agent_zip_code","mgmt_agent_zip4_code",
    "mgmt_agent_main_phone_number","mgmt_agent_main_fax_number","mgmt_contact_email_text",
    "0BR Units","1BR Units","2BR Units","3BR Units","4BR Units","5BR+ Units",
]

def build_master(hap_rents, renewal1, hapinfo, reac, rents_uas):
    rents = shape_hap_rents(hap_rents)
    r1    = shape_renewal1(renewal1)
    hi    = shape_hapinfo(hapinfo)
    rc    = shape_reac(reac)
    ru    = shape_rents_uas(rents_uas)

    # spine
    spine = rents[["contract_number","property_id","property_name_text",
                   "rent_to_FMR_ratio","program_type_name","contract_doc_type_code"]].copy()

    # J1 + renewal1
    r1cols = {"renewal_option_name":"R1_renewal_option_name","tracs_effective_date":"R1_tracs_effective_date",
              "tracs_overall_expiration_date":"R1_tracs_overall_expiration_date","contract_doc_type_code":"R1_contract_doc_type_code",
              "program_type_name":"R1_program_type_name","assisted_units_count":"R1_assisted_units_count",
              "maximum_contract_unit_count":"R1_max_contract_units"}
    r1sel = r1[["contract_number"]+list(r1cols.keys())].rename(columns=r1cols)
    m = spine.merge(r1sel, on="contract_number", how="left")

    # J2 + HAPInfo (property_id). city/state/county get HI_ prefix; rest keep name
    hi2 = hi.rename(columns={"city_name_text":"HI_city_name_text","state_code":"HI_state_code","county_name_text":"HI_county_name_text"})
    hi_keep = ["property_id","property_phone_number","address_line1_text","address_line2_text",
        "HI_city_name_text","HI_state_code","zip_code","zip4_code","county_code","HI_county_name_text","msa_code",
        "msa_name_text","congressional_district_code","placed_base_city_name_text","property_total_unit_count",
        "property_category_name","primary_financing_type","associated_financing_Number","is_insured_ind","is_202_811_ind",
        "is_hud_held_ind","is_hud_owned_ind","is_hospital_ind","is_nursing_home_ind","is_board_and_care_ind",
        "is_assisted_living_ind","is_refinanced_ind","is_221d3_ind","is_221d4_ind","is_236_ind","is_non_insured_ind",
        "is_bmir_ind","is_risk_sharing_ind","is_mip_ind","is_co_insured_ind","is_opportunity_zone_ind","ownership_effective_date",
        "owner_participant_id","owner_company_type","owner_individual_first_name","owner_individual_middle_name",
        "owner_individual_last_name","owner_individual_full_name","owner_individual_title_text","owner_organization_name",
        "owner_address_line1","owner_address_line2","owner_city_name","owner_state_code","owner_zip_code","owner_zip4_code",
        "owner_main_phone_number_text","owner_main_fax_number_text","owner_contact_email_text","mgmt_agent_participant_id",
        "mgmt_agent_company_type","mgmt_agent_indv_first_name","mgmt_agent_indv_last_name","mgmt_agent_indv_middle_name",
        "mgmt_agent_full_name","mgmt_agent_indv_title_text","mgmt_agent_org_name","mgmt_agent_address_line1",
        "mgmt_agent_address_line2","mgmt_agent_city_name","mgmt_agent_state_code","mgmt_agent_zip_code","mgmt_agent_zip4_code",
        "mgmt_agent_main_phone_number","mgmt_agent_main_fax_number","mgmt_contact_email_text"]
    hi_keep = [c for c in hi_keep if c in hi2.columns]
    m = m.merge(hi2[hi_keep], on="property_id", how="left")

    # J3 + REACScores (property_id)
    rc2 = rc.rename(columns={"Inspection Score1":"Last REAC 1","Inspection Score2":"Last REAC 2","Inspection Score3":"Last REAC 3",
                             "Release Date 1":"Last REAC Date 1","Release Date 2":"Last REAC Date 2","Release Date 3":"Last REAC Date 3",
                             "Last REAC Numeric 1":"Last REAC 1 Numeric","Last REAC Numeric 2":"Last REAC 2 Numeric","Last REAC Numeric 3":"Last REAC 3 Numeric"})
    rc_keep = ["property_id","Last REAC 1","Last REAC 2","Last REAC 3","Last REAC Date 1","Last REAC Date 2","Last REAC Date 3",
               "Last REAC 1 Numeric","Last REAC 2 Numeric","Last REAC 3 Numeric"]
    m = m.merge(rc2[rc_keep], on="property_id", how="left")

    # J4 + RentsAndUAs (contract_number)
    ru2 = ru.rename(columns={"1BR_Rent":"1BR Rent","2BR_Rent":"2BR Rent","3BR_Rent":"3BR Rent","4BR_Rent":"4BR Rent","0BR_Rent":"0BR Rent","5plusBR_Rent":"5BR+ Rent",
                             "1BR_UA":"1BR UA","2BR_UA":"2BR UA","3BR_UA":"3BR UA","4BR_UA":"4BR UA","0BR_UA":"0BR UA","5plusBR_UA":"5BR+ UA",
                             "0BR_Units":"0BR Units","1BR_Units":"1BR Units","2BR_Units":"2BR Units","3BR_Units":"3BR Units","4BR_Units":"4BR Units","5plusBR_Units":"5BR+ Units"})
    m = m.merge(ru2, on="contract_number", how="left")

    # computed columns
    m["Rent-to-FMR Ratio"] = pd.to_numeric(m["rent_to_FMR_ratio"], errors="coerce")/100
    m["Property ID2"] = m["property_id"]
    m["city_name_text"] = m["HI_city_name_text"]
    m["state_code"] = m["HI_state_code"]
    m["county_name_text"] = m["HI_county_name_text"]

    # renames to final labels
    m = m.rename(columns={
        "property_name_text":"Property Name","R1_assisted_units_count":"Total Units",
        "HI_city_name_text":"City","HI_state_code":"State","HI_county_name_text":"County",
        "R1_renewal_option_name":"Renewal Type","R1_program_type_name":"Program Type","R1_contract_doc_type_code":"Doc Type",
        "R1_tracs_effective_date":"HAP Effective Date","R1_tracs_overall_expiration_date":"HAP Expiration Date",
        "R1_max_contract_units":"Contract Units","property_id":"Property ID","contract_number":"Contract Number"})

    for col in FINAL_ORDER:
        if col not in m.columns: m[col] = np.nan
    return m[FINAL_ORDER]

# ---------- row-level rents & UA detail (both HUD sheets combined) ----------
DETAIL_COLS = ["property_id","contract_number","assistance_bedroom_count",
               "assistance_unit_count","contract_rent_amount",
               "fair_market_rent_amount","utility_allowance_amount"]

def rents_ua_detail(df):
    """The full row-level Rents & UAs table (all sheets, source order),
    typed for a clean spreadsheet."""
    d = df.copy()
    d["property_id"] = d["property_id"].map(_t)
    d["contract_number"] = d["contract_number"].map(_t)
    d = d[d["contract_number"].notna() & (d["contract_number"] != "")]
    for c in ["assistance_bedroom_count","assistance_unit_count",
              "contract_rent_amount","fair_market_rent_amount",
              "utility_allowance_amount"]:
        d[c] = d[c].map(_num)
    return d[DETAIL_COLS].reset_index(drop=True)

# ---------- live HUD fetch (runs on the user's machine) ----------
HTTP_HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}

def _download(name, url, folder):
    import requests, os
    ext = ".xls" if url.lower().endswith(".xls") else ".xlsx"
    dest = os.path.join(folder, name + ext)
    r = requests.get(url, headers=HTTP_HEADERS, timeout=300)
    r.raise_for_status()
    with open(dest, "wb") as f:
        f.write(r.content)
    return dest

def _read_excel_any(path):
    """Read a HUD workbook. HUD's export tool sometimes splits ONE table
    across several sheets (e.g. contractsrentutilityamt.xlsx); any extra
    sheets whose header row matches the first sheet are continuations and
    are concatenated in order. Sheets with different headers are ignored,
    so single-sheet files behave exactly as before."""
    engine = "xlrd" if path.lower().endswith(".xls") else "openpyxl"
    sheets = pd.read_excel(path, sheet_name=None, engine=engine, dtype=str)
    frames = list(sheets.values())
    first_cols = list(frames[0].columns)
    same = [f for f in frames if list(f.columns) == first_cols]
    if len(same) > 1:
        return pd.concat(same, ignore_index=True)
    return frames[0]

def fetch_and_build(progress=None):
    """Download the 6 HUD files and return the finished master DataFrame.
       progress: optional callable(str) for status messages."""
    import tempfile, os
    def say(m):
        if progress: progress(m)
    tmp = tempfile.mkdtemp(prefix="hap_")
    try:
        files = {}
        for name, url in SOURCE_URLS.items():
            say(f"downloading {name} …")
            files[name] = _download(name, url, tmp)
        say("reading files …")
        dfs = {n: _read_excel_any(p) for n, p in files.items()}
        say("building master table (applying the rent fix) …")
        master = build_master(dfs["HAPRents"], dfs["HAPRenewal1"],
                              dfs["HAPInfo"], dfs["REACScores"], dfs["RentsAndUAs"])
        return master
    finally:
        try:
            for f in os.listdir(tmp): os.remove(os.path.join(tmp, f))
            os.rmdir(tmp)
        except Exception:
            pass
