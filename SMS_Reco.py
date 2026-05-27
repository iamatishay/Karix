"""
SMS Finance Analysis — Streamlit Application
=============================================
Tabs:
  1. Upload & Run
  2. Summary (Vol/Rev reconciliation + Remarks)
  3. Rate Comparison
  4. Volume & Revenue Detail
  5. Ex Rate Validation
  6. Download Results

FIX: Rate comparison now correctly surfaces mismatches from SMS_NLD and SMS_ILD
     by falling back to ESME-only matching when country-level join finds no strategy rate.
"""

import io
import numpy as np
import pandas as pd
import streamlit as st

st.markdown(
    """
    <style>
    .app-header {
        display: flex;
        align-items: center;
        padding: 0.5rem 1rem;
        background: white;
    }

    .logo {
        height: 32px;
        width: auto;
        object-fit: contain;
    }
    </style>

    <div class="app-header">
        <img src="https://cdn.prod.website-files.com/68c6698e3517c4af35b889cf/68e2ae5bad4b483249840e02_Karix-%201200x630.png" class="logo">
    </div>
    """,
    unsafe_allow_html=True
)

# ──────────────────────────────────────────────
# PAGE CONFIG
# ──────────────────────────────────────────────
st.set_page_config(
    page_title="SMS Finance Reconciliation",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ──────────────────────────────────────────────
# CUSTOM CSS
# ──────────────────────────────────────────────
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600&family=Syne:wght@700;800&display=swap');

    :root {
        --karix-blue: #132c7a;
        --karix-blue-light: #1d46b3;
        --karix-pink: #c63d8f;
        --bg-light: #f5f7fb;
        --card-bg: #ffffff;
        --border: #d6dbe7;
        --text-dark: #1f2937;
        --text-light: #6b7280;
    }

    html, body, [class*="css"] {
        font-family: 'DM Sans', sans-serif;
        background: var(--bg-light);
        color: var(--text-dark);
    }

    h1, h2, h3 {
        font-family: 'Syne', sans-serif;
    }

    .main {
        background: var(--bg-light);
        color: var(--text-dark);
    }

    /* Header */
    .app-header {
        background: linear-gradient(
            135deg,
            var(--karix-blue) 0%,
            var(--karix-blue-light) 100%
        );
        padding: 1.6rem 2rem;
        border-radius: 0 0 18px 18px;
        margin-bottom: 1.5rem;
        box-shadow: 0 4px 14px rgba(0,0,0,0.08);
    }

    .app-title {
        font-family: 'Syne', sans-serif;
        font-size: 2rem;
        font-weight: 800;
        color: white;
        letter-spacing: 0.5px;
    }

    .app-sub {
        color: #dbe5ff;
        font-size: 0.92rem;
        margin-top: 0.2rem;
    }

    /* Tabs */
    .stTabs [data-baseweb="tab-list"] {
        gap: 6px;
        border-bottom: 2px solid var(--border);
    }

    .stTabs [data-baseweb="tab"] {
        background: white;
        border: 1px solid var(--border);
        border-radius: 10px 10px 0 0;
        color: var(--text-light);
        font-size: 0.9rem;
        font-weight: 500;
        padding: 0.55rem 1.2rem;
    }

    .stTabs [aria-selected="true"] {
        background: #eef3ff;
        border-color: var(--karix-blue);
        color: var(--karix-blue);
        font-weight: 600;
    }

    /* Dataframe */
    .stDataFrame {
        border: 1px solid var(--border);
        border-radius: 12px;
        overflow: hidden;
    }

    /* Buttons */
    .stDownloadButton button,
    .stButton button {
        background: var(--karix-blue);
        color: white;
        border: none;
        border-radius: 8px;
        font-weight: 600;
        padding: 0.55rem 1.4rem;
        transition: 0.2s ease;
    }

    .stDownloadButton button:hover,
    .stButton button:hover {
        background: var(--karix-pink);
        transform: translateY(-1px);
    }

    /* File uploader */
    div[data-testid="stFileUploader"] {
        background: white;
        border: 2px dashed var(--karix-blue-light);
        border-radius: 14px;
        padding: 1.2rem;
    }

    /* Cards / metric containers */
    div[data-testid="metric-container"] {
        background: white;
        border: 1px solid var(--border);
        padding: 1rem;
        border-radius: 14px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.04);
    }

    /* Sidebar */
    section[data-testid="stSidebar"] {
        background: white;
        border-right: 1px solid var(--border);
    }

    </style>
    """,
    unsafe_allow_html=True,
)
st.markdown("""
<style>
.app-header{
    display:flex;
    flex-direction:column;
    align-items:flex-start;
    gap:4px;
    padding:10px 0;
}

.app-title{
    font-size:32px;
    font-weight:700;
}

.app-sub{
    font-size:15px;
    color:#666;
}
</style>
""", unsafe_allow_html=True)

st.markdown(
    """
    <div class="app-header">
      <div class="app-title">📊 SMS Finance Reconciliation</div>
      <div class="app-sub">Reconcile Finance vs Strategy — Volume · Revenue · Rates · FX Conversion</div>
    </div>
    """,
    unsafe_allow_html=True,
)

# ──────────────────────────────────────────────
# HELPERS
# ──────────────────────────────────────────────

def clean_country(x):
    if pd.isna(x):
        return np.nan
    x = str(x).upper().strip()
    mapping = {
        "UAE":   "UNITED ARAB EMIRATES",
        "USA":   "UNITED STATES",
        "US":    "UNITED STATES",
        "UK":    "UNITED KINGDOM",
        "INDIA": "INDIA",
    }
    return mapping.get(x, x)


def _safe_val(val):
    """Convert NaN / Inf to None so xlsxwriter never calls write_number() on them."""
    if val is None:
        return None
    try:
        if isinstance(val, float) and (np.isnan(val) or np.isinf(val)):
            return None
    except Exception:
        pass
    return val


# ──────────────────────────────────────────────
# CORE ANALYSIS
# ──────────────────────────────────────────────

def run_analysis(fin, sms_data, sms_nld, sms_ild):
    """Returns (summary, rate_df, final, ex_issues)."""

    for df in [fin, sms_data, sms_nld, sms_ild]:
        df.columns = df.columns.str.strip()

    # ── Finance SMS ──
    fin["ESME_prefix"] = (
        fin["ESME"].astype(str)
        .str.replace(".0", "", regex=False)
        .str[:6]
    )
    fin["Country_clean"] = (
        fin["DESCRIPTION"].astype(str)
        .str.upper().str.strip()
        .apply(clean_country)
    )
    if "Media" in fin.columns:
        fin.loc[fin["Media"].astype(str).str.strip().str.upper() == "SMS", "Country_clean"] = "INDIA"
    for c in ["Qty", "INR Amount", "Rate", "Amount", "Ex Rate", "INR Rate"]:
        if c in fin.columns:
            fin[c] = pd.to_numeric(fin[c], errors="coerce")
    fin = fin[fin["Qty"].notna() & fin["INR Amount"].notna()].copy()

    # ── Ex Rate Validation ──
    ex_issues = pd.DataFrame()
    if all(c in fin.columns for c in ["Amount", "Ex Rate", "INR Amount"]):
        fin["Expected_INR"] = (fin["Amount"] * fin["Ex Rate"]).round(2)
        fin["INR_Diff"]     = (fin["INR Amount"] - fin["Expected_INR"]).round(2)
        fin["INR_Diff_Pct"] = np.where(
            fin["Expected_INR"].abs() > 0,
            ((fin["INR_Diff"] / fin["Expected_INR"].abs()) * 100).round(2),
            np.nan,
        )
        TOLERANCE = 0.5   # ₹
        keep = [c for c in [
            "Entity", "Invoice Month", "CUSTOMERACCOUNT", "FREETEXTNUMBER",
            "ESME", "Group Name", "Sales Manager", "Country_clean",
            "Currency", "Qty", "Rate", "Amount", "Ex Rate",
            "INR Rate", "INR Amount", "Expected_INR", "INR_Diff", "INR_Diff_Pct",
        ] if c in fin.columns]
        ex_issues = fin[fin["INR_Diff"].abs() > TOLERANCE][keep].copy()
        ex_issues = ex_issues.rename(columns={"Country_clean": "Country"})

    # ── SMS_Data ──
    sms_data["ESME_prefix"]   = sms_data["ESMEADDR"].astype(str).str.replace(".0", "", regex=False).str[:6]
    sms_data["Country_clean"] = sms_data["COUNTRY_KARIX"].apply(clean_country)
    sms_data["Volume"]        = pd.to_numeric(sms_data["BILLABLE"],          errors="coerce")
    sms_data["Revenue"]       = pd.to_numeric(sms_data["BR_TOTAL__REVENUE"], errors="coerce")
    sms_data["Rate"]          = pd.to_numeric(sms_data["BR__RATE"],          errors="coerce")
    sms_data["Currency"]      = sms_data["BR__CURRENCY"]
    sms_data["Parent Org"]    = sms_data["PARENT_ORGNAME"].astype(str).str.strip()

    # ── SMS_NLD ──
    sms_nld["ESME_prefix"]   = sms_nld["ESMEADDR_SUPERADMIN"].astype(str).str.replace(".0", "", regex=False).str[:6]
    sms_nld["Country_clean"] = "INDIA"
    sms_nld["Volume"]        = pd.to_numeric(sms_nld["BILLABLE_CNT"],      errors="coerce") * 1_000_000
    sms_nld["Revenue"]       = pd.to_numeric(sms_nld["BR_TOTAL__REVENUE"], errors="coerce")
    sms_nld["Rate"]          = pd.to_numeric(sms_nld["BR_RATE"],           errors="coerce")
    sms_nld["Currency"]      = sms_nld["BR_CURRENCY"]
    sms_nld["Parent Org"]    = sms_nld["PARENT_ORGNAME"].astype(str).str.strip()

    # ── SMS_ILD ──
    sms_ild["ESME_prefix"]   = sms_ild["ESMEADDR_SUPERADMIN"].astype(str).str.replace(".0", "", regex=False).str[:6]
    sms_ild["Country_clean"] = "INDIA"
    sms_ild["Volume"]        = pd.to_numeric(sms_ild["BILLABLE_CNT"], errors="coerce")
    sms_ild["Revenue"]       = pd.to_numeric(sms_ild["TOTAL_REV"],    errors="coerce")
    sms_ild["Rate"]          = pd.to_numeric(sms_ild["BR_RATE"],      errors="coerce")
    sms_ild["Currency"]      = sms_ild["BR_CURRENCY"]
    sms_ild["Parent Org"]    = sms_ild["PARENT_ORGNAME"].astype(str).str.strip()

    # ── Combine strategy ──
    cols = ["ESME_prefix", "Country_clean", "Volume", "Revenue",
            "Rate", "Currency", "Parent Org"]
    strategy = pd.concat(
        [sms_data[cols], sms_nld[cols], sms_ild[cols]],
        ignore_index=True,
    )

    parent_map = (
        strategy[["ESME_prefix", "Parent Org"]]
        .dropna(subset=["Parent Org"])
        .drop_duplicates("ESME_prefix")
        .set_index("ESME_prefix")["Parent Org"]
        .to_dict()
    )

    # ── Rate comparison — Parent Org × Country wise ──

    # Map ESME → Parent Org on Finance side using strategy parent_map + Group Name fallback
    fin["Parent Org"] = fin["ESME_prefix"].map(parent_map)
    if "Group Name" in fin.columns:
        fin["Parent Org"] = fin["Parent Org"].fillna(fin["Group Name"].astype(str).str.strip())
    fin["Parent Org"] = fin["Parent Org"].fillna(fin["ESME_prefix"])

    fin_rate = (
        fin.groupby(["Parent Org", "Country_clean"])
        .agg(Rate=("Rate", "mean"), Currency=("Currency", "first"))
        .reset_index()
        .rename(columns={"Rate": "Finance Rate", "Currency": "Finance Currency"})
    )

    # Strategy: attach Parent Org then group by Parent Org + Country
    sms_data["Parent Org"] = sms_data["PARENT_ORGNAME"].astype(str).str.strip()
    sms_nld["Parent Org"]  = sms_nld["PARENT_ORGNAME"].astype(str).str.strip()
    sms_ild["Parent Org"]  = sms_ild["PARENT_ORGNAME"].astype(str).str.strip()

    str_rate_data = (
        sms_data[["Parent Org", "Country_clean", "Rate", "Currency"]]
        .groupby(["Parent Org", "Country_clean"])
        .agg(Rate=("Rate", "mean"), Currency=("Currency", "first"))
        .reset_index()
    )
    str_rate_nld = (
        sms_nld[["Parent Org", "Rate", "Currency"]]
        .assign(Country_clean="INDIA")
        .groupby(["Parent Org", "Country_clean"])
        .agg(Rate=("Rate", "mean"), Currency=("Currency", "first"))
        .reset_index()
    )
    str_rate_ild = (
        sms_ild[["Parent Org", "Rate", "Currency"]]
        .assign(Country_clean="INDIA")
        .groupby(["Parent Org", "Country_clean"])
        .agg(Rate=("Rate", "mean"), Currency=("Currency", "first"))
        .reset_index()
    )

    str_rate_all = (
        pd.concat([str_rate_data, str_rate_nld, str_rate_ild], ignore_index=True)
        .groupby(["Parent Org", "Country_clean"])
        .agg(Rate=("Rate", "mean"), Currency=("Currency", "first"))
        .reset_index()
        .rename(columns={"Rate": "Strategy Rate", "Currency": "Strategy Currency"})
    )

    rate_df = pd.merge(fin_rate, str_rate_all,
                       on=["Parent Org", "Country_clean"], how="outer")

    rate_df = rate_df.rename(columns={"Country_clean": "Country"})

    def compare_rates(row):
        f_missing = pd.isna(row["Finance Rate"])
        s_missing = pd.isna(row["Strategy Rate"])
        if f_missing and not s_missing:
            return pd.Series([np.nan, row["Strategy Rate"], np.nan, "No Rate in Finance"])
        if s_missing and not f_missing:
            return pd.Series([row["Finance Rate"], np.nan, np.nan, "No Rate in Strategy"])
        try:
            f  = round(float(row["Finance Rate"]),  4)
            s  = round(float(row["Strategy Rate"]), 4)
            mp = (min(f, s) / max(f, s)) * 100 if max(f, s) != 0 else 100.0
            return pd.Series([f, s, round(mp, 2),
                               "Match" if mp >= 95 else "Mismatch (<95%)"])
        except Exception:
            return pd.Series([row["Finance Rate"], row["Strategy Rate"],
                               np.nan, "No Rate in Finance/Strategy"])

    rate_df[["Finance Rate", "Strategy Rate", "Match %", "Status"]] = \
        rate_df.apply(compare_rates, axis=1)

    # Drop rows where both Finance and Strategy rates are blank
    rate_df = rate_df[
        ~(rate_df["Finance Rate"].isna() & rate_df["Strategy Rate"].isna())
    ].copy()

    # Reorder columns
    col_order = [c for c in [
        "Parent Org", "Country",
        "Finance Rate", "Finance Currency",
        "Strategy Rate", "Strategy Currency",
        "Match %", "Status",
    ] if c in rate_df.columns]
    rate_df = rate_df[col_order].sort_values(["Parent Org", "Country"]).reset_index(drop=True)

    # ── Volume & Revenue ──
    fin_vr = (
        fin.groupby("ESME_prefix")
        .agg(Volume=("Qty", "sum"), Revenue=("INR Amount", "sum"))
        .reset_index()
    )
    fin_vr["Finance Volume"]  = (fin_vr["Volume"]  / 1_000_000).round(2)
    fin_vr["Finance Revenue"] = (fin_vr["Revenue"] / 10_000_000).round(2)

    str_vr = (
        strategy.groupby("ESME_prefix")
        .agg(Volume=("Volume", "sum"), Revenue=("Revenue", "sum"))
        .reset_index()
    )
    str_vr["Strategy Volume"]  = (str_vr["Volume"]  / 1_000_000).round(2)
    str_vr["Strategy Revenue"] = (str_vr["Revenue"] / 10_000_000).round(2)

    final = pd.merge(fin_vr, str_vr, on="ESME_prefix", how="outer")
    for c in ["Finance Volume", "Finance Revenue",
              "Strategy Volume", "Strategy Revenue"]:
        final[c] = final[c].fillna(0).round(2)
    final = final.rename(columns={"ESME_prefix": "Super ESME"})

    fin_parent = (
        fin.groupby("ESME_prefix")["Group Name"].first().to_dict()
        if "Group Name" in fin.columns else {}
    )
    final["Parent Org"] = final["Super ESME"].map(parent_map)
    final["Parent Org"] = final["Parent Org"].replace(["", "nan", "None"], np.nan)
    final["Parent Org"] = final["Parent Org"].fillna(
        final["Super ESME"].map(fin_parent)
    )
    final["Parent Org"] = final["Parent Org"].fillna(final["Super ESME"])
    final = final[[
        "Super ESME", "Parent Org",
        "Finance Volume", "Finance Revenue",
        "Strategy Volume", "Strategy Revenue",
    ]]

    # ── Remarks for ESME-level detail (reuse rate_map built later, stub for now) ──
    # Will be computed after rate_map is available; attach after summary block below.

    # ── Summary ──
    summary = final.copy()
    summary["Parent Org"] = (
        summary["Parent Org"]
        .replace(["", "nan", "None"], np.nan)
        .fillna(summary["Super ESME"])
    )
    summary = (
        summary.groupby("Parent Org", as_index=False)
        .agg(
            Finance_Volume=("Finance Volume",   "sum"),
            Finance_Revenue=("Finance Revenue", "sum"),
            Strategy_Volume=("Strategy Volume", "sum"),
            Strategy_Revenue=("Strategy Revenue", "sum"),
        )
    )
    summary = summary.rename(columns={
        "Finance_Volume":   "Finance Volume",
        "Finance_Revenue":  "Finance Revenue",
        "Strategy_Volume":  "Strategy Volume",
        "Strategy_Revenue": "Strategy Revenue",
    })
    for c in ["Finance Volume", "Finance Revenue",
              "Strategy Volume", "Strategy Revenue"]:
        summary[c] = summary[c].round(2)

    summary["Vol. Diff"] = (
        summary["Finance Volume"] - summary["Strategy Volume"]
    ).round(2)
    summary["Rev. Diff"] = (
        summary["Finance Revenue"] - summary["Strategy Revenue"]
    ).round(2)

    # Rate-mismatch map for remarks
    rate_mismatch = rate_df[
        (rate_df["Status"] == "Mismatch (<95%)") &
        rate_df["Finance Rate"].notna() &
        rate_df["Strategy Rate"].notna()
    ].copy()
    if not rate_mismatch.empty:
        rate_mismatch["Rate Detail"] = rate_mismatch.apply(
            lambda x: (
                f"{x['Country']} "
                f"(Fin: {round(float(x['Finance Rate']),4)} {x.get('Finance Currency','')} "
                f"vs Str: {round(float(x['Strategy Rate']),4)} {x.get('Strategy Currency','')})"
            ), axis=1,
        )
        rate_map = (
            rate_mismatch.groupby("Parent Org")["Rate Detail"]
            .apply(lambda x: " | ".join(sorted(set(x))))
            .to_dict()
        )
    else:
        rate_map = {}

    VOL_TH = 0.01
    REV_TH = 0.001

    def build_remarks(row):
        if row["Finance Volume"] == 0 and row["Strategy Volume"] == 0:
            return "No Traffic"
        if abs(row["Vol. Diff"]) <= VOL_TH and abs(row["Rev. Diff"]) <= REV_TH:
            return "Match"
        r = []
        if abs(row["Vol. Diff"]) > VOL_TH:
            if row["Finance Volume"] == 0 and row["Strategy Volume"] > 0:
                r.append("No Reporting of Volume")
            elif row["Finance Volume"] < row["Strategy Volume"]:
                r.append("Under Reporting of Volume")
            else:
                r.append("Over Reporting of Volume")
        if abs(row["Rev. Diff"]) > REV_TH:
            mm = rate_map.get(row["Parent Org"])
            r.append(f"Rate Mismatch — {mm}" if mm else "Revenue Mismatch")
        return " | ".join(r)

    summary["Remarks"] = summary.apply(build_remarks, axis=1)
    summary = summary[[
        "Parent Org",
        "Finance Volume",  "Finance Revenue",
        "Strategy Volume", "Strategy Revenue",
        "Vol. Diff", "Rev. Diff",
        "Remarks",
    ]].sort_values("Vol. Diff").reset_index(drop=True)

    # ── ESME-level Remarks for Vol & Revenue detail tab ──
    final["Vol. Diff"] = (final["Finance Volume"] - final["Strategy Volume"]).round(2)
    final["Rev. Diff"] = (final["Finance Revenue"] - final["Strategy Revenue"]).round(2)

    def build_esme_remarks(row):
        if row["Finance Volume"] == 0 and row["Strategy Volume"] == 0:
            return "No Traffic"
        if abs(row["Vol. Diff"]) <= VOL_TH and abs(row["Rev. Diff"]) <= REV_TH:
            return "Match"
        r = []
        if abs(row["Vol. Diff"]) > VOL_TH:
            if row["Finance Volume"] == 0 and row["Strategy Volume"] > 0:
                r.append("No Reporting of Volume")
            elif row["Finance Volume"] < row["Strategy Volume"]:
                r.append("Under Reporting of Volume")
            else:
                r.append("Over Reporting of Volume")
        if abs(row["Rev. Diff"]) > REV_TH:
            mm = rate_map.get(row["Parent Org"])
            r.append(f"Rate Mismatch — {mm}" if mm else "Revenue Mismatch")
        return " | ".join(r)

    final["Remarks"] = final.apply(build_esme_remarks, axis=1)
    final = final[[
        "Super ESME", "Parent Org",
        "Finance Volume", "Finance Revenue",
        "Strategy Volume", "Strategy Revenue",
        "Vol. Diff", "Rev. Diff",
        "Remarks",
    ]]

    return summary, rate_df, final, ex_issues


# ──────────────────────────────────────────────
# SAMPLE DATA
# ──────────────────────────────────────────────

def generate_sample_excel() -> bytes:
    buf = io.BytesIO()

    fin_cols = [
        "Entity", "Invoice Month", "CUSTOMERACCOUNT", "FREETEXTNUMBER",
        "Proforma", "Date", "ESME", "User Name", "Group Name", "Sales Manager",
        "Revenue Month 2", "Revenue Month", "Currency", "Qty", "Rate",
        "Amount", "Ex Rate", "INR Rate", "INR Amount",
        "GL", "Media", "GL Code", "DESCRIPTION", "Media2", "GL Code3",
    ]
    fin_data = [
        ["TANLA", "2024-04", "CUST001", "FTN001", "PRO001", "2024-04-01",
         100101, "Alice", "Group A", "SM1", "Apr-24", "Apr-24",
         "USD", 50000, 0.0050, 250.0, 83.50, 83.50, 20875.0,
         "GL1", "SMS", "GC1", "UNITED STATES", "SMS2", "GC3"],
        ["TANLA", "2024-04", "CUST002", "FTN002", "PRO002", "2024-04-01",
         200202, "Bob", "Group B", "SM2", "Apr-24", "Apr-24",
         "AED", 30000, 0.0080, 240.0, 22.70, 22.70, 5448.0,
         "GL2", "SMS", "GC2", "UNITED ARAB EMIRATES", "SMS2", "GC3"],
        # Deliberate Ex Rate mismatch: 60 × 83.50 = 5010, but INR Amount = 999
        ["TANLA", "2024-04", "CUST003", "FTN003", "PRO003", "2024-04-02",
         300303, "Carol", "Group C", "SM3", "Apr-24", "Apr-24",
         "USD", 10000, 0.0060, 60.0, 83.50, 83.50, 999.0,
         "GL3", "SMS", "GC3", "INDIA", "SMS2", "GC3"],
        ["TANLA", "2024-04", "CUST004", "FTN004", "PRO004", "2024-04-03",
         400404, "Dave", "Group A", "SM1", "Apr-24", "Apr-24",
         "INR", 80000, 0.0045, 360.0, 1.0, 1.0, 360.0,
         "GL1", "SMS", "GC1", "INDIA", "SMS2", "GC3"],
        # NLD ESME with different rate in Finance vs NLD
        ["TANLA", "2024-04", "CUST005", "FTN005", "PRO005", "2024-04-04",
         500505, "Eve", "Group D", "SM4", "Apr-24", "Apr-24",
         "INR", 20000, 0.0055, 110.0, 1.0, 1.0, 110.0,
         "GL1", "SMS", "GC1", "INDIA", "SMS2", "GC3"],
        # ILD ESME with deliberate rate mismatch
        ["TANLA", "2024-04", "CUST006", "FTN006", "PRO006", "2024-04-05",
         600606, "Frank", "Group E", "SM5", "Apr-24", "Apr-24",
         "INR", 15000, 0.0090, 135.0, 1.0, 1.0, 135.0,
         "GL2", "SMS", "GC2", "INDIA", "SMS2", "GC3"],
    ]
    fin_df = pd.DataFrame(fin_data, columns=fin_cols)

    sms_data_df = pd.DataFrame([
        [100101, "United States", 50000, 250.0,  0.0050, "USD", "Group A"],
        [200202, "UAE",           30000, 240.0,  0.0080, "AED", "Group B"],
        [300303, "India",         10000,  60.0,  0.0060, "USD", "Group C"],
    ], columns=["ESMEADDR", "COUNTRY_KARIX", "BILLABLE",
                "BR_TOTAL__REVENUE", "BR__RATE", "BR__CURRENCY", "PARENT_ORGNAME"])

    sms_nld_df = pd.DataFrame([
        [400404, 0.08, 360.0, 0.0045, "INR", "Group A"],
        [100101, 0.05, 125.0, 0.0050, "USD", "Group A"],
        # NLD ESME 500505 with a different rate → should show mismatch vs Finance 0.0055
        [500505, 0.02,  90.0, 0.0030, "INR", "Group D"],
    ], columns=["ESMEADDR_SUPERADMIN", "BILLABLE_CNT",
                "BR_TOTAL__REVENUE", "BR_RATE", "BR_CURRENCY", "PARENT_ORGNAME"])

    sms_ild_df = pd.DataFrame([
        [200202, 5000,  40.0, 0.0080, "AED", "Group B"],
        # ILD ESME 600606 with different rate → should show mismatch vs Finance 0.0090
        [600606, 3000, 180.0, 0.0060, "INR", "Group E"],
    ], columns=["ESMEADDR_SUPERADMIN", "BILLABLE_CNT",
                "TOTAL_REV", "BR_RATE", "BR_CURRENCY", "PARENT_ORGNAME"])

    with pd.ExcelWriter(buf, engine="xlsxwriter") as w:
        fin_df.to_excel(     w, sheet_name="Finance SMS", index=False)
        sms_data_df.to_excel(w, sheet_name="SMS_Data",    index=False)
        sms_nld_df.to_excel( w, sheet_name="SMS_NLD",     index=False)
        sms_ild_df.to_excel( w, sheet_name="SMS_ILD",     index=False)

    return buf.getvalue()


# ──────────────────────────────────────────────
# EXCEL EXPORT  ← NaN / Inf safe
# ──────────────────────────────────────────────

def export_to_excel(summary, rate_df, final, ex_issues) -> bytes:
    buf = io.BytesIO()

    with pd.ExcelWriter(
        buf,
        engine="xlsxwriter",
        engine_kwargs={"options": {"nan_inf_to_errors": True}},
    ) as writer:
        wb = writer.book

        # ── Formats ──
        H  = wb.add_format({"bold": True, "bg_color": "#1F4E79",
                             "font_color": "white", "border": 1, "align": "center"})
        N2 = wb.add_format({"num_format": "#,##0.00",   "border": 1})
        N4 = wb.add_format({"num_format": "#,##0.0000", "border": 1})
        T  = wb.add_format({"border": 1})
        MF = wb.add_format({"bg_color": "#C6EFCE", "font_color": "#006100", "border": 1})
        UF = wb.add_format({"bg_color": "#FFEB9C", "font_color": "#9C6500", "border": 1})
        OF = wb.add_format({"bg_color": "#DDEBF7", "font_color": "#1F4E79", "border": 1})
        NF = wb.add_format({"bg_color": "#FFC7CE", "font_color": "#9C0006", "border": 1})
        RF = wb.add_format({"bg_color": "#FCE4D6", "font_color": "#833C00", "border": 1})
        EF = wb.add_format({"bg_color": "#FFD7D7", "font_color": "#C00000",
                             "num_format": "#,##0.00",  "border": 1})
        # Rate mismatch highlight for Rate Comparison sheet
        NTF = wb.add_format({"bg_color": "#D9D9D9", "font_color": "#595959", "border": 1})  # grey — No Traffic
        RMF = wb.add_format({"bg_color": "#FFC7CE", "font_color": "#9C0006",
                              "num_format": "#,##0.0000", "border": 1})   # red  — rate mismatch
        NRF = wb.add_format({"bg_color": "#FFF2CC", "font_color": "#7F6000", "border": 1})  # amber — no rate
        MMF = wb.add_format({"bg_color": "#C6EFCE", "font_color": "#006100", "border": 1})  # green — match

        rev_cols = {
            "Finance Volume", "Finance Revenue",
            "Strategy Volume", "Strategy Revenue",
            "Vol. Diff", "Rev. Diff",
            "INR_Diff", "INR_Diff_Pct",
            "Amount", "INR Rate", "INR Amount", "Expected_INR",
            "Match %",
        }
        rate_cols = {"Finance Rate", "Strategy Rate", "Rate", "Ex Rate"}

        def rfmt(v):
            v = str(v) if v is not None else ""
            if v == "No Traffic":          return NTF
            if "No Reporting"    in v: return NF
            if "Under Reporting" in v: return UF
            if "Over Reporting"  in v: return OF
            if "Rate Mismatch"   in v: return RF
            if v == "Match":           return MF
            return T

        widths = {
            "Parent Org": 40, "Country": 24,
            "Finance Volume": 16,  "Finance Revenue": 16,
            "Strategy Volume": 16, "Strategy Revenue": 16,
            "Vol. Diff": 14,       "Rev. Diff": 14,
            "Finance Rate": 14,    "Strategy Rate": 14,
            "Finance Currency": 14, "Strategy Currency": 14,
            "Match %": 12,         "Status": 20,
            "INR_Diff": 14,        "INR_Diff_Pct": 14,
            "Expected_INR": 16,    "Remarks": 65,
        }

        ex_out = (
            ex_issues if not ex_issues.empty
            else pd.DataFrame(columns=["No Ex Rate issues found"])
        )

        sheets = [
            ("Summary",          summary),
            ("Rate Comparison",  rate_df),
            ("Volume & Revenue", final),
            ("Ex Rate Issues",   ex_out),
        ]

        for sheet_name, df in sheets:
            df_out = df.copy()

            for c in df_out.columns:
                if c in rev_cols and pd.api.types.is_numeric_dtype(df_out[c]):
                    df_out[c] = df_out[c].round(2)

            df_out.to_excel(writer, sheet_name=sheet_name, index=False)
            ws   = writer.sheets[sheet_name]
            cols = list(df_out.columns)

            for ci, col in enumerate(cols):
                ws.write(0, ci, col, H)
                ws.set_column(ci, ci, widths.get(col, 18))

            for ri, row in df_out.iterrows():
                for ci, col in enumerate(cols):
                    val = _safe_val(row[col])

                    if sheet_name == "Summary" and col == "Remarks":
                        fmt = rfmt(val)
                    elif sheet_name in ("Summary", "Volume & Revenue") and col == "Remarks":
                        fmt = rfmt(val)
                    elif sheet_name == "Rate Comparison" and col == "Status":
                        status_str = str(val) if val is not None else ""
                        if status_str == "Match":
                            fmt = MMF
                        elif "Mismatch" in status_str:
                            fmt = NF
                        elif "No Rate" in status_str:
                            fmt = NRF
                        else:
                            fmt = T
                    elif sheet_name == "Ex Rate Issues" and col in ("INR_Diff", "INR_Diff_Pct"):
                        fmt = EF
                    elif col in rate_cols:
                        # Highlight mismatched rate cells in Rate Comparison sheet
                        if sheet_name == "Rate Comparison":
                            status_val = _safe_val(row.get("Status", ""))
                            fmt = RMF if str(status_val) == "Mismatch (<95%)" else N4
                        else:
                            fmt = N4
                    elif col in rev_cols:
                        fmt = N2
                    else:
                        fmt = T

                    if val is None:
                        ws.write_blank(ri + 1, ci, None, fmt)
                    else:
                        ws.write(ri + 1, ci, val, fmt)

            ws.freeze_panes(1, 0)

    return buf.getvalue()


# ──────────────────────────────────────────────
# MAIN UI
# ──────────────────────────────────────────────

tabs = st.tabs([
    "📂 Upload & Run",
    "📋 Summary",
    "📈 Rate Comparison",
    "🗂 Vol & Revenue",
    "💱 Ex Rate Issues",
    "⬇️ Download",
])


def require_results():
    if "results" not in st.session_state:
        st.warning("⬅️ Please upload and process a file in the **Upload & Run** tab first.")
        return None
    return st.session_state["results"]


# ── TAB 0: Upload ──
with tabs[0]:
    st.markdown("### Upload your SMS Finance Excel file")
    st.markdown(
        "The file must contain **4 sheets**: "
        "`Finance SMS`, `SMS_Data`, `SMS_NLD`, `SMS_ILD`."
    )

    col1, col2 = st.columns([2, 1])
    with col1:
        uploaded = st.file_uploader(
            "Drop your Excel file here", type=["xlsx", "xls"]
        )
    with col2:
        st.markdown("<br>", unsafe_allow_html=True)
        if "sample_excel" not in st.session_state:
            st.session_state["sample_excel"] = generate_sample_excel()
        st.download_button(
            "⬇️ Download Sample Format",
            data=st.session_state["sample_excel"],
            file_name="SMS_Finance_Sample.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        st.caption("Sample includes deliberate NLD/ILD rate mismatches and an Ex Rate mismatch row.")

    if uploaded:
        with st.spinner("Reading sheets…"):
            try:
                fin_raw      = pd.read_excel(uploaded, sheet_name="Finance SMS")
                sms_data_raw = pd.read_excel(uploaded, sheet_name="SMS_Data")
                sms_nld_raw  = pd.read_excel(uploaded, sheet_name="SMS_NLD")
                sms_ild_raw  = pd.read_excel(uploaded, sheet_name="SMS_ILD")
                st.success("✅ All 4 sheets loaded successfully.")
            except Exception as e:
                st.error(f"Could not read sheets: {e}")
                st.stop()

        with st.spinner("Running analysis…"):
            try:
                summary, rate_df, final, ex_issues = run_analysis(
                    fin_raw, sms_data_raw, sms_nld_raw, sms_ild_raw
                )
                st.session_state["results"] = (summary, rate_df, final, ex_issues)
            except Exception as e:
                st.error(f"Analysis error: {e}")
                st.stop()

        st.markdown("#### Quick Stats")
        m1, m2, m3, m4, m5, m6 = st.columns(6)
        m1.metric("Finance Volume (Mn)",    f"{final['Finance Volume'].sum():,.2f}")
        m2.metric("Strategy Volume (Mn)",   f"{final['Strategy Volume'].sum():,.2f}")
        m3.metric("Finance Revenue (Cr)",   f"{final['Finance Revenue'].sum():,.2f}")
        m4.metric("Strategy Revenue (Cr)",  f"{final['Strategy Revenue'].sum():,.2f}")
        m5.metric("Rate Mismatches",
                  int((rate_df["Status"] == "Mismatch (<95%)").sum()))
        m6.metric("Ex Rate Issues",         len(ex_issues))
        st.info("👆 Navigate the tabs above to explore results, or go to Download.")
    else:
        st.info("Upload a file above to begin. No file yet? Download the sample format first.")


# ── TAB 1: Summary ──
with tabs[1]:
    res = require_results()
    if res:
        summary, _, _, _ = res
        st.markdown("### Reconciliation Summary")
        total      = len(summary)
        matched    = (summary["Remarks"] == "Match").sum()
        mismatched = total - matched
        c1, c2, c3 = st.columns(3)
        c1.metric("Total Entities", total)
        c2.metric("✅ Match",       matched)
        c3.metric("⚠️ Issues",      mismatched)

        choices = ["All"] + sorted(summary["Remarks"].dropna().unique().tolist())
        flt = st.selectbox("Filter by Remark", choices)
        df_show = summary if flt == "All" else summary[summary["Remarks"] == flt]
        st.dataframe(df_show.reset_index(drop=True),
                     use_container_width=True, height=500)


# ── TAB 2: Rate Comparison ──
with tabs[2]:
    res = require_results()
    if res:
        _, rate_df, _, _ = res
        st.markdown("### Rate Comparison — Finance vs Strategy")
        st.caption(
            "Rates compared at **Parent Org × Country** level. "
            "NLD and ILD rows are grouped under INDIA."
        )
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Total Rows",             len(rate_df))
        c2.metric("✅ Match",               (rate_df["Status"] == "Match").sum())
        c3.metric("❌ Mismatch",            (rate_df["Status"] == "Mismatch (<95%)").sum())
        c4.metric("⚠️ No Rate (one side)",  rate_df["Status"].str.startswith("No Rate").sum())

        flt = st.selectbox("Filter by Status",
                           ["All", "Match", "Mismatch (<95%)",
                            "No Rate in Finance", "No Rate in Strategy",
                            "No Rate in Finance/Strategy"])
        df_show = rate_df if flt == "All" else rate_df[rate_df["Status"] == flt]
        st.dataframe(df_show.reset_index(drop=True),
                     use_container_width=True, height=500)


# ── TAB 3: Volume & Revenue ──
with tabs[3]:
    res = require_results()
    if res:
        _, _, final, _ = res
        st.markdown("### Volume & Revenue Detail (by Super ESME)")
        st.caption("Includes Vol. Diff, Rev. Diff and Remarks per ESME for detailed analysis.")
        search = st.text_input("Search by ESME / Parent Org")
        df_show = final
        if search:
            mask = (
                final["Super ESME"].astype(str).str.contains(search, case=False, na=False) |
                final["Parent Org"].astype(str).str.contains(search, case=False, na=False)
            )
            df_show = final[mask]
        st.dataframe(df_show.reset_index(drop=True),
                     use_container_width=True, height=500)


# ── TAB 4: Ex Rate Issues ──
with tabs[4]:
    res = require_results()
    if res:
        _, _, _, ex_issues = res
        st.markdown("### 💱 Exchange Rate Conversion Issues")
        st.markdown(
            "Rows where **`INR Amount ≠ Amount × Ex Rate`** (tolerance ₹0.50). "
            "Finance has not applied the exchange rate correctly.\n\n"
            "- **Expected INR** = `Amount × Ex Rate`  \n"
            "- **INR Diff** = `INR Amount − Expected INR`  \n"
            "- **INR Diff %** = deviation as % of expected"
        )
        if ex_issues.empty:
            st.success("🎉 No exchange rate conversion issues found.")
        else:
            st.error(f"⚠️ {len(ex_issues)} row(s) with incorrect INR conversion detected.")

            col_order = [c for c in [
                "Entity", "Invoice Month", "Group Name", "Country",
                "ESME", "Currency", "Qty", "Rate",
                "Amount", "Ex Rate", "INR Rate",
                "INR Amount", "Expected_INR", "INR_Diff", "INR_Diff_Pct",
            ] if c in ex_issues.columns]

            fmt_map = {
                "Amount":       "{:,.2f}",
                "INR Amount":   "{:,.2f}",
                "Expected_INR": "{:,.2f}",
                "INR_Diff":     "{:+,.2f}",
                "INR_Diff_Pct": "{:+.2f}%",
                "Ex Rate":      "{:.4f}",
            }
            active_fmt = {k: v for k, v in fmt_map.items() if k in col_order}

            def highlight_diff(val):
                try:
                    if abs(float(val)) > 0.5:
                        return "background-color:#5a1e02; color:#f78166"
                except Exception:
                    pass
                return ""

            styled = (
                ex_issues[col_order]
                .reset_index(drop=True)
                .style.map(highlight_diff, subset=["INR_Diff"])
                .format(active_fmt)
            )
            st.dataframe(styled, use_container_width=True, height=480)

            total_diff = (
                ex_issues["INR_Diff"].sum()
                if "INR_Diff" in ex_issues.columns else 0
            )
            st.metric("Total INR Difference (₹)", f"{total_diff:+,.2f}")


# ── TAB 5: Download ──
with tabs[5]:
    res = require_results()
    if res:
        summary, rate_df, final, ex_issues = res
        st.markdown("### Export Full Analysis to Excel")
        try:
            excel_bytes = export_to_excel(summary, rate_df, final, ex_issues)
            st.download_button(
                "⬇️ Download SMS Finance Analysis.xlsx",
                data=excel_bytes,
                file_name="SMS_Finance_Analysis.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            st.markdown(
                "The Excel contains 4 colour-coded sheets: "
                "**Summary · Rate Comparison · Volume & Revenue · Ex Rate Issues**  \n"
                "- Rate Comparison highlights mismatched rows in red (including NLD/ILD ESMEs)  \n"
                "- Revenue & GM columns are rounded to **2 decimal places**."
            )
        except Exception as e:
            st.error(f"Export failed: {e}")
    else:
        st.info("Upload and process a file first.")
