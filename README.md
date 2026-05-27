# 📊 SMS Finance Analyzer

A Streamlit app to reconcile **Finance vs Strategy** data for SMS billing — covering Volume, Revenue, Rates, and FX conversion checks.

## Features

- **Summary** — Parent Org level reconciliation with remarks (Match / Under Reporting / Over Reporting / Rate Mismatch / No Traffic)
- **Rate Comparison** — Finance vs Strategy rates at Parent Org × Country level
- **Volume & Revenue Detail** — ESME-level breakdown with remarks
- **Ex Rate Issues** — Detects rows where `INR Amount ≠ Amount × Ex Rate`
- **Excel Export** — Colour-coded 4-sheet workbook

## Setup

```bash
pip install -r requirements.txt
streamlit run sms_finance_analyzer.py
```

## Input Format

Upload an Excel file with exactly 4 sheets:

| Sheet | Key Columns |
|-------|-------------|
| `Finance SMS` | ESME, Qty, Rate, Amount, Ex Rate, INR Amount, Currency, DESCRIPTION, Media, Group Name |
| `SMS_Data` | ESMEADDR, COUNTRY_KARIX, BILLABLE, BR_TOTAL__REVENUE, BR__RATE, BR__CURRENCY, PARENT_ORGNAME |
| `SMS_NLD` | ESMEADDR_SUPERADMIN, BILLABLE_CNT, BR_TOTAL__REVENUE, BR_RATE, BR_CURRENCY, PARENT_ORGNAME |
| `SMS_ILD` | ESMEADDR_SUPERADMIN, BILLABLE_CNT, TOTAL_REV, BR_RATE, BR_CURRENCY, PARENT_ORGNAME |

Download the sample format from inside the app to get started.
