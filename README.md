# Technical Loss Automation Simulator

A Streamlit app that computes **feeder-wise technical loss** by comparing
feeder-head energy import against the sum of consumer billing consumption,
and flags data-quality issues that distort the loss number.

## What it does

```
Loss (kWh) = Total Feeder Consumption − Total Consumer Consumption
Loss (%)   = Loss (kWh) / Total Feeder Consumption × 100
```

For each feeder, the app:
1. Loads the feeder-head energy import for the month (file upload or API fetch).
2. Loads the WC installation master to map consumers → feeder codes.
3. Loads current + previous month billing reports and computes each
   consumer's monthly delta (`net_energy_current − net_energy_previous`).
4. Sums consumer consumption per feeder and subtracts it from feeder
   consumption to get loss, in kWh and %.
5. Surfaces an expandable Excel export (feeder rows collapse to show
   their consumers) plus per-feeder stats in the UI.

## Data-quality flags

Because loss is a residual, bad data inflates or deflates it silently.
The app tags and reports four causes separately so you can audit which
one is driving a feeder's number, rather than just trusting the total:

| Flag | Cause | Effect on loss |
|---|---|---|
| **Unmatched to Feeder** | Billed consumer not found in the installation master for any feeder | Consumption excluded → loss inflated |
| **No Previous Reading** | No prior-month bill to diff against, delta is `NaN` | Consumption dropped → loss inflated |
| **Exporting Consumers** | Current reading < previous reading (meter swap/rollover) | Consumption goes negative → loss deflated |
| **Installed, Never Billed** | In the installation master but absent from this month's billing entirely | Not counted anywhere, reported for audit |

## Inputs

| File | Format | Notes |
|---|---|---|
| Installation Master (WC) | CSV | Up to ~1.6 GB, read in chunks. Needs `Feeder Code`, `Feeder Name`, `Consumer Number (KNO)`, `DT Code` |
| Feeder Consumption | CSV or JSON, one per month | Or fetched live from an API endpoint (see below) |
| Billing Cycle Report | CSV, one per month | Needs current **and** previous month for the delta calc |

### Feeder Consumption via API

Instead of uploading files, you can fetch feeder consumption directly:
paste a URL template containing `{year}`, `{month}`, `{subdivision}`
placeholders (as path segments or query params), pick a year and one or
more months, and the app calls the endpoint and parses the same
`{"MonthlyEnergy": [...]}` shape as the file upload path. Never paste an
API key/token into the URL field — the app has a spot to add an auth
header if your endpoint needs one.

## Output

A single Excel workbook (`feeder_wise_loss_<month>.xlsx`) with:
- **Feeder-wise Report** — summary sheet, one row per feeder, plus an
  appended section listing exporting (negative-delta) consumers.
- **Feeder + Consumer Detail** — the same feeder rows, each with its
  consumers grouped underneath and collapsed (click `+` to expand).
- **Flag – Unmatched to Feeder / No Previous Reading / Exporting
  Consumers / Installed Never Billed** — one sheet per flag, for
  root-causing a specific feeder's number.

## Setup

```bash
pip install streamlit pandas requests openpyxl
streamlit run technical_loss_app.py
```

A `.streamlit/config.toml` ships alongside the app to raise Streamlit's
upload cap (needed for the ~1.6 GB installation file):

```toml
[server]
maxUploadSize = 2000
```

## Usage

1. Upload the installation master CSV.
2. Upload feeder consumption file(s) (or fetch via API) and label each by month.
3. Upload billing report CSVs for the current and previous month, labeled by month.
4. Pick the current feeder-consumption month, current billing month, and
   previous billing month (baseline for the delta).
5. Click **Run Loss Calculation**.
6. Review the summary metrics, per-feeder table, and flag counts in the UI.
7. Select a feeder from the dropdown for its individual stats.
8. Download the full Excel report.

## Tech stack

Python, Streamlit, pandas, openpyxl, requests.

---
Genus Jodhpur · [github.com/rajvyas72002](https://github.com/rajvyas72002)
