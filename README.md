# Giordano's DTC Daily Revenue Reporting

This repository generates a daily Excel report for Giordano's DTC business using Shopify, Klaviyo, and Meta Ads API data.

The report is written to:

```text
reports/daily_revenue.xlsx
```

The workflow is designed to run in GitHub Actions every day and commit the updated Excel file back into the repository, so the latest report can be accessed from the repo locally after a `git pull`.

## Repository Structure

```text
.
├── .github/workflows/daily_report.yml
├── data/
│   └── daily_revenue.csv
├── reports/
│   └── daily_revenue.xlsx
├── scripts/
│   └── daily_revenue_report.py
├── requirements.txt
└── README.md
```

## Report Output

The Excel workbook contains three sheets:

- `Daily Performance`: daily normalized performance by channel
- `Summary`: totals, channel mix, ROAS, AOV, and variance
- `Sources & Notes`: source definitions and calculation notes

The daily table includes:

- `date`
- `Shopify revenue`
- `Meta revenue`
- `Email revenue`
- `SMS revenue`
- `Meta spend`
- `Orders`
- `AOV`
- `Owned revenue`
- `Other / unattributed revenue`
- `Target revenue`
- `Variance vs target`
- `Meta purchases`

## Required GitHub Secrets

In GitHub, go to:

`Settings` → `Secrets and variables` → `Actions` → `Secrets` → `New repository secret`

Add these secrets:

```text
SHOPIFY_API_KEY
SHOPIFY_PASSWORD
KLAVIYO_API_KEY
META_ACCESS_TOKEN
```

Notes:

- For Shopify custom apps, `SHOPIFY_PASSWORD` can be the Admin API access token such as `shpat_...`.
- If using an older private app, use the private app API key and password.
- Secrets should never be committed to the repo.

## Optional GitHub Variables

In GitHub, go to:

`Settings` → `Secrets and variables` → `Actions` → `Variables` → `New repository variable`

Optional variables:

```text
SHOPIFY_STORE=giordanos-frozen-pizza.myshopify.com
META_AD_ACCOUNT=act_1271027757410529
REPORT_START_DATE=2026-03-03
REPORT_END_DATE=2026-04-23
DAILY_REVENUE_TARGET=30000
KLAVIYO_CONVERSION_METRIC_ID=
```

If `REPORT_END_DATE` is omitted, the script uses yesterday in `America/Chicago`.

If `KLAVIYO_CONVERSION_METRIC_ID` is omitted, the script looks for the Shopify `Placed Order` metric in Klaviyo.

## Running Locally

Install dependencies:

```bash
python -m pip install -r requirements.txt
```

Set environment variables.

PowerShell:

```powershell
$env:SHOPIFY_API_KEY="your-shopify-api-key"
$env:SHOPIFY_PASSWORD="your-shopify-password-or-admin-api-token"
$env:KLAVIYO_API_KEY="your-klaviyo-api-key"
$env:META_ACCESS_TOKEN="your-meta-access-token"

# Optional
$env:SHOPIFY_STORE="giordanos-frozen-pizza.myshopify.com"
$env:META_AD_ACCOUNT="act_1271027757410529"
$env:REPORT_START_DATE="2026-03-03"
$env:DAILY_REVENUE_TARGET="30000"
```

Run the report:

```bash
python scripts/daily_revenue_report.py
```

The script writes:

```text
reports/daily_revenue.xlsx
data/daily_revenue.csv
```

## Triggering Manually in GitHub

Go to:

`Actions` → `Daily Revenue Report` → `Run workflow`

This runs the same process as the scheduled job and commits the updated report if the file changed.

## Schedule

The workflow runs daily using this cron expression:

```yaml
cron: "0 14 * * *"
```

GitHub Actions cron is UTC. `14:00 UTC` is 8:00 AM CST. During daylight saving time, that same UTC schedule runs at 9:00 AM CDT unless the workflow cron is adjusted seasonally.

## Calculations

```text
Owned revenue = Email revenue + SMS revenue
Other / unattributed revenue = Shopify revenue - Meta revenue - Email revenue - SMS revenue
Variance vs target = Shopify revenue - DAILY_REVENUE_TARGET
```

Attribution can overlap across platforms, so `Other / unattributed revenue` should be treated as a directional residual, not a perfect source-of-truth attribution bucket.
