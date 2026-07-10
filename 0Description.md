## Workflow

Ghseet Link: https://docs.google.com/spreadsheets/d/1Pcq9g05IxxtKSkdYEn19FSZkNd1J0y30LzQZ7k4lwJU/edit?gid=1920223642#gid=1920223642

## Owner / PIC
- Michelle — michellemy0051@kzgroup.biz (owner)
- Kin — pham0003@kzgroup.biz (backup PIC)

## Purpose
This Apps Script project identifies and reports potentially suspicious players in Thailand (Timezone: GMT+7). It runs BigQuery-based analytics, computes percentile thresholds, and writes filtered and aggregated results into Google Sheets for operator review. All amounts are expressed in THB.

Primary filters used to surface users include:
- Low win/loss ratio (example threshold: ≤ 0.1)
- Low recent turnover (past 2 months) relative to thresholds computed from percentiles

Additional risk indicators include:
- High number of successful deposits (> 10) with little betting activity
- High deposit attempts (> 20) combined with a low deposit success ratio (≤ 0.1)
- High account registration rate in a short time (e.g., ≥ 10 registrations/day)
- High number of distinct registered brands

Result prioritisation:
1. Users with lifetime turnover (`ltTurnover`) below 1 are listed first.
2. Within the same `ltTurnover` bucket, users are ordered by recent turnover (ascending).

## Suspicious-player categorisation logic

| Category | When to use | Core signals | Example criteria |
|---|---:|---|---|
| Watchlist | High deposit attempts but little successful play | deposit attempts high, low deposit success, low turnover | deposit attempts ≥ 20; deposit success ratio ≤ 10%; turnover ≤ 1; win/loss ≤ 10% |
| Watchlist | Many brands but low activity | brand count high, low turnover | brand count ≥ 25; turnover ≤ 1; win/loss ≤ 10% |
| Suspicious | Deposited but minimal betting | deposit amount > 0, very low bets | deposit amount > 0; turnover ≤ 1.5; win/loss ≤ 5% |
| Suspicious | Low deposit success rate with many timeouts | low success ratio, timeout-dominated failures | deposit success ratio ≤ 10%; ≥ 80% failed deposits are timeouts; turnover ≤ 1.5 |
| Suspicious | Many brands but low betting | high brand usage, low play | brand count ≥ 20; turnover ≤ 1.5; win/loss ≤ 5% |

## Workflow
- Source: BigQuery realtime tables (e.g. kz-dp-prod.kz_pg_to_bq_realtime.*) queried from Apps Script.
- Processing: Apps Script SQL runners compute percentiles, aggregate lifetime metrics, and classify users.
- Output: Results written to Google Sheets tabs for operator review and further action.

## Tabs (important)
- `Suspicious_Player` — reduced reporting view of flagged users and status cells used by operators.
- `Percentile_phone_email` — percentile thresholds and summary metrics used to compute decision cutoffs.
- `Raw_Lt_phone_email` — full per-phone/email aggregated dataset (lifetime / long-term metrics) for detailed analysis.
- `RT_Raw_Data_v4`, `RT_Report_v4`, `User_Result` — supporting raw/reporting sheets used by workflows.

## Key scripts / triggers
- `Button.js` exposes `doGet(e)` webapp endpoints and helper functions (short/async task runners) used by UI buttons.
- Percentile & classification: `Percentile_Phone.js`, `Percentile_Sus_Phone_Email.js`, and related helper scripts perform the BigQuery queries and sheet writes.
- `duplicate_replace_sheet.js` and `phone_email_lt_sus_player.js` contain helper tasks for creating snapshot sheets and categorisation pipelines.

## Prerequisites & permissions
- BigQuery advanced service enabled in Apps Script (see `appsscript.json` — BigQuery is enabled).
- OAuth scopes required: BigQuery, Drive, Spreadsheets, UrlFetch, ScriptApp (listed in `appsscript.json`).
- The Apps Script deployment must run with an account that has access to the target Sheets and the `kz-dp-prod` BigQuery project.

## How to run / common operator actions
- Trigger from the Google Sheet UI menu or call the webapp `doGet` endpoint with an `action` parameter (see `Button.js`). Typical actions: `refresh_hardcoded_data`, `duplicate_sheet`.
- For long tasks the project uses time-based triggers to queue and run background jobs (see functions `startAsyncTask` / `runShortTask`).

## Troubleshooting & signals
- Status cells (e.g., `C4`, `C5`, `C6`) show `RUNNING`, `COMPLETED`, or `ERROR` and are the first place to check.
- Check Apps Script execution logs for stack traces and SQL logged queries.
- If sheets are not populated, confirm Apps Script webapp deployment and that the executing account has BigQuery access.

## TODO(owner)
- Confirm the deployed webapp URL and paste it here: `TODO(owner): webapp URL`.
- Document which Google Sheet IDs are the canonical control and result files: `TODO(owner): control sheet URL`.
- Review and confirm exact numeric thresholds (percentile indexes and static thresholds) used in the SQL — replace example numbers with the official values if different.
