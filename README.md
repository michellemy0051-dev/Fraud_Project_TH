# Fraud Project TH

## Overview
This Apps Script project is used to analyze suspicious-player behavior for the Thailand market and surface likely fraud or abuse patterns in Google Sheets. It pulls data from BigQuery, computes percentile-based thresholds, and writes the results into several reporting sheets for fraud team to review.

The project is designed for a Google Sheets-based workflow with Apps Script automation. It is currently configured for the Thailand market with the country prefix `TH` and uses BigQuery tables under the `kz-dp-prod` project.

## Owner / PIC
- Michelle — michellemy0051@kzgroup.biz (owner)
- Kin — pham0003@kzgroup.biz (backup PIC)

## Main purpose
The workflow identifies users who may be involved in suspicious activity by combining:
- deposit and withdrawal behavior,
- betting / turnover patterns,
- registration speed and brand spread,
- percentile-based risk thresholds,
- manual review outcomes stored in the `User_Result` sheet.

All monetary values are handled in THB.

## Main sheets used by the workflow
- `Suspicious_Player` — main operator-facing report sheet.
- `Percentile_phone_email` — percentile threshold output sheet.
- `Raw_Lt_phone_email` — raw details of extracted dataset.
- `User_Result` — manual review and exclusion results used to suppress or monitor users.

## Current automation flow
1. The Apps Script reads date inputs from the relevant report sheet (`B3` and `B4`).
2. It runs BigQuery queries to pull member, deposit, withdrawal, promotion, turnover, and session data.
3. Percentile thresholds are calculated and written to the percentile sheet.
4. The raw and report sheets are populated with risk metrics and categorization results.
5. Operators can review flagged users and update the `User_Result` sheet for future exclusions or monitoring.

## Key Apps Script files
- `const.js` — shared constants such as project IDs, country prefix, timezone offset, and minimum amount values.
- `Utils.js` — helper functions for date formatting, BigQuery streaming, member-group mapping, and sheet formatting.
- `Button.js` — web app entry points and button-driven automation helpers.
- `Suspicious_Player.js` — main suspicious-player report generation and snapshot sheet creation.
- `phone_email_lt_percentile.js` — percentile workflow for the lifetime/raw analysis flow.
- `phone_email_lt_raw_data.js` — raw-data workflow for the long-term analysis flow.
- `phone_email_lt_sus_player.js` — categorization and review logic for the lifetime analysis flow.
- `duplicate_replace_sheet.js` — creates a copy/snapshot of the suspicious-player sheet for historical reporting.

## Available actions
The web app endpoint in `Button.js` supports the following actions:
- `duplicate_sheet`


These are used by the Apps Script UI/button wrappers to trigger the main processing routines.

## Prerequisites and permissions
- Apps Script project with BigQuery advanced service enabled.
- OAuth scopes for BigQuery, Drive, Spreadsheets, UrlFetch, and ScriptApp.
- The executing Apps Script account must have access to the target Google Sheets and the BigQuery datasets used by the queries.

## Troubleshooting
- Check the status cells such as `C4`, `C5`, or `C6` on the relevant report sheet for run status.
- Review Apps Script execution logs for BigQuery query failures or unexpected runtime errors.
- Confirm that the sheet names referenced by the scripts still exist and match the current workbook.
- Validate that the deployment/web app is active if the button-based actions are not responding.

## Notes
- The project currently uses the Thailand market configuration and BigQuery tables under the `kz-dp-prod` environment.
- The workflow is still evolving, so some scripts are legacy/alternative paths that may overlap with the newer RT-based flow.
- Review and update threshold values in the sheet cells if the business rules change.
