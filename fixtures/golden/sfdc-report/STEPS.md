# Salesforce report exports (manual, dev org)

Users upload Salesforce report exports into Load Records, so both export flavours are needed. Report export is
UI-only (no API), so the rows are loaded by script and the export is done by hand.

1. Load the canonical-shaped Opportunity rows into the dev org:
   `node fixtures/sfdc/load-report-rows.mjs` (uses `pnpm sf:api`; idempotent - rows are tagged
   `Description` prefix `[xlsx-engine canonical]`).
2. In the dev org: Reports > New Report > Opportunities > filter `Opportunity Name contains "xlsx-engine"`,
   add columns Opportunity Name, Close Date, Amount, Probability (%), Stage, Private, Next Step, Description,
   Created Date. Save as "xlsx-engine canonical".
3. Export > **Formatted Report** > Excel Format .xlsx → `fixtures/golden/sfdc-report/canonical.formatted.xlsx`
   (title rows, grouping/footer rows, merged headers).
4. Export > **Details Only** > Excel Format .xlsx → `fixtures/golden/sfdc-report/canonical.details.xlsx`.
5. Register both:
   `node fixtures/register.mjs golden/sfdc-report/canonical.details.xlsx --id golden-sfdc-report-details --generator "Salesforce report export" --provenance "manual:golden/sfdc-report/STEPS.md" --license "Salesforce export (dev org data)" --tags kind:golden,generator:sfdc-report,EC-SFDC-REPORT-DETAILS`
   `node fixtures/register.mjs golden/sfdc-report/canonical.formatted.xlsx --id golden-sfdc-report-formatted --generator "Salesforce report export" --provenance "manual:golden/sfdc-report/STEPS.md" --license "Salesforce export (dev org data)" --tags kind:golden,generator:sfdc-report,EC-SFDC-REPORT-FORMATTED,EC-TITLE-AND-FOOTER-ROWS`
6. Run `npm run oracle -- --tag generator:sfdc-report --label sfdc`. There is no ground-truth dump for these
   (Salesforce formats values), so the interesting output is the reader dumps: header row position, merged
   title rows, footer rows, how dates/currency/percent are typed.
