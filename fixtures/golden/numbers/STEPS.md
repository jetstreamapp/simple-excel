# Numbers golden (manual)

Automated export was attempted twice (`oracle/numbers/export-xlsx.applescript`, `export-front.applescript`):
on Numbers 14.4 the AppleScript `open` verb and the LaunchServices import both yielded an empty document, so
the exported workbook was a 5 KB single empty sheet. Produce this golden by hand:

1. Open `fixtures/golden/exceljs/canonical.xlsx` in Numbers (File > Open). Note any import warning banner
   ("Some features aren't supported") and write it down below.
2. File > Export To > Excel… > keep "Include a summary worksheet" OFF > Next > save as
   `fixtures/golden/numbers/canonical.from-exceljs.xlsx`.
3. Register it:
   `node fixtures/register.mjs golden/numbers/canonical.from-exceljs.xlsx --id golden-canonical-numbers --generator "Numbers 14.4" --provenance "manual:golden/numbers/STEPS.md" --license "Apple Numbers export" --tags kind:golden,generator:numbers --expected canonical/canonical.json --notes "<import warnings seen>"`
4. Run `npm run oracle -- --fixtures golden-canonical-numbers --label numbers`.

Import warnings observed: _(fill in)_

Unsupported formulas were replaced by the last calculated value

Hidden sheets were made visible
Sheet: Hidden
