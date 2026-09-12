# macOS permissions for the local oracle

The first `npm run oracle --` run that touches Excel or Numbers triggers macOS prompts:

- **Automation**: "Terminal (or your IDE) wants to control Microsoft Excel / Numbers" → Allow.
  If you clicked Deny, re-enable it in System Settings > Privacy & Security > Automation.
- The Excel oracle only opens, reads cell values and closes without saving, so the read-only Excel license is
  sufficient. A repair dialog ("We found a problem with some content…") blocks the script; the runner reports
  HUNG after the timeout - dismiss the dialog by hand and treat the file as REPAIRED.
- LibreOffice runs headless with its own profile under `.generated/lo-profile` and needs no permissions.
- Numbers automation produced empty exports on Numbers 14.4; see `fixtures/golden/numbers/STEPS.md`.
