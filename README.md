# gfh-inventory-aging-processor

Standalone GFH Telecom Tkinter tool: `GFH_Inventory_Aging_Processor.pyw`.
Builds a Windows EXE automatically via GitHub Actions on every push to `main`
(uploaded directly to the repo's Releases page — no Actions storage used),
and via CircleCI once the repo is connected at circleci.com.

## Exclusions panel (Stores / IMEIs)

Click **Exclusions** in the app to open a paste-box panel:

- **Stores to exclude** and **IMEIs to exclude** are removed from the Google
  Sheets upload (district tabs + Executive Dashboard) and from the district
  `.xlsx` files that get emailed via Outlook. One entry per line; store names
  match case-insensitively (equal or contains either way), IMEIs match
  digits-only.
- **Always-blocked IMEIs** are removed from *both* pipelines at cleaning time
  (this box took over the old hardcoded block list, pre-seeded with it).

Everything persists to `gfh_aging_exclusions.json` next to the app and
survives restarts. A district whose devices are all excluded still gets an
empty tab / emailed report so managers know it was checked.
