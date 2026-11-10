# Inventory Aging Processor

Cleans inventory **aging reports**, emails the district files out, and
uploads results to Google Sheets. Sender/district contacts are configurable
in-app (seeded on first run, editable in Settings) and persist to JSON.

## Highlights
- Aging report cleanup and district split
- Email dispatch with configurable sender identity
- Google Sheets upload
- Windows EXE via GitHub Actions on push.

## Tab 2: WhatsApp — Aged Devices
A second, independent tab. Pick a raw Inventory Aging export, click
**Load Districts**, then for each district tick it and type the WhatsApp
Desktop group name to send it to (typed by hand — no fixed list). Set the
**Aged (days) or more** cutoff, optionally add store name patterns to
**Store Exclusions** (same case-insensitive match as the main Exclusions
panel, but only affects this tab), then **Save Mapping** or **Send to
WhatsApp Desktop**. Everything (checks, group names, aged-days, excluded
stores) persists to `gfh_aging_whatsapp_config.json` next to the app.

Each ticked district becomes one GFH-branded screenshot image — navy header
bar with the GFH logo and a centered title (same layout as every other
GFH/VDS app window), then a table of DISTRICT, STORE, PRODUCT, IMEI,
AGE IN COMPANY, PO RECEIVING DATE — pasted straight into that WhatsApp
Desktop group's chat and sent.

**Requires `pywinauto`** (`pip install pywinauto`, included in
`requirements.txt`) and the WhatsApp Desktop Windows app (not WhatsApp Web)
signed in. If a send step can't find a UI element (WhatsApp updates its
interface from time to time), the Activity Log says exactly which step
failed — the image is still on the clipboard and the chat is already open,
so finishing with Ctrl+V and Enter takes one second by hand.
