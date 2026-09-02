# MSpace-Time Counter_v2

Open `index.html` in a browser to use the front-desk console.

Shared operations workbook: [MSpace-TimeCounter_v2 — Mlang Live Operations](https://docs.google.com/spreadsheets/d/1Cs269YewMz6mPNK4ye7qzZMpVfg3rm6NWLsbPVfIS_I/edit)

- Select **Check-in**, tap/type an RFID UID, and complete the customer form for a new customer. Known customers start their timer immediately.
- Select **Check-out**, tap/type the active customer’s RFID UID, and the app calculates the amount from elapsed time.
- Default rate: ₱30/hour, prorated per minute. Students and teachers receive a configurable 10% discount (₱27/hour by default).
- Data is stored locally in the browser until the Google Sheets connector is configured.
- The workbook is organized into **Dashboard**, **Live Monitor**, **Sessions**, **Customers**, and **Settings** with MSpace black/gold styling, ₱30/hour pricing, 10% student/teacher discount, and prorated-per-minute billing.
- Use **Export .CSV** for a file, or **Open Google Sheet** to open the shared income report.

## Realtime Google Sheets setup

1. Open the shared workbook and choose **Extensions → Apps Script**.
2. Replace the editor contents with [`apps-script/Code.gs`](apps-script/Code.gs), then save.
3. Deploy → **New deployment** → **Web app**. Run as your account. Choose the narrowest access setting your Google account supports; the deployed URL can read/write the operations sheet, so treat it as sensitive.
4. Open the web app, click **Rate settings**, keep the rate at `30` and discount at `10`, then paste the deployed Web app URL when prompted.
5. Repeat the URL setup on the staff computer and on the monitoring computer. The app polls the shared sheet every 10 seconds, while Google Sheets itself remains the source of truth.

If the connector is not configured, the app remains usable in local mode and keeps the existing RFID, checkout, payment popup, and CSV features.

The included `Mspace_logo.png` is the supplied MSpace logo used by the app.
