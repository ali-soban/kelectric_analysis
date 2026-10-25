# KE net-metering bill dashboard

Turns your K-Electric (Karachi) **net-metering** bill PDFs into a dashboard: units exported and
imported, what you paid, what solar saved you, your credit balance with KE, and the charge
breakdown, with trends over time.

Unofficial, and not affiliated with K-Electric. Always check the numbers against your actual bill.

Only the bill layout KE introduced in **Aug-2025** is supported. Older bills are skipped.

## Two ways to use it

### 1. In the browser (easiest)

Open **https://ke-net-metering.pages.dev**, drop in your bill PDFs, and the dashboard appears.
Everything is read inside your browser. Nothing is uploaded, and the page's security policy
blocks all outgoing requests. Source is in [`web/`](web/).

### 2. On your PC (monthly automation, Windows)

Downloads the month's bill, fills an Excel workbook and rebuilds a local HTML dashboard.

1. `pip install pdfplumber openpyxl`
2. Double-click **Run KE Update.bat** (or run `python ke_update.py`).
3. If this month's bill isn't saved yet, KE's bill page opens. Solve the CAPTCHA yourself,
   download the PDF, and press Enter. The bill is moved from Downloads into the bills folder
   as `YYMM.pdf` (2610 = the Oct-2026 bill).
4. The dashboard opens. Optionally type each month's solar production (from your inverter app)
   into column C of `KE Bills Auto.xlsx`, save, and run again.

Bills are kept in `~/Desktop/ke_bills` by default. Set the `KE_BILLS_DIR` environment variable
to use another folder. The dashboard loads Chart.js from a CDN, so it needs internet to draw.

If a bill can't be read, the run says which one and why. It never guesses a number.

## Privacy

Bills contain your name, address and account number. This repository contains no bills. Keep
yours local, and **don't attach bill PDFs to issues** unless you've redacted them.

## Tests

Add your own bills to the bills folder, then run `npm test` in `web/`. It checks that the browser
parser (`web/parse.js`) gives the same results as the Python one (`parse_bill()` in `ke_update.py`)
on every bill you have. Bills aren't included in this repo because they're personal.

## License

[MIT](LICENSE). The bundled libraries in `web/vendor/` (Chart.js, pdf.js) keep their own licenses.
