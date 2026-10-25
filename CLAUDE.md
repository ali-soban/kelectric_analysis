# KE net-metering bill pipeline

Turns K-Electric (Karachi) net-metering bill PDFs into a workbook and an HTML dashboard.
Two front ends share one parser: the Python pipeline (`ke_update.py`) and a browser-only page (`web/`).

## Files

| Path | Role |
|---|---|
| `ke_update.py` | Pipeline: fetch → import → extract → workbook → dashboard. `parse_bill()` and `derive()` are the core |
| `dashboard_template.html` | Dashboard page; `__DATA__` is replaced with the JSON rows |
| `Run KE Update.bat` | Windows entry point |
| `web/` | Browser-only version. `parse.js` is a JS port of `parse_bill()`/`derive()`: change both together, then run `npm test` in `web/` (compares JS vs Python on every bill in the bills folder). Serve over http (ES modules don't load from `file://`) |

Bills folder: `$KE_BILLS_DIR`, default `~/Desktop/ke_bills`. Files are named `YYMM.pdf` by bill month.

## Rules

- Never commit bills, workbooks or dashboards: they hold personal data.
- Don't automate the KE site's CAPTCHA.
- Validate, don't guess: `parse_bill()` raises `BillError` on any mismatch rather than write a doubtful number.
- `web/` must stay offline-only: libraries are self-hosted in `web/vendor/`, and `_headers` sets
  a CSP with `connect-src 'none'`. Don't add CDN scripts, analytics or `fetch` calls.
- Only the Aug-2025+ bill layout is parsed. Older bills raise `BillError` and are skipped.
- If KE changes the layout: run `npm test` in `web/`, inspect the failing bill's `pdfplumber`
  text, and adjust the regexes in `parse_bill()` and `parse.js` together.

## Parser notes (new layout)

- Header line `: Sep-26 : 07-Sep-2026` gives the bill month. Text extraction works; no OCR.
- KE's "Import Energy" is energy KE imports from the house, i.e. solar **export**.
- A mid-period tariff change adds `Variable Off Peak (New)` / `Fixed Charges (New)` lines.
- Payable/credit: use what is printed, in this order: `Outstanding Balance` (credit) →
  `Amount Payable within Due Date` → computed from `Previous (Recoverable) Dues + Current Bill − NM Benefit`,
  rounded **up**. If that net is ~0 the PDF was downloaded after paying: rebuild from `min(0, Opening Balance)`.
- Don't trust `Opening Balance` generally (wrong on some bills), the "Remaining Dues" line, or the
  payment-history table.
- `charge_breakdown()` returns None unless the lines sum to the Current Bill Amount (±1).
- Quarterly PPP months (Dec/Mar/Jun/Sep) settle units via the NAPP line; carry-forward units are 0 then.
