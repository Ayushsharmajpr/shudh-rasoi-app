# Shudh Rasoi — Tiffin Business Manager (v3)

A web app for running a tiffin / meal-subscription business: customers and PG houses, daily skipped-meal and leave tracking,
monthly invoices (PDF with UPI QR), payments and dues, expenses, savings, month closing, reports and Excel export.

This release is built on V2.0 (V1.0's functionality is a subset of it — see `docs/COMPARISON.md`) and adds:
sign-in, a complete invoice editor, safe invoice deletion, single-customer invoice creation with auto-fill, global search,
a redesigned interface, and a built-in user guide. **Your existing database is used as-is** (`tiffin.db`).

## Quick start

Requires Python 3.10+.

| System | Run |
|---|---|
| Windows | double-click `start_app.bat` |
| macOS / Linux | `./start_app.sh` |
| Manual | `pip install -r requirements.txt` then `python app.py` |

Open <http://127.0.0.1:5000>.

**First sign-in:** username `admin`, password `admin`. A yellow banner stays on every page until you change them under
**Account** (top-right menu → *Account & password*). Do this first.

## Your data

* The database is the single file **`tiffin.db`** in this folder — it is your original backup file, with all records intact.
  On first start the app adds a few things *without touching existing rows*: a `users` table (sign-in), four nullable columns on
  `invoices` (`bill_to_name`, `bill_to_area`, `bill_to_contact`, `updated_at`), an `invoice_lines` table, and some indexes.
* To use a different file: set the environment variable `TIFFIN_DB=/path/to/file.db`.
* **Back up regularly**: *Settings & backup → Download backup now*. Restoring a backup keeps the current sign-in account, so a
  restore can never lock you out or reset your password. (A backup file contains your password *hash*, never the password.)

## What's new in this version

* **Sign-in & security** — hashed passwords (scrypt), 12-hour sessions, CSRF protection on every form, lockout for 5 minutes after
  5 wrong passwords, secure cookie flags, security headers/CSP, no outside network calls (fonts are system fonts), every page,
  download and API route requires sign-in. Change username/password in *Account*; the password change signs out other sessions.
* **Invoices (the two reported problems are fixed)**
  * *Edit invoice* now edits everything: invoice & due dates, customer details printed on the invoice (optionally also updating the
    customer record), subscription amount (with "use calculated amount"), extra orders (date/description/amount, remove, add unbilled
    ones), extra lines (description × quantity × price; a negative price is a discount), and notes. Totals update live and are
    re-calculated on the server; payments are untouched.
  * *Delete invoice* is available from the invoice page with a confirmation dialog. If payments exist you choose: keep them as
    advance credit (default) or delete them too. The invoice number is never reused; the saved snapshot is kept as "voided".
  * Why it felt broken before: V2 hid both buttons whenever the month was closed (all September invoices) or any payment existed
    (11 of 15 invoices), and the amount form sat far down the page. Now both are always reachable; for a **closed month** the page
    explains why and offers **Unlock month** (reopens the month; close it again from Month history).
  * *New invoice* for one customer + month, pre-filled from the plan, rate, leave days and unbilled extras.
  * Invoice list: search, month/status filters (Unpaid / Partial / Overdue / Paid), sorting, totals, paging, one-click
    *Record payment*, print (opens the PDF in the browser's viewer) and download.
* **Less manual work** — auto-filled new invoices, one-click payments for the exact balance, "use calculated amount", quick-action
  buttons, global search (`Ctrl+K` or `/`), unbilled extras attach automatically, double-submit protection on every form.
* **Help built in** — first-run tour, `?` help drawer on every screen (press `?`), tooltips on important fields, and a full
  *Help & user guide* page.

## Deploying beyond one computer

`python app.py` uses the production-grade **waitress** server automatically. Environment variables:

| Variable | Meaning |
|---|---|
| `TIFFIN_DB` | path to the database file (default `./tiffin.db`) |
| `TIFFIN_HOST` / `TIFFIN_PORT` | bind address / port (default `127.0.0.1` / `5000`) |
| `TIFFIN_HTTPS=1` | mark the session cookie `Secure` — set this when serving over HTTPS |

Keep the default host (`127.0.0.1`) unless you put the app behind a reverse proxy (nginx, Caddy…) that provides HTTPS, or are
on a trusted private network. Run **one** application process (SQLite, single writer). The session secret is generated on first
run in `.secret_key` — keep that file private and don't copy it between installations.

## Tests

```
python -m unittest discover -s tests                       # 37 billing / rate / snapshot / month-lock tests
QA_SOURCE_DB=tiffin.db python tests/qa_invoices_and_auth.py   # invoice edit/create/delete/unlock, sign-in, CSRF, lockout, restore
QA_SOURCE_DB=tiffin.db python tests/qa_crud.py                # create/edit/delete across every module, exports, data integrity
```
All of them work on a throw-away copy of the database. `tests/qa_browser.py` is a Playwright browser walk-through (see its header).

## Known limits / assumptions (please read)

* **Taxes/discounts:** the original data model has no tax or discount fields. They are supported as *extra lines* (negative price =
  discount; add GST as a line) — there is no automatic tax calculation.
* **An invoice can't be moved to a different customer** (its calculation snapshot belongs to the original customer). Delete it and
  create a new one, or edit the *customer details printed on the invoice*.
* **Closed months stay protected** by the original database triggers; editing a closed month's invoice requires *Unlock month*.
* The sign-in is a single administrator account (the request specified one). There are no roles or multiple users.
* Login lockout counters live in memory and reset when the app restarts.
* The first-sign-in hint (admin/admin) is shown on the login page only while the default password has not been changed.
* UPI QR generation for invoices when only a UPI ID is configured needs the `qrcode` package (in `requirements.txt`); your data already
  has an uploaded QR image.
