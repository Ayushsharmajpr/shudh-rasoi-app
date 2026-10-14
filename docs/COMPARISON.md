# V1.0 vs V2.0 vs this release

Method: read both code bases (V1 `app.py` 2,822 lines; V2 4,575 lines), diffed their routes and templates, opened both against the
provided database, and ran V2's invoice paths on a copy of your data.

## V1.0 → V2.0
**Added in V2** (not in V1): month history with *close / reopen* and permanent month snapshots; customer **rate history** (a rate
change applies from a chosen month; past months keep their rate); invoice **snapshots** (how each invoice was calculated, kept even
if records change later); database triggers that lock closed months; **logo and payment-QR uploads** and a UPI QR on invoices; month
JSON export; automatic month closing and recurring-expense creation; a friendly error page.

**Kept from V1:** every V1 route still exists in V2 (customers, PG houses, daily tracking, extra orders, invoices, payments, dues,
expenses, savings, personal income, cash & bank, reports, monthly summary, settings, backup/restore, Excel export).

**Problems found in V2 (all fixed here)**
| Finding | Evidence | Fix |
|---|---|---|
| "Edit amount" appears not to work | Backend worked, but the button is hidden when the month is closed (all 7 September invoices) and the form opens far down the page | Full editor page; closed months show an explicit *Unlock month* path |
| "Delete invoice" missing | Hidden whenever any payment exists (11 of 15 invoices) or the month is closed; server also refused | Delete is always offered; payments can be kept as advance credit or deleted; confirmation dialog |
| Only the subscription amount was editable | no date, due date, customer, extras, notes, lines | Complete invoice editor |
| No way to create one invoice | only "generate for whole month" | *New invoice* with auto-fill |
| No authentication | every page and the full database download were open | Sign-in, CSRF, lockout, headers |
| Customer-details edits couldn't reach the PDF | PDF read a frozen snapshot | Invoice-level "bill to" fields used by view and PDF |
| Notes were stored but never printed | PDF builder ignored them | Printed on the PDF |
| External web-font request | UI called Google Fonts | removed (works offline) |
| Restore would have wiped a sign-in account | (new risk once sign-in exists) | Restore keeps the current account |

## Regression check
All V1 and V2 routes are present in this release (`/billing`, invoice pages, `edit-amount` kept as a compatibility shortcut).
V2's 37 automated tests pass (their client now signs in), plus the new suites listed in the README.

## Honest scope note
The interface was redesigned (new shell, design tokens, tables, forms, dialogs, dark mode, mobile layout) and the invoice, sign-in,
account and help screens were rebuilt from scratch. The other screens (customers, daily tracking, expenses, reports, settings…) keep
V2's page structure and logic but inherit the new design system; their templates were not individually redesigned.
