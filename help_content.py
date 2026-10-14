"""Plain-language help for every screen. Used by the Help drawer, the Help & guide page and
the first-run tour. Every statement here describes behaviour that exists in the app."""

TOUR = [
    {"title": "Welcome to your business manager",
     "body": "This app keeps your tiffin business in one place: customers, daily deliveries, monthly invoices, payments, expenses and profit. "
             "Everything is calculated for you from what you enter. This 1-minute tour shows where things are."},
    {"title": "1 · Add your customers",
     "body": "Open Customers and add each person with their meal plan and monthly rate. If they live in a PG house, add the PG first under PG houses and link them."},
    {"title": "2 · Mark skipped meals and leaves",
     "body": "Daily tracking lists every customer. Tap a meal to switch it between Delivered, Skipped and Leave. "
             "You only need to touch the exceptions — everything else counts as delivered."},
    {"title": "3 · Create invoices",
     "body": "In Invoices, press Generate invoices once a month to bill everyone at once, or New invoice for a single customer. "
             "Open any invoice to change its amount, dates, items, notes or customer details, or to delete it."},
    {"title": "4 · Record payments",
     "body": "Press Record payment on an invoice when money arrives. Dues shows who still owes you. Expenses, Savings and Reports then work out your profit automatically."},
    {"title": "5 · Keep your data safe",
     "body": "Open Settings & backup to download a backup now and then. Open Account to change your username and password — "
             "do this first if you are still using the starting sign-in. You can reopen this tour any time from Help."},
]

# One entry per screen (key = Flask endpoint name).
PAGE_HELP = {
    "dashboard": {
        "title": "Dashboard",
        "what": "A live summary of your business for the month you pick: money in, money out, profit, who owes you, and anything that needs attention.",
        "steps": ["Use the month picker to look at another month.",
                  "Read the alerts at the top first — they point to overdue invoices and unusual balances.",
                  "Click any card's link (for example View dues) to jump to the detail behind the number."],
        "tips": ["Nothing here is typed in by hand. If a number looks wrong, fix the invoice, payment or expense behind it."]},
    "customers": {
        "title": "Customers",
        "what": "Everyone you serve. Each customer has a meal plan, a rate, a status and optional PG house.",
        "steps": ["Fill in the form to add a customer: name, category, meals and monthly rate (or a rate per meal).",
                  "Use search, filters and sorting to find people quickly.",
                  "Open a customer to see their invoices, payments and balance. Use Edit to change details or rates."],
        "tips": ["Active customers are billed every month. Paused customers are credited for the days they were away. Inactive customers are not billed.",
                 "To change a rate, use Edit and pick the month it starts from — earlier months keep their old rate."]},
    "edit_customer": {
        "title": "Edit customer",
        "what": "Change a customer's details. Plan and rate changes take effect from the month you choose.",
        "steps": ["Change the details and press Save.",
                  "To change the rate, choose 'applies from' month. Months that are already invoiced can't be changed."],
        "tips": ["The rate history lists every plan the customer has had. A change that hasn't started yet can be removed."]},
    "customer_detail": {
        "title": "Customer detail",
        "what": "One customer's invoices, payments, balance and pause history in one place.",
        "steps": ["Check the balance at the top to see what they owe or have paid in advance.",
                  "Use the tabs to switch between Invoices, Payments and Extra orders.",
                  "Use Pause deliveries to record a holiday: the paused days are credited on the bill."],
        "tips": ["Payments that aren't linked to an invoice are treated as advance credit and reduce what the customer owes."]},
    "pgs": {
        "title": "PG houses",
        "what": "Paying-guest houses you deliver to. Linking customers to a PG groups them together.",
        "steps": ["Add the PG's name, address and contact person.", "Link customers to it from the Customers screen."],
        "tips": ["A PG can't be deleted while customers are linked to it."]},
    "daily": {
        "title": "Daily tracking",
        "what": "Record exceptions to the normal delivery routine, day by day. Customers whose meals you don't touch count as delivered.",
        "steps": ["Pick the date (use the arrows to step through days).",
                  "Tap a meal button to cycle: Delivered → Skipped → Leave.",
                  "Switch to the Month tab to see the whole month at a glance."],
        "tips": ["Skipped = one missed delivery. It does not reduce the bill.",
                 "Leave = the customer was away for the day. Leave days are credited on the next invoice.",
                 "You can't edit a month that has been closed. Reopen it from Month history if you must."]},
    "extra_orders": {
        "title": "Extra orders",
        "what": "One-off charges: a guest tiffin, an extra box, a special order. For customers or walk-in guests.",
        "steps": ["Choose a customer (or type a guest name), add the amount and press Add.",
                  "Mark it Paid if the money was received straight away — it is then not billed again.",
                  "Unpaid extras are added to that customer's next invoice automatically."],
        "tips": ["On an invoice you can also attach or remove unbilled extras while editing."]},
    "billing": {
        "title": "Invoices",
        "what": "All invoices. Generate a whole month in one click, create a single invoice, search, and record payments quickly.",
        "steps": ["Generate invoices: choose the month and press the button. Customers already billed for that month are skipped.",
                  "New invoice: bill one customer for a month; amounts are pre-filled but editable.",
                  "Click an invoice number to view, edit, print or delete it. Use Record payment in a row for a quick payment."],
        "tips": ["Status: Unpaid = nothing received, Partial = some received, Paid = fully paid. Overdue means the due date has passed and a balance remains.",
                 "A padlock means that month is closed and its invoices are locked until you unlock the month."]},
    "invoice_new": {
        "title": "New invoice",
        "what": "Create an invoice for one customer and month. Everything is pre-filled from the customer's plan, leave days and unbilled extras.",
        "steps": ["Pick the customer and month — the amounts fill in by themselves.",
                  "Change the subscription amount if you want to charge something different.",
                  "Tick the extra orders to include, and add any extra lines (a negative price makes a discount).",
                  "Press Create invoice."],
        "tips": ["A customer can have only one invoice per month. If one exists you'll be offered a link to edit it."]},
    "invoice_detail": {
        "title": "Invoice",
        "what": "The invoice as your customer sees it, plus how it was calculated and the payments received.",
        "steps": ["Edit invoice changes the amount, dates, customer details, items and notes.",
                  "Record payment adds money received against this invoice (up to the balance due).",
                  "Download PDF or Print gives you the invoice to send."],
        "tips": ["Deleting an invoice always asks you to confirm. If payments were recorded you choose to keep them as advance credit or delete them too.",
                 "If the month is closed, press Unlock month first; this is recorded and the old month snapshot is kept."]},
    "invoice_edit": {
        "title": "Edit invoice",
        "what": "Change anything on the invoice. The total is recalculated as you type and saved when you press Save changes.",
        "steps": ["Dates: set the invoice date and the due date.",
                  "Customer details: the name, area and contact printed on this invoice. Tick the box to update the customer record too.",
                  "Subscription: type a new amount, or press Use calculated amount to go back to the automatic figure.",
                  "Extra orders: change their date, description or amount, or remove them from this invoice.",
                  "Other lines: add items with a quantity and price. A negative price is a discount.",
                  "Save changes."],
        "tips": ["Removed extra orders return to Extra orders as unbilled and will be picked up by the next invoice.",
                 "Taxes aren't calculated automatically. If you charge GST, add it as a line."]},
    "payments": {
        "title": "Payments",
        "what": "Every payment received. Link a payment to an invoice, or record it as an advance.",
        "steps": ["Choose the customer, amount, date and how it was paid.",
                  "Pick an invoice to settle, or leave it empty for an advance payment.",
                  "Use the month filter to look at one month's payments."],
        "tips": ["Advance payments automatically reduce what the customer owes on Dues."]},
    "dues": {
        "title": "Dues",
        "what": "Who owes you money and how much, after counting advance payments.",
        "steps": ["The list is sorted so the customer who owes the most is first.", "Click a customer to see their invoices and record a payment."],
        "tips": ["Every customer who has ever been billed is listed; green 0 means they are fully paid up. Advance payments are already subtracted."]},
    "expenses": {
        "title": "Expenses",
        "what": "Money you spend on the business (and personal spending, kept separate). Recurring expenses such as rent are added for you every month.",
        "steps": ["Add an expense: date, category, amount and how it was paid.",
                  "Set up a recurring expense once and it is created automatically on the day you choose.",
                  "Use the filters to review a month or a category."],
        "tips": ["Business and Personal expenses are reported separately in Reports."]},
    "savings": {
        "title": "Savings & FD",
        "what": "Money you set aside from the business, including transfers into a fixed deposit.",
        "steps": ["Add a deposit or an FD transfer with its date and amount.", "Edit or delete entries with the row buttons."],
        "tips": ["Savings reduce 'remaining available funds' on the dashboard."]},
    "personal_income": {
        "title": "Personal income",
        "what": "Money coming in from outside the tiffin business (for example another job). Kept apart from business revenue.",
        "steps": ["Add the source, amount and date."], "tips": []},
    "accounts": {
        "title": "Cash & bank",
        "what": "How much sits in Cash, UPI and Bank, worked out from payments in and money spent.",
        "steps": ["Set opening balances in Settings if you started with money already in an account."],
        "tips": ["If a balance is negative, check the opening balance or recent entries."]},
    "reports": {
        "title": "Profit & loss",
        "what": "Revenue minus expenses for any month, with a breakdown by payment mode.",
        "steps": ["Choose a month to see its profit and loss.", "Use Export to Excel in the sidebar for a full workbook."], "tips": []},
    "monthly": {
        "title": "Monthly summary",
        "what": "One row per month: billed, received, expenses and profit.",
        "steps": ["Click a month for the full breakdown."], "tips": []},
    "history": {
        "title": "Month history",
        "what": "Every month with its permanent snapshot. Finished months are closed automatically a few days after they end.",
        "steps": ["Open a month to see all its figures and entries.", "Download a month as JSON if you want a copy."],
        "tips": ["Closed means locked: payments, expenses, invoices and tiffin records dated in that month can't be changed.",
                 "If you must correct a closed month, open it and use Reopen month. Close it again when you finish."]},
    "history_month": {
        "title": "Month detail",
        "what": "The full record of one month.",
        "steps": ["Close month saves a permanent snapshot and locks the month.", "Reopen month unlocks it for corrections (the old snapshot is kept)."], "tips": []},
    "settings_page": {
        "title": "Settings & backup",
        "what": "Business details, invoice options, logo, payment QR, backups and restore.",
        "steps": ["Business details appear on every invoice.", "Download backup now saves your whole database as one file.",
                  "Restore replaces your data with a backup file — your current data is saved as a safety copy first. Your sign-in account is kept."],
        "tips": ["Keep a backup somewhere other than this computer (USB drive or cloud folder)."]},
    "account": {
        "title": "Account",
        "what": "Change the username and password you sign in with.",
        "steps": ["Enter your current password.", "Type a new username, a new password, or both.", "Press Update sign-in details."],
        "tips": ["Use at least 8 characters, and don't reuse the starting password.", "Signing in again is not required after changing."]},
    "help_page": {
        "title": "Help & user guide", "what": "Guides for every screen and common jobs.", "steps": [], "tips": []},
}

GUIDE_SECTIONS = [
    {"id": "start", "title": "Getting started",
     "intro": "A first-time setup takes about ten minutes.",
     "steps": ["Sign in, then open Account and change the starting username and password.",
               "Open Settings & backup and enter your business name, address, phone, and UPI ID or QR code. These print on invoices.",
               "Add PG houses (if you deliver to any), then add your customers with their meals and rates.",
               "Use Daily tracking for skipped meals and leave days as they happen.",
               "At month end, open Invoices and press Generate invoices."]},
    {"id": "customers", "title": "Managing customers",
     "intro": "Customers drive everything else: meals, rate and status decide what they are billed.",
     "steps": ["Add: Customers → fill the form → Add customer.",
               "Edit: open the customer's row and choose Edit. For a new rate, choose the month it starts from.",
               "Pause: use Pause on the customer when they go away. The paused days are credited.",
               "Remove: Delete is only for mistakes. To stop billing someone, set them Inactive instead."]},
    {"id": "daily", "title": "Daily tracking",
     "intro": "Only record what differs from normal.",
     "steps": ["Open Daily tracking and choose the date.", "Tap a meal to cycle Delivered → Skipped → Leave.",
               "Skipped meals are only recorded; they don't reduce the bill. Leave days do."]},
    {"id": "invoices", "title": "Creating and editing invoices",
     "intro": "Invoices can be generated in bulk or created one at a time, and every part can be edited.",
     "steps": ["Bulk: Invoices → choose the month → Generate invoices.",
               "Single: Invoices → New invoice → choose customer and month → check amounts → Create invoice.",
               "Edit: open the invoice → Edit invoice → change dates, customer details, subscription amount, extra orders or add lines → Save changes.",
               "Delete: open the invoice → press Delete → confirm. If payments exist, choose whether to keep them as advance credit.",
               "Print or download: open the invoice and use Print or Download PDF."]},
    {"id": "payments", "title": "Payments and dues",
     "intro": "Money received is recorded against invoices so balances stay right.",
     "steps": ["From the Invoices list press Record payment, confirm the amount, and save.", "A payment can't be more than the invoice balance. For advance money use Payments without an invoice.",
               "Dues lists everyone who still owes money."]},
    {"id": "money", "title": "Expenses, savings and reports",
     "intro": "Profit is revenue minus business expenses.",
     "steps": ["Log expenses as you spend; set up recurring ones for rent and similar bills.",
               "Savings & FD and Personal income are tracked separately from business profit.",
               "Reports and Monthly summary show profit by month."]},
    {"id": "months", "title": "Closed months",
     "intro": "To protect your records, a finished month is locked a few days after it ends.",
     "steps": ["A closed month can't be edited. You'll see a padlock and a message.",
               "To correct it: open the invoice (or Month history) and press Unlock / Reopen month.",
               "When you finish, close the month again from Month history."]},
    {"id": "backup", "title": "Backups and your account",
     "intro": "Your data lives in one file, tiffin.db.",
     "steps": ["Settings & backup → Download backup now. Do this regularly.", "Restore uploads a backup; your current data is first saved as a safety copy.",
               "Account → change username/password. Sign out with the button in the top-right menu."]},
]

GLOSSARY = [
    ("Unpaid / Partial / Paid", "Nothing received / some received / fully received for an invoice."),
    ("Overdue", "The due date has passed and a balance is still unpaid."),
    ("Active / Paused / Inactive", "Billed normally / billed with paused days credited / not billed."),
    ("Delivered / Skipped / Leave", "Normal meal / a missed delivery (no bill reduction) / full day away (credited)."),
    ("Advance credit", "A payment not tied to an invoice; it reduces what the customer owes."),
    ("Closed month", "A finished month that has been saved and locked. Reopen it to make corrections."),
    ("Adjusted amount", "An invoice whose subscription amount was changed from the automatic calculation."),
    ("Extra order", "A one-off charge such as a guest tiffin."),
]

SIGN_IN_ERRORS = [
    ("Incorrect username or password", "Check Caps Lock. After 5 wrong attempts sign-in pauses for 5 minutes."),
    ("Your session expired", "Sign in again. Sessions end after 12 hours or when the sign-in details change."),
]
