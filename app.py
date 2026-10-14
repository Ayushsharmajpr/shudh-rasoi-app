"""
Shudh Rasoi — Tiffin Business Management System
A local business management app: customers/PGs, meal plans, daily tiffin
tracking, billing & invoices, payments & dues, expenses, and reports.
Runs entirely on your own machine with a local SQLite database (tiffin.db).
"""
import sqlite3
import os
import io
import shutil
import threading
import math
import re
import json
import glob
import logging
import secrets
import tempfile
import calendar as calmod
from xml.sax.saxutils import escape as xml_escape
from datetime import date, datetime, timedelta
from flask import (
    Flask, render_template, request, redirect, url_for, jsonify, flash,
    send_file, abort, g, has_request_context, Response, session, jsonify
)
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.platypus import Image as RLImage
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, KeepTogether
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_RIGHT, TA_LEFT, TA_CENTER
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
# TIFFIN_DB lets you keep the database somewhere else (e.g. a synced folder);
# by default it sits next to the app exactly as before.
DB_PATH = os.environ.get("TIFFIN_DB") or os.path.join(BASE_DIR, "tiffin.db")
BACKUP_DIR = os.path.join(os.path.dirname(os.path.abspath(DB_PATH)), "backups")
AUTO_BACKUPS_TO_KEEP = 14
log = logging.getLogger("shudh_rasoi")


def today_date():
    """Single source of 'today' for the whole app (makes date logic testable)."""
    return datetime.now().date()

_FONT_REGULAR = "Helvetica"
_FONT_BOLD = "Helvetica-Bold"
try:
    pdfmetrics.registerFont(TTFont("DejaVuSans", os.path.join(BASE_DIR, "static", "fonts", "DejaVuSans.ttf")))
    pdfmetrics.registerFont(TTFont("DejaVuSans-Bold", os.path.join(BASE_DIR, "static", "fonts", "DejaVuSans-Bold.ttf")))
    _FONT_REGULAR = "DejaVuSans"
    _FONT_BOLD = "DejaVuSans-Bold"
except Exception:
    pass

EXPENSE_CATEGORIES = [
    "Raw Material / Groceries", "Gas Cylinder", "Staff Salary / Wages",
    "Delivery / Transport", "Packaging (Tiffin/Boxes)", "Rent",
    "Electricity & Utilities", "Maintenance & Repairs", "Marketing",
    "Miscellaneous",
]
PERSONAL_EXPENSE_CATEGORIES = [
    "Household", "Groceries (Personal)", "Medical", "Education", "Family Support",
    "Entertainment", "Shopping", "Travel", "Bills & Utilities (Personal)", "Miscellaneous",
]
PERSONAL_INCOME_SOURCES = [
    "Salary (Other Job)", "Freelance / Side Work", "Rental Income", "Family Contribution",
    "Gift / Bonus", "Interest / Investment", "Other",
]
MAX_AMOUNT = 10_000_000
PAYMENT_MODES = ["Cash", "UPI", "Bank Transfer", "Other"]
CUSTOMER_CATEGORIES = ["PG", "Regular", "Extra"]
CUSTOMER_STATUSES = ["Active", "Paused", "Inactive"]
MEALS = ["breakfast", "lunch", "dinner"]
MEAL_LABELS = {"breakfast": "Breakfast", "lunch": "Lunch", "dinner": "Dinner"}
DAILY_STATUSES = ["Delivered", "Skipped", "Leave"]
EXPENSE_TYPES = ["Business", "Personal"]
SAVINGS_TYPES = ["Deposit", "FD Transfer"]

class MonthLocked(Exception):
    """Raised when an action would change a month that has been closed (snapshotted)."""
    def __init__(self, ym):
        self.ym = ym
        super().__init__(f"{ym} is closed")


def _load_secret_key():
    env_key = os.environ.get("TIFFIN_SECRET_KEY")
    if env_key:
        return env_key
    key_path = os.path.join(BASE_DIR, ".secret_key")
    try:
        with open(key_path, "r", encoding="utf-8") as fh:
            stored = fh.read().strip()
            if stored:
                return stored
    except OSError:
        pass
    new_key = secrets.token_hex(32)
    try:
        with open(key_path, "w", encoding="utf-8") as fh:
            fh.write(new_key)
    except OSError:
        pass  # fall back to a per-process key; only flash messages are affected
    return new_key


app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 64 * 1024 * 1024   # restore uploads up to 64 MB
app.secret_key = _load_secret_key()

_last_automation_check = None
_db_ready = False


def _ensure_db_ready():
    """Migrate the database once per process, before anything reads it."""
    global _db_ready
    if not _db_ready:
        init_db()  # idempotent; makes `flask run` / gunicorn work without a manual step
        _db_ready = True


# Authentication guard + CSRF + security headers. Registered first so that nothing
# (not even the daily automation) runs for a visitor who has not signed in.
import auth as _auth
_auth.init_auth(app, lambda: get_db(), _ensure_db_ready)


@app.before_request
def _run_daily_automation():
    """Make sure the database is migrated, then — once per day, on the first request
    after midnight — catch up due recurring expenses, refresh each customer's current
    rate, close any finished months, and take a rolling safety-net backup."""
    global _last_automation_check
    _ensure_db_ready()
    today_str = today_date().isoformat()
    if _last_automation_check == today_str:
        return
    _last_automation_check = today_str
    try:
        conn = get_db()
        try:
            run_recurring_expenses(conn)
            sync_customer_current_plans(conn)
            auto_close_elapsed_months(conn)
            conn.commit()
        finally:
            conn.close()
        run_auto_backup()
    except Exception:  # never let a background automation hiccup block a page from loading
        log.exception("daily automation failed")


@app.teardown_request
def _close_request_connections(exc):
    """Every connection opened during a request is closed (and any half-finished
    transaction rolled back) even if the view raised, so one failed save can never
    leave the database locked for the next request."""
    for conn in g.pop("_db_conns", []):
        try:
            conn.rollback()
        except sqlite3.Error:
            pass
        try:
            conn.close()
        except sqlite3.Error:
            pass



# ---------------------------------------------------------------------------
# Request guard: same-site check + one place that validates common form fields
# ---------------------------------------------------------------------------
_DATE_FIELDS = {"date": "Date", "join_date": "Join date", "pause_from": "Pause start date", "pause_to": "Pause end date"}
_MONTH_FIELDS = {"month": "Month", "effective_month": "Effective month"}
_PHONE_FIELDS = {"bill_to_contact": "Contact number", "contact": "Contact number", "contact_number": "Contact number", "business_phone": "Phone number"}


def valid_iso_date(v):
    try:
        d = date.fromisoformat(v)
        return 2000 <= d.year <= 2100
    except (ValueError, TypeError):
        return False


def valid_phone(v):
    digits = re.sub(r"\D", "", v)          # several numbers may be listed, e.g. "98290 11111 / 98290 22222"
    return bool(re.fullmatch(r"[0-9+()\-\s,/.]+", v)) and 7 <= len(digits) <= 30


def _same_site():
    if os.environ.get("TIFFIN_DISABLE_ORIGIN_CHECK") == "1":
        return True
    src = request.headers.get("Origin") or request.headers.get("Referer")
    if not src:
        return True                      # non-browser clients (tests, curl) send neither
    from urllib.parse import urlparse
    host = urlparse(src).netloc
    return host in (request.host, request.headers.get("X-Forwarded-Host", ""))


@app.before_request
def _guard_requests():
    if request.method != "POST":
        return None
    if not _same_site():
        abort(403)
    errors = []
    for key in request.form:
        for val in request.form.getlist(key):
            v = (val or "").strip()
            if not v:
                continue
            if len(v) > 2000:
                errors.append("One of the fields is too long.")
            elif key in _DATE_FIELDS and not valid_iso_date(v):
                errors.append(f"{_DATE_FIELDS[key]} isn't a valid date.")
            elif key in _MONTH_FIELDS and not valid_ym(v):
                errors.append(f"{_MONTH_FIELDS[key]} isn't valid.")
            elif key == "payment_mode" and v not in PAYMENT_MODES:
                errors.append("Please choose a valid payment mode.")
            elif key in _PHONE_FIELDS and not valid_phone(v):
                errors.append(f"{_PHONE_FIELDS[key]} isn't valid — enter at least 7 digits (spaces, +, dashes and commas are fine).")
    if errors:
        for e in dict.fromkeys(errors):
            flash(e, "error")
        return redirect(request.referrer or url_for("dashboard"))
    return None


def _error_page(code, title, message):
    if request.path == "/daily/toggle" or request.is_json:
        return jsonify({"ok": False, "error": title}), code
    return render_template("error.html", code=code, title=title, message=message), code


@app.errorhandler(404)
def _e404(e):
    return _error_page(404, "Page not found", "We couldn't find what you were looking for. It may have been removed or the link is wrong.")


@app.errorhandler(405)
def _e405(e):
    return _error_page(405, "Action not allowed", "That action isn't available from this page.")


@app.errorhandler(403)
def _e403(e):
    return _error_page(403, "Request blocked", "That request came from another website, so it was blocked for your safety. Please use the buttons inside the app.")


@app.errorhandler(413)
def _e413(e):
    return _error_page(413, "File too large", "That file is too large to upload. Backups can be up to 64 MB.")


@app.errorhandler(500)
def _e500(e):
    log.exception("unhandled error")
    return _error_page(500, "Something went wrong", "An unexpected error occurred. Your data is safe — please go back and try again. If it keeps happening, restore from a recent backup in Settings.")


@app.route("/favicon.ico")
def favicon():
    if not session.get("uid"):
        return ("", 204)
    meta = get_brand_meta().get("logo")
    return redirect(meta["url"]) if meta else ("", 204)


@app.errorhandler(MonthLocked)
def _handle_month_locked(e):
    flash(f"{month_label(e.ym)} is closed, so it can't be changed. Open Month history and use "
          f"\"Reopen month\" first if you really need to edit it.", "error")
    return redirect(request.referrer or url_for("history_month", ym=e.ym))


@app.errorhandler(sqlite3.IntegrityError)
def _handle_integrity_error(e):
    msg = str(e)
    if msg.startswith("MONTH_LOCKED"):
        flash("That entry belongs to a closed month, so it can't be added, changed or removed. "
              "Open Month history and use \"Reopen month\" first if you really need to edit it.", "error")
    elif msg.startswith("RATE_LOCKED"):
        flash("That rate change would alter a month that is already invoiced for this customer. "
              "Choose a later month, or delete the invoice first.", "error")
    elif msg.startswith("SNAPSHOT_IMMUTABLE"):
        flash("Saved month history is permanent and can't be edited.", "error")
    else:
        log.exception("database constraint failed")
        flash("Couldn't save that — it conflicts with existing data.", "error")
    return redirect(request.referrer or url_for("dashboard"))


def cur_symbol():
    return get_settings().get("currency_symbol", "₹")


def get_settings():
    try:
        conn = get_db()
        rows = conn.execute("SELECT key, value FROM settings").fetchall()
        conn.close()
        return {r["key"]: r["value"] for r in rows}
    except Exception:
        return {"currency_symbol": "₹", "business_name": "Shudh Rasoi"}


# ---------------------------------------------------------------------------
# Brand assets: company logo + payment QR (stored inside the database so they
# travel with every backup / restore)
# ---------------------------------------------------------------------------
BRAND_KINDS = {"logo": "Company logo", "qr": "Payment QR code"}
MAX_BRAND_UPLOAD = 4 * 1024 * 1024          # 4 MB
BRAND_MAX_SIDE = {"logo": 800, "qr": 1000}


def _ensure_brand_table(conn):
    conn.execute(
        "CREATE TABLE IF NOT EXISTS brand_assets (kind TEXT PRIMARY KEY, mime TEXT NOT NULL, "
        "data BLOB NOT NULL, filename TEXT, width INTEGER, height INTEGER, updated_at TEXT)")


def get_brand_meta():
    """{'logo': {...}|None, 'qr': {...}|None} without loading the image bytes."""
    out = {"logo": None, "qr": None}
    try:
        conn = get_db()
        _ensure_brand_table(conn)
        for r in conn.execute("SELECT kind, filename, width, height, updated_at, length(data) AS size FROM brand_assets"):
            v = (r["updated_at"] or "0").replace(" ", "").replace(":", "").replace("-", "")
            out[r["kind"]] = {"url": url_for("brand_image", kind=r["kind"], v=v), "filename": r["filename"],
                              "width": r["width"], "height": r["height"], "updated_at": r["updated_at"],
                              "size_kb": round((r["size"] or 0) / 1024)}
        conn.close()
    except Exception:
        log.exception("could not read brand assets")
    return out


def get_brand_bytes(conn, kind):
    _ensure_brand_table(conn)
    r = conn.execute("SELECT data, mime, width, height FROM brand_assets WHERE kind=?", (kind,)).fetchone()
    return r


def process_brand_image(file_storage, kind):
    """Validate and normalise an uploaded image. Returns (png_bytes, w, h, error)."""
    from PIL import Image, ImageOps, UnidentifiedImageError
    raw = file_storage.stream.read(MAX_BRAND_UPLOAD + 1)
    if not raw:
        return None, 0, 0, "Please choose an image file first."
    if len(raw) > MAX_BRAND_UPLOAD:
        return None, 0, 0, "That file is larger than 4 MB. Please choose a smaller image."
    try:
        im = Image.open(io.BytesIO(raw))
        fmt = (im.format or "").upper()
        if fmt not in ("PNG", "JPEG", "WEBP", "GIF", "MPO"):
            return None, 0, 0, "Unsupported file type. Please upload a PNG, JPG or WEBP image."
        if im.size[0] * im.size[1] > 40_000_000:
            return None, 0, 0, "That image is too large in pixels. Please use a smaller one."
        im.load()
        im = ImageOps.exif_transpose(im)
    except (UnidentifiedImageError, OSError, ValueError, Image.DecompressionBombError):
        return None, 0, 0, "That file doesn't look like a valid image."
    if min(im.size) < 40:
        return None, 0, 0, "That image is too small (minimum 40 px). Please use a clearer one."
    side = BRAND_MAX_SIDE[kind]
    if max(im.size) > side:
        im.thumbnail((side, side), Image.LANCZOS)
    if kind == "qr":                       # scanners need a solid white background
        rgba = im.convert("RGBA")
        bg = Image.new("RGB", rgba.size, (255, 255, 255))
        bg.paste(rgba, mask=rgba.split()[3])
        im = bg
    else:
        im = im.convert("RGBA")
    out = io.BytesIO()
    im.save(out, format="PNG", optimize=True)
    return out.getvalue(), im.size[0], im.size[1], None


DEFAULT_ACCOUNT_NAME = "Jyoti Sharma"


def payment_instruction(settings):
    return (f"Please check that the QR code shows {account_name_of(settings)} before making the payment. "
            "After payment, please send the payment screenshot to our WhatsApp number.")
UPI_ID_RE = re.compile(r"^[A-Za-z0-9._\-]{2,64}@[A-Za-z][A-Za-z0-9.\-]{1,30}$")


def account_name_of(settings):
    return (settings.get("upi_account_name") or "").strip() or DEFAULT_ACCOUNT_NAME


def valid_upi_id(v):
    return bool(v and UPI_ID_RE.match(v.strip()))


def make_upi_qr(upi_id, payee, amount=None, note=None):
    """Return (png_bytes, w, h) for a standard UPI payment QR.
    The payee name (pn) is embedded, so UPI apps show it when the code is scanned."""
    import qrcode
    from urllib.parse import quote
    uri = f"upi://pay?pa={quote(upi_id.strip(), safe='@')}&pn={quote(payee)}"
    if amount and amount > 0:
        uri += f"&am={amount:.2f}"
    uri += "&cu=INR"
    if note:
        uri += f"&tn={quote(note)}"
    qr = qrcode.QRCode(error_correction=qrcode.constants.ERROR_CORRECT_M, box_size=10, border=2)
    qr.add_data(uri)
    qr.make(fit=True)
    im = qr.make_image(fill_color="black", back_color="white").convert("RGB")
    out = io.BytesIO()
    im.save(out, format="PNG")
    return out.getvalue(), im.size[0], im.size[1]


def resolve_invoice_qr(conn, settings, invoice_number, balance_due):
    """Pick the QR for an invoice: generated from the UPI ID (carries the payee name and,
    while unpaid, the amount) or the uploaded image, according to Settings."""
    upi = (settings.get("upi_id") or "").strip()
    uploaded = get_brand_bytes(conn, "qr")
    if uploaded:
        use = "upload"
    elif valid_upi_id(upi):
        use = "generate"
    else:
        return None
    if use == "upload":
        return {"kind": "upload", "data": uploaded["data"], "w": uploaded["width"] or 1, "h": uploaded["height"] or 1}
    data, w, h = make_upi_qr(upi, account_name_of(settings),
                             amount=(balance_due if balance_due and balance_due > 0.005 else None), note=invoice_number)
    return {"kind": "generate", "data": data, "w": w, "h": h}


def brand_flowable(conn, kind, max_w_mm, max_h_mm):
    """ReportLab image scaled to fit the box (aspect ratio kept), or None."""
    r = get_brand_bytes(conn, kind)
    if not r:
        return None
    w, h = r["width"] or 1, r["height"] or 1
    scale = min(max_w_mm * mm / w, max_h_mm * mm / h)
    return RLImage(io.BytesIO(r["data"]), width=w * scale, height=h * scale)


@app.context_processor
def inject_globals():
    settings = get_settings()
    return {
        "today_label": datetime.now().strftime("%A, %d %b %Y"),
        "biz": settings,
        "brand": get_brand_meta(),
        "cur": settings.get("currency_symbol", "₹"),
    }


# ---------------------------------------------------------------------------
# Database connection
# ---------------------------------------------------------------------------
def get_db():
    conn = sqlite3.connect(DB_PATH, timeout=15)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    conn.execute("PRAGMA busy_timeout = 15000")
    if has_request_context():
        g.setdefault("_db_conns", []).append(conn)
    return conn


def _has_column(conn, table, col):
    cols = [r["name"] for r in conn.execute(f"PRAGMA table_info({table})").fetchall()]
    return col in cols


def _add_column_if_missing(conn, table, coldef):
    col = coldef.split()[0]
    if not _has_column(conn, table, col):
        conn.execute(f"ALTER TABLE {table} ADD COLUMN {coldef}")


# ---------------------------------------------------------------------------
# Schema v2: rate history, pauses, permanent snapshots (all additive)
# ---------------------------------------------------------------------------
SCHEMA_V2 = """
CREATE TABLE IF NOT EXISTS customer_rates (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    customer_id INTEGER NOT NULL,
    effective_month TEXT NOT NULL,              -- 'YYYY-MM': this plan/rate applies from this month onward
    monthly_rate REAL NOT NULL DEFAULT 0,
    rate_breakfast REAL NOT NULL DEFAULT 0,
    rate_lunch REAL NOT NULL DEFAULT 0,
    rate_dinner REAL NOT NULL DEFAULT 0,
    meal_breakfast INTEGER NOT NULL DEFAULT 0,
    meal_lunch INTEGER NOT NULL DEFAULT 0,
    meal_dinner INTEGER NOT NULL DEFAULT 0,
    note TEXT,
    created_at TEXT NOT NULL,
    UNIQUE (customer_id, effective_month),
    FOREIGN KEY (customer_id) REFERENCES customers(id)
);

CREATE TABLE IF NOT EXISTS customer_pauses (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    customer_id INTEGER NOT NULL,
    pause_from TEXT NOT NULL,
    pause_to TEXT,                              -- NULL = still paused
    created_at TEXT NOT NULL,
    FOREIGN KEY (customer_id) REFERENCES customers(id)
);

CREATE TABLE IF NOT EXISTS invoice_snapshots (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    invoice_id INTEGER NOT NULL,                -- deliberately no FK: voided snapshots outlive a deleted invoice
    invoice_number TEXT NOT NULL,
    customer_id INTEGER NOT NULL,
    ym TEXT NOT NULL,
    period_start TEXT NOT NULL,
    period_end TEXT NOT NULL,
    rate_monthly REAL NOT NULL DEFAULT 0,
    rate_effective_from TEXT,
    rate_mode TEXT,
    daily_rate REAL NOT NULL DEFAULT 0,
    days_in_period INTEGER NOT NULL DEFAULT 0,
    active_days INTEGER NOT NULL DEFAULT 0,
    leave_days INTEGER NOT NULL DEFAULT 0,
    billable_days INTEGER NOT NULL DEFAULT 0,
    meals_per_day INTEGER NOT NULL DEFAULT 0,
    scheduled_meals INTEGER NOT NULL DEFAULT 0,
    delivered_meals INTEGER NOT NULL DEFAULT 0,
    skipped_meals INTEGER NOT NULL DEFAULT 0,
    leave_meals INTEGER NOT NULL DEFAULT 0,
    calculated_amount REAL NOT NULL DEFAULT 0,
    extra_amount REAL NOT NULL DEFAULT 0,
    final_subscription_amount REAL NOT NULL DEFAULT 0,
    final_total REAL NOT NULL DEFAULT 0,
    status TEXT NOT NULL DEFAULT 'active',      -- 'active' | 'voided'
    source TEXT NOT NULL DEFAULT 'generated',   -- 'generated' | 'migrated'
    snapshot_json TEXT NOT NULL,
    created_at TEXT NOT NULL,
    voided_at TEXT
);
CREATE UNIQUE INDEX IF NOT EXISTS uq_invoice_snapshot_active
    ON invoice_snapshots (customer_id, ym) WHERE status = 'active';
CREATE INDEX IF NOT EXISTS ix_invoice_snapshot_invoice ON invoice_snapshots (invoice_id);

CREATE TABLE IF NOT EXISTS month_snapshots (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    ym TEXT NOT NULL,
    version INTEGER NOT NULL DEFAULT 1,
    status TEXT NOT NULL DEFAULT 'active',      -- 'active' (locks the month) | 'superseded' (month was reopened)
    closed_by TEXT NOT NULL DEFAULT 'auto',     -- 'auto' | 'manual'
    closed_at TEXT NOT NULL,
    superseded_at TEXT,
    billed REAL NOT NULL DEFAULT 0,
    revenue REAL NOT NULL DEFAULT 0,
    expense REAL NOT NULL DEFAULT 0,
    profit REAL NOT NULL DEFAULT 0,
    personal_expense REAL NOT NULL DEFAULT 0,
    savings_deposits REAL NOT NULL DEFAULT 0,
    remaining_funds REAL NOT NULL DEFAULT 0,
    snapshot_json TEXT NOT NULL,
    UNIQUE (ym, version)
);
CREATE UNIQUE INDEX IF NOT EXISTS uq_month_snapshot_active
    ON month_snapshots (ym) WHERE status = 'active';
"""

LOCKED_DATE_TABLES = ["revenue", "expenses", "extra_orders", "savings", "personal_income", "daily_log"]
_INVOICE_SNAPSHOT_FROZEN_COLUMNS = [
    "invoice_id", "invoice_number", "customer_id", "ym", "period_start", "period_end",
    "rate_monthly", "rate_effective_from", "rate_mode", "daily_rate", "days_in_period",
    "active_days", "leave_days", "billable_days", "meals_per_day", "scheduled_meals",
    "delivered_meals", "skipped_meals", "leave_meals", "calculated_amount", "extra_amount",
    "source", "snapshot_json", "created_at",
]
_MONTH_SNAPSHOT_FROZEN_COLUMNS = [
    "ym", "version", "closed_by", "closed_at", "billed", "revenue", "expense", "profit",
    "personal_expense", "savings_deposits", "remaining_funds", "snapshot_json",
]


def _closed_expr(date_expr):
    return ("EXISTS (SELECT 1 FROM month_snapshots "
            f"WHERE ym = strftime('%Y-%m', {date_expr}) AND status = 'active')")


def install_integrity_triggers(conn):
    """Database-level guarantees (they hold no matter which code path touches the data):
      * rows dated inside a closed month can't be inserted, changed or deleted
      * saved snapshots can never be edited or deleted (only voided / superseded)
      * a customer's rate history can't be changed for months they've already been invoiced for
    Triggers are dropped and re-created on every start so upgrades always get the latest rules."""
    stmts = []

    def add(name, timing, table, when, message, of_cols=None):
        stmts.append(f"DROP TRIGGER IF EXISTS {name}")
        of = f" OF {', '.join(of_cols)}" if of_cols else ""
        stmts.append(
            f"CREATE TRIGGER {name} BEFORE {timing}{of} ON {table} WHEN {when} "
            f"BEGIN SELECT RAISE(ABORT, '{message}'); END")

    for t in LOCKED_DATE_TABLES:
        add(f"lock_{t}_ins", "INSERT", t, _closed_expr("NEW.date"), "MONTH_LOCKED")
        add(f"lock_{t}_upd", "UPDATE", t, f"{_closed_expr('OLD.date')} OR {_closed_expr('NEW.date')}", "MONTH_LOCKED")
        add(f"lock_{t}_del", "DELETE", t, _closed_expr("OLD.date"), "MONTH_LOCKED")

    add("lock_invoices_ins", "INSERT", "invoices", _closed_expr("NEW.period_start"), "MONTH_LOCKED")
    add("lock_invoices_upd", "UPDATE", "invoices",
        f"{_closed_expr('OLD.period_start')} OR {_closed_expr('NEW.period_start')}", "MONTH_LOCKED")
    add("lock_invoices_del", "DELETE", "invoices", _closed_expr("OLD.period_start"), "MONTH_LOCKED")

    _ps = "(SELECT period_start FROM invoices WHERE id={}.invoice_id)"
    add("lock_invoice_lines_ins", "INSERT", "invoice_lines", _closed_expr(_ps.format("NEW")), "MONTH_LOCKED")
    add("lock_invoice_lines_upd", "UPDATE", "invoice_lines", _closed_expr(_ps.format("OLD")), "MONTH_LOCKED")
    add("lock_invoice_lines_del", "DELETE", "invoice_lines", _closed_expr(_ps.format("OLD")), "MONTH_LOCKED")

    def rate_locked(prefix):
        return ("EXISTS (SELECT 1 FROM invoice_snapshots s "
                f"WHERE s.customer_id = {prefix}.customer_id AND s.status = 'active' "
                f"AND s.ym >= {prefix}.effective_month)")
    add("lock_rates_ins", "INSERT", "customer_rates", rate_locked("NEW"), "RATE_LOCKED")
    add("lock_rates_upd", "UPDATE", "customer_rates", f"{rate_locked('OLD')} OR {rate_locked('NEW')}", "RATE_LOCKED")
    add("lock_rates_del", "DELETE", "customer_rates", rate_locked("OLD"), "RATE_LOCKED")

    add("freeze_invoice_snapshots_upd", "UPDATE", "invoice_snapshots", "1", "SNAPSHOT_IMMUTABLE",
        of_cols=_INVOICE_SNAPSHOT_FROZEN_COLUMNS)
    add("freeze_invoice_snapshots_del", "DELETE", "invoice_snapshots", "1", "SNAPSHOT_IMMUTABLE")
    add("freeze_month_snapshots_upd", "UPDATE", "month_snapshots", "1", "SNAPSHOT_IMMUTABLE",
        of_cols=_MONTH_SNAPSHOT_FROZEN_COLUMNS)
    add("freeze_month_snapshots_del", "DELETE", "month_snapshots", "1", "SNAPSHOT_IMMUTABLE")

    for stmt in stmts:
        conn.execute(stmt)


# ---------------------------------------------------------------------------
# Schema creation + migration (non-destructive — preserves existing data)
# ---------------------------------------------------------------------------
def init_db():
    conn = get_db()
    # WAL mode lets readers and writers work without blocking each other, which is
    # the standard fix for "database is locked" errors in a Flask + SQLite app.
    # It's a persistent setting stored in the database file, so this only needs
    # to run once, but is safe to run on every startup.
    conn.execute("PRAGMA journal_mode=WAL")

    conn.executescript(
        """
        CREATE TABLE IF NOT EXISTS pgs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            address TEXT,
            contact_person TEXT,
            contact_number TEXT,
            notes TEXT
        );

        CREATE TABLE IF NOT EXISTS customers (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            type TEXT NOT NULL DEFAULT 'Regular',
            area TEXT,
            contact TEXT,
            plan TEXT,
            monthly_rate REAL NOT NULL DEFAULT 0,
            join_date TEXT,
            status TEXT NOT NULL DEFAULT 'Active'
        );

        CREATE TABLE IF NOT EXISTS revenue (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            date TEXT NOT NULL,
            customer_id INTEGER NOT NULL,
            amount REAL NOT NULL,
            payment_mode TEXT,
            notes TEXT,
            FOREIGN KEY (customer_id) REFERENCES customers(id)
        );

        CREATE TABLE IF NOT EXISTS expenses (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            date TEXT NOT NULL,
            category TEXT NOT NULL,
            description TEXT,
            amount REAL NOT NULL,
            payment_mode TEXT
        );

        CREATE TABLE IF NOT EXISTS daily_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            customer_id INTEGER NOT NULL,
            date TEXT NOT NULL,
            meal TEXT NOT NULL,
            status TEXT NOT NULL,
            notes TEXT,
            UNIQUE(customer_id, date, meal),
            FOREIGN KEY (customer_id) REFERENCES customers(id)
        );

        CREATE TABLE IF NOT EXISTS extra_orders (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            date TEXT NOT NULL,
            customer_id INTEGER,
            guest_name TEXT,
            description TEXT,
            amount REAL NOT NULL,
            payment_status TEXT NOT NULL DEFAULT 'Unpaid',
            payment_mode TEXT,
            invoice_id INTEGER,
            FOREIGN KEY (customer_id) REFERENCES customers(id)
        );

        CREATE TABLE IF NOT EXISTS invoices (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            invoice_number TEXT UNIQUE NOT NULL,
            customer_id INTEGER NOT NULL,
            period_start TEXT NOT NULL,
            period_end TEXT NOT NULL,
            subscription_amount REAL NOT NULL DEFAULT 0,
            extra_amount REAL NOT NULL DEFAULT 0,
            total_amount REAL NOT NULL DEFAULT 0,
            generated_date TEXT NOT NULL,
            due_date TEXT,
            notes TEXT,
            FOREIGN KEY (customer_id) REFERENCES customers(id)
        );

        CREATE TABLE IF NOT EXISTS recurring_expenses (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            category TEXT NOT NULL,
            description TEXT,
            amount REAL NOT NULL,
            day_of_month INTEGER NOT NULL DEFAULT 1,
            active INTEGER NOT NULL DEFAULT 1,
            last_generated_month TEXT,
            payment_mode TEXT NOT NULL DEFAULT 'Cash'
        );

        CREATE TABLE IF NOT EXISTS settings (
            key TEXT PRIMARY KEY,
            value TEXT
        );

        CREATE TABLE IF NOT EXISTS brand_assets (
            kind TEXT PRIMARY KEY,          -- 'logo' or 'qr'
            mime TEXT NOT NULL,
            data BLOB NOT NULL,
            filename TEXT,
            width INTEGER,
            height INTEGER,
            updated_at TEXT
        );

        CREATE TABLE IF NOT EXISTS savings (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            date TEXT NOT NULL,
            amount REAL NOT NULL,
            entry_type TEXT NOT NULL DEFAULT 'Deposit',
            description TEXT,
            payment_mode TEXT
        );

        CREATE TABLE IF NOT EXISTS personal_income (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            date TEXT NOT NULL,
            source TEXT NOT NULL,
            amount REAL NOT NULL,
            description TEXT,
            payment_mode TEXT
        );
        """
    )
    conn.commit()

    conn.executescript(SCHEMA_V2)
    conn.commit()

    _add_column_if_missing(conn, "customers", "category TEXT")
    _add_column_if_missing(conn, "customers", "pg_id INTEGER")
    _add_column_if_missing(conn, "customers", "meal_breakfast INTEGER DEFAULT 0")
    _add_column_if_missing(conn, "customers", "meal_lunch INTEGER DEFAULT 0")
    _add_column_if_missing(conn, "customers", "meal_dinner INTEGER DEFAULT 0")
    _add_column_if_missing(conn, "customers", "rate_breakfast REAL DEFAULT 0")
    _add_column_if_missing(conn, "customers", "rate_lunch REAL DEFAULT 0")
    _add_column_if_missing(conn, "customers", "rate_dinner REAL DEFAULT 0")
    _add_column_if_missing(conn, "customers", "pause_from TEXT")
    _add_column_if_missing(conn, "customers", "pause_to TEXT")
    _add_column_if_missing(conn, "customers", "notes TEXT")
    _add_column_if_missing(conn, "revenue", "invoice_id INTEGER")
    _add_column_if_missing(conn, "expenses", "recurring_expense_id INTEGER")
    _add_column_if_missing(conn, "recurring_expenses", "payment_mode TEXT DEFAULT 'Cash'")
    _add_column_if_missing(conn, "expenses", "expense_type TEXT DEFAULT 'Business'")
    conn.execute("UPDATE expenses SET expense_type='Business' WHERE expense_type IS NULL")
    conn.commit()

    legacy_rows = conn.execute(
        "SELECT id, type, plan FROM customers WHERE category IS NULL"
    ).fetchall()
    for row in legacy_rows:
        cat = "PG" if (row["type"] or "").strip() == "PG" else "Regular"
        plan_text = (row["plan"] or "").lower()
        mb = 1 if "breakfast" in plan_text else 0
        ml = 1 if "lunch" in plan_text else 0
        md = 1 if "dinner" in plan_text else 0
        if not (mb or ml or md):
            ml = 1
        conn.execute(
            """UPDATE customers SET category=?, meal_breakfast=?, meal_lunch=?, meal_dinner=?
               WHERE id=?""",
            (cat, mb, ml, md, row["id"]),
        )
    conn.commit()

    defaults = {
        "business_name": "Shudh Rasoi",
        "business_address": "",
        "business_phone": "",
        "gst_number": "",
        "currency_symbol": "₹",
        "invoice_prefix": "INV",
        "invoice_due_days": "7",
        "opening_balance_cash": "0",
        "opening_balance_upi": "0",
        "opening_balance_bank_transfer": "0",
        "opening_balance_other": "0",
        "auto_close_months": "1",
        "month_close_grace_days": "5",
    }
    for k, v in defaults.items():
        conn.execute("INSERT OR IGNORE INTO settings (key, value) VALUES (?,?)", (k, v))
    conn.commit()

    cust_count = conn.execute("SELECT COUNT(*) FROM customers").fetchone()[0]
    pg_count = conn.execute("SELECT COUNT(*) FROM pgs").fetchone()[0]
    if cust_count == 0 and pg_count == 0 and os.environ.get("TIFFIN_DEMO") == "1":   # demo data only on request
        conn.execute(
            """INSERT INTO pgs (name, address, contact_person, contact_number, notes)
               VALUES (?,?,?,?,?)""",
            ("Om Boys PG", "Malviya Nagar, Jaipur", "Mr. Om Prakash", "9829099999",
             "Sample PG — feel free to edit or delete"),
        )
        pg_id = conn.execute("SELECT id FROM pgs WHERE name='Om Boys PG'").fetchone()["id"]
        conn.execute(
            """INSERT INTO customers (name, type, category, pg_id, area, contact, plan,
                                       monthly_rate, meal_breakfast, meal_lunch, meal_dinner,
                                       join_date, status)
               VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)""",
            ("Rahul Sharma", "PG", "PG", pg_id, "Om Boys PG, Malviya Nagar", "9829011111",
             "Lunch + Dinner", 3600, 0, 1, 1, today_date().replace(day=1).isoformat(), "Active"),
        )
        conn.execute(
            """INSERT INTO customers (name, type, category, area, contact, plan,
                                       monthly_rate, meal_breakfast, meal_lunch, meal_dinner,
                                       join_date, status)
               VALUES (?,?,?,?,?,?,?,?,?,?,?,?)""",
            ("Priya Verma", "Individual", "Regular", "", "9829022222", "Dinner Only",
             1800, 0, 0, 1, today_date().replace(day=1).isoformat(), "Active"),
        )
        conn.commit()
        c1 = conn.execute("SELECT id FROM customers WHERE name='Rahul Sharma'").fetchone()["id"]
        c2 = conn.execute("SELECT id FROM customers WHERE name='Priya Verma'").fetchone()["id"]
        conn.execute(
            "INSERT INTO revenue (date, customer_id, amount, payment_mode, notes) VALUES (?,?,?,?,?)",
            (today_date().replace(day=1).isoformat(), c1, 3600, "UPI", "Sample advance payment"),
        )
        conn.execute(
            "INSERT INTO revenue (date, customer_id, amount, payment_mode, notes) VALUES (?,?,?,?,?)",
            (today_date().replace(day=1).isoformat(), c2, 1800, "Cash", "Sample advance payment"),
        )
        conn.execute(
            """INSERT INTO expenses (date, category, description, amount, payment_mode)
               VALUES (?,?,?,?,?)""",
            (today_date().replace(day=1).isoformat(), "Raw Material / Groceries",
             "Vegetables & atta - weekly (sample)", 4200, "Cash"),
        )
        conn.execute(
            """INSERT INTO expenses (date, category, description, amount, payment_mode)
               VALUES (?,?,?,?,?)""",
            (today_date().isoformat(), "Gas Cylinder", "Commercial LPG refill (sample)", 1150, "UPI"),
        )
        conn.commit()

    # --- sign-in + invoice-editing schema (additive; existing rows are untouched) ----
    _auth.ensure_auth_schema(conn)
    ensure_invoice_edit_schema(conn)
    conn.commit()

    # --- rate history, pauses, invoice snapshots, month locking ------------------
    seed_rate_history(conn)
    seed_pause_history(conn)
    conn.commit()
    backfill_invoice_snapshots(conn)
    install_integrity_triggers(conn)
    conn.commit()
    sync_customer_current_plans(conn)
    auto_close_elapsed_months(conn)
    conn.commit()
    conn.close()


# ---------------------------------------------------------------------------
# Date / month helpers
# ---------------------------------------------------------------------------
def prev_month_str(ym):
    """Given 'YYYY-MM', return the previous month as 'YYYY-MM'."""
    y, m = map(int, ym.split("-"))
    m -= 1
    if m == 0:
        m = 12
        y -= 1
    return f"{y:04d}-{m:02d}"


def month_bounds(ym=None):
    """Return (start_date, end_date, ym_string) for a given 'YYYY-MM' or current month.
    Falls back to the current month for a missing, malformed, or out-of-range ym (e.g. a
    hand-edited URL or a stale bookmark) instead of crashing the page."""
    today = today_date()
    y, m = today.year, today.month
    if ym:
        try:
            yy, mm = map(int, str(ym).split("-"))
            date(yy, mm, 1)  # validates the year/month combination is real
            y, m = yy, mm
        except (ValueError, TypeError):
            pass
    start = date(y, m, 1)
    last_day = calmod.monthrange(y, m)[1]
    end = date(y, m, last_day)
    return start, end, f"{y:04d}-{m:02d}"


def current_ym():
    return today_date().strftime("%Y-%m")


def month_label(ym):
    y, m = map(int, ym.split("-"))
    return f"{calmod.month_name[m][:3]} {y}"


def daterange(d1, d2):
    for n in range((d2 - d1).days + 1):
        yield d1 + timedelta(days=n)


# ---------------------------------------------------------------------------
# Recurring expenses — auto-generate due entries (called on relevant page loads)
# ---------------------------------------------------------------------------
def run_recurring_expenses(conn):
    today = today_date()
    ym = today.strftime("%Y-%m")
    rows = conn.execute("SELECT * FROM recurring_expenses WHERE active=1").fetchall()
    for r in rows:
        if r["last_generated_month"] == ym:
            continue
        if today.day < r["day_of_month"]:
            continue
        conn.execute(
            """INSERT INTO expenses (date, category, description, amount, payment_mode, recurring_expense_id)
               VALUES (?,?,?,?,?,?)""",
            (today.isoformat(), r["category"],
             (r["description"] or "") + " (auto-generated recurring expense)",
             r["amount"], r["payment_mode"] or "Cash", r["id"]),
        )
        conn.execute(
            "UPDATE recurring_expenses SET last_generated_month=? WHERE id=?", (ym, r["id"])
        )
    conn.commit()


def consistent_db_copy(src_path, dst_path):
    """Copy a SQLite database using SQLite's own backup API. Unlike copying the file, this
    includes everything still sitting in the write-ahead log, so a backup is never missing
    the most recent entries."""
    src = sqlite3.connect(src_path)
    try:
        dst = sqlite3.connect(dst_path)
        try:
            src.backup(dst)
        finally:
            dst.close()
    finally:
        src.close()


def run_auto_backup():
    """Silently keep a rolling daily safety-net backup of the database, independent of
    the manual Settings > Backup download. Runs at most once per day; keeps the most
    recent AUTO_BACKUPS_TO_KEEP copies and prunes older ones. Never raises — a backup
    failure should never take down the app."""
    try:
        if not os.path.exists(DB_PATH):
            return
        os.makedirs(BACKUP_DIR, exist_ok=True)
        today_str = today_date().isoformat()
        target = os.path.join(BACKUP_DIR, f"tiffin_auto_{today_str}.db")
        if os.path.exists(target):
            return  # already backed up today
        consistent_db_copy(DB_PATH, target)
        existing = sorted(glob.glob(os.path.join(BACKUP_DIR, "tiffin_auto_*.db")))
        for old_file in existing[:-AUTO_BACKUPS_TO_KEEP]:
            os.remove(old_file)
    except (OSError, sqlite3.Error):
        pass  # best-effort — a backup hiccup should never break the app


# ---------------------------------------------------------------------------
# Form validation helpers — every monetary/numeric form field is parsed through
# these so bad input (blank, non-numeric, negative) always produces a friendly
# flash message and a redirect, never a 500 error page.
# ---------------------------------------------------------------------------
def parse_amount(raw, allow_zero=False, field_label="Amount"):
    """Parse a monetary amount from form input. Returns (value, error_message) —
    exactly one of the two is None. Rejects blank/non-numeric/negative input;
    rejects zero unless allow_zero=True."""
    if raw is None or str(raw).strip() == "":
        return None, f"Please enter {('an' if field_label[0].lower() in 'aeiou' else 'a')} {field_label.lower()}."
    try:
        value = float(raw)
    except (TypeError, ValueError):
        return None, f"{field_label} must be a valid number."
    if not math.isfinite(value):
        return None, f"{field_label} must be a valid number."
    if value > MAX_AMOUNT:
        return None, f"{field_label} is too large (maximum ₹1,00,00,000)."
    if value < 0:
        return None, f"{field_label} cannot be negative."
    if value == 0 and not allow_zero:
        return None, f"{field_label} must be greater than zero."
    return round(value, 2), None


def parse_int(raw, default=None, min_value=None, max_value=None, field_label="Value"):
    """Parse an integer from form input. Returns (value, error_message)."""
    if raw is None or str(raw).strip() == "":
        if default is not None:
            return default, None
        return None, f"Please enter a {field_label.lower()}."
    try:
        value = int(raw)
    except (TypeError, ValueError):
        return None, f"{field_label} must be a whole number."
    if min_value is not None and value < min_value:
        return None, f"{field_label} must be {min_value} or more."
    if max_value is not None and value > max_value:
        return None, f"{field_label} must be {max_value} or less."
    return value, None


def safe_float(raw, default=0.0):
    """Best-effort float parse for values already stored in the database (e.g. settings).
    Never raises — falls back to `default` on anything invalid, so a legacy or corrupted
    stored value can never crash a page. (Form input should go through parse_amount
    instead, so the user gets a clear error at save time.)"""
    try:
        return float(raw)
    except (TypeError, ValueError):
        return default


def safe_int(raw, default=0):
    """Best-effort int parse for values already stored in the database. Never raises."""
    try:
        return int(raw)
    except (TypeError, ValueError):
        return default


# ---------------------------------------------------------------------------
# Customer meal / subscription helpers
# ---------------------------------------------------------------------------
def customer_meals(c):
    meals = []
    if c["meal_breakfast"]:
        meals.append("breakfast")
    if c["meal_lunch"]:
        meals.append("lunch")
    if c["meal_dinner"]:
        meals.append("dinner")
    return meals


def plan_label(c):
    meals = [MEAL_LABELS[m] for m in customer_meals(c)]
    return " + ".join(meals) if meals else "No meals selected"


def effective_monthly_rate(c):
    per_meal_sum = (c["rate_breakfast"] or 0) + (c["rate_lunch"] or 0) + (c["rate_dinner"] or 0)
    return per_meal_sum if per_meal_sum > 0 else (c["monthly_rate"] or 0)


def _parse_iso(s):
    return datetime.strptime(s, "%Y-%m-%d").date()


def _now_iso():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def next_month_str(ym):
    y, m = map(int, ym.split("-"))
    m += 1
    if m == 13:
        m = 1
        y += 1
    return f"{y:04d}-{m:02d}"


def shift_month(ym, delta):
    for _ in range(abs(delta)):
        ym = next_month_str(ym) if delta > 0 else prev_month_str(ym)
    return ym


def valid_ym(raw):
    try:
        yy, mm = map(int, str(raw).split("-"))
        date(yy, mm, 1)
        return f"{yy:04d}-{mm:02d}"
    except (ValueError, TypeError):
        return None


def months_between(first_ym, last_ym):
    out, ym = [], first_ym
    while ym <= last_ym:
        out.append(ym)
        ym = next_month_str(ym)
    return out


# ---------------------------------------------------------------------------
# Closed-month helpers
# ---------------------------------------------------------------------------
def is_month_closed(conn, ym):
    return conn.execute(
        "SELECT 1 FROM month_snapshots WHERE ym=? AND status='active'", (ym,)
    ).fetchone() is not None


def latest_closed_month(conn):
    return conn.execute(
        "SELECT MAX(ym) FROM month_snapshots WHERE status='active'"
    ).fetchone()[0]


def assert_month_open(conn, date_str):
    if date_str and len(date_str) >= 7 and is_month_closed(conn, date_str[:7]):
        raise MonthLocked(date_str[:7])


# ---------------------------------------------------------------------------
# Rate / plan history — a customer's meals + rates are versioned by effective month
# ---------------------------------------------------------------------------
PLAN_FIELDS = ("monthly_rate", "rate_breakfast", "rate_lunch", "rate_dinner",
               "meal_breakfast", "meal_lunch", "meal_dinner")


def plan_for_month(conn, customer_id, ym):
    """The plan/rate row in force for `ym`: the latest row whose effective_month <= ym.
    Months before a customer's first row use that first row (they had no earlier plan)."""
    row = conn.execute(
        """SELECT * FROM customer_rates WHERE customer_id=? AND effective_month<=?
           ORDER BY effective_month DESC LIMIT 1""", (customer_id, ym)).fetchone()
    if row is None:
        row = conn.execute(
            """SELECT * FROM customer_rates WHERE customer_id=?
               ORDER BY effective_month ASC LIMIT 1""", (customer_id,)).fetchone()
    if row is not None:
        return dict(row)
    c = conn.execute("SELECT * FROM customers WHERE id=?", (customer_id,)).fetchone()
    d = {k: (c[k] if c is not None and c[k] is not None else 0) for k in PLAN_FIELDS}
    d.update({"customer_id": customer_id, "effective_month": None, "id": None})
    return d


def plan_overlay(conn, c, ym):
    """A copy of the customer row with the plan/rates that applied in `ym`."""
    d = dict(c)
    p = plan_for_month(conn, c["id"], ym)
    for k in PLAN_FIELDS:
        d[k] = p[k] or 0
    d["plan_effective_from"] = p.get("effective_month")
    d["plan"] = plan_label(d) if customer_meals(d) else ""
    return d


def _plan_values_from(row):
    return {k: (row[k] or 0) for k in PLAN_FIELDS}


def seed_rate_history(conn, customer_id=None):
    """Guarantee every customer has at least one rate-history row (migration + new customers)."""
    sql = "SELECT * FROM customers" + (" WHERE id=?" if customer_id else "")
    for c in conn.execute(sql, (customer_id,) if customer_id else ()).fetchall():
        if conn.execute("SELECT 1 FROM customer_rates WHERE customer_id=? LIMIT 1", (c["id"],)).fetchone():
            continue
        months = []
        jd = (c["join_date"] or "")[:7]
        if valid_ym(jd):
            months.append(jd)
        first_inv = conn.execute(
            "SELECT MIN(substr(period_start,1,7)) FROM invoices WHERE customer_id=?", (c["id"],)
        ).fetchone()[0]
        if first_inv:
            months.append(first_inv)
        eff = min(months) if months else current_ym()
        vals = _plan_values_from(c)
        conn.execute(
            """INSERT INTO customer_rates (customer_id, effective_month, monthly_rate, rate_breakfast,
                   rate_lunch, rate_dinner, meal_breakfast, meal_lunch, meal_dinner, note, created_at)
               VALUES (?,?,?,?,?,?,?,?,?,?,?)""",
            (c["id"], eff, vals["monthly_rate"], vals["rate_breakfast"], vals["rate_lunch"],
             vals["rate_dinner"], vals["meal_breakfast"], vals["meal_lunch"], vals["meal_dinner"],
             "Initial plan", _now_iso()))


def set_plan_from_month(conn, customer_id, eff_month, plan, note=None):
    """Create or replace the plan row that takes effect from `eff_month`. Rows for later
    months are untouched, so a change only ever affects eff_month up to the next change."""
    existing = conn.execute(
        "SELECT id FROM customer_rates WHERE customer_id=? AND effective_month=?",
        (customer_id, eff_month)).fetchone()
    cols = ", ".join(f"{k}=?" for k in PLAN_FIELDS)
    vals = [plan[k] or 0 for k in PLAN_FIELDS]
    if existing:
        conn.execute(f"UPDATE customer_rates SET {cols}, note=? WHERE id=?", vals + [note, existing["id"]])
    else:
        conn.execute(
            f"""INSERT INTO customer_rates (customer_id, effective_month, {', '.join(PLAN_FIELDS)}, note, created_at)
                VALUES (?,?,{','.join('?' * len(PLAN_FIELDS))},?,?)""",
            [customer_id, eff_month] + vals + [note, _now_iso()])


def earliest_editable_month(conn, customer_id):
    """First month a plan/rate change may take effect from: after every closed month and
    after the last month this customer has already been invoiced for. None = no limit."""
    candidates = []
    lc = latest_closed_month(conn)
    if lc:
        candidates.append(next_month_str(lc))
    last_inv = conn.execute(
        "SELECT MAX(ym) FROM invoice_snapshots WHERE customer_id=? AND status='active'", (customer_id,)
    ).fetchone()[0]
    if last_inv:
        candidates.append(next_month_str(last_inv))
    return max(candidates) if candidates else None


def rate_history_rows(conn, customer_id):
    cur_ym = current_ym()
    rows = conn.execute(
        "SELECT * FROM customer_rates WHERE customer_id=? ORDER BY effective_month DESC", (customer_id,)
    ).fetchall()
    in_force = plan_for_month(conn, customer_id, cur_ym).get("id")
    last_inv = conn.execute(
        "SELECT MAX(ym) FROM invoice_snapshots WHERE customer_id=? AND status='active'", (customer_id,)
    ).fetchone()[0]
    out = []
    for r in rows:
        d = dict(r)
        d["effective_rate"] = effective_monthly_rate(d)
        d["plan_text"] = plan_label(d) if customer_meals(d) else "No meals"
        d["status"] = "Current" if d["id"] == in_force else ("Scheduled" if d["effective_month"] > cur_ym else "Past")
        d["removable"] = (d["effective_month"] > cur_ym and (not last_inv or d["effective_month"] > last_inv)
                          and len(rows) > 1)
        d["label"] = month_label(d["effective_month"])
        out.append(d)
    return out


def sync_customer_current_plans(conn, customer_id=None):
    """Keep customers.* (used for display/search) mirroring the plan in force this month."""
    ym = current_ym()
    sql = "SELECT * FROM customers" + (" WHERE id=?" if customer_id else "")
    for c in conn.execute(sql, (customer_id,) if customer_id else ()).fetchall():
        p = plan_for_month(conn, c["id"], ym)
        label = plan_label(p) if customer_meals(p) else ""
        same = all((c[k] or 0) == (p[k] or 0) for k in PLAN_FIELDS) and (c["plan"] or "") == label
        if same:
            continue
        conn.execute(
            f"UPDATE customers SET {', '.join(k + '=?' for k in PLAN_FIELDS)}, plan=? WHERE id=?",
            [p[k] or 0 for k in PLAN_FIELDS] + [label, c["id"]])


# ---------------------------------------------------------------------------
# Pause history — remembered even after the customer resumes
# ---------------------------------------------------------------------------
def seed_pause_history(conn):
    for c in conn.execute(
        "SELECT id, pause_from, pause_to FROM customers WHERE status='Paused' AND pause_from IS NOT NULL"
    ).fetchall():
        if not conn.execute("SELECT 1 FROM customer_pauses WHERE customer_id=?", (c["id"],)).fetchone():
            conn.execute(
                "INSERT INTO customer_pauses (customer_id, pause_from, pause_to, created_at) VALUES (?,?,?,?)",
                (c["id"], c["pause_from"], c["pause_to"], _now_iso()))


def get_pause_ranges(conn, customer_id):
    return [(r["pause_from"], r["pause_to"]) for r in conn.execute(
        "SELECT pause_from, pause_to FROM customer_pauses WHERE customer_id=? ORDER BY pause_from",
        (customer_id,)).fetchall()]


def start_pause(conn, customer_id, pause_from, pause_to):
    end_open_pauses(conn, customer_id, pause_from)
    conn.execute(
        "INSERT INTO customer_pauses (customer_id, pause_from, pause_to, created_at) VALUES (?,?,?,?)",
        (customer_id, pause_from, pause_to, _now_iso()))


def end_open_pauses(conn, customer_id, from_day_iso):
    """Close every pause still running on/after `from_day_iso` the day before it
    (or drop it if it hadn't started yet)."""
    last_paused_day = (_parse_iso(from_day_iso) - timedelta(days=1)).isoformat()
    rows = conn.execute(
        "SELECT * FROM customer_pauses WHERE customer_id=? AND (pause_to IS NULL OR pause_to>=?)",
        (customer_id, from_day_iso)).fetchall()
    for r in rows:
        if r["pause_from"] >= from_day_iso:
            conn.execute("DELETE FROM customer_pauses WHERE id=?", (r["id"],))
        else:
            conn.execute("UPDATE customer_pauses SET pause_to=? WHERE id=?", (last_paused_day, r["id"]))


# ---------------------------------------------------------------------------
# Tiffin counting + billing calculation (one implementation used everywhere)
# ---------------------------------------------------------------------------
def tiffin_counts(meals, join_date, start, end, exceptions, pauses, cap_end=None):
    """Meal-by-meal status over a period, from first principles.

    Default for every scheduled meal is Delivered; daily_log exceptions (Skipped / Leave)
    and pause ranges override that. Never counts days before the join date. `cap_end` stops
    the count early (used for the live tracker of the month in progress)."""
    last = end if cap_end is None else min(end, cap_end)
    active_start = start
    if join_date and join_date > start.isoformat():
        try:
            active_start = max(start, _parse_iso(join_date))
        except ValueError:
            pass
    out = {"meals": list(meals), "active_start": active_start.isoformat(), "days": 0,
           "delivered": 0, "skipped": 0, "leave": 0, "scheduled": 0, "leave_days": 0,
           "breakdown": [], "not_joined_yet": False}
    if active_start > last:
        out["not_joined_yet"] = bool(join_date) and join_date > last.isoformat()
        return out
    ex_map = {(e["date"], e["meal"]): e["status"] for e in exceptions}
    leave_dates = {e["date"] for e in exceptions if e["status"] == "Leave"}
    for d in daterange(active_start, last):
        d_iso = d.isoformat()
        on_pause = any(pf <= d_iso <= (pt or "9999-12-31") for pf, pt in pauses)
        statuses = {}
        for m in meals:
            status = "Leave" if on_pause else ex_map.get((d_iso, m), "Delivered")
            statuses[m] = status
            if status == "Delivered":
                out["delivered"] += 1
            elif status == "Skipped":
                out["skipped"] += 1
            else:
                out["leave"] += 1
        out["scheduled"] += len(meals)
        out["days"] += 1
        if on_pause or d_iso in leave_dates:
            out["leave_days"] += 1
        out["breakdown"].append({"date": d_iso, "statuses": statuses})
    return out


def compute_billing_calc(conn, c, ym):
    """Everything about how a customer's bill for `ym` is worked out — using the plan and
    rate that were in force *that month* (never the customer's current rate)."""
    start, end, ym = month_bounds(ym)
    plan = plan_overlay(conn, c, ym)
    meals = customer_meals(plan)
    rate = effective_monthly_rate(plan)
    per_meal = ((plan["rate_breakfast"] or 0) + (plan["rate_lunch"] or 0) + (plan["rate_dinner"] or 0)) > 0
    exceptions = [dict(r) for r in conn.execute(
        "SELECT date, meal, status FROM daily_log WHERE customer_id=? AND date BETWEEN ? AND ? ORDER BY date, meal",
        (c["id"], start.isoformat(), end.isoformat())).fetchall()]
    pauses = get_pause_ranges(conn, c["id"])
    t = tiffin_counts(meals, plan.get("join_date"), start, end, exceptions, pauses)
    days_in_period = (end - start).days + 1
    billable = max(t["days"] - t["leave_days"], 0)
    daily_rate = (rate / days_in_period) if days_in_period else 0.0
    amount = round(max(daily_rate * billable, 0), 2) if rate > 0 else 0.0
    return {
        "ym": ym,
        "period_start": start.isoformat(), "period_end": end.isoformat(),
        "days_in_period": days_in_period,
        "join_date": plan.get("join_date"),
        "active_start": t["active_start"], "active_days": t["days"],
        "leave_days": t["leave_days"], "billable_days": billable,
        "meals": meals, "meals_per_day": len(meals),
        "scheduled_meals": t["scheduled"], "delivered_meals": t["delivered"],
        "skipped_meals": t["skipped"], "leave_meals": t["leave"],
        "rate": {
            "mode": "per_meal" if per_meal else "flat", "monthly": rate,
            "flat": plan["monthly_rate"] or 0, "breakfast": plan["rate_breakfast"] or 0,
            "lunch": plan["rate_lunch"] or 0, "dinner": plan["rate_dinner"] or 0,
            "effective_from": plan.get("plan_effective_from"),
        },
        "daily_rate": daily_rate,
        "calculated_amount": amount,
        "inputs": {
            "exceptions": [e for e in exceptions if e["date"] >= t["active_start"]],
            "pauses": [list(p) for p in pauses],
        },
    }


def compute_subscription_amount(conn, c, start, end):
    """Backwards-compatible wrapper — the month's prorated subscription charge."""
    return compute_billing_calc(conn, c, start.strftime("%Y-%m"))["calculated_amount"]


def get_monthly_tiffin_summary(conn, ym):
    """Per-customer tiffin totals for a month. Closed months are served from their saved
    snapshot; open months are computed live (counted only up to today)."""
    start, end, ym = month_bounds(ym)
    snap = get_month_snapshot(conn, ym)
    if snap:
        return tiffin_rows_from_month_data(snap["data"])
    today = today_date()
    effective_end = min(end, today)
    rows = []
    for c in conn.execute("SELECT * FROM customers WHERE status IN ('Active','Paused') ORDER BY name").fetchall():
        plan = plan_overlay(conn, c, ym)
        meals = customer_meals(plan)
        if not meals:
            continue
        rows.append(_tiffin_row(conn, c["id"], c["name"], plan, meals, start, end, effective_end, ym))
    rows.sort(key=lambda r: -r["delivered"])
    return rows


def _tiffin_row(conn, cid, name, plan, meals, start, end, effective_end, ym):
    base = {"id": cid, "name": name, "meals": meals, "delivered": 0, "skipped": 0, "leave": 0,
            "total_tracked_days": 0, "daily_breakdown": [], "not_joined_yet": False}
    if effective_end < start:
        return base
    exceptions = [dict(r) for r in conn.execute(
        "SELECT date, meal, status FROM daily_log WHERE customer_id=? AND date BETWEEN ? AND ?",
        (cid, start.isoformat(), end.isoformat())).fetchall()]
    t = tiffin_counts(meals, plan.get("join_date"), start, end, exceptions,
                      get_pause_ranges(conn, cid), cap_end=effective_end)
    if t["days"] == 0:
        base.update({"not_joined_yet": True, "join_date": plan.get("join_date")})
        return base
    base.update({"delivered": t["delivered"], "skipped": t["skipped"], "leave": t["leave"],
                 "total_tracked_days": t["days"], "daily_breakdown": t["breakdown"]})
    return base


def tiffin_rows_from_month_data(data):
    """Rebuild the monthly tracker rows of a closed month from its saved snapshot."""
    start, end = _parse_iso(data["start"]), _parse_iso(data["end"])
    rows = []
    for r in data["customers"]:
        t_in = r["tracker"]
        if not t_in["meals"]:
            continue
        t = tiffin_counts(t_in["meals"], t_in.get("join_date"), start, end, t_in["exceptions"],
                          [tuple(p) for p in t_in["pauses"]])
        base = {"id": r["id"], "name": r["name"], "meals": t_in["meals"], "delivered": 0, "skipped": 0,
                "leave": 0, "total_tracked_days": 0, "daily_breakdown": [], "not_joined_yet": False}
        if t["days"] == 0:
            base.update({"not_joined_yet": True, "join_date": t_in.get("join_date")})
        else:
            base.update({"delivered": t["delivered"], "skipped": t["skipped"], "leave": t["leave"],
                         "total_tracked_days": t["days"], "daily_breakdown": t["breakdown"]})
        rows.append(base)
    rows.sort(key=lambda r: -r["delivered"])
    return rows


# ---------------------------------------------------------------------------
# Invoice snapshots — the permanent record of how each invoice was calculated
# ---------------------------------------------------------------------------
def create_invoice_snapshot(conn, invoice_id, invoice_number, c, calc, extra_rows,
                            subscription_amount, extra_amount, total, generated_date,
                            source="generated"):
    pg_name = None
    if c["pg_id"]:
        pg = conn.execute("SELECT name FROM pgs WHERE id=?", (c["pg_id"],)).fetchone()
        pg_name = pg["name"] if pg else None
    projected_from = None
    if source == "generated" and calc["period_end"] > generated_date:
        projected_from = (_parse_iso(generated_date) + timedelta(days=1)).isoformat()
    payload = {
        "schema": 1,
        "calc": calc,
        "customer": {"name": c["name"], "category": c["category"], "area": c["area"],
                     "contact": c["contact"], "pg_name": pg_name},
        "extras": [{"id": e["id"], "date": e["date"], "description": e["description"],
                    "amount": e["amount"]} for e in extra_rows],
        "generated_date": generated_date,
        "projected_from": projected_from,
        "reconstructed": source == "migrated",
    }
    r = calc["rate"]
    conn.execute(
        """INSERT INTO invoice_snapshots
           (invoice_id, invoice_number, customer_id, ym, period_start, period_end, rate_monthly,
            rate_effective_from, rate_mode, daily_rate, days_in_period, active_days, leave_days,
            billable_days, meals_per_day, scheduled_meals, delivered_meals, skipped_meals, leave_meals,
            calculated_amount, extra_amount, final_subscription_amount, final_total, status, source,
            snapshot_json, created_at)
           VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,'active',?,?,?)""",
        (invoice_id, invoice_number, c["id"], calc["ym"], calc["period_start"], calc["period_end"],
         r["monthly"], r["effective_from"], r["mode"], calc["daily_rate"], calc["days_in_period"],
         calc["active_days"], calc["leave_days"], calc["billable_days"], calc["meals_per_day"],
         calc["scheduled_meals"], calc["delivered_meals"], calc["skipped_meals"], calc["leave_meals"],
         calc["calculated_amount"], extra_amount, subscription_amount, total, source,
         json.dumps(payload), _now_iso()))


def get_invoice_snapshot(conn, invoice_id):
    row = conn.execute(
        "SELECT * FROM invoice_snapshots WHERE invoice_id=? AND status='active' ORDER BY id DESC LIMIT 1",
        (invoice_id,)).fetchone()
    if not row:
        return None
    d = dict(row)
    d["data"] = json.loads(d["snapshot_json"])
    return d


def backfill_invoice_snapshots(conn):
    """Invoices created before snapshots existed get one now, rebuilt from the records
    and rate history as they stand — the invoice's own stored amounts stay authoritative."""
    for inv in conn.execute(
        """SELECT i.* FROM invoices i WHERE NOT EXISTS
           (SELECT 1 FROM invoice_snapshots s WHERE s.invoice_id=i.id AND s.status='active')"""
    ).fetchall():
        c = conn.execute("SELECT * FROM customers WHERE id=?", (inv["customer_id"],)).fetchone()
        if not c:
            continue
        calc = compute_billing_calc(conn, c, inv["period_start"][:7])
        extras = conn.execute("SELECT * FROM extra_orders WHERE invoice_id=?", (inv["id"],)).fetchall()
        create_invoice_snapshot(conn, inv["id"], inv["invoice_number"], c, calc, extras,
                                inv["subscription_amount"], inv["extra_amount"], inv["total_amount"],
                                inv["generated_date"], source="migrated")


def void_invoice_snapshot(conn, invoice_id):
    conn.execute(
        "UPDATE invoice_snapshots SET status='voided', voided_at=? WHERE invoice_id=? AND status='active'",
        (_now_iso(), invoice_id))


def invoice_summary(snap):
    """Display-ready monthly summary for an invoice, built only from its snapshot."""
    calc = snap["data"]["calc"]
    rate = calc["rate"]
    meal_names = [MEAL_LABELS[m] for m in calc["meals"]]
    if rate["mode"] == "per_meal":
        rate_lines = [(MEAL_LABELS[m], rate[m]) for m in MEALS if rate[m] > 0]
    else:
        rate_lines = [("Flat monthly rate", rate["flat"])]
    exceptions = [e for e in calc["inputs"]["exceptions"] if e["status"] in ("Skipped", "Leave")]
    return {
        "ym": calc["ym"], "label": month_label(calc["ym"]),
        "period_start": calc["period_start"], "period_end": calc["period_end"],
        "days_in_period": calc["days_in_period"], "active_days": calc["active_days"],
        "active_start": calc["active_start"],
        "leave_days": calc["leave_days"], "billable_days": calc["billable_days"],
        "plan_text": " + ".join(meal_names) if meal_names else "No meals selected",
        "meals_per_day": calc["meals_per_day"],
        "scheduled": calc["scheduled_meals"], "delivered": calc["delivered_meals"],
        "skipped": calc["skipped_meals"], "leave_meals": calc["leave_meals"],
        "rate_monthly": rate["monthly"], "rate_mode": rate["mode"], "rate_lines": rate_lines,
        "rate_effective_from": rate["effective_from"],
        "rate_effective_label": month_label(rate["effective_from"]) if rate["effective_from"] else None,
        "daily_rate": calc["daily_rate"],
        "calculated_amount": calc["calculated_amount"],
        "final_subscription": snap["final_subscription_amount"],
        "adjusted": abs(snap["final_subscription_amount"] - calc["calculated_amount"]) > 0.005,
        "extras": snap["data"]["extras"], "extra_amount": snap["extra_amount"],
        "final_total": snap["final_total"],
        "projected_from": snap["data"].get("projected_from"),
        "generated_date": snap["data"].get("generated_date"),
        "generated_label": _fmt_pdf_date(snap["data"].get("generated_date")),
        "projected_label": _fmt_pdf_date(snap["data"].get("projected_from")) if snap["data"].get("projected_from") else None,
        "active_start_label": _fmt_pdf_date(calc["active_start"]),
        "reconstructed": snap["data"].get("reconstructed", False),
        "exceptions": exceptions,
        "customer": snap["data"]["customer"],
    }


def invoice_records_drifted(conn, snap):
    """True if tiffin records for this invoice's month were changed after it was generated."""
    calc = snap["data"]["calc"]
    now_rows = conn.execute(
        "SELECT date, meal, status FROM daily_log WHERE customer_id=? AND date BETWEEN ? AND ? "
        "AND date>=? ORDER BY date, meal",
        (snap["customer_id"], calc["period_start"], calc["period_end"], calc["active_start"])).fetchall()
    now = [(r["date"], r["meal"], r["status"]) for r in now_rows]
    then = [(e["date"], e["meal"], e["status"]) for e in calc["inputs"]["exceptions"]]
    return now != then


# ---------------------------------------------------------------------------
# Month snapshots — the permanent monthly record, and the lock that protects it
# ---------------------------------------------------------------------------
def get_month_snapshot(conn, ym):
    row = conn.execute(
        "SELECT * FROM month_snapshots WHERE ym=? AND status='active'", (ym,)).fetchone()
    if not row:
        return None
    d = dict(row)
    d["data"] = json.loads(d["snapshot_json"])
    return d


def first_data_month(conn):
    """Earliest month with any real activity (entries or invoices)."""
    parts = [f"SELECT MIN(substr(date,1,7)) AS m FROM {t}" for t in LOCKED_DATE_TABLES]
    parts.append("SELECT MIN(substr(period_start,1,7)) AS m FROM invoices")
    row = conn.execute("SELECT MIN(m) FROM (" + " UNION ALL ".join(parts) + ")").fetchone()
    return row[0] if row and row[0] else None


def _sum(conn, sql, params=()):
    return conn.execute(sql, params).fetchone()[0] or 0


def build_month_data(conn, ym):
    """Everything about one month, as plain JSON-able data: totals, per-customer calculations
    (rate, tiffins taken/skipped, final amount), and every individual entry."""
    start, end, ym = month_bounds(ym)
    s, e = start.isoformat(), end.isoformat()
    rng = (s, e)

    billed = _sum(conn, "SELECT SUM(total_amount) FROM invoices WHERE period_start=?", (s,))
    received = _sum(conn, "SELECT SUM(amount) FROM revenue WHERE date BETWEEN ? AND ?", rng)
    extra_paid = _sum(conn, "SELECT SUM(amount) FROM extra_orders WHERE payment_status='Paid' AND date BETWEEN ? AND ?", rng)
    business_expense = _sum(conn, "SELECT SUM(amount) FROM expenses WHERE expense_type='Business' AND date BETWEEN ? AND ?", rng)
    personal_expense = _sum(conn, "SELECT SUM(amount) FROM expenses WHERE expense_type='Personal' AND date BETWEEN ? AND ?", rng)
    savings_deposits = _sum(conn, "SELECT SUM(amount) FROM savings WHERE entry_type='Deposit' AND date BETWEEN ? AND ?", rng)
    personal_income = _sum(conn, "SELECT SUM(amount) FROM personal_income WHERE date BETWEEN ? AND ?", rng)
    revenue = received + extra_paid
    profit = revenue - business_expense
    kpis = {
        "billed": billed, "payments_received": received, "extra_income": extra_paid,
        "revenue": revenue, "business_expense": business_expense, "profit": profit,
        "margin": (profit / revenue * 100) if revenue else 0,
        "personal_expense": personal_expense, "savings_deposits": savings_deposits,
        "personal_income": personal_income,
        "remaining_funds": revenue - business_expense - personal_expense - savings_deposits,
    }

    cum_rev = (_sum(conn, "SELECT SUM(amount) FROM revenue WHERE date<=?", (e,))
               + _sum(conn, "SELECT SUM(amount) FROM extra_orders WHERE payment_status='Paid' AND date<=?", (e,)))
    cum_exp = _sum(conn, "SELECT SUM(amount) FROM expenses WHERE expense_type='Business' AND date<=?", (e,))
    cum_pers = _sum(conn, "SELECT SUM(amount) FROM expenses WHERE expense_type='Personal' AND date<=?", (e,))
    cum_sav = _sum(conn, "SELECT SUM(amount) FROM savings WHERE entry_type='Deposit' AND date<=?", (e,))
    cumulative = {
        "revenue": cum_rev, "business_expense": cum_exp, "profit": cum_rev - cum_exp,
        "personal_expense": cum_pers, "savings_deposits": cum_sav,
        "remaining_funds": cum_rev - cum_exp - cum_pers - cum_sav,
    }

    by_cat = [{"category": r["category"], "total": r["total"]} for r in conn.execute(
        """SELECT category, SUM(amount) AS total FROM expenses
           WHERE expense_type='Business' AND date BETWEEN ? AND ? GROUP BY category ORDER BY total DESC""",
        rng).fetchall()]

    # --- per-customer calculations -----------------------------------------
    customers = []
    for c in conn.execute(
        """SELECT c.*, p.name AS pg_name FROM customers c LEFT JOIN pgs p ON p.id=c.pg_id ORDER BY c.name"""
    ).fetchall():
        snap = conn.execute(
            "SELECT * FROM invoice_snapshots WHERE customer_id=? AND ym=? AND status='active'",
            (c["id"], ym)).fetchone()
        joined = not c["join_date"] or c["join_date"] <= e
        if not snap and (not joined or c["status"] == "Inactive"):
            continue
        live = compute_billing_calc(conn, c, ym)
        row = {
            "id": c["id"], "name": c["name"], "category": c["category"], "pg_name": c["pg_name"],
            "invoiced": bool(snap), "invoice_id": None, "invoice_number": None,
            "invoice_total": 0, "invoice_paid_at_close": 0, "records_changed": False,
            "calc": live, "calc_live": None,
            "received": _sum(conn, "SELECT SUM(amount) FROM revenue WHERE customer_id=? AND date BETWEEN ? AND ?",
                             (c["id"], s, e)),
        }
        if snap:
            sd = json.loads(snap["snapshot_json"])
            row.update({"invoice_id": snap["invoice_id"], "invoice_number": snap["invoice_number"],
                        "invoice_total": snap["final_total"],
                        "invoice_paid_at_close": invoice_paid_amount(conn, snap["invoice_id"]),
                        "calc": sd["calc"], "calc_live": live,
                        "records_changed": sd["calc"]["inputs"]["exceptions"] != live["inputs"]["exceptions"]})
        if not row["invoiced"] and not live["scheduled_meals"] and not live["calculated_amount"] \
                and not row["received"]:
            continue
        live_calc = row["calc_live"] or row["calc"]
        row["tracker"] = {"meals": live_calc["meals"], "join_date": live_calc["join_date"],
                          "exceptions": live_calc["inputs"]["exceptions"], "pauses": live_calc["inputs"]["pauses"]}
        customers.append(row)

    # --- individual entries ---------------------------------------------------
    def rows(sql, params=rng):
        return [dict(r) for r in conn.execute(sql, params).fetchall()]

    entries = {
        "payments": rows("""SELECT r.date, c.name AS customer, r.amount, r.payment_mode, r.notes, i.invoice_number
                            FROM revenue r JOIN customers c ON c.id=r.customer_id
                            LEFT JOIN invoices i ON i.id=r.invoice_id
                            WHERE r.date BETWEEN ? AND ? ORDER BY r.date, r.id"""),
        "extra_orders": rows("""SELECT e.date, COALESCE(c.name, e.guest_name, 'Walk-in') AS who, e.description,
                                       e.amount, e.payment_status, e.payment_mode, e.invoice_id
                                FROM extra_orders e LEFT JOIN customers c ON c.id=e.customer_id
                                WHERE e.date BETWEEN ? AND ? ORDER BY e.date, e.id"""),
        "business_expenses": rows("""SELECT date, category, description, amount, payment_mode FROM expenses
                                     WHERE expense_type='Business' AND date BETWEEN ? AND ? ORDER BY date, id"""),
        "personal_expenses": rows("""SELECT date, category, description, amount, payment_mode FROM expenses
                                     WHERE expense_type='Personal' AND date BETWEEN ? AND ? ORDER BY date, id"""),
        "savings": rows("""SELECT date, entry_type, amount, description, payment_mode FROM savings
                           WHERE date BETWEEN ? AND ? ORDER BY date, id"""),
        "personal_income": rows("""SELECT date, source, amount, description, payment_mode FROM personal_income
                                   WHERE date BETWEEN ? AND ? ORDER BY date, id"""),
        "invoices": rows("""SELECT s.invoice_number, c.name AS customer, s.rate_monthly, s.delivered_meals,
                                   s.skipped_meals, s.scheduled_meals, s.final_subscription_amount,
                                   s.extra_amount, s.final_total
                            FROM invoice_snapshots s JOIN customers c ON c.id=s.customer_id
                            WHERE s.ym=? AND s.status='active' ORDER BY c.name""", (ym,)),
        "tiffin_records": rows("""SELECT d.date, c.name AS customer, d.meal, d.status FROM daily_log d
                                  JOIN customers c ON c.id=d.customer_id
                                  WHERE d.date BETWEEN ? AND ? ORDER BY d.date, c.name, d.meal"""),
    }

    # --- rate / plan changes taking effect this month -------------------------
    rate_changes = []
    for r in conn.execute(
        """SELECT r.*, c.name AS customer FROM customer_rates r JOIN customers c ON c.id=r.customer_id
           WHERE r.effective_month=? ORDER BY c.name""", (ym,)).fetchall():
        prev = conn.execute(
            """SELECT * FROM customer_rates WHERE customer_id=? AND effective_month<? 
               ORDER BY effective_month DESC LIMIT 1""", (r["customer_id"], ym)).fetchone()
        if prev:
            rate_changes.append({
                "customer": r["customer"], "from_rate": effective_monthly_rate(prev),
                "to_rate": effective_monthly_rate(r),
                "plan_changed": [prev[k] for k in ("meal_breakfast", "meal_lunch", "meal_dinner")] !=
                                [r[k] for k in ("meal_breakfast", "meal_lunch", "meal_dinner")],
                "to_plan": plan_label(r) if customer_meals(r) else "No meals"})

    outstanding = sum(d["outstanding"] for d in get_all_dues(conn))
    active_count = sum(1 for r in customers if r["calc"]["scheduled_meals"] or r["invoiced"])
    return {
        "schema": 1, "ym": ym, "label": month_label(ym), "start": s, "end": e,
        "built_at": _now_iso(),
        "kpis": kpis, "cumulative": cumulative,
        "outstanding_dues": outstanding, "active_customers": active_count,
        "expense_by_category": by_cat,
        "cash_breakdown": get_cash_breakdown(conn, start, end),
        "customers": customers, "entries": entries, "rate_changes": rate_changes,
    }


def close_month(conn, ym, closed_by="manual"):
    """Save the permanent snapshot of a finished month and lock it."""
    start, end, ym = month_bounds(ym)
    if end >= today_date():
        raise ValueError("This month hasn't finished yet, so it can't be closed.")
    if is_month_closed(conn, ym):
        return None
    backfill_invoice_snapshots(conn)
    data = build_month_data(conn, ym)
    k = data["kpis"]
    version = _sum(conn, "SELECT MAX(version) FROM month_snapshots WHERE ym=?", (ym,)) + 1
    conn.execute(
        """INSERT INTO month_snapshots (ym, version, status, closed_by, closed_at, billed, revenue, expense,
               profit, personal_expense, savings_deposits, remaining_funds, snapshot_json)
           VALUES (?,?,'active',?,?,?,?,?,?,?,?,?,?)""",
        (ym, version, closed_by, _now_iso(), k["billed"], k["revenue"], k["business_expense"], k["profit"],
         k["personal_expense"], k["savings_deposits"], k["remaining_funds"], json.dumps(data)))
    conn.commit()
    return version


def reopen_month(conn, ym):
    """Unlock a month for corrections. The old snapshot is kept (as superseded) for the record."""
    cur = conn.execute(
        "UPDATE month_snapshots SET status='superseded', superseded_at=? WHERE ym=? AND status='active'",
        (_now_iso(), ym))
    conn.commit()
    return cur.rowcount > 0


def auto_close_elapsed_months(conn):
    """Close every finished month once its grace period has passed. A month that has ever
    been closed-and-reopened is never auto-closed again — it waits for a manual close."""
    settings = get_settings()
    if settings.get("auto_close_months", "1") != "1":
        return []
    grace = max(safe_int(settings.get("month_close_grace_days"), default=5), 0)
    first = first_data_month(conn)
    if not first:
        return []
    cutoff = today_date() - timedelta(days=grace)
    closed = []
    for ym in months_between(first, current_ym()):
        end = month_bounds(ym)[1]
        if end >= cutoff:
            break
        if conn.execute("SELECT 1 FROM month_snapshots WHERE ym=? LIMIT 1", (ym,)).fetchone():
            continue
        if close_month(conn, ym, closed_by="auto"):
            closed.append(ym)
    return closed


def history_months(conn):
    """One summary line per month from the first activity to now (newest first)."""
    first = first_data_month(conn)
    cur_ym = current_ym()
    if not first:
        first = cur_ym
    out = []
    for ym in reversed(months_between(min(first, cur_ym), cur_ym)):
        snap = conn.execute("SELECT * FROM month_snapshots WHERE ym=? AND status='active'", (ym,)).fetchone()
        if snap:
            out.append({"ym": ym, "label": month_label(ym), "state": "closed", "billed": snap["billed"],
                        "revenue": snap["revenue"], "expense": snap["expense"], "profit": snap["profit"],
                        "margin": (snap["profit"] / snap["revenue"] * 100) if snap["revenue"] else 0,
                        "closed_at": snap["closed_at"], "version": snap["version"]})
        else:
            k = build_month_data(conn, ym)["kpis"]
            reopened = conn.execute("SELECT 1 FROM month_snapshots WHERE ym=? LIMIT 1", (ym,)).fetchone()
            out.append({"ym": ym, "label": month_label(ym),
                        "state": "current" if ym == cur_ym else ("reopened" if reopened else "open"),
                        "billed": k["billed"], "revenue": k["revenue"], "expense": k["business_expense"],
                        "profit": k["profit"], "margin": k["margin"], "closed_at": None, "version": None})
    return out


# ---------------------------------------------------------------------------
# Invoice numbering & generation
# ---------------------------------------------------------------------------
def next_invoice_number(conn):
    """Next invoice number for the current year, based on the highest sequence number
    ever used — including invoices that were later deleted (their saved snapshots keep the
    number) — so a number is never reissued."""
    settings = get_settings()
    prefix = settings.get("invoice_prefix", "INV")
    year = today_date().year
    like_pattern = f"{prefix}-{year}-%"
    numbers = [r[0] for r in conn.execute(
        "SELECT invoice_number FROM invoices WHERE invoice_number LIKE ? "
        "UNION SELECT invoice_number FROM invoice_snapshots WHERE invoice_number LIKE ?",
        (like_pattern, like_pattern)).fetchall()]
    max_seq = 0
    for n in numbers:
        tail = n.rsplit("-", 1)[-1]
        if tail.isdigit():
            max_seq = max(max_seq, int(tail))
    return f"{prefix}-{year}-{max_seq + 1:04d}"


def generate_invoices_for_month(conn, ym):
    """Create invoices for a month. Each invoice is billed at the rate/plan that applied in
    *that* month and is saved together with a permanent snapshot of how it was calculated."""
    start, end, ym = month_bounds(ym)
    if is_month_closed(conn, ym):
        raise MonthLocked(ym)
    customers = conn.execute(
        "SELECT * FROM customers WHERE status IN ('Active','Paused')"
    ).fetchall()
    settings = get_settings()
    due_days = safe_int(settings.get("invoice_due_days"), default=7)
    created, skipped = [], []
    generated_date = today_date().isoformat()

    for c in customers:
        existing = conn.execute(
            "SELECT id FROM invoices WHERE customer_id=? AND period_start=? AND period_end=?",
            (c["id"], start.isoformat(), end.isoformat()),
        ).fetchone()
        if existing:
            skipped.append(c["name"])
            continue

        calc = compute_billing_calc(conn, c, ym)
        subscription_amount = calc["calculated_amount"]

        # Only extras that are still unpaid roll into the invoice — an extra already paid
        # in cash is revenue already, and billing it again would charge the customer twice.
        extra_rows = conn.execute(
            """SELECT * FROM extra_orders
               WHERE customer_id=? AND invoice_id IS NULL AND payment_status='Unpaid'
                 AND date BETWEEN ? AND ?""",
            (c["id"], start.isoformat(), end.isoformat()),
        ).fetchall()
        extra_amount = sum(r["amount"] for r in extra_rows)

        if subscription_amount <= 0 and extra_amount <= 0:
            skipped.append(c["name"])
            continue

        invoice_number = next_invoice_number(conn)
        total = round(subscription_amount + extra_amount, 2)
        due_date = (end + timedelta(days=due_days)).isoformat()
        cur = conn.execute(
            """INSERT INTO invoices (invoice_number, customer_id, period_start, period_end,
                                      subscription_amount, extra_amount, total_amount,
                                      generated_date, due_date)
               VALUES (?,?,?,?,?,?,?,?,?)""",
            (invoice_number, c["id"], start.isoformat(), end.isoformat(),
             subscription_amount, extra_amount, total, generated_date, due_date),
        )
        invoice_id = cur.lastrowid
        for r in extra_rows:
            conn.execute("UPDATE extra_orders SET invoice_id=? WHERE id=?", (invoice_id, r["id"]))
        create_invoice_snapshot(conn, invoice_id, invoice_number, c, calc, extra_rows,
                                subscription_amount, extra_amount, total, generated_date)
        created.append((c["name"], invoice_number, total))

    conn.commit()
    return created, skipped


def invoice_paid_amount(conn, invoice_id):
    return conn.execute(
        "SELECT COALESCE(SUM(amount),0) FROM revenue WHERE invoice_id=?", (invoice_id,)
    ).fetchone()[0]


def invoice_status(total, paid):
    if paid <= 0:
        return "Unpaid"
    if paid < total:
        return "Partial"
    return "Paid"


# ---------------------------------------------------------------------------
# Financial summaries — customer dues, KPIs, P&L
# ---------------------------------------------------------------------------
def get_customer_financials(conn, customer_id):
    invoices = conn.execute(
        "SELECT * FROM invoices WHERE customer_id=? ORDER BY period_start DESC", (customer_id,)
    ).fetchall()
    total_billed = sum(i["total_amount"] for i in invoices)
    total_paid_on_invoices = sum(invoice_paid_amount(conn, i["id"]) for i in invoices)

    unlinked_payments = conn.execute(
        "SELECT COALESCE(SUM(amount),0) FROM revenue WHERE customer_id=? AND invoice_id IS NULL",
        (customer_id,),
    ).fetchone()[0]

    unbilled_unpaid_extras = conn.execute(
        """SELECT COALESCE(SUM(amount),0) FROM extra_orders
           WHERE customer_id=? AND invoice_id IS NULL AND payment_status='Unpaid'""",
        (customer_id,),
    ).fetchone()[0]

    outstanding = round(
        total_billed - total_paid_on_invoices - unlinked_payments + unbilled_unpaid_extras, 2
    )
    return {
        "invoices": invoices,
        "total_billed": total_billed,
        "total_paid_on_invoices": total_paid_on_invoices,
        "unlinked_payments": unlinked_payments,
        "unbilled_unpaid_extras": unbilled_unpaid_extras,
        "outstanding": max(outstanding, 0),
        "credit_balance": abs(min(outstanding, 0)),
    }


def get_all_dues(conn):
    customers = conn.execute("SELECT * FROM customers ORDER BY name").fetchall()
    result = []
    for c in customers:
        fin = get_customer_financials(conn, c["id"])
        if fin["outstanding"] > 0 or fin["total_billed"] > 0:
            result.append({
                "id": c["id"], "name": c["name"], "category": c["category"],
                "status": c["status"], "outstanding": fin["outstanding"],
                "total_billed": fin["total_billed"],
            })
    result.sort(key=lambda r: -r["outstanding"])
    return result


def get_savings_summary(conn):
    """Savings/FD balances — principal only, never any interest calculation."""
    total_deposits = conn.execute(
        "SELECT COALESCE(SUM(amount),0) FROM savings WHERE entry_type='Deposit'"
    ).fetchone()[0]
    total_fd_transfers = conn.execute(
        "SELECT COALESCE(SUM(amount),0) FROM savings WHERE entry_type='FD Transfer'"
    ).fetchone()[0]
    # FD Transfer just reallocates money already deposited into the general savings pool
    # into an FD — it never creates new money and is never double-counted against revenue.
    pool_balance = round(total_deposits - total_fd_transfers, 2)
    return {
        "total_deposits": total_deposits,
        "fd_balance": total_fd_transfers,
        "pool_balance": max(pool_balance, 0),
    }


def get_kpis(conn, ym=None):
    start, end, ym = month_bounds(ym)

    total_billed_all = conn.execute("SELECT COALESCE(SUM(total_amount),0) FROM invoices").fetchone()[0]
    total_received_all = conn.execute("SELECT COALESCE(SUM(amount),0) FROM revenue").fetchone()[0]
    total_extra_paid_all = conn.execute(
        "SELECT COALESCE(SUM(amount),0) FROM extra_orders WHERE payment_status='Paid'"
    ).fetchone()[0]
    total_revenue = total_received_all + total_extra_paid_all

    # Business (operational) expenses only — personal expenses are tracked separately
    # and must never be counted as an operating cost of the tiffin business.
    total_expense = conn.execute(
        "SELECT COALESCE(SUM(amount),0) FROM expenses WHERE expense_type='Business'"
    ).fetchone()[0]
    total_personal_expense = conn.execute(
        "SELECT COALESCE(SUM(amount),0) FROM expenses WHERE expense_type='Personal'"
    ).fetchone()[0]

    net_profit = total_revenue - total_expense
    margin = (net_profit / total_revenue * 100) if total_revenue else 0
    active_customers = conn.execute(
        "SELECT COUNT(*) FROM customers WHERE status='Active'"
    ).fetchone()[0]

    month_received = conn.execute(
        "SELECT COALESCE(SUM(amount),0) FROM revenue WHERE date BETWEEN ? AND ?",
        (start.isoformat(), end.isoformat()),
    ).fetchone()[0]
    month_extra_paid = conn.execute(
        "SELECT COALESCE(SUM(amount),0) FROM extra_orders WHERE payment_status='Paid' AND date BETWEEN ? AND ?",
        (start.isoformat(), end.isoformat()),
    ).fetchone()[0]
    month_revenue = month_received + month_extra_paid
    month_expense = conn.execute(
        "SELECT COALESCE(SUM(amount),0) FROM expenses WHERE expense_type='Business' AND date BETWEEN ? AND ?",
        (start.isoformat(), end.isoformat()),
    ).fetchone()[0]
    month_personal_expense = conn.execute(
        "SELECT COALESCE(SUM(amount),0) FROM expenses WHERE expense_type='Personal' AND date BETWEEN ? AND ?",
        (start.isoformat(), end.isoformat()),
    ).fetchone()[0]
    month_billed = conn.execute(
        "SELECT COALESCE(SUM(total_amount),0) FROM invoices WHERE period_start=?",
        (start.isoformat(),),
    ).fetchone()[0]
    month_profit = month_revenue - month_expense
    month_margin = (month_profit / month_revenue * 100) if month_revenue else 0

    total_outstanding = sum(d["outstanding"] for d in get_all_dues(conn))

    savings = get_savings_summary(conn)
    month_savings_deposits = conn.execute(
        "SELECT COALESCE(SUM(amount),0) FROM savings WHERE entry_type='Deposit' AND date BETWEEN ? AND ?",
        (start.isoformat(), end.isoformat()),
    ).fetchone()[0]

    # Consistent financial model:
    # Revenue -> Business Expenses -> Personal Expenses -> Savings/FD Allocation -> Remaining Available Funds
    # Savings deposits reduce available cash but are never treated as an operating expense.
    remaining_funds = total_revenue - total_expense - total_personal_expense - savings["total_deposits"]
    month_remaining_funds = month_revenue - month_expense - month_personal_expense - month_savings_deposits

    return {
        "total_revenue": total_revenue, "total_expense": total_expense,
        "net_profit": net_profit, "margin": margin,
        "active_customers": active_customers,
        "month_revenue": month_revenue, "month_expense": month_expense,
        "month_profit": month_profit, "month_margin": month_margin,
        "month_billed": month_billed,
        "pending_total": total_outstanding,
        "extra_income_total": total_extra_paid_all,
        "extra_income_month": month_extra_paid,
        "total_personal_expense": total_personal_expense,
        "month_personal_expense": month_personal_expense,
        "savings_total_deposits": savings["total_deposits"],
        "savings_fd_balance": savings["fd_balance"],
        "savings_pool_balance": savings["pool_balance"],
        "month_savings_deposits": month_savings_deposits,
        "remaining_funds": remaining_funds,
        "month_remaining_funds": month_remaining_funds,
        "ym": ym,
    }


def get_monthly_summary(conn, months_back=12):
    rev_rows = conn.execute(
        "SELECT strftime('%Y-%m', date) AS ym, SUM(amount) AS total FROM revenue GROUP BY ym"
    ).fetchall()
    rev_by_month = {r["ym"]: r["total"] for r in rev_rows}

    extra_rows = conn.execute(
        """SELECT strftime('%Y-%m', date) AS ym, SUM(amount) AS total FROM extra_orders
           WHERE payment_status='Paid' GROUP BY ym"""
    ).fetchall()
    extra_by_month = {r["ym"]: r["total"] for r in extra_rows}

    exp_rows = conn.execute(
        "SELECT strftime('%Y-%m', date) AS ym, SUM(amount) AS total FROM expenses WHERE expense_type='Business' GROUP BY ym"
    ).fetchall()
    exp_by_month = {r["ym"]: r["total"] for r in exp_rows}

    all_months = sorted(set(rev_by_month) | set(exp_by_month) | set(extra_by_month))
    if not all_months:
        return []

    summary = []
    for ym in all_months:
        rev = (rev_by_month.get(ym, 0) or 0) + (extra_by_month.get(ym, 0) or 0)
        exp = exp_by_month.get(ym, 0) or 0
        profit = rev - exp
        margin = (profit / rev * 100) if rev else 0
        summary.append({
            "ym": ym, "label": month_label(ym), "revenue": rev, "expense": exp,
            "profit": profit, "margin": margin,
        })
    return summary[-months_back:]


def get_expense_by_category(conn, ym=None):
    start, end, ym = month_bounds(ym)
    rows = conn.execute(
        """SELECT category, SUM(amount) AS total FROM expenses
           WHERE expense_type='Business' AND date BETWEEN ? AND ? GROUP BY category ORDER BY total DESC""",
        (start.isoformat(), end.isoformat()),
    ).fetchall()
    data = [{"category": r["category"], "total": r["total"]} for r in rows]
    grand_total = sum(d["total"] for d in data)
    return data, grand_total, ym


def get_cash_breakdown(conn, start=None, end=None):
    """Real cash/UPI/bank movement — includes business expenses, personal expenses,
    and savings deposits as outflows, since all three genuinely move money out of
    the same physical cash/UPI/bank balances. Business-only figures are reported
    separately elsewhere (P&L, expense-by-category)."""
    where = "1=1"
    params = []
    if start and end:
        where = "date BETWEEN ? AND ?"
        params = [start.isoformat(), end.isoformat()]
    rev_rows = conn.execute(
        f"SELECT payment_mode, COALESCE(SUM(amount),0) t FROM revenue WHERE {where} GROUP BY payment_mode",
        params,
    ).fetchall()
    extra_rows = conn.execute(
        f"SELECT payment_mode, COALESCE(SUM(amount),0) t FROM extra_orders WHERE payment_status='Paid' AND {where} GROUP BY payment_mode",
        params,
    ).fetchall()
    exp_rows = conn.execute(
        f"SELECT payment_mode, COALESCE(SUM(amount),0) t FROM expenses WHERE {where} GROUP BY payment_mode",
        params,
    ).fetchall()
    savings_rows = conn.execute(
        f"SELECT payment_mode, COALESCE(SUM(amount),0) t FROM savings WHERE entry_type='Deposit' AND {where} GROUP BY payment_mode",
        params,
    ).fetchall()

    modes = {}
    for r in rev_rows:
        modes.setdefault(r["payment_mode"] or "Other", {"in": 0, "out": 0})
        modes[r["payment_mode"] or "Other"]["in"] += r["t"]
    for r in extra_rows:
        modes.setdefault(r["payment_mode"] or "Other", {"in": 0, "out": 0})
        modes[r["payment_mode"] or "Other"]["in"] += r["t"]
    for r in exp_rows:
        modes.setdefault(r["payment_mode"] or "Other", {"in": 0, "out": 0})
        modes[r["payment_mode"] or "Other"]["out"] += r["t"]
    for r in savings_rows:
        modes.setdefault(r["payment_mode"] or "Other", {"in": 0, "out": 0})
        modes[r["payment_mode"] or "Other"]["out"] += r["t"]

    result = []
    for mode, v in modes.items():
        result.append({"mode": mode, "in": v["in"], "out": v["out"], "net": v["in"] - v["out"]})
    result.sort(key=lambda r: r["mode"])
    return result


_OPENING_BALANCE_KEYS = {
    "Cash": "opening_balance_cash",
    "UPI": "opening_balance_upi",
    "Bank Transfer": "opening_balance_bank_transfer",
    "Other": "opening_balance_other",
}


def get_account_balances(conn):
    """Real running balance per payment mode: opening balance + everything that has
    ever moved through it (business revenue, extra orders, personal income in;
    business expenses, personal expenses, savings deposits out). This answers
    'how much money do I actually have right now', not just monthly flow."""
    settings = get_settings()
    flow = {m["mode"]: m for m in get_cash_breakdown(conn)}  # all-time, no date filter

    income_rows = conn.execute(
        "SELECT payment_mode, COALESCE(SUM(amount),0) t FROM personal_income GROUP BY payment_mode"
    ).fetchall()
    income_by_mode = {r["payment_mode"] or "Other": r["t"] for r in income_rows}

    result = []
    for mode, setting_key in _OPENING_BALANCE_KEYS.items():
        opening = safe_float(settings.get(setting_key), default=0.0)
        in_amt = flow.get(mode, {"in": 0})["in"] + income_by_mode.get(mode, 0)
        out_amt = flow.get(mode, {"out": 0})["out"]
        balance = round(opening + in_amt - out_amt, 2)
        result.append({
            "mode": mode, "opening": opening, "in": in_amt, "out": out_amt, "balance": balance,
        })
    return result


def get_alerts(conn):
    """Automated advisor-style flags computed purely from existing data — no manual
    entry required. Surfaced on the Dashboard so nothing important gets missed."""
    alerts = []
    today = today_date().isoformat()
    cur = get_settings().get("currency_symbol", "₹")

    overdue = conn.execute(
        "SELECT i.id, i.invoice_number, i.due_date, i.total_amount, c.name FROM invoices i "
        "JOIN customers c ON c.id=i.customer_id WHERE i.due_date IS NOT NULL AND i.due_date < ?",
        (today,),
    ).fetchall()
    overdue_unpaid = []
    for inv in overdue:
        paid = invoice_paid_amount(conn, inv["id"])
        if paid < inv["total_amount"]:
            overdue_unpaid.append(inv)
    if overdue_unpaid:
        total_overdue = sum(i["total_amount"] - invoice_paid_amount(conn, i["id"]) for i in overdue_unpaid)
        alerts.append({
            "level": "warning",
            "message": f"{len(overdue_unpaid)} invoice(s) are overdue, totalling {cur}{total_overdue:,.0f} unpaid.",
            "link": "/dues", "link_label": "View dues",
        })

    dues = get_all_dues(conn)
    top_due = sorted(dues, key=lambda d: -d["outstanding"])[:1]
    if top_due and top_due[0]["outstanding"] > 0:
        d = top_due[0]
        alerts.append({
            "level": "info",
            "message": f"{d['name']} has the highest outstanding balance: {cur}{d['outstanding']:,.0f}.",
            "link": f"/customers/{d['id']}", "link_label": "View customer",
        })

    todayd = today_date()
    recurring = conn.execute("SELECT * FROM recurring_expenses WHERE active=1").fetchall()
    for r in recurring:
        days_until = r["day_of_month"] - todayd.day
        if 0 <= days_until <= 3:
            alerts.append({
                "level": "info",
                "message": f"Recurring expense \"{r['category']}\" ({cur}{r['amount']:,.0f}) is due in "
                           f"{days_until} day(s).",
                "link": "/expenses", "link_label": "View expenses",
            })

    kpis = get_kpis(conn)
    if kpis["remaining_funds"] < 0:
        alerts.append({
            "level": "warning",
            "message": f"Remaining available funds are negative ({cur}{kpis['remaining_funds']:,.0f}) — "
                       f"expenses, personal spending and savings together exceed revenue.",
            "link": "/reports", "link_label": "View reports",
        })

    balances = get_account_balances(conn)
    for b in balances:
        if b["balance"] < 0:
            alerts.append({
                "level": "warning",
                "message": f"{b['mode']} balance is negative ({cur}{b['balance']:,.0f}) — "
                           f"check your opening balance in Settings or review recent transactions.",
                "link": "/accounts", "link_label": "View accounts",
            })

    return alerts


# ---------------------------------------------------------------------------
# PDF invoice generation
# ---------------------------------------------------------------------------
def _fmt_pdf_date(iso):
    try:
        return _parse_iso(iso).strftime("%d %b %Y")
    except (ValueError, TypeError):
        return iso or "-"


# ---- Devanagari (Hindi) text in PDFs ---------------------------------------
# ReportLab cannot shape Indic scripts (vowel signs and conjuncts come out wrong), so any run
# of Hindi text is rendered to a small transparent image with a real shaping engine (Pillow + RAQM)
# and placed inline. English text stays real, selectable PDF text.
_DEVA_RE = re.compile(r"[^\s]*[\u0900-\u097F][^\s]*(?:\s+[^\s]*[\u0900-\u097F][^\s]*)*")
_pdf_local = threading.local()
_DEVA_FONTS = {False: os.path.join(BASE_DIR, "static", "fonts", "NotoSansDevanagari-Regular.ttf"),
               True: os.path.join(BASE_DIR, "static", "fonts", "NotoSansDevanagari-Bold.ttf")}


def _deva_image(text, size, bold, color):
    from PIL import Image, ImageDraw, ImageFont, features
    if not features.check("raqm") or not os.path.exists(_DEVA_FONTS[bold]):
        return None
    scale, pad = 6, 2
    font = ImageFont.truetype(_DEVA_FONTS[bold], int(size * scale), layout_engine=ImageFont.Layout.RAQM)
    l, t, r, b = font.getbbox(text, anchor="ls")
    asc, desc = -t, max(b, 0)
    w, h = int(r - min(l, 0)) + 2 * pad, int(asc + desc) + 2 * pad
    im = Image.new("RGBA", (w, h), (255, 255, 255, 0))
    ImageDraw.Draw(im).text((pad - min(l, 0), asc + pad), text, font=font, fill=color, anchor="ls")
    d = getattr(_pdf_local, "dir", None) or tempfile.mkdtemp(prefix="sr_pdf_")
    _pdf_local.dir = d
    path = os.path.join(d, f"t{len(os.listdir(d))}.png")
    im.save(path)
    return path, w / scale, h / scale, (desc + pad) / scale


def pdf_text(text, size=10, bold=False, color="#1C2420"):
    """XML-escaped text for a ReportLab Paragraph; Hindi runs become inline shaped images."""
    t = "" if text is None else str(text)
    if not _DEVA_RE.search(t):
        return xml_escape(t)
    out, last = [], 0
    for m in _DEVA_RE.finditer(t):
        out.append(xml_escape(t[last:m.start()]))
        img = _deva_image(m.group(0), size, bold, color)
        out.append(f'<img src="{img[0]}" width="{img[1]:.2f}" height="{img[2]:.2f}" valign="-{img[3]:.2f}"/>'
                   if img else xml_escape(m.group(0)))
        last = m.end()
    out.append(xml_escape(t[last:]))
    return "".join(out)


def build_invoice_pdf(conn, invoice_id):
    _pdf_local.dir = None
    try:
        return _build_invoice_pdf(conn, invoice_id)
    finally:
        d = getattr(_pdf_local, "dir", None)
        if d:
            shutil.rmtree(d, ignore_errors=True)
        _pdf_local.dir = None


def _build_invoice_pdf(conn, invoice_id):
    inv = conn.execute("SELECT * FROM invoices WHERE id=?", (invoice_id,)).fetchone()
    if not inv:
        return None
    c = conn.execute("SELECT * FROM customers WHERE id=?", (inv["customer_id"],)).fetchone()
    snap = get_invoice_snapshot(conn, invoice_id)
    if snap is None:
        backfill_invoice_snapshots(conn)
        conn.commit()
        snap = get_invoice_snapshot(conn, invoice_id)
    summ = invoice_summary(snap)
    cust = summ["customer"]  # name/area as they were when the invoice was issued
    extras = conn.execute("SELECT * FROM extra_orders WHERE invoice_id=?", (invoice_id,)).fetchall()
    paid = invoice_paid_amount(conn, invoice_id)
    status = invoice_status(inv["total_amount"], paid)
    settings = get_settings()
    cur = settings.get("currency_symbol", "₹")
    esc = xml_escape

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, topMargin=12 * mm, bottomMargin=10 * mm,
                             leftMargin=18 * mm, rightMargin=18 * mm)
    styles = getSampleStyleSheet()
    green = colors.HexColor("#1E3226")
    marigold = colors.HexColor("#D9932E")
    ink_soft = colors.HexColor("#6B6152")
    cream = colors.HexColor("#F8F3E6")
    rule = colors.HexColor("#DED0AF")

    title_style = ParagraphStyle("title", parent=styles["Title"], textColor=green, fontSize=18, spaceAfter=0, fontName=_FONT_BOLD, alignment=TA_LEFT, leading=22)
    biz_style = ParagraphStyle("biz", parent=styles["Normal"], fontSize=9, textColor=ink_soft, leading=13, fontName=_FONT_REGULAR)
    label_style = ParagraphStyle("label", parent=styles["Normal"], fontSize=9, textColor=ink_soft, fontName=_FONT_REGULAR)
    normal_style = ParagraphStyle("normal", parent=styles["Normal"], fontSize=10, leading=14, fontName=_FONT_REGULAR)
    right_style = ParagraphStyle("right", parent=normal_style, alignment=TA_RIGHT, fontName=_FONT_BOLD)
    cell_style = ParagraphStyle("cell", parent=normal_style, fontSize=9.5, leading=13)
    cell_bold = ParagraphStyle("cellb", parent=cell_style, fontName=_FONT_BOLD)
    cell_right = ParagraphStyle("cellr", parent=cell_style, alignment=TA_RIGHT)
    cell_right_bold = ParagraphStyle("cellrb", parent=cell_right, fontName=_FONT_BOLD)
    section_style = ParagraphStyle("section", parent=normal_style, fontName=_FONT_BOLD, fontSize=11,
                                   textColor=green, spaceBefore=4, spaceAfter=5)
    note_style = ParagraphStyle("note", parent=styles["Normal"], fontSize=8.5, leading=12,
                                textColor=ink_soft, fontName=_FONT_REGULAR)
    stat_num = ParagraphStyle("statnum", parent=normal_style, fontName=_FONT_BOLD, fontSize=15,
                              alignment=TA_CENTER, textColor=green, leading=18)
    stat_lbl = ParagraphStyle("statlbl", parent=label_style, alignment=TA_CENTER, fontSize=8.5, leading=11)

    elements = []
    name_par = Paragraph(f"<b>{pdf_text(settings.get('business_name','Shudh Rasoi'), 18, True, '#14532d')}</b>", title_style)
    logo_img = brand_flowable(conn, "logo", 40, 15)
    if logo_img is not None:
        left_cell = Table([[logo_img, name_par]], colWidths=[logo_img.drawWidth + 4 * mm, 114 * mm - logo_img.drawWidth])
        left_cell.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("LEFTPADDING", (0, 0), (-1, -1), 0),
                                        ("RIGHTPADDING", (0, 0), (-1, -1), 0), ("TOPPADDING", (0, 0), (-1, -1), 0),
                                        ("BOTTOMPADDING", (0, 0), (-1, -1), 0)]))
    else:
        left_cell = name_par
    header_data = [[
        left_cell,
        Paragraph(f"<b>INVOICE</b><br/>{esc(inv['invoice_number'])}", right_style),
    ]]
    header_table = Table(header_data, colWidths=[118 * mm, 52 * mm])
    header_table.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP")]))
    elements.append(header_table)

    biz_lines = []
    if settings.get("business_address"):
        biz_lines.append(pdf_text(settings["business_address"], 9, False, "#6B6152"))
    if settings.get("business_phone"):
        biz_lines.append("Phone: " + esc(settings["business_phone"]))
    if settings.get("gst_number"):
        biz_lines.append("GSTIN: " + esc(settings["gst_number"]))
    if biz_lines:
        elements.append(Paragraph("<br/>".join(biz_lines), biz_style))
    elements.append(Spacer(1, 8))

    bt = _resolve_bill_to(_invoice_row(conn, invoice_id), cust)   # invoice-level edits win
    bill_to_lines = [f"<b>{pdf_text(bt['name'] or c['name'], 10, True)}</b>"]
    if bt.get("pg_name"):
        bill_to_lines.append(pdf_text(bt["pg_name"]))
    if bt.get("area"):
        bill_to_lines.append(pdf_text(bt["area"]))
    if bt.get("contact"):
        bill_to_lines.append("Contact: " + esc(bt["contact"]))

    meta_data = [
        [Paragraph("<b>Bill To</b>", label_style), Paragraph("<b>Invoice Details</b>", label_style)],
        [Paragraph("<br/>".join(bill_to_lines), normal_style),
         Paragraph(
             f"Period: {_fmt_pdf_date(inv['period_start'])} to {_fmt_pdf_date(inv['period_end'])}<br/>"
             f"Generated: {_fmt_pdf_date(inv['generated_date'])}<br/>"
             f"Due date: {_fmt_pdf_date(inv['due_date']) if inv['due_date'] else '-'}<br/>"
             f"Status: <b>{status}</b>",
             normal_style)],
    ]
    meta_table = Table(meta_data, colWidths=[85 * mm, 85 * mm])
    meta_table.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 6),
    ]))
    elements.append(meta_table)
    elements.append(Spacer(1, 8))

    # ---- Monthly summary: how this bill was worked out -------------------------------
    elements.append(Paragraph(f"Monthly summary — {summ['label']}", section_style))

    stats = Table([
        [Paragraph(f"{summ['scheduled']}", stat_num), Paragraph(f"{summ['delivered']}", stat_num),
         Paragraph(f"{summ['skipped']}", stat_num), Paragraph(f"{summ['leave_meals']}", stat_num)],
        [Paragraph("Total meals<br/>in plan", stat_lbl), Paragraph("Tiffins<br/>taken", stat_lbl),
         Paragraph("Tiffins<br/>skipped", stat_lbl), Paragraph("Meals on leave<br/>(credited)", stat_lbl)],
    ], colWidths=[42.5 * mm] * 4)
    stats.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), cream),
        ("BOX", (0, 0), (-1, -1), 0.5, rule),
        ("LINEAFTER", (0, 0), (-2, -1), 0.5, rule),
        ("TOPPADDING", (0, 0), (-1, 0), 8), ("BOTTOMPADDING", (0, 1), (-1, 1), 7),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 0), ("TOPPADDING", (0, 1), (-1, 1), 0),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))
    elements.append(stats)
    elements.append(Spacer(1, 6))

    service_txt = f"{summ['active_days']} of {summ['days_in_period']} days"
    if summ["active_days"] < summ["days_in_period"]:
        service_txt += f" (started {_fmt_pdf_date(summ['active_start'])})"
    rate_unit = " per month"
    if summ["rate_mode"] == "per_meal":
        parts = " + ".join(f"{n} {cur}{v:,.2f}" for n, v in summ["rate_lines"])
        rate_txt = f"{cur}{summ['rate_monthly']:,.2f}{rate_unit} ({parts})"
    else:
        rate_txt = f"{cur}{summ['rate_monthly']:,.2f}{rate_unit} (flat monthly rate)"
    if summ["rate_effective_label"]:
        rate_txt += f"<br/><font size='8' color='#6B6152'>Rate in effect from {esc(summ['rate_effective_label'])}</font>"

    rows = [
        ["Plan", f"{esc(summ['plan_text'])} ({summ['meals_per_day']} meal{'s' if summ['meals_per_day'] != 1 else ''} per day)"],
        ["Service days", service_txt],
        ["Leave days credited", f"{summ['leave_days']} day{'s' if summ['leave_days'] != 1 else ''}"],
        ["Billable days", f"{summ['billable_days']} of {summ['days_in_period']}"],
        ["Applicable rate", rate_txt],
        ["Daily rate", f"{cur}{summ['rate_monthly']:,.2f} ÷ {summ['days_in_period']} days = {cur}{summ['daily_rate']:,.2f} per day"],
        ["Subscription", f"{cur}{summ['daily_rate']:,.2f} × {summ['billable_days']} billable days = <b>{cur}{summ['calculated_amount']:,.2f}</b>"],
    ]
    if summ["adjusted"]:
        rows.append(["Adjustment", f"Subscription amount adjusted from {cur}{summ['calculated_amount']:,.2f} "
                                   f"to <b>{cur}{summ['final_subscription']:,.2f}</b>"])
    sum_table = Table([[Paragraph(esc(k), cell_bold), Paragraph(v, cell_style)] for k, v in rows],
                      colWidths=[46 * mm, 124 * mm])
    sum_table.setStyle(TableStyle([
        ("LINEBELOW", (0, 0), (-1, -2), 0.4, rule),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    elements.append(sum_table)
    elements.append(Spacer(1, 4))

    notes = ["Skipped tiffins are not deducted from the bill. Only full leave days are credited."]
    if summ["projected_from"]:
        notes.append(f"This invoice was issued on {_fmt_pdf_date(summ['generated_date'])}, before the month ended; "
                     f"days from {_fmt_pdf_date(summ['projected_from'])} are counted as delivered.")
    if summ["reconstructed"]:
        notes.append("Tiffin counts were rebuilt from the delivery records on file.")
    elements.append(Paragraph(" ".join(notes), note_style))

    if summ["exceptions"]:
        by_status = {"Skipped": [], "Leave": []}
        for e in summ["exceptions"]:
            by_status[e["status"]].append(f"{_parse_iso(e['date']).strftime('%d %b')} ({MEAL_LABELS.get(e['meal'], e['meal'])})")
        lines = []
        for st in ("Skipped", "Leave"):
            if by_status[st]:
                lines.append(f"<b>{st}:</b> " + ", ".join(by_status[st]))
        elements.append(Spacer(1, 3))
        elements.append(Paragraph("<br/>".join(lines), note_style))
    elements.append(Spacer(1, 12))

    # ---- Charges ----------------------------------------------------------------------
    line_items = [[Paragraph("Description", ParagraphStyle("h", parent=cell_bold, textColor=colors.white)),
                   Paragraph("Amount", ParagraphStyle("h2", parent=cell_right_bold, textColor=colors.white))]]
    line_items.append([
        Paragraph(f"Tiffin subscription — {esc(summ['plan_text'])} ({_fmt_pdf_date(inv['period_start'])} to "
                  f"{_fmt_pdf_date(inv['period_end'])})", cell_style),
        Paragraph(f"{cur}{inv['subscription_amount']:,.2f}", cell_right),
    ])
    for e in extras:
        desc = e["description"] or "Extra order"
        line_items.append([Paragraph(f"{pdf_text(desc)} ({_fmt_pdf_date(e['date'])})", cell_style),
                           Paragraph(f"{cur}{e['amount']:,.2f}", cell_right)])

    for ln in _invoice_lines(conn, invoice_id):
        qty = ln["quantity"]
        qty_txt = f"{qty:g} × {cur}{ln['unit_price']:,.2f} — " if (abs(qty - 1) > 1e-9) else ""
        line_items.append([Paragraph(f"{qty_txt}{pdf_text(ln['description'])}", cell_style),
                           Paragraph(f"{'-' if ln['amount'] < 0 else ''}{cur}{abs(ln['amount']):,.2f}", cell_right)])

    items_table = Table(line_items, colWidths=[130 * mm, 40 * mm])
    items_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), green),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("LINEBELOW", (0, 0), (-1, -2), 0.5, rule),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, cream]),
    ]))
    elements.append(items_table)
    elements.append(Spacer(1, 8))

    balance_due = round(inv["total_amount"] - paid, 2)
    totals_data = [
        ["Total", f"{cur}{inv['total_amount']:,.2f}"],
        ["Paid", f"{cur}{paid:,.2f}"],
        ["Balance Due", f"{cur}{balance_due:,.2f}"],
    ]
    totals_table = Table(totals_data, colWidths=[130 * mm, 40 * mm])
    totals_table.setStyle(TableStyle([
        ("ALIGN", (1, 0), (1, -1), "RIGHT"),
        ("FONTNAME", (0, 0), (-1, -1), _FONT_REGULAR),
        ("FONTSIZE", (0, 0), (-1, -1), 10),
        ("FONTNAME", (0, 2), (-1, 2), _FONT_BOLD),
        ("TEXTCOLOR", (0, 2), (-1, 2), green),
        ("LINEABOVE", (0, 2), (-1, 2), 1, green),
        ("TOPPADDING", (0, 2), (-1, 2), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))
    elements.append(totals_table)
    elements.append(Spacer(1, 8))
    if inv["notes"]:
        elements.append(Paragraph("<b>Notes:</b> " + pdf_text(inv["notes"]), note_style))
        elements.append(Spacer(1, 8))
    qr_info = resolve_invoice_qr(conn, settings, inv["invoice_number"], balance_due)
    if qr_info is not None:
        box = 32 * mm
        sc = min(box / qr_info["w"], box / qr_info["h"])
        qr_img = RLImage(io.BytesIO(qr_info["data"]), width=qr_info["w"] * sc, height=qr_info["h"] * sc)
        acct_style = ParagraphStyle("acct", parent=normal_style, fontName=_FONT_BOLD, fontSize=8.5, leading=11,
                                    textColor=colors.HexColor("#D00000"), alignment=TA_CENTER)
        qr_col = Table([[qr_img], [Paragraph("Account Name: " + pdf_text(account_name_of(settings), 8.5, True, "#D00000"), acct_style)]],
                       colWidths=[60 * mm])
        qr_col.setStyle(TableStyle([("ALIGN", (0, 0), (-1, -1), "CENTER"), ("VALIGN", (0, 0), (-1, -1), "TOP"),
                                    ("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                                    ("TOPPADDING", (0, 0), (-1, -1), 0), ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
                                    ("TOPPADDING", (0, 1), (0, 1), 3)]))
        if balance_due > 0.005:
            pay_head = "Scan to pay"
            pay_sub = f"Amount due: <b>{cur}{balance_due:,.2f}</b>"
        else:
            pay_head = "Payment QR"
            pay_sub = "This invoice is paid in full. Thank you!"
        pay_lines = [Paragraph(pay_head, ParagraphStyle("payh", parent=normal_style, fontSize=13, textColor=green, fontName=_FONT_BOLD)),
                     Spacer(1, 3), Paragraph(pay_sub, normal_style)]
        if valid_upi_id(settings.get("upi_id") or ""):
            pay_lines += [Spacer(1, 3), Paragraph("UPI ID: <b>" + esc(settings["upi_id"].strip()) + "</b>", normal_style)]
        pay_lines += [Spacer(1, 3), Paragraph("Scan with any UPI app (GPay, PhonePe, Paytm, BHIM).", note_style)]
        instr_style = ParagraphStyle("instr", parent=normal_style, fontName=_FONT_BOLD, fontSize=9, leading=12.5,
                                     textColor=colors.HexColor("#1C2420"), alignment=TA_CENTER)
        pay_table = Table([[qr_col, pay_lines], [Paragraph(pdf_text(payment_instruction(settings), 9, True), instr_style), ""]],
                          colWidths=[66 * mm, 104 * mm])
        pay_table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), cream), ("BOX", (0, 0), (-1, -1), 0.5, rule),
            ("SPAN", (0, 1), (1, 1)),
            ("LINEABOVE", (0, 1), (1, 1), 0.5, rule),
            ("VALIGN", (0, 0), (-1, 0), "MIDDLE"), ("LEFTPADDING", (0, 0), (0, 0), 3 * mm),
            ("TOPPADDING", (0, 0), (-1, 0), 3 * mm), ("BOTTOMPADDING", (0, 0), (-1, 0), 3 * mm),
            ("TOPPADDING", (0, 1), (-1, 1), 2.5 * mm), ("BOTTOMPADDING", (0, 1), (-1, 1), 2.5 * mm),
            ("LEFTPADDING", (0, 1), (-1, 1), 5 * mm), ("RIGHTPADDING", (0, 1), (-1, 1), 5 * mm),
        ]))
        elements.append(KeepTogether(pay_table))
    elements.append(Spacer(1, 8))
    elements.append(Paragraph(
        "Thank you for your business. For any billing queries, please contact us directly.",
        ParagraphStyle("footer", parent=styles["Normal"], fontSize=8.5, textColor=ink_soft, fontName=_FONT_REGULAR),
    ))

    doc.build(elements)
    buf.seek(0)
    return buf, inv["invoice_number"]


# ---------------------------------------------------------------------------
# Routes: Dashboard
# ---------------------------------------------------------------------------
@app.route("/")
def dashboard():
    conn = get_db()
    run_recurring_expenses(conn)
    ym = request.args.get("month") or current_ym()
    kpis = get_kpis(conn, ym)
    monthly = get_monthly_summary(conn, months_back=12)
    by_category, cat_total, cat_ym = get_expense_by_category(conn, ym)
    month_options = month_selector_options()
    alerts = get_alerts(conn)
    conn.close()

    # Month-over-month trend for the "this month" figures — compares the selected
    # month against the one immediately before it, using the same monthly summary
    # already computed above (no extra queries needed).
    by_ym = {m["ym"]: m for m in monthly}
    this_month_stats = by_ym.get(ym)
    prev_month_stats = by_ym.get(prev_month_str(ym))
    trend = {}
    if this_month_stats and prev_month_stats:
        for field in ("revenue", "expense", "profit"):
            cur_val = this_month_stats[field]
            prev_val = prev_month_stats[field]
            trend[field] = {"delta": abs(cur_val - prev_val), "up": cur_val >= prev_val}

    return render_template(
        "dashboard.html", kpis=kpis, monthly=monthly, by_category=by_category,
        cat_total=cat_total, selected_month=ym, month_options=month_options,
        alerts=alerts, trend=trend,
    )


def last_n_months(n):
    """Backward-looking month list (kept for internal use where only past months make sense)."""
    result = []
    today = today_date()
    y, m = today.year, today.month
    for _ in range(n):
        ym = f"{y:04d}-{m:02d}"
        result.append({"ym": ym, "label": month_label(ym)})
        m -= 1
        if m == 0:
            m = 12
            y -= 1
    return result


def month_selector_options(back=18, forward=12):
    """Month list spanning both past and future — used for every month picker in the
    app (Dashboard, Billing, Reports, Daily/Monthly Tracker, Payments, Expenses, Extra
    Orders) so users are never blocked from navigating to a future or past period."""
    result = []
    today = today_date()
    y, m = today.year, today.month
    # walk forward first (furthest future last becomes first after reverse)
    fy, fm = y, m
    for _ in range(forward):
        fm += 1
        if fm == 13:
            fm = 1
            fy += 1
        ym = f"{fy:04d}-{fm:02d}"
        result.append({"ym": ym, "label": month_label(ym)})
    result.reverse()
    result.append({"ym": f"{y:04d}-{m:02d}", "label": month_label(f"{y:04d}-{m:02d}") + " (current)"})
    by, bm = y, m
    for _ in range(back):
        bm -= 1
        if bm == 0:
            bm = 12
            by -= 1
        ym = f"{by:04d}-{bm:02d}"
        result.append({"ym": ym, "label": month_label(ym)})
    return result


# ---------------------------------------------------------------------------
# Routes: PGs
# ---------------------------------------------------------------------------
@app.route("/pgs", methods=["GET", "POST"])
def pgs():
    conn = get_db()
    if request.method == "POST":
        name = request.form.get("name", "").strip()
        if not name:
            flash("PG name is required.", "error")
        elif conn.execute("SELECT 1 FROM pgs WHERE lower(name)=lower(?)", (name,)).fetchone():
            flash(f'A PG named "{name}" already exists.', "error")
        else:
            conn.execute(
                """INSERT INTO pgs (name, address, contact_person, contact_number, notes)
                   VALUES (?,?,?,?,?)""",
                (name, request.form.get("address", "").strip(),
                 request.form.get("contact_person", "").strip(),
                 request.form.get("contact_number", "").strip(),
                 request.form.get("notes", "").strip()),
            )
            conn.commit()
            flash(f'Added PG "{name}".', "success")
        conn.close()
        return redirect(url_for("pgs"))

    rows = conn.execute("SELECT * FROM pgs ORDER BY name").fetchall()
    result = []
    for p in rows:
        count = conn.execute(
            "SELECT COUNT(*) FROM customers WHERE pg_id=? AND status!='Inactive'", (p["id"],)
        ).fetchone()[0]
        d = dict(p)
        d["customer_count"] = count
        result.append(d)
    conn.close()
    return render_template("pgs.html", pgs=result)


@app.route("/pgs/<int:pid>/edit", methods=["GET", "POST"])
def edit_pg(pid):
    conn = get_db()
    pg = conn.execute("SELECT * FROM pgs WHERE id=?", (pid,)).fetchone()
    if not pg:
        conn.close()
        abort(404)
    if request.method == "POST":
        name = request.form.get("name", "").strip()
        if not name:
            flash("PG name is required.", "error")
        elif conn.execute("SELECT 1 FROM pgs WHERE lower(name)=lower(?) AND id!=?", (name, pid)).fetchone():
            flash(f'A PG named "{name}" already exists.', "error")
        else:
            conn.execute(
                """UPDATE pgs SET name=?, address=?, contact_person=?, contact_number=?, notes=?
                   WHERE id=?""",
                (name, request.form.get("address", "").strip(),
                 request.form.get("contact_person", "").strip(),
                 request.form.get("contact_number", "").strip(),
                 request.form.get("notes", "").strip(), pid),
            )
            conn.commit()
            flash("PG updated.", "success")
            conn.close()
            return redirect(url_for("pgs"))
    conn.close()
    return render_template("edit_pg.html", pg=pg)


@app.route("/pgs/<int:pid>/delete", methods=["POST"])
def delete_pg(pid):
    conn = get_db()
    count = conn.execute("SELECT COUNT(*) FROM customers WHERE pg_id=?", (pid,)).fetchone()[0]
    if count > 0:
        flash("Can't remove this PG — customers are still linked to it. Reassign or remove them first.", "error")
    else:
        conn.execute("DELETE FROM pgs WHERE id=?", (pid,))
        conn.commit()
        flash("PG removed.", "success")
    conn.close()
    return redirect(url_for("pgs"))


# ---------------------------------------------------------------------------
# Routes: Customers
# ---------------------------------------------------------------------------
SORT_COLUMNS = {"name": "name", "rate": "monthly_rate", "join": "join_date"}


@app.route("/customers", methods=["GET", "POST"])
def customers():
    conn = get_db()
    if request.method == "POST":
        name = request.form.get("name", "").strip()
        errors = []
        if not name:
            errors.append("Customer name is required.")
        elif conn.execute("SELECT 1 FROM customers WHERE lower(name)=lower(?)", (name,)).fetchone():
            errors.append(f'A customer named "{name}" already exists. If it is a different person, add the area to the name (for example "{name} – Sector 5").')
        if request.form.get("category", "Regular") not in CUSTOMER_CATEGORIES:
            errors.append("Please choose a valid category.")
        if request.form.get("status", "Active") not in CUSTOMER_STATUSES:
            errors.append("Please choose a valid status.")
        _pg = request.form.get("pg_id")
        if _pg and not conn.execute("SELECT 1 FROM pgs WHERE id=?", (_pg,)).fetchone():
            errors.append("The selected PG no longer exists.")
        monthly_rate, mr_err = parse_amount(request.form.get("monthly_rate") or "0", allow_zero=True, field_label="Monthly rate")
        if mr_err:
            errors.append(mr_err)
        rb, rb_err = parse_amount(request.form.get("rate_breakfast") or "0", allow_zero=True, field_label="Breakfast rate")
        if rb_err:
            errors.append(rb_err)
        rl, rl_err = parse_amount(request.form.get("rate_lunch") or "0", allow_zero=True, field_label="Lunch rate")
        if rl_err:
            errors.append(rl_err)
        rd, rd_err = parse_amount(request.form.get("rate_dinner") or "0", allow_zero=True, field_label="Dinner rate")
        if rd_err:
            errors.append(rd_err)
        meal_b = 1 if request.form.get("meal_breakfast") else 0
        meal_l = 1 if request.form.get("meal_lunch") else 0
        meal_d = 1 if request.form.get("meal_dinner") else 0
        if not (meal_b or meal_l or meal_d) and request.form.get("category") != "Extra":
            errors.append("Select at least one meal (Breakfast / Lunch / Dinner).")

        if errors:
            for e in errors:
                flash(e, "error")
        else:
            pg_id = request.form.get("pg_id") or None
            category = request.form.get("category", "Regular")
            plan_parts = []
            if meal_b: plan_parts.append("Breakfast")
            if meal_l: plan_parts.append("Lunch")
            if meal_d: plan_parts.append("Dinner")
            cur = conn.execute(
                """INSERT INTO customers
                   (name, type, category, pg_id, area, contact, plan, monthly_rate,
                    meal_breakfast, meal_lunch, meal_dinner, rate_breakfast, rate_lunch, rate_dinner,
                    join_date, status, notes)
                   VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                (name, category, category, pg_id,
                 request.form.get("area", "").strip(), request.form.get("contact", "").strip(),
                 " + ".join(plan_parts), monthly_rate, meal_b, meal_l, meal_d, rb, rl, rd,
                 request.form.get("join_date") or today_date().isoformat(),
                 request.form.get("status", "Active"), request.form.get("notes", "").strip()),
            )
            seed_rate_history(conn, cur.lastrowid)
            conn.commit()
            flash(f'Added "{name}" to your customer list.', "success")
            conn.close()
            return redirect(url_for("customers"))
        conn.close()
        return redirect(url_for("customers"))

    q = request.args.get("q", "").strip()
    category_filter = request.args.get("category", "")
    status_filter = request.args.get("status", "")
    sort = request.args.get("sort", "name")
    direction = request.args.get("dir", "asc")

    sql = "SELECT c.*, p.name AS pg_name FROM customers c LEFT JOIN pgs p ON p.id=c.pg_id WHERE 1=1"
    params = []
    if q:
        sql += " AND (c.name LIKE ? OR c.contact LIKE ? OR c.area LIKE ?)"
        params += [f"%{q}%", f"%{q}%", f"%{q}%"]
    if category_filter:
        sql += " AND c.category=?"
        params.append(category_filter)
    if status_filter:
        sql += " AND c.status=?"
        params.append(status_filter)
    col = SORT_COLUMNS.get(sort, "name")
    sql += f" ORDER BY c.{col} {'DESC' if direction=='desc' else 'ASC'}"

    customers_raw = conn.execute(sql, params).fetchall()
    month_start, month_end = month_bounds()[:2]
    rows = []
    for c in customers_raw:
        received = conn.execute(
            "SELECT COALESCE(SUM(amount),0) FROM revenue WHERE customer_id=? AND date BETWEEN ? AND ?",
            (c["id"], month_start.isoformat(), month_end.isoformat()),
        ).fetchone()[0]
        fin = get_customer_financials(conn, c["id"])
        row = dict(c)
        row["received_this_month"] = received
        row["pending"] = fin["outstanding"]
        row["plan_display"] = plan_label(c)
        rows.append(row)

    pgs_list = conn.execute("SELECT id, name FROM pgs ORDER BY name").fetchall()
    pending_total = sum(r["pending"] for r in rows)
    conn.close()
    return render_template(
        "customers.html", customers=rows, pending_total=pending_total,
        categories=CUSTOMER_CATEGORIES, statuses=CUSTOMER_STATUSES, pgs=pgs_list,
        q=q, category_filter=category_filter, status_filter=status_filter,
        sort=sort, direction=direction,
    )


@app.route("/customers/<int:cid>/edit", methods=["GET", "POST"])
def edit_customer(cid):
    conn = get_db()
    c = conn.execute("SELECT * FROM customers WHERE id=?", (cid,)).fetchone()
    if not c:
        conn.close()
        abort(404)
    if request.method == "POST":
        name = request.form.get("name", "").strip()
        errors = []
        if not name:
            errors.append("Customer name is required.")
        elif conn.execute("SELECT 1 FROM customers WHERE lower(name)=lower(?) AND id!=?", (name, cid)).fetchone():
            errors.append(f'Another customer is already named "{name}". Add the area to tell them apart.')
        if request.form.get("category", c["category"]) not in CUSTOMER_CATEGORIES:
            errors.append("Please choose a valid category.")
        _pg = request.form.get("pg_id")
        if _pg and not conn.execute("SELECT 1 FROM pgs WHERE id=?", (_pg,)).fetchone():
            errors.append("The selected PG no longer exists.")
        monthly_rate, mr_err = parse_amount(request.form.get("monthly_rate") or "0", allow_zero=True, field_label="Monthly rate")
        if mr_err:
            errors.append(mr_err)
        rb, rb_err = parse_amount(request.form.get("rate_breakfast") or "0", allow_zero=True, field_label="Breakfast rate")
        if rb_err:
            errors.append(rb_err)
        rl, rl_err = parse_amount(request.form.get("rate_lunch") or "0", allow_zero=True, field_label="Lunch rate")
        if rl_err:
            errors.append(rl_err)
        rd, rd_err = parse_amount(request.form.get("rate_dinner") or "0", allow_zero=True, field_label="Dinner rate")
        if rd_err:
            errors.append(rd_err)
        if errors:
            for e in errors:
                flash(e, "error")
            conn.close()
            return redirect(url_for("edit_customer", cid=cid))
        meal_b = 1 if request.form.get("meal_breakfast") else 0
        meal_l = 1 if request.form.get("meal_lunch") else 0
        meal_d = 1 if request.form.get("meal_dinner") else 0
        plan_parts = []
        if meal_b: plan_parts.append("Breakfast")
        if meal_l: plan_parts.append("Lunch")
        if meal_d: plan_parts.append("Dinner")
        pg_id = request.form.get("pg_id") or None
        new_join_date = request.form.get("join_date") or c["join_date"]

        # --- plan / rate change: versioned by effective month -------------------------
        new_plan = {"monthly_rate": monthly_rate, "rate_breakfast": rb, "rate_lunch": rl,
                    "rate_dinner": rd, "meal_breakfast": meal_b, "meal_lunch": meal_l,
                    "meal_dinner": meal_d}
        plan_changed = any(round(float(new_plan[k] or 0), 2) != round(float(c[k] or 0), 2) for k in PLAN_FIELDS)
        eff_month = valid_ym(request.form.get("effective_month")) or current_ym()
        if plan_changed:
            floor = earliest_editable_month(conn, cid)
            if floor and eff_month < floor:
                flash(f"This customer's rate can only be changed from {month_label(floor)} onward — "
                      f"earlier months are already invoiced or closed and stay exactly as they were.", "error")
                conn.close()
                return redirect(url_for("edit_customer", cid=cid))

        # --- join date moved later: records before it are removed, but never from closed months
        removed = 0
        if new_join_date and new_join_date != c["join_date"]:
            stale = conn.execute(
                "SELECT date FROM daily_log WHERE customer_id=? AND date < ?", (cid, new_join_date)).fetchall()
            for r in stale:
                assert_month_open(conn, r["date"])
            removed = len(stale)
            conn.execute("DELETE FROM daily_log WHERE customer_id=? AND date < ?", (cid, new_join_date))

        conn.execute(
            """UPDATE customers SET name=?, category=?, type=?, pg_id=?, area=?, contact=?,
               join_date=?, notes=? WHERE id=?""",
            (name, request.form.get("category", "Regular"), request.form.get("category", "Regular"),
             pg_id, request.form.get("area", "").strip(), request.form.get("contact", "").strip(),
             new_join_date, request.form.get("notes", "").strip(), cid),
        )
        if plan_changed:
            set_plan_from_month(conn, cid, eff_month, new_plan,
                                note=(request.form.get("rate_note") or "").strip() or None)
        sync_customer_current_plans(conn, cid)
        conn.commit()
        extra_msg = ""
        if plan_changed:
            extra_msg = (f" New plan/rate applies from {month_label(eff_month)} onward; "
                         f"earlier months keep their previous rate.")
        if removed:
            flash(f"Customer updated. Removed {removed} tiffin record(s) that were dated before the new join date.{extra_msg}", "success")
        else:
            flash(f"Customer updated.{extra_msg}", "success")
        conn.close()
        return redirect(url_for("customers"))

    pgs_list = conn.execute("SELECT id, name FROM pgs ORDER BY name").fetchall()
    cur_ym = current_ym()
    floor = earliest_editable_month(conn, cid)
    first_choice = max(shift_month(cur_ym, -6), floor) if floor else shift_month(cur_ym, -6)
    first_choice = min(first_choice, cur_ym) if not floor or floor <= cur_ym else floor
    effective_options = [(m, month_label(m)) for m in months_between(first_choice, shift_month(max(cur_ym, first_choice), 12))]
    default_effective = max(cur_ym, floor) if floor else cur_ym
    history = rate_history_rows(conn, cid)
    conn.close()
    return render_template("edit_customer.html", c=c, pgs=pgs_list,
                            categories=CUSTOMER_CATEGORIES, effective_options=effective_options,
                            default_effective=default_effective, rate_history=history,
                            rate_floor_label=month_label(floor) if floor else None)


@app.route("/customers/<int:cid>/rates/<int:rid>/delete", methods=["POST"])
def delete_scheduled_rate(cid, rid):
    """Remove a rate change that hasn't taken effect yet (it only ever applied to future months)."""
    conn = get_db()
    row = conn.execute("SELECT * FROM customer_rates WHERE id=? AND customer_id=?", (rid, cid)).fetchone()
    if not row:
        conn.close()
        abort(404)
    matching = [h for h in rate_history_rows(conn, cid) if h["id"] == rid]
    if not matching or not matching[0]["removable"]:
        flash("Only a scheduled (future) rate change can be removed — past and current rates are part of the record.", "error")
    else:
        conn.execute("DELETE FROM customer_rates WHERE id=?", (rid,))
        sync_customer_current_plans(conn, cid)
        conn.commit()
        flash(f"Scheduled change from {month_label(row['effective_month'])} removed.", "success")
    conn.close()
    return redirect(url_for("edit_customer", cid=cid))


@app.route("/customers/<int:cid>")
def customer_detail(cid):
    conn = get_db()
    c = conn.execute(
        "SELECT c.*, p.name AS pg_name FROM customers c LEFT JOIN pgs p ON p.id=c.pg_id WHERE c.id=?",
        (cid,),
    ).fetchone()
    if not c:
        conn.close()
        abort(404)
    fin = get_customer_financials(conn, cid)
    invoices_with_status = []
    for inv in fin["invoices"]:
        paid = invoice_paid_amount(conn, inv["id"])
        d = dict(inv)
        d["paid"] = paid
        d["balance"] = round(inv["total_amount"] - paid, 2)
        d["status"] = invoice_status(inv["total_amount"], paid)
        invoices_with_status.append(d)
    payments = conn.execute(
        "SELECT * FROM revenue WHERE customer_id=? ORDER BY date DESC, id DESC", (cid,)
    ).fetchall()
    extras = conn.execute(
        "SELECT * FROM extra_orders WHERE customer_id=? ORDER BY date DESC, id DESC", (cid,)
    ).fetchall()
    history = rate_history_rows(conn, cid)
    conn.close()
    return render_template(
        "customer_detail.html", c=c, fin=fin, invoices=invoices_with_status,
        payments=payments, extras=extras, plan_display=plan_label(c), rate_history=history,
    )


@app.route("/customers/<int:cid>/delete", methods=["POST"])
def delete_customer(cid):
    conn = get_db()
    row = conn.execute("SELECT COUNT(*) FROM revenue WHERE customer_id=?", (cid,)).fetchone()[0]
    inv_count = conn.execute("SELECT COUNT(*) FROM invoices WHERE customer_id=?", (cid,)).fetchone()[0]
    extra_count = conn.execute("SELECT COUNT(*) FROM extra_orders WHERE customer_id=?", (cid,)).fetchone()[0]
    daily_count = conn.execute("SELECT COUNT(*) FROM daily_log WHERE customer_id=?", (cid,)).fetchone()[0]
    if row > 0 or inv_count > 0 or extra_count > 0 or daily_count > 0:
        flash("Can't remove this customer — they already have payment, billing, extra order, or daily tracking "
              "history. Set status to Inactive instead.", "error")
    else:
        conn.execute("DELETE FROM customer_rates WHERE customer_id=?", (cid,))
        conn.execute("DELETE FROM customer_pauses WHERE customer_id=?", (cid,))
        conn.execute("DELETE FROM customers WHERE id=?", (cid,))
        conn.commit()
        flash("Customer removed.", "success")
    conn.close()
    return redirect(url_for("customers"))


@app.route("/customers/<int:cid>/status", methods=["POST"])
def set_customer_status(cid):
    conn = get_db()
    new_status = request.form.get("status")
    if new_status not in CUSTOMER_STATUSES:
        flash("Invalid status.", "error")
    else:
        pause_from = request.form.get("pause_from") or None
        pause_to = request.form.get("pause_to") or None
        if new_status == "Paused":
            pf = pause_from or today_date().isoformat()
            if pause_to and pause_to < pf:
                flash("The pause end date can't be before its start date.", "error")
                conn.close()
                return redirect(request.referrer or url_for("customers"))
            assert_month_open(conn, pf)
            start_pause(conn, cid, pf, pause_to)
            conn.execute(
                "UPDATE customers SET status=?, pause_from=?, pause_to=? WHERE id=?",
                (new_status, pf, pause_to, cid),
            )
        else:
            # The pause that just ended is kept in pause history, so the days the customer
            # was away are still credited when their bill for that month is worked out.
            end_open_pauses(conn, cid, today_date().isoformat())
            conn.execute(
                "UPDATE customers SET status=?, pause_from=NULL, pause_to=NULL WHERE id=?",
                (new_status, cid),
            )
        conn.commit()
        flash(f"Customer status set to {new_status}.", "success")
    conn.close()
    return redirect(request.referrer or url_for("customers"))


# ---------------------------------------------------------------------------
# Routes: Daily tiffin tracking
# ---------------------------------------------------------------------------
@app.route("/daily")
def daily():
    conn = get_db()
    sel_date = request.args.get("date") or today_date().isoformat()
    try:
        sel_date = _parse_iso(sel_date).isoformat()
    except ValueError:
        sel_date = today_date().isoformat()
    sel_ym = sel_date[:7]
    customers_raw = conn.execute(
        "SELECT * FROM customers WHERE status IN ('Active','Paused') ORDER BY name"
    ).fetchall()

    exceptions = conn.execute(
        "SELECT customer_id, meal, status FROM daily_log WHERE date=?", (sel_date,)
    ).fetchall()
    ex_map = {(r["customer_id"], r["meal"]): r["status"] for r in exceptions}

    rows = []
    for c in customers_raw:
        c = plan_overlay(conn, c, sel_ym)  # the meals this customer had in the selected month
        meals = customer_meals(c)
        if not meals:
            continue
        not_joined_yet = bool(c["join_date"]) and sel_date < c["join_date"]
        if not_joined_yet:
            rows.append({"id": c["id"], "name": c["name"], "meals": meals,
                         "meal_status": {}, "on_pause": False, "not_joined_yet": True,
                         "join_date": c["join_date"]})
            continue
        on_pause = any(pf <= sel_date <= (pt or "9999-12-31")
                       for pf, pt in get_pause_ranges(conn, c["id"]))
        meal_status = {}
        for m in meals:
            if on_pause:
                meal_status[m] = "Leave"
            else:
                meal_status[m] = ex_map.get((c["id"], m), "Delivered")
        rows.append({"id": c["id"], "name": c["name"], "meals": meals,
                      "meal_status": meal_status, "on_pause": on_pause, "not_joined_yet": False})

    sel_month_input = request.args.get("month") or sel_date[:7]
    _, _, sel_month = month_bounds(sel_month_input)
    monthly_rows = get_monthly_tiffin_summary(conn, sel_month)
    month_options = month_selector_options()
    month_locked = is_month_closed(conn, sel_ym)

    conn.close()
    return render_template(
        "daily.html", rows=rows, sel_date=sel_date, meal_labels=MEAL_LABELS,
        month_locked=month_locked, sel_date_month_label=month_label(sel_ym), sel_date_ym=sel_ym,
        today=today_date().isoformat(), sel_month=sel_month, sel_month_label=month_label(sel_month),
        monthly_rows=monthly_rows, month_options=month_options, current_ym=current_ym(),
    )


@app.route("/daily/toggle", methods=["POST"])
def daily_toggle():
    data = request.get_json(force=True)
    customer_id = data.get("customer_id")
    d = data.get("date")
    meal = data.get("meal")
    new_status = data.get("status")
    try:
        _parse_iso(d)
    except (ValueError, TypeError):
        return jsonify({"ok": False, "error": "invalid date"}), 400
    if new_status not in DAILY_STATUSES:
        return jsonify({"ok": False, "error": "invalid status"}), 400
    if meal not in ("breakfast", "lunch", "dinner"):
        return jsonify({"ok": False, "error": "invalid meal"}), 400
    conn = get_db()

    # Backend validation — never trust the frontend for data integrity.
    cust = conn.execute("SELECT join_date, meal_breakfast, meal_lunch, meal_dinner FROM customers WHERE id=?", (customer_id,)).fetchone()
    if not cust:
        conn.close()
        return jsonify({"ok": False, "error": "customer not found"}), 404
    if not cust["meal_" + meal]:
        conn.close()
        return jsonify({"ok": False, "error": f"This customer's plan doesn't include {meal}."}), 400
    if cust["join_date"] and d < cust["join_date"]:
        conn.close()
        return jsonify({"ok": False, "error": f"This customer joined on {cust['join_date']} — can't record tiffins before that date."}), 400
    if is_month_closed(conn, d[:7]):
        conn.close()
        return jsonify({"ok": False, "error": f"{month_label(d[:7])} is closed — its tiffin records are locked. "
                                              f"Reopen the month from Month history to correct it."}), 409

    if new_status == "Delivered":
        conn.execute("DELETE FROM daily_log WHERE customer_id=? AND date=? AND meal=?",
                     (customer_id, d, meal))
    else:
        existing = conn.execute(
            "SELECT id FROM daily_log WHERE customer_id=? AND date=? AND meal=?",
            (customer_id, d, meal),
        ).fetchone()
        if existing:
            conn.execute("UPDATE daily_log SET status=? WHERE id=?", (new_status, existing["id"]))
        else:
            conn.execute(
                "INSERT INTO daily_log (customer_id, date, meal, status) VALUES (?,?,?,?)",
                (customer_id, d, meal, new_status),
            )
    conn.commit()
    conn.close()
    return jsonify({"ok": True, "status": new_status})


# ---------------------------------------------------------------------------
# Routes: Extra orders
# ---------------------------------------------------------------------------
@app.route("/extra-orders", methods=["GET", "POST"])
def extra_orders():
    conn = get_db()
    if request.method == "POST":
        amount, amount_error = parse_amount(request.form.get("amount"))
        customer_id = request.form.get("customer_id") or None
        guest_name = request.form.get("guest_name", "").strip()
        if customer_id == "__guest__":
            customer_id = None
            if not guest_name:
                flash("Enter a name for the walk-in / guest order.", "error")
                conn.close()
                return redirect(url_for("extra_orders"))
        who_error = None
        if not customer_id and not guest_name:
            who_error = "Choose a customer, or enter the walk-in / guest name."
        elif customer_id and not conn.execute("SELECT 1 FROM customers WHERE id=?", (customer_id,)).fetchone():
            who_error = "That customer no longer exists."
        if amount_error or who_error:
            flash(amount_error or who_error, "error")
        else:
            paid_now = request.form.get("paid_now")
            conn.execute(
                """INSERT INTO extra_orders (date, customer_id, guest_name, description, amount,
                                              payment_status, payment_mode)
                   VALUES (?,?,?,?,?,?,?)""",
                (request.form.get("date") or today_date().isoformat(),
                 customer_id, guest_name if not customer_id else None,
                 request.form.get("description", "").strip(), amount,
                 "Paid" if paid_now else "Unpaid",
                 request.form.get("payment_mode", "Cash") if paid_now else None),
            )
            conn.commit()
            flash("Extra order logged.", "success")
        conn.close()
        return redirect(url_for("extra_orders"))

    month_filter = request.args.get("month", "")
    status_filter = request.args.get("status", "")
    sql = """SELECT e.*, c.name AS customer_name FROM extra_orders e
              LEFT JOIN customers c ON c.id=e.customer_id WHERE 1=1"""
    params = []
    if month_filter:
        sql += " AND strftime('%Y-%m', e.date)=?"
        params.append(month_filter)
    if status_filter:
        sql += " AND e.payment_status=?"
        params.append(status_filter)
    sql += " ORDER BY e.date DESC, e.id DESC"
    entries = conn.execute(sql, params).fetchall()
    total = sum(e["amount"] for e in entries)
    unpaid_total = sum(e["amount"] for e in entries if e["payment_status"] == "Unpaid")

    customer_list = conn.execute(
        "SELECT id, name FROM customers WHERE status='Active' ORDER BY name"
    ).fetchall()
    month_options = month_selector_options()
    conn.close()
    return render_template(
        "extra_orders.html", entries=entries, total=total, unpaid_total=unpaid_total,
        customer_list=customer_list, payment_modes=PAYMENT_MODES,
        today=today_date().isoformat(), month_filter=month_filter, status_filter=status_filter,
        month_options=month_options,
    )


@app.route("/extra-orders/<int:eid>/mark-paid", methods=["POST"])
def mark_extra_paid(eid):
    conn = get_db()
    row = conn.execute("SELECT invoice_id FROM extra_orders WHERE id=?", (eid,)).fetchone()
    if row and row["invoice_id"]:
        flash("This extra order is already on an invoice — record the payment against that invoice "
              "instead, so it isn't counted as income twice.", "error")
        conn.close()
        return redirect(url_for("extra_orders"))
    conn.execute(
        "UPDATE extra_orders SET payment_status='Paid', payment_mode=? WHERE id=?",
        (request.form.get("payment_mode", "Cash"), eid),
    )
    conn.commit()
    conn.close()
    flash("Marked as paid.", "success")
    return redirect(url_for("extra_orders"))


@app.route("/extra-orders/<int:eid>/delete", methods=["POST"])
def delete_extra_order(eid):
    conn = get_db()
    linked = conn.execute("SELECT invoice_id FROM extra_orders WHERE id=?", (eid,)).fetchone()
    if linked and linked["invoice_id"]:
        flash("Can't remove — this extra order is already part of a generated invoice.", "error")
    else:
        conn.execute("DELETE FROM extra_orders WHERE id=?", (eid,))
        conn.commit()
        flash("Extra order removed.", "success")
    conn.close()
    return redirect(url_for("extra_orders"))


# ---------------------------------------------------------------------------
# Routes: Billing & Invoices
# ---------------------------------------------------------------------------
# ---------------------------------------------------------------------------
# Invoice management — list, create, view, edit, delete, print
# ---------------------------------------------------------------------------
INVOICE_SORTS = {
    "date": ("Newest first", lambda r: (r["generated_date"], r["id"]), True),
    "oldest": ("Oldest first", lambda r: (r["generated_date"], r["id"]), False),
    "number": ("Invoice number", lambda r: r["invoice_number"], False),
    "customer": ("Customer A–Z", lambda r: (r["customer_name"] or "").lower(), False),
    "balance": ("Highest balance", lambda r: r["balance"], True),
    "total": ("Highest total", lambda r: r["total_amount"], True),
    "due": ("Due date", lambda r: (r["due_date"] or "9999-99-99", r["id"]), False),
}
INVOICES_PER_PAGE = 40


def ensure_invoice_edit_schema(conn):
    """Additive schema for the full invoice editor. Existing invoices are untouched."""
    _add_column_if_missing(conn, "invoices", "bill_to_name TEXT")
    _add_column_if_missing(conn, "invoices", "bill_to_area TEXT")
    _add_column_if_missing(conn, "invoices", "bill_to_contact TEXT")
    _add_column_if_missing(conn, "invoices", "updated_at TEXT")
    conn.execute(
        """CREATE TABLE IF NOT EXISTS invoice_lines (
               id INTEGER PRIMARY KEY AUTOINCREMENT,
               invoice_id INTEGER NOT NULL,
               description TEXT NOT NULL,
               quantity REAL NOT NULL DEFAULT 1,
               unit_price REAL NOT NULL DEFAULT 0,
               amount REAL NOT NULL DEFAULT 0,
               sort_order INTEGER NOT NULL DEFAULT 0,
               created_at TEXT,
               FOREIGN KEY (invoice_id) REFERENCES invoices(id)
           )""")
    conn.execute("CREATE INDEX IF NOT EXISTS ix_invoice_lines_invoice ON invoice_lines (invoice_id)")
    conn.execute("CREATE INDEX IF NOT EXISTS ix_revenue_invoice ON revenue (invoice_id)")
    conn.execute("CREATE INDEX IF NOT EXISTS ix_extra_orders_invoice ON extra_orders (invoice_id)")


def _invoice_row(conn, iid):
    row = conn.execute(
        "SELECT i.*, c.name AS customer_name, c.area AS customer_area, c.contact AS customer_contact "
        "FROM invoices i JOIN customers c ON c.id=i.customer_id WHERE i.id=?", (iid,)).fetchone()
    if not row:
        abort(404)
    return row


def _resolve_bill_to(inv, snap_customer=None):
    """What is printed under 'Bill to': an invoice-level override wins, then the details as
    they were when the invoice was issued, then the current customer record."""
    snap_customer = snap_customer or {}
    return {
        "name": inv["bill_to_name"] or snap_customer.get("name") or inv["customer_name"],
        "area": inv["bill_to_area"] if inv["bill_to_area"] is not None else (snap_customer.get("area") or inv["customer_area"] or ""),
        "contact": inv["bill_to_contact"] if inv["bill_to_contact"] is not None else (snap_customer.get("contact") or inv["customer_contact"] or ""),
        "pg_name": snap_customer.get("pg_name"),
    }


def _invoice_lines(conn, iid):
    return conn.execute("SELECT * FROM invoice_lines WHERE invoice_id=? ORDER BY sort_order, id", (iid,)).fetchall()


def sync_invoice_totals(conn, iid, subscription_amount=None):
    """Single place that keeps an invoice's totals consistent:
       extra_amount = linked extra orders + custom lines; total = subscription + extra_amount.
       The saved snapshot's *final* figures follow (its original calculation stays frozen)."""
    inv = conn.execute("SELECT * FROM invoices WHERE id=?", (iid,)).fetchone()
    sub = inv["subscription_amount"] if subscription_amount is None else subscription_amount
    extras = conn.execute("SELECT COALESCE(SUM(amount),0) FROM extra_orders WHERE invoice_id=?", (iid,)).fetchone()[0]
    lines = conn.execute("SELECT COALESCE(SUM(amount),0) FROM invoice_lines WHERE invoice_id=?", (iid,)).fetchone()[0]
    extra_amount = round(extras + lines, 2)
    total = round(sub + extra_amount, 2)
    conn.execute("UPDATE invoices SET subscription_amount=?, extra_amount=?, total_amount=?, updated_at=? WHERE id=?",
                 (round(sub, 2), extra_amount, total, _now_iso(), iid))
    conn.execute(
        "UPDATE invoice_snapshots SET final_subscription_amount=?, final_total=? WHERE invoice_id=? AND status='active'",
        (round(sub, 2), total, iid))
    return total


def _parse_signed(raw, label, max_abs=MAX_AMOUNT):
    if raw is None or str(raw).strip() == "":
        return None, f"{label} is required."
    try:
        v = float(raw)
    except (TypeError, ValueError):
        return None, f"{label} must be a number."
    if not math.isfinite(v) or abs(v) > max_abs:
        return None, f"{label} is out of range."
    return round(v, 2), None


def _parse_line_rows(form):
    """Custom line items (description × quantity × price). A negative price is a discount.
    Returns (rows, errors); completely blank rows are ignored."""
    ids = form.getlist("line_id")
    descs = form.getlist("line_desc")
    qtys = form.getlist("line_qty")
    prices = form.getlist("line_price")
    rows, errors = [], []
    for n in range(max(len(descs), len(qtys), len(prices))):
        d = (descs[n] if n < len(descs) else "").strip()
        q_raw = (qtys[n] if n < len(qtys) else "").strip()
        p_raw = (prices[n] if n < len(prices) else "").strip()
        lid = (ids[n] if n < len(ids) else "").strip()
        if not d and not p_raw:
            continue
        if not d:
            errors.append(f"Line {n + 1}: please add a description.")
            continue
        if len(d) > 200:
            errors.append(f"Line {n + 1}: the description is too long (200 characters maximum).")
            continue
        q, qe = _parse_signed(q_raw or "1", f"Line {n + 1} quantity", 1_000_000)
        p, pe = _parse_signed(p_raw, f"Line {n + 1} price")
        if qe or pe:
            errors.append(qe or pe)
            continue
        if q <= 0:
            errors.append(f"Line {n + 1}: the quantity must be more than zero.")
            continue
        rows.append({"id": int(lid) if lid.isdigit() else None, "description": d, "quantity": q,
                     "unit_price": p, "amount": round(q * p, 2)})
    return rows, errors


def invoice_bundle(conn, iid):
    inv = _invoice_row(conn, iid)
    snap = get_invoice_snapshot(conn, iid)
    if snap is None:
        backfill_invoice_snapshots(conn)
        conn.commit()
        snap = get_invoice_snapshot(conn, iid)
    summary = invoice_summary(snap) if snap else None
    extras = conn.execute("SELECT * FROM extra_orders WHERE invoice_id=? ORDER BY date, id", (iid,)).fetchall()
    payments = conn.execute("SELECT * FROM revenue WHERE invoice_id=? ORDER BY date DESC, id DESC", (iid,)).fetchall()
    paid = round(sum(p["amount"] for p in payments), 2)
    balance = round(inv["total_amount"] - paid, 2)
    ym = inv["period_start"][:7]
    return {
        "inv": inv, "snap": snap, "summary": summary,
        "bill_to": _resolve_bill_to(inv, summary["customer"] if summary else None),
        "extras": extras, "lines": _invoice_lines(conn, iid), "payments": payments,
        "paid": paid, "balance": balance, "status": invoice_status(inv["total_amount"], paid),
        "overdue": bool(balance > 0 and inv["due_date"] and inv["due_date"] < today_date().isoformat()),
        "ym": ym, "month_locked": is_month_closed(conn, ym), "month_name": month_label(ym),
        "drifted": invoice_records_drifted(conn, snap) if snap else False,
    }


def build_invoice_preview(conn, customer_id, ym):
    """Everything a new invoice can be pre-filled with, from data already on file."""
    c = conn.execute("SELECT * FROM customers WHERE id=?", (customer_id,)).fetchone()
    if not c:
        return None
    start, end, ym = month_bounds(ym)
    calc = compute_billing_calc(conn, c, ym)
    existing = conn.execute(
        "SELECT id, invoice_number FROM invoices WHERE customer_id=? AND period_start=? AND period_end=?",
        (customer_id, start.isoformat(), end.isoformat())).fetchone()
    extras = conn.execute(
        """SELECT id, date, description, amount FROM extra_orders
           WHERE customer_id=? AND invoice_id IS NULL AND payment_status='Unpaid' AND date<=?
           ORDER BY date, id""", (customer_id, end.isoformat())).fetchall()
    due_days = safe_int(get_settings().get("invoice_due_days"), default=7)
    return {
        "ym": ym, "month_label": month_label(ym), "period_start": start.isoformat(), "period_end": end.isoformat(),
        "customer": {"id": c["id"], "name": c["name"], "area": c["area"] or "", "contact": c["contact"] or ""},
        "plan_text": plan_label(c), "calculated_amount": round(calc["calculated_amount"], 2),
        "billable_days": calc["billable_days"], "days_in_period": calc["days_in_period"],
        "leave_days": calc["leave_days"], "daily_rate": round(calc["daily_rate"], 2),
        "monthly_rate": round(calc["rate"]["monthly"], 2),
        "extras": [{"id": e["id"], "date": e["date"], "description": e["description"] or "Extra order",
                    "amount": e["amount"], "in_month": start.isoformat() <= e["date"] <= end.isoformat()} for e in extras],
        "due_date": (end + timedelta(days=due_days)).isoformat(),
        "existing": ({"id": existing["id"], "number": existing["invoice_number"]} if existing else None),
        "month_closed": is_month_closed(conn, ym),
    }


# ---- list ------------------------------------------------------------------
@app.route("/billing", methods=["GET", "POST"])
def billing():
    conn = get_db()
    if request.method == "POST":
        ym = valid_ym(request.form.get("month")) if request.form.get("month") else current_ym()
        if not ym:
            flash("Please choose a valid month.", "error")
            return redirect(url_for("billing"))
        created, skipped = generate_invoices_for_month(conn, ym)
        if created:
            flash(f"Generated {len(created)} invoice(s) for {month_label(ym)}.", "success")
        else:
            flash(f"No new invoices to generate for {month_label(ym)} — all customers are already billed or nothing is due.", "error")
        return redirect(url_for("billing", month=ym))

    q = (request.args.get("q") or "").strip()[:80]
    month_filter = request.args.get("month", "")
    month_filter = valid_ym(month_filter) if month_filter else ""
    status_filter = request.args.get("status", "")
    if status_filter not in ("Unpaid", "Partial", "Paid", "Overdue"):
        status_filter = ""
    sort = request.args.get("sort", "date")
    if sort not in INVOICE_SORTS:
        sort = "date"
    page = max(safe_int(request.args.get("page"), default=1), 1)

    sql = """SELECT i.*, c.name AS customer_name, COALESCE(p.paid,0) AS paid
             FROM invoices i JOIN customers c ON c.id=i.customer_id
             LEFT JOIN (SELECT invoice_id, SUM(amount) AS paid FROM revenue
                        WHERE invoice_id IS NOT NULL GROUP BY invoice_id) p ON p.invoice_id=i.id
             WHERE 1=1"""
    params = []
    if month_filter:
        sql += " AND strftime('%Y-%m', i.period_start)=?"
        params.append(month_filter)
    if q:
        sql += " AND (i.invoice_number LIKE ? OR c.name LIKE ? OR COALESCE(i.bill_to_name,'') LIKE ?)"
        params += [f"%{q}%"] * 3
    today_iso = today_date().isoformat()
    rows = []
    for r in conn.execute(sql, params).fetchall():
        d = dict(r)
        d["balance"] = round(r["total_amount"] - r["paid"], 2)
        d["status"] = invoice_status(r["total_amount"], r["paid"])
        d["overdue"] = bool(d["balance"] > 0 and r["due_date"] and r["due_date"] < today_iso)
        if status_filter == "Overdue" and not d["overdue"]:
            continue
        if status_filter in ("Unpaid", "Partial", "Paid") and d["status"] != status_filter:
            continue
        rows.append(d)
    label, keyf, rev = INVOICE_SORTS[sort]
    rows.sort(key=keyf, reverse=rev)
    totals = {
        "count": len(rows),
        "billed": round(sum(r["total_amount"] for r in rows), 2),
        "collected": round(sum(min(r["paid"], r["total_amount"]) for r in rows), 2),
        "outstanding": round(sum(r["balance"] for r in rows if r["balance"] > 0), 2),
        "overdue": sum(1 for r in rows if r["overdue"]),
    }
    pages = max(math.ceil(len(rows) / INVOICES_PER_PAGE), 1)
    page = min(page, pages)
    page_rows = rows[(page - 1) * INVOICES_PER_PAGE: page * INVOICES_PER_PAGE]
    closed = {r[0] for r in conn.execute("SELECT ym FROM month_snapshots WHERE status='active'").fetchall()}
    for r in page_rows:
        r["locked"] = r["period_start"][:7] in closed
    has_any = conn.execute("SELECT 1 FROM invoices LIMIT 1").fetchone() is not None
    return render_template(
        "billing.html", invoices=page_rows, totals=totals, q=q, month_filter=month_filter,
        status_filter=status_filter, sort=sort, sorts=[(k, v[0]) for k, v in INVOICE_SORTS.items()],
        page=page, pages=pages, has_any=has_any, month_options=month_selector_options(),
        current_ym=current_ym(), pay_date=today_iso)


# ---- create ----------------------------------------------------------------
@app.route("/invoices/preview")
def invoice_preview():
    conn = get_db()
    cid = safe_int(request.args.get("customer_id"), default=0)
    ym = valid_ym(request.args.get("month") or "")
    if not cid or not ym:
        return jsonify({"ok": False, "error": "Choose a customer and a month."}), 400
    data = build_invoice_preview(conn, cid, ym)
    if not data:
        return jsonify({"ok": False, "error": "Customer not found."}), 404
    return jsonify({"ok": True, **data})


def _customers_for_picker(conn):
    return conn.execute(
        "SELECT id, name, status, area FROM customers ORDER BY CASE status WHEN 'Active' THEN 0 WHEN 'Paused' THEN 1 ELSE 2 END, name COLLATE NOCASE"
    ).fetchall()


@app.route("/invoices/new", methods=["GET", "POST"])
def invoice_new():
    conn = get_db()
    customers_list = _customers_for_picker(conn)
    errors, posted = [], None
    if request.method == "POST":
        f = request.form
        posted = f
        cid = safe_int(f.get("customer_id"), default=0)
        ym = valid_ym(f.get("month") or "")
        c = conn.execute("SELECT * FROM customers WHERE id=?", (cid,)).fetchone()
        if not c:
            errors.append("Please choose a customer.")
        if not ym:
            errors.append("Please choose the month this invoice is for.")
        if not errors and is_month_closed(conn, ym):
            errors.append(f"{month_label(ym)} is closed, so new invoices can't be added to it. "
                          "Reopen the month from Month history first if you really need to.")
        start = end = None
        if not errors:
            start, end, ym = month_bounds(ym)
            dup = conn.execute("SELECT id, invoice_number FROM invoices WHERE customer_id=? AND period_start=? AND period_end=?",
                               (cid, start.isoformat(), end.isoformat())).fetchone()
            if dup:
                errors.append(f"{c['name']} already has invoice {dup['invoice_number']} for {month_label(ym)}. "
                              "Open it to edit instead of creating a second one.")
        gen = (f.get("generated_date") or today_date().isoformat()).strip()
        due = (f.get("due_date") or "").strip()
        if not valid_iso_date(gen):
            errors.append("The invoice date isn't a valid date.")
        if due and not valid_iso_date(due):
            errors.append("The due date isn't a valid date.")
        if due and valid_iso_date(gen) and valid_iso_date(due) and due < gen:
            errors.append("The due date can't be earlier than the invoice date.")
        sub, se = parse_amount(f.get("subscription_amount") or "0", allow_zero=True, field_label="Subscription amount")
        if se:
            errors.append(se)
        line_rows, le = _parse_line_rows(f)
        errors += le
        attach = [safe_int(x) for x in f.getlist("attach_extra") if str(x).isdigit()]
        extra_rows = []
        if not errors and attach:
            qs = ",".join("?" * len(attach))
            extra_rows = conn.execute(
                f"SELECT * FROM extra_orders WHERE id IN ({qs}) AND customer_id=? AND invoice_id IS NULL", attach + [cid]).fetchall()
        if not errors:
            extra_total = round(sum(e["amount"] for e in extra_rows) + sum(r["amount"] for r in line_rows), 2)
            total = round(sub + extra_total, 2)
            if total <= 0:
                errors.append("The invoice total must be more than zero.")
            elif total < 0:
                errors.append("The invoice total can't be negative.")
        if not errors:
            calc = compute_billing_calc(conn, c, ym)
            number = next_invoice_number(conn)
            cur = conn.execute(
                """INSERT INTO invoices (invoice_number, customer_id, period_start, period_end, subscription_amount,
                       extra_amount, total_amount, generated_date, due_date, notes, bill_to_name, bill_to_area,
                       bill_to_contact, updated_at) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                (number, cid, start.isoformat(), end.isoformat(), sub, extra_total, total, gen, due or None,
                 (f.get("notes") or "").strip()[:1000] or None, (f.get("bill_to_name") or "").strip()[:120] or None,
                 (f.get("bill_to_area") or "").strip()[:200] or None,
                 (f.get("bill_to_contact") or "").strip()[:60] or None,
                 _now_iso()))
            iid = cur.lastrowid
            for e in extra_rows:
                conn.execute("UPDATE extra_orders SET invoice_id=? WHERE id=?", (iid, e["id"]))
            for n, r in enumerate(line_rows):
                conn.execute(
                    "INSERT INTO invoice_lines (invoice_id, description, quantity, unit_price, amount, sort_order, created_at) "
                    "VALUES (?,?,?,?,?,?,?)", (iid, r["description"], r["quantity"], r["unit_price"], r["amount"], n, _now_iso()))
            create_invoice_snapshot(conn, iid, number, c, calc, extra_rows, sub, extra_total, total, gen)
            conn.commit()
            flash(f"Invoice {number} created for {c['name']} — {month_label(ym)}.", "success")
            return redirect(url_for("invoice_detail", iid=iid))
    pre_cid = safe_int(request.args.get("customer_id"), default=0)
    pre_ym = valid_ym(request.args.get("month") or "") or current_ym()
    return render_template("invoice_new.html", customers=customers_list, errors=errors, posted=posted,
                           month_options=month_selector_options(), pre_cid=pre_cid, pre_ym=pre_ym,
                           today=today_date().isoformat())


# ---- view / pdf / pay -------------------------------------------------------
@app.route("/invoices/<int:iid>")
def invoice_detail(iid):
    conn = get_db()
    b = invoice_bundle(conn, iid)
    _st = get_settings()
    inv = b["inv"]
    has_qr = resolve_invoice_qr(conn, _st, inv["invoice_number"], max(b["balance"], 0)) is not None
    return render_template(
        "invoice_detail.html", **b, pay_qr_url=(url_for("invoice_qr", iid=iid) if has_qr else None),
        payment_instruction=payment_instruction(_st), account_name=account_name_of(_st),
        upi_ok=valid_upi_id(_st.get("upi_id") or ""), pay_date=today_date().isoformat(),
        plan_display=b["summary"]["plan_text"] if b["summary"] else plan_label(
            conn.execute("SELECT * FROM customers WHERE id=?", (inv["customer_id"],)).fetchone()))


@app.route("/invoices/<int:iid>/pdf")
def invoice_pdf(iid):
    conn = get_db()
    result = build_invoice_pdf(conn, iid)
    if not result:
        abort(404)
    buf, number = result
    inline = request.args.get("inline") == "1"      # inline = open in the browser's viewer (for printing)
    return send_file(buf, as_attachment=not inline, download_name=f"{number}.pdf", mimetype="application/pdf")


@app.route("/invoices/<int:iid>/pay", methods=["POST"])
def pay_invoice(iid):
    conn = get_db()
    inv = conn.execute("SELECT * FROM invoices WHERE id=?", (iid,)).fetchone()
    if not inv:
        abort(404)
    back = request.form.get("back") if (request.form.get("back") or "").startswith("/") and not (request.form.get("back") or "").startswith("//") else None
    dest = back or url_for("invoice_detail", iid=iid)
    amount, amount_error = parse_amount(request.form.get("amount"), field_label="Payment amount")
    balance_now = round(inv["total_amount"] - invoice_paid_amount(conn, iid), 2)
    if amount_error:
        flash(amount_error, "error")
    elif amount > balance_now + 0.005:
        flash(f"That's more than the balance due ({cur_symbol()}{balance_now:,.2f}). Enter the exact amount received — "
              "for an advance payment use the Payments page.", "error")
    else:
        conn.execute(
            "INSERT INTO revenue (date, customer_id, amount, payment_mode, notes, invoice_id) VALUES (?,?,?,?,?,?)",
            (request.form.get("date") or today_date().isoformat(), inv["customer_id"], amount,
             request.form.get("payment_mode", "Cash"), request.form.get("notes", "").strip(), iid))
        conn.commit()
        flash(f"Payment recorded: {cur_symbol()}{amount:,.0f} against {inv['invoice_number']}.", "success")
    return redirect(dest)


# ---- edit ------------------------------------------------------------------
def _edit_context(conn, iid, posted=None, errors=None):
    b = invoice_bundle(conn, iid)
    inv = b["inv"]
    unbilled = conn.execute(
        """SELECT id, date, description, amount FROM extra_orders
           WHERE customer_id=? AND invoice_id IS NULL AND payment_status='Unpaid' ORDER BY date, id""",
        (inv["customer_id"],)).fetchall()
    ctx = dict(b)
    ctx.update(unbilled=unbilled, errors=errors or [], posted=posted, today=today_date().isoformat(),
               calculated=(b["summary"]["calculated_amount"] if b["summary"] else None))
    return ctx


@app.route("/invoices/<int:iid>/edit", methods=["GET", "POST"])
def invoice_edit(iid):
    conn = get_db()
    inv = _invoice_row(conn, iid)
    ym = inv["period_start"][:7]
    if request.method == "GET":
        return render_template("invoice_edit.html", **_edit_context(conn, iid))

    f = request.form
    errors = []
    if is_month_closed(conn, ym):
        errors.append(f"{month_label(ym)} is closed, so this invoice is locked. Use \"Unlock {month_label(ym)}\" first.")
    gen = (f.get("generated_date") or "").strip()
    due = (f.get("due_date") or "").strip()
    if not valid_iso_date(gen):
        errors.append("The invoice date isn't a valid date.")
    if due and not valid_iso_date(due):
        errors.append("The due date isn't a valid date.")
    if due and valid_iso_date(gen) and valid_iso_date(due) and due < gen:
        errors.append("The due date can't be earlier than the invoice date.")
    sub, se = parse_amount(f.get("subscription_amount"), allow_zero=True, field_label="Subscription amount")
    if se:
        errors.append(se)
    notes = (f.get("notes") or "").strip()
    if len(notes) > 1000:
        errors.append("The notes are too long (1000 characters maximum).")
    bill_name = (f.get("bill_to_name") or "").strip()
    bill_area = (f.get("bill_to_area") or "").strip()
    bill_contact = (f.get("bill_to_contact") or "").strip()
    if not bill_name:
        errors.append("The customer name can't be empty.")
    if len(bill_name) > 120 or len(bill_area) > 200 or len(bill_contact) > 60:
        errors.append("One of the customer details is too long.")

    # existing extra orders on this invoice
    own = {r["id"]: r for r in conn.execute("SELECT * FROM extra_orders WHERE invoice_id=?", (iid,)).fetchall()}
    extra_updates, extra_removed = [], set()
    removed = {safe_int(x) for x in f.getlist("extra_remove")}
    for n, eid_raw in enumerate(f.getlist("extra_id")):
        eid = safe_int(eid_raw, default=0)
        if eid not in own:
            continue
        if eid in removed:
            extra_removed.add(eid)
            continue
        amt, ae = parse_amount((f.getlist("extra_amount") + [""] * 99)[n], field_label=f"Extra order {n + 1} amount")
        edate = (f.getlist("extra_date") + [""] * 99)[n].strip()
        edesc = (f.getlist("extra_desc") + [""] * 99)[n].strip()
        if ae:
            errors.append(ae)
        elif not valid_iso_date(edate):
            errors.append(f"Extra order {n + 1}: the date isn't valid.")
        elif len(edesc) > 200:
            errors.append(f"Extra order {n + 1}: the description is too long.")
        else:
            extra_updates.append((eid, edate, edesc, amt))
    line_rows, le = _parse_line_rows(f)
    errors += le
    keep_line_ids = {r["id"] for r in line_rows if r["id"]}
    attach = [safe_int(x) for x in f.getlist("attach_extra") if str(x).isdigit()]
    attach_rows = []
    if attach and not errors:
        qs = ",".join("?" * len(attach))
        attach_rows = conn.execute(
            f"SELECT * FROM extra_orders WHERE id IN ({qs}) AND customer_id=? AND invoice_id IS NULL", attach + [inv["customer_id"]]).fetchall()
    if not errors:
        extras_total = sum(a for (_, _, _, a) in extra_updates) + sum(r["amount"] for r in attach_rows)
        total_preview = round(sub + extras_total + sum(r["amount"] for r in line_rows), 2)
        if total_preview < 0:
            errors.append("The invoice total can't be negative. Check your discount lines.")
        elif total_preview == 0:
            errors.append("The invoice total would be zero. If this invoice isn't needed, delete it instead.")
    if errors:
        return render_template("invoice_edit.html", **_edit_context(conn, iid, posted=f, errors=errors)), 400

    # --- write (the invoices row first: a closed month aborts everything here) ---
    conn.execute(
        """UPDATE invoices SET generated_date=?, due_date=?, notes=?, bill_to_name=?, bill_to_area=?, bill_to_contact=?,
               updated_at=? WHERE id=?""",
        (gen, due or None, notes or None, bill_name, bill_area, bill_contact, _now_iso(), iid))
    for eid in extra_removed:
        conn.execute("UPDATE extra_orders SET invoice_id=NULL WHERE id=?", (eid,))
    for eid, edate, edesc, amt in extra_updates:
        conn.execute("UPDATE extra_orders SET date=?, description=?, amount=? WHERE id=?", (edate, edesc, amt, eid))
    for e in attach_rows:
        conn.execute("UPDATE extra_orders SET invoice_id=? WHERE id=?", (iid, e["id"]))
    existing_line_ids = {r["id"] for r in _invoice_lines(conn, iid)}
    for lid in existing_line_ids - keep_line_ids:
        conn.execute("DELETE FROM invoice_lines WHERE id=? AND invoice_id=?", (lid, iid))
    for n, r in enumerate(line_rows):
        if r["id"] and r["id"] in existing_line_ids:
            conn.execute("UPDATE invoice_lines SET description=?, quantity=?, unit_price=?, amount=?, sort_order=? WHERE id=?",
                         (r["description"], r["quantity"], r["unit_price"], r["amount"], n, r["id"]))
        else:
            conn.execute(
                "INSERT INTO invoice_lines (invoice_id, description, quantity, unit_price, amount, sort_order, created_at) VALUES (?,?,?,?,?,?,?)",
                (iid, r["description"], r["quantity"], r["unit_price"], r["amount"], n, _now_iso()))
    if f.get("update_customer") == "1":
        conn.execute("UPDATE customers SET name=?, area=?, contact=? WHERE id=?",
                     (bill_name, bill_area, bill_contact, inv["customer_id"]))
    total = sync_invoice_totals(conn, iid, subscription_amount=sub)
    conn.commit()
    paid = invoice_paid_amount(conn, iid)
    msg = f"Invoice {inv['invoice_number']} saved — new total {cur_symbol()}{total:,.2f}."
    if paid > total + 0.005:
        msg += f" Note: {cur_symbol()}{paid:,.0f} was already paid, which is more than the new total — the customer is in credit."
    flash(msg, "success")
    return redirect(url_for("invoice_detail", iid=iid))


@app.route("/invoices/<int:iid>/edit-amount", methods=["POST"])
def edit_invoice_amount(iid):
    """Quick amount-only edit (kept for older bookmarks/forms). The full editor is /invoices/<id>/edit."""
    conn = get_db()
    inv = _invoice_row(conn, iid)
    new_sub, err = parse_amount(request.form.get("subscription_amount"), allow_zero=True, field_label="Subscription amount")
    if err:
        flash(err, "error")
        return redirect(url_for("invoice_detail", iid=iid))
    if is_month_closed(conn, inv["period_start"][:7]):
        raise MonthLocked(inv["period_start"][:7])
    conn.execute("UPDATE invoices SET updated_at=? WHERE id=?", (_now_iso(), iid))
    total = sync_invoice_totals(conn, iid, subscription_amount=new_sub)
    conn.commit()
    flash(f"Subscription amount updated — new total {cur_symbol()}{total:,.2f}.", "success")
    return redirect(url_for("invoice_detail", iid=iid))


@app.route("/invoices/<int:iid>/unlock", methods=["POST"])
def invoice_unlock(iid):
    conn = get_db()
    inv = _invoice_row(conn, iid)
    ym = inv["period_start"][:7]
    if reopen_month(conn, ym):
        flash(f"{month_label(ym)} is unlocked for corrections. The old month snapshot is kept on file; "
              "close the month again from Month history when you're done.", "success")
    else:
        flash(f"{month_label(ym)} wasn't locked.", "error")
    nxt = request.form.get("next")
    return redirect(url_for("invoice_edit", iid=iid) if nxt == "edit" else url_for("invoice_detail", iid=iid))


# ---- delete ----------------------------------------------------------------
@app.route("/invoices/<int:iid>/delete", methods=["POST"])
def delete_invoice(iid):
    conn = get_db()
    inv = _invoice_row(conn, iid)
    ym = inv["period_start"][:7]
    if is_month_closed(conn, ym):
        flash(f"{month_label(ym)} is closed, so this invoice is locked. Unlock the month first, then delete it.", "error")
        return redirect(url_for("invoice_detail", iid=iid))
    pays = conn.execute("SELECT id, amount FROM revenue WHERE invoice_id=?", (iid,)).fetchall()
    choice = request.form.get("payments", "")
    if pays and choice not in ("keep", "delete"):
        flash("This invoice has payments recorded. Choose what should happen to them, then delete again.", "error")
        return redirect(url_for("invoice_detail", iid=iid))
    # write the invoices row first so a locked month aborts before anything else changes
    conn.execute("UPDATE invoices SET updated_at=? WHERE id=?", (_now_iso(), iid))
    if pays:
        if choice == "delete":
            conn.execute("DELETE FROM revenue WHERE invoice_id=?", (iid,))
        else:
            conn.execute("UPDATE revenue SET invoice_id=NULL WHERE invoice_id=?", (iid,))
    conn.execute("UPDATE extra_orders SET invoice_id=NULL WHERE invoice_id=?", (iid,))
    conn.execute("DELETE FROM invoice_lines WHERE invoice_id=?", (iid,))
    void_invoice_snapshot(conn, iid)       # kept on file as 'voided'; the number is never reused
    conn.execute("DELETE FROM invoices WHERE id=?", (iid,))
    conn.commit()
    total_paid = round(sum(p["amount"] for p in pays), 2)
    if pays and choice == "keep":
        flash(f"Invoice {inv['invoice_number']} deleted. Its {cur_symbol()}{total_paid:,.0f} of payments were kept as "
              f"advance credit for {inv['customer_name']}.", "success")
    elif pays:
        flash(f"Invoice {inv['invoice_number']} and its {cur_symbol()}{total_paid:,.0f} of payments were deleted.", "success")
    else:
        flash(f"Invoice {inv['invoice_number']} deleted. You can generate or create it again any time.", "success")
    return redirect(url_for("billing"))


# ---------------------------------------------------------------------------
# Help, guided tour and global search
# ---------------------------------------------------------------------------
import help_content as _help

PAGE_LINKS = [
    ("Dashboard", "dashboard", "Overview, alerts"), ("Customers", "customers", "People and plans"),
    ("PG houses", "pgs", "Paying-guest houses"), ("Daily tracking", "daily", "Skipped meals and leave"),
    ("Extra orders", "extra_orders", "One-off charges"), ("Invoices", "billing", "Generate, edit, print"),
    ("New invoice", "invoice_new", "Bill one customer"), ("Payments", "payments", "Money received"),
    ("Dues", "dues", "Who owes you"), ("Expenses", "expenses", "Spending and recurring bills"),
    ("Savings & FD", "savings", "Money set aside"), ("Personal income", "personal_income", "Outside income"),
    ("Cash & bank", "accounts", "Balances by mode"), ("Profit & loss", "reports", "Monthly P&L"),
    ("Monthly summary", "monthly", "Month by month"), ("Month history", "history", "Closed months"),
    ("Settings & backup", "settings_page", "Business details, backups"), ("Account", "account", "Username and password"),
    ("Help & user guide", "help_page", "How to use the app"),
]


@app.context_processor
def _inject_help():
    ep = request.endpoint if has_request_context() else None
    return {"page_help": _help.PAGE_HELP.get(ep), "tour_steps": _help.TOUR}


@app.route("/help")
def help_page():
    return render_template("help.html", sections=_help.GUIDE_SECTIONS, glossary=_help.GLOSSARY,
                           pages=_help.PAGE_HELP, tour_steps=_help.TOUR)


@app.route("/api/search")
def api_search():
    q = (request.args.get("q") or "").strip()[:60]
    out = {"pages": [], "customers": [], "invoices": []}
    ql = q.lower()
    out["pages"] = [{"label": l, "url": url_for(ep), "hint": h} for l, ep, h in PAGE_LINKS
                    if not ql or ql in l.lower() or ql in h.lower()][:8]
    if len(q) >= 2:
        conn = get_db()
        like = f"%{q}%"
        out["customers"] = [{"label": r["name"], "hint": " · ".join(x for x in (r["status"], r["area"]) if x),
                             "url": url_for("customer_detail", cid=r["id"])}
                            for r in conn.execute("SELECT id, name, status, area FROM customers WHERE name LIKE ? OR contact LIKE ? "
                                                  "ORDER BY name COLLATE NOCASE LIMIT 6", (like, like)).fetchall()]
        out["invoices"] = [{"label": r["invoice_number"], "hint": f"{r['name']} · {month_label(r['period_start'][:7])} · {cur_symbol()}{r['total_amount']:,.0f}",
                            "url": url_for("invoice_detail", iid=r["id"])}
                           for r in conn.execute(
                               "SELECT i.id, i.invoice_number, i.period_start, i.total_amount, c.name FROM invoices i "
                               "JOIN customers c ON c.id=i.customer_id WHERE i.invoice_number LIKE ? OR c.name LIKE ? "
                               "ORDER BY i.id DESC LIMIT 6", (like, like)).fetchall()]
    return jsonify(out)


# ---------------------------------------------------------------------------
# Routes: Payments (general payment log — was "Revenue log")
# ---------------------------------------------------------------------------
@app.route("/payments", methods=["GET", "POST"])
def payments():
    conn = get_db()
    if request.method == "POST":
        customer_id = request.form.get("customer_id")
        amount, amount_error = parse_amount(request.form.get("amount"))
        if not customer_id:
            flash("Please choose a customer.", "error")
        elif amount_error:
            flash(amount_error, "error")
        else:
            invoice_id = request.form.get("invoice_id") or None
            conn.execute(
                """INSERT INTO revenue (date, customer_id, amount, payment_mode, notes, invoice_id)
                   VALUES (?,?,?,?,?,?)""",
                (request.form.get("date") or today_date().isoformat(), int(customer_id),
                 amount, request.form.get("payment_mode", "Cash"),
                 request.form.get("notes", "").strip(), invoice_id),
            )
            conn.commit()
            flash("Payment logged.", "success")
        conn.close()
        return redirect(url_for("payments"))

    month_filter = request.args.get("month", "")
    sql = """SELECT r.id, r.date, r.amount, r.payment_mode, r.notes, r.invoice_id,
                    c.name AS customer_name, c.category AS customer_category,
                    i.invoice_number
             FROM revenue r JOIN customers c ON c.id = r.customer_id
             LEFT JOIN invoices i ON i.id = r.invoice_id WHERE 1=1"""
    params = []
    if month_filter:
        sql += " AND strftime('%Y-%m', r.date)=?"
        params.append(month_filter)
    sql += " ORDER BY r.date DESC, r.id DESC"
    entries = conn.execute(sql, params).fetchall()
    customer_list = conn.execute(
        "SELECT id, name FROM customers WHERE status!='Inactive' ORDER BY name"
    ).fetchall()
    total = sum(e["amount"] for e in entries)
    month_options = month_selector_options()
    conn.close()
    return render_template(
        "payments.html", entries=entries, customer_list=customer_list,
        total=total, payment_modes=PAYMENT_MODES, today=today_date().isoformat(),
        month_filter=month_filter, month_options=month_options,
    )


@app.route("/payments/<int:rid>/delete", methods=["POST"])
def delete_payment(rid):
    conn = get_db()
    conn.execute("DELETE FROM revenue WHERE id=?", (rid,))
    conn.commit()
    conn.close()
    flash("Entry removed.", "success")
    return redirect(url_for("payments"))


# ---------------------------------------------------------------------------
# Routes: Dues / Outstanding
# ---------------------------------------------------------------------------
@app.route("/dues")
def dues():
    conn = get_db()
    rows = get_all_dues(conn)
    total_outstanding = sum(r["outstanding"] for r in rows)
    conn.close()
    return render_template("dues.html", rows=rows, total_outstanding=total_outstanding)


# ---------------------------------------------------------------------------
# Routes: Expense log + recurring expenses
# ---------------------------------------------------------------------------
@app.route("/expenses", methods=["GET", "POST"])
def expenses():
    conn = get_db()
    run_recurring_expenses(conn)
    if request.method == "POST":
        amount, amount_error = parse_amount(request.form.get("amount"))
        expense_type = request.form.get("expense_type", "Business")
        _cat = request.form.get("category", "Miscellaneous")
        if expense_type not in EXPENSE_TYPES:
            amount_error = amount_error or "Please choose a valid expense type."
        elif _cat not in EXPENSE_CATEGORIES + PERSONAL_EXPENSE_CATEGORIES:
            amount_error = amount_error or "Please choose a category."
        if amount_error:
            flash(amount_error, "error")
        else:
            default_cat = "Miscellaneous"
            conn.execute(
                """INSERT INTO expenses (date, category, description, amount, payment_mode, expense_type)
                   VALUES (?,?,?,?,?,?)""",
                (request.form.get("date") or today_date().isoformat(),
                 request.form.get("category", default_cat),
                 request.form.get("description", "").strip(), amount,
                 request.form.get("payment_mode", "Cash"), expense_type),
            )
            conn.commit()
            flash(f"{expense_type} expense logged.", "success")
        conn.close()
        return redirect(url_for("expenses"))

    month_filter = request.args.get("month", "")

    def fetch_entries(etype):
        sql = "SELECT * FROM expenses WHERE expense_type=?"
        params = [etype]
        if month_filter:
            sql += " AND strftime('%Y-%m', date)=?"
            params.append(month_filter)
        sql += " ORDER BY date DESC, id DESC"
        return conn.execute(sql, params).fetchall()

    business_entries = fetch_entries("Business")
    personal_entries = fetch_entries("Personal")
    business_total = sum(e["amount"] for e in business_entries)
    personal_total = sum(e["amount"] for e in personal_entries)

    recurring = conn.execute("SELECT * FROM recurring_expenses ORDER BY active DESC, category").fetchall()
    month_options = month_selector_options()
    conn.close()
    return render_template(
        "expenses.html",
        business_entries=business_entries, personal_entries=personal_entries,
        business_total=business_total, personal_total=personal_total,
        categories=EXPENSE_CATEGORIES, personal_categories=PERSONAL_EXPENSE_CATEGORIES,
        payment_modes=PAYMENT_MODES, today=today_date().isoformat(), recurring=recurring,
        month_filter=month_filter, month_options=month_options,
    )


@app.route("/expenses/<int:eid>/edit", methods=["GET", "POST"])
def edit_expense(eid):
    conn = get_db()
    e = conn.execute("SELECT * FROM expenses WHERE id=?", (eid,)).fetchone()
    if not e:
        conn.close()
        abort(404)
    if request.method == "POST":
        amount, amount_error = parse_amount(request.form.get("amount"))
        if amount_error:
            flash(amount_error, "error")
            conn.close()
            return redirect(url_for("edit_expense", eid=eid))
        conn.execute(
            """UPDATE expenses SET date=?, category=?, description=?, amount=?, payment_mode=?
               WHERE id=?""",
            (request.form.get("date"), request.form.get("category"),
             request.form.get("description", "").strip(), amount,
             request.form.get("payment_mode", "Cash"), eid),
        )
        conn.commit()
        flash("Expense updated.", "success")
        conn.close()
        return redirect(url_for("expenses"))
    cats = PERSONAL_EXPENSE_CATEGORIES if e["expense_type"] == "Personal" else EXPENSE_CATEGORIES
    conn.close()
    return render_template("edit_expense.html", e=e, categories=cats, payment_modes=PAYMENT_MODES)


@app.route("/expenses/<int:eid>/delete", methods=["POST"])
def delete_expense(eid):
    conn = get_db()
    conn.execute("DELETE FROM expenses WHERE id=?", (eid,))
    conn.commit()
    conn.close()
    flash("Entry removed.", "success")
    return redirect(url_for("expenses"))


@app.route("/expenses/recurring/add", methods=["POST"])
def add_recurring_expense():
    conn = get_db()
    amount, amount_error = parse_amount(request.form.get("amount"), field_label="Recurring expense amount")
    day, day_error = parse_int(request.form.get("day_of_month"), default=1, min_value=1, max_value=28, field_label="Day of month")
    if amount_error:
        flash(amount_error, "error")
    elif day_error:
        flash(day_error, "error")
    elif request.form.get("category", "Miscellaneous") not in EXPENSE_CATEGORIES + PERSONAL_EXPENSE_CATEGORIES:
        flash("Please choose a category.", "error")
    else:
        conn.execute(
            """INSERT INTO recurring_expenses (category, description, amount, day_of_month, active, payment_mode)
               VALUES (?,?,?,?,1,?)""",
            (request.form.get("category", "Miscellaneous"),
             request.form.get("description", "").strip(), amount, day,
             request.form.get("payment_mode", "Cash")),
        )
        conn.commit()
        flash("Recurring expense added — it will auto-generate each month from the day you set.", "success")
    conn.close()
    return redirect(url_for("expenses"))


@app.route("/expenses/recurring/<int:rid>/toggle", methods=["POST"])
def toggle_recurring_expense(rid):
    conn = get_db()
    cur = conn.execute("SELECT active FROM recurring_expenses WHERE id=?", (rid,)).fetchone()
    if cur:
        conn.execute("UPDATE recurring_expenses SET active=? WHERE id=?", (0 if cur["active"] else 1, rid))
        conn.commit()
    conn.close()
    return redirect(url_for("expenses"))


@app.route("/expenses/recurring/<int:rid>/delete", methods=["POST"])
def delete_recurring_expense(rid):
    conn = get_db()
    conn.execute("DELETE FROM recurring_expenses WHERE id=?", (rid,))
    conn.commit()
    conn.close()
    flash("Recurring expense removed.", "success")
    return redirect(url_for("expenses"))


# ---------------------------------------------------------------------------
# Routes: Monthly summary
# ---------------------------------------------------------------------------
@app.route("/monthly")
def monthly():
    conn = get_db()
    summary = get_monthly_summary(conn, months_back=24)
    conn.close()
    return render_template("monthly.html", summary=list(reversed(summary)))


@app.route("/monthly/<ym>")
def monthly_detail(ym):
    conn = get_db()
    start, end, ym = month_bounds(ym)
    revenue_rows = conn.execute(
        """SELECT r.*, c.name AS customer_name FROM revenue r JOIN customers c ON c.id=r.customer_id
           WHERE r.date BETWEEN ? AND ? ORDER BY r.date""",
        (start.isoformat(), end.isoformat()),
    ).fetchall()
    extra_rows = conn.execute(
        """SELECT e.*, c.name AS customer_name FROM extra_orders e LEFT JOIN customers c ON c.id=e.customer_id
           WHERE e.date BETWEEN ? AND ? AND e.payment_status='Paid' ORDER BY e.date""",
        (start.isoformat(), end.isoformat()),
    ).fetchall()
    expense_rows = conn.execute(
        "SELECT * FROM expenses WHERE expense_type='Business' AND date BETWEEN ? AND ? ORDER BY date",
        (start.isoformat(), end.isoformat()),
    ).fetchall()
    personal_rows = conn.execute(
        "SELECT * FROM expenses WHERE expense_type='Personal' AND date BETWEEN ? AND ? ORDER BY date",
        (start.isoformat(), end.isoformat()),
    ).fetchall()
    month_closed = is_month_closed(conn, ym)
    invoice_rows = conn.execute(
        """SELECT i.*, c.name AS customer_name FROM invoices i JOIN customers c ON c.id=i.customer_id
           WHERE i.period_start=? ORDER BY c.name""",
        (start.isoformat(),),
    ).fetchall()

    total_revenue = sum(r["amount"] for r in revenue_rows) + sum(e["amount"] for e in extra_rows)
    total_expense = sum(e["amount"] for e in expense_rows)
    total_billed = sum(i["total_amount"] for i in invoice_rows)
    conn.close()
    return render_template(
        "monthly_detail.html", ym=ym, label=month_label(ym), start=start, end=end,
        revenue_rows=revenue_rows, extra_rows=extra_rows, expense_rows=expense_rows,
        personal_rows=personal_rows, month_closed=month_closed,
        invoice_rows=invoice_rows, total_revenue=total_revenue, total_expense=total_expense,
        total_billed=total_billed, net_profit=total_revenue - total_expense,
    )


# ---------------------------------------------------------------------------
# Routes: Reports (P&L + cash/UPI/bank breakdown)
# ---------------------------------------------------------------------------
@app.route("/reports")
def reports():
    conn = get_db()
    ym = request.args.get("month") or current_ym()
    start, end, ym = month_bounds(ym)

    received = conn.execute(
        "SELECT COALESCE(SUM(amount),0) FROM revenue WHERE date BETWEEN ? AND ?",
        (start.isoformat(), end.isoformat()),
    ).fetchone()[0]
    extra_paid = conn.execute(
        "SELECT COALESCE(SUM(amount),0) FROM extra_orders WHERE payment_status='Paid' AND date BETWEEN ? AND ?",
        (start.isoformat(), end.isoformat()),
    ).fetchone()[0]
    billed = conn.execute(
        "SELECT COALESCE(SUM(total_amount),0) FROM invoices WHERE period_start=?",
        (start.isoformat(),),
    ).fetchone()[0]
    by_category, expense_total, _ = get_expense_by_category(conn, ym)
    cash_breakdown = get_cash_breakdown(conn, start, end)

    personal_expense_total = conn.execute(
        "SELECT COALESCE(SUM(amount),0) FROM expenses WHERE expense_type='Personal' AND date BETWEEN ? AND ?",
        (start.isoformat(), end.isoformat()),
    ).fetchone()[0]
    savings_deposits = conn.execute(
        "SELECT COALESCE(SUM(amount),0) FROM savings WHERE entry_type='Deposit' AND date BETWEEN ? AND ?",
        (start.isoformat(), end.isoformat()),
    ).fetchone()[0]

    total_revenue = received + extra_paid
    net_profit = total_revenue - expense_total  # pure business P&L, unaffected by personal/savings
    margin = (net_profit / total_revenue * 100) if total_revenue else 0
    remaining_funds = total_revenue - expense_total - personal_expense_total - savings_deposits

    month_options = month_selector_options()
    conn.close()
    return render_template(
        "reports.html", ym=ym, label=month_label(ym), received=received, extra_paid=extra_paid,
        total_revenue=total_revenue, billed=billed, by_category=by_category,
        expense_total=expense_total, cash_breakdown=cash_breakdown, net_profit=net_profit,
        margin=margin, month_options=month_options,
        personal_expense_total=personal_expense_total, savings_deposits=savings_deposits,
        remaining_funds=remaining_funds,
    )


# ---------------------------------------------------------------------------
# Routes: Settings
# ---------------------------------------------------------------------------
SETTINGS_FIELDS = ["business_name", "business_address", "business_phone", "gst_number",
                    "currency_symbol", "invoice_prefix", "invoice_due_days",
                    "opening_balance_cash", "opening_balance_upi",
                    "opening_balance_bank_transfer", "opening_balance_other",
                    "auto_close_months", "month_close_grace_days", "upi_id", "upi_account_name"]
NUMERIC_SETTINGS_FIELDS = {
    "invoice_due_days": ("int", "Invoice due days"),
    "month_close_grace_days": ("int", "Days before a finished month is closed"),
    "opening_balance_cash": ("amount", "Cash opening balance"),
    "opening_balance_upi": ("amount", "UPI opening balance"),
    "opening_balance_bank_transfer": ("amount", "Bank Transfer opening balance"),
    "opening_balance_other": ("amount", "Other opening balance"),
}



@app.route("/invoices/<int:iid>/pay-qr.png")
def invoice_qr(iid):
    conn = get_db()
    inv = conn.execute("SELECT * FROM invoices WHERE id=?", (iid,)).fetchone()
    if not inv:
        conn.close()
        abort(404)
    bal = round(inv["total_amount"] - invoice_paid_amount(conn, iid), 2)
    info = resolve_invoice_qr(conn, get_settings(), inv["invoice_number"], bal)
    conn.close()
    if not info:
        abort(404)
    r = Response(info["data"], mimetype="image/png")
    r.headers["Cache-Control"] = "no-cache"
    return r


@app.route("/brand/<kind>")
def brand_image(kind):
    if kind not in BRAND_KINDS:
        abort(404)
    conn = get_db()
    r = get_brand_bytes(conn, kind)
    conn.close()
    if not r:
        abort(404)
    resp = Response(r["data"], mimetype=r["mime"])
    resp.headers["Cache-Control"] = "public, max-age=31536000, immutable" if request.args.get("v") else "no-cache"
    return resp


@app.route("/settings/brand/<kind>", methods=["POST"])
def upload_brand(kind):
    if kind not in BRAND_KINDS:
        abort(404)
    f = request.files.get("file")
    if not f or not f.filename:
        flash("Please choose an image file first.", "error")
        return redirect(url_for("settings_page") + "#brand")
    data, w, h, err = process_brand_image(f, kind)
    if err:
        flash(err, "error")
        return redirect(url_for("settings_page") + "#brand")
    conn = get_db()
    _ensure_brand_table(conn)
    existed = conn.execute("SELECT 1 FROM brand_assets WHERE kind=?", (kind,)).fetchone()
    conn.execute(
        "INSERT INTO brand_assets (kind, mime, data, filename, width, height, updated_at) VALUES (?,?,?,?,?,?,?) "
        "ON CONFLICT(kind) DO UPDATE SET mime=excluded.mime, data=excluded.data, filename=excluded.filename, "
        "width=excluded.width, height=excluded.height, updated_at=excluded.updated_at",
        (kind, "image/png", data, os.path.basename(f.filename)[:120], w, h, datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")))
    conn.commit()
    conn.close()
    flash(f"{BRAND_KINDS[kind]} {'replaced' if existed else 'uploaded'}.", "success")
    if kind == "qr" and abs(w - h) > 0.25 * max(w, h):
        flash("Heads up: that QR image isn't square. Make sure the whole code is visible so it scans properly.", "error")
    return redirect(url_for("settings_page") + "#brand")


@app.route("/settings/brand/<kind>/remove", methods=["POST"])
def remove_brand(kind):
    if kind not in BRAND_KINDS:
        abort(404)
    conn = get_db()
    _ensure_brand_table(conn)
    conn.execute("DELETE FROM brand_assets WHERE kind=?", (kind,))
    conn.commit()
    conn.close()
    flash(f"{BRAND_KINDS[kind]} removed.", "success")
    return redirect(url_for("settings_page") + "#brand")


@app.route("/savings", methods=["GET", "POST"])
def savings():
    conn = get_db()
    if request.method == "POST":
        amt, amount_error = parse_amount(request.form.get("amount"))
        entry_type = request.form.get("entry_type", "Deposit")
        if entry_type not in SAVINGS_TYPES:
            amount_error = amount_error or "Please choose a valid entry type."
        if amount_error:
            flash(amount_error, "error")
        else:
            if entry_type == "FD Transfer":
                pool = get_savings_summary(conn)["pool_balance"]
                if amt > pool:
                    flash(f"Can't transfer {get_settings().get('currency_symbol','₹')}{amt:,.0f} to FD — only "
                          f"{get_settings().get('currency_symbol','₹')}{pool:,.0f} is available in your savings pool.", "error")
                    conn.close()
                    return redirect(url_for("savings"))
            conn.execute(
                """INSERT INTO savings (date, amount, entry_type, description, payment_mode)
                   VALUES (?,?,?,?,?)""",
                (request.form.get("date") or today_date().isoformat(), amt, entry_type,
                 request.form.get("description", "").strip(),
                 request.form.get("payment_mode", "Cash") if entry_type == "Deposit" else None),
            )
            conn.commit()
            flash(f"{entry_type} of {get_settings().get('currency_symbol','₹')}{amt:,.0f} logged.", "success")
        conn.close()
        return redirect(url_for("savings"))

    entries = conn.execute("SELECT * FROM savings ORDER BY date DESC, id DESC").fetchall()
    summary = get_savings_summary(conn)
    conn.close()
    return render_template("savings.html", entries=entries, summary=summary,
                            savings_types=SAVINGS_TYPES, payment_modes=PAYMENT_MODES,
                            today=today_date().isoformat())


@app.route("/savings/<int:sid>/edit", methods=["GET", "POST"])
def edit_savings(sid):
    conn = get_db()
    s = conn.execute("SELECT * FROM savings WHERE id=?", (sid,)).fetchone()
    if not s:
        conn.close()
        abort(404)
    if request.method == "POST":
        amt, amount_error = parse_amount(request.form.get("amount"))
        if amount_error:
            flash(amount_error, "error")
            conn.close()
            return redirect(url_for("edit_savings", sid=sid))
        entry_type = request.form.get("entry_type", s["entry_type"])
        if entry_type == "FD Transfer":
            # Validate against the pool balance excluding this entry's own current amount
            other_deposits = conn.execute(
                "SELECT COALESCE(SUM(amount),0) FROM savings WHERE entry_type='Deposit' AND id!=?", (sid,)
            ).fetchone()[0]
            other_fd = conn.execute(
                "SELECT COALESCE(SUM(amount),0) FROM savings WHERE entry_type='FD Transfer' AND id!=?", (sid,)
            ).fetchone()[0]
            available_pool = other_deposits - other_fd
            if amt > available_pool:
                flash(f"Can't set this FD transfer to {get_settings().get('currency_symbol','₹')}{amt:,.0f} — "
                      f"only {get_settings().get('currency_symbol','₹')}{max(available_pool,0):,.0f} would be available in the savings pool.", "error")
                conn.close()
                return redirect(url_for("edit_savings", sid=sid))
        conn.execute(
            """UPDATE savings SET date=?, amount=?, entry_type=?, description=?, payment_mode=? WHERE id=?""",
            (request.form.get("date"), amt, entry_type, request.form.get("description", "").strip(),
             request.form.get("payment_mode") if entry_type == "Deposit" else None, sid),
        )
        conn.commit()
        flash("Savings entry updated.", "success")
        conn.close()
        return redirect(url_for("savings"))
    conn.close()
    return render_template("edit_savings.html", s=s, savings_types=SAVINGS_TYPES, payment_modes=PAYMENT_MODES)


@app.route("/savings/<int:sid>/delete", methods=["POST"])
def delete_savings(sid):
    conn = get_db()
    conn.execute("DELETE FROM savings WHERE id=?", (sid,))
    conn.commit()
    conn.close()
    flash("Savings entry removed.", "success")
    return redirect(url_for("savings"))


# ---------------------------------------------------------------------------
# Routes: Personal income (salary, side income — separate from tiffin business revenue)
# ---------------------------------------------------------------------------
@app.route("/personal-income", methods=["GET", "POST"])
def personal_income():
    conn = get_db()
    if request.method == "POST":
        amount, amount_error = parse_amount(request.form.get("amount"))
        source = request.form.get("source", "").strip()
        if not source:
            flash("Please enter a source.", "error")
        elif amount_error:
            flash(amount_error, "error")
        else:
            conn.execute(
                """INSERT INTO personal_income (date, source, amount, description, payment_mode)
                   VALUES (?,?,?,?,?)""",
                (request.form.get("date") or today_date().isoformat(), source,
                 amount, request.form.get("description", "").strip(),
                 request.form.get("payment_mode", "Cash")),
            )
            conn.commit()
            flash("Personal income logged.", "success")
        conn.close()
        return redirect(url_for("personal_income"))

    entries = conn.execute("SELECT * FROM personal_income ORDER BY date DESC, id DESC").fetchall()
    total = sum(e["amount"] for e in entries)
    conn.close()
    return render_template("personal_income.html", entries=entries, total=total,
                            payment_modes=PAYMENT_MODES, today=today_date().isoformat())


@app.route("/personal-income/<int:pid>/edit", methods=["GET", "POST"])
def edit_personal_income(pid):
    conn = get_db()
    e = conn.execute("SELECT * FROM personal_income WHERE id=?", (pid,)).fetchone()
    if not e:
        conn.close()
        abort(404)
    if request.method == "POST":
        amount, amount_error = parse_amount(request.form.get("amount"))
        if amount_error:
            flash(amount_error, "error")
            conn.close()
            return redirect(url_for("edit_personal_income", pid=pid))
        conn.execute(
            """UPDATE personal_income SET date=?, source=?, amount=?, description=?, payment_mode=?
               WHERE id=?""",
            (request.form.get("date"), request.form.get("source", "").strip(), amount,
             request.form.get("description", "").strip(), request.form.get("payment_mode", "Cash"), pid),
        )
        conn.commit()
        flash("Personal income entry updated.", "success")
        conn.close()
        return redirect(url_for("personal_income"))
    conn.close()
    return render_template("edit_personal_income.html", e=e, payment_modes=PAYMENT_MODES)


@app.route("/personal-income/<int:pid>/delete", methods=["POST"])
def delete_personal_income(pid):
    conn = get_db()
    conn.execute("DELETE FROM personal_income WHERE id=?", (pid,))
    conn.commit()
    conn.close()
    flash("Personal income entry removed.", "success")
    return redirect(url_for("personal_income"))


# ---------------------------------------------------------------------------
# Routes: Accounts — real running balance per payment mode
# ---------------------------------------------------------------------------
@app.route("/accounts")
def accounts():
    conn = get_db()
    balances = get_account_balances(conn)
    total_balance = sum(b["balance"] for b in balances)
    conn.close()
    return render_template("accounts.html", balances=balances, total_balance=total_balance)


@app.route("/settings", methods=["GET", "POST"])
def settings_page():
    conn = get_db()
    if request.method == "POST":
        errors = []
        clean_values = {}
        for field in SETTINGS_FIELDS:
            if field not in request.form:
                continue  # each form on the page only saves the fields it actually contains
            val = request.form.get(field, "").strip()
            if field == "auto_close_months":
                val = "1" if val == "1" else "0"
            if field == "upi_id":
                val = val.replace(" ", "")
                if val and not valid_upi_id(val):
                    errors.append("That UPI ID doesn't look right. Use the format name@bank — for a mobile number "
                                  "add your UPI handle, e.g. 9876543210@ybl.")
                    continue
            if field == "upi_account_name":
                val = " ".join(val.split())[:60] or DEFAULT_ACCOUNT_NAME

            if field in NUMERIC_SETTINGS_FIELDS:
                kind, label = NUMERIC_SETTINGS_FIELDS[field]
                if kind == "int":
                    parsed, err = parse_int(val or "0", min_value=0, field_label=label)
                else:
                    parsed, err = parse_amount(val or "0", allow_zero=True, field_label=label)
                if err:
                    errors.append(err)
                else:
                    clean_values[field] = str(parsed)
            else:
                clean_values[field] = val

        if errors:
            for e in errors:
                flash(e, "error")
            conn.close()
            return redirect(url_for("settings_page"))

        for field, val in clean_values.items():
            conn.execute(
                "INSERT INTO settings (key, value) VALUES (?,?) "
                "ON CONFLICT(key) DO UPDATE SET value=excluded.value",
                (field, val),
            )
        conn.commit()
        flash("Settings saved.", "success")
        conn.close()
        return redirect(url_for("settings_page"))

    current = get_settings()
    latest_inv = conn.execute("SELECT id FROM invoices ORDER BY id DESC LIMIT 1").fetchone()
    conn.close()
    auto_backups = sorted(glob.glob(os.path.join(BACKUP_DIR, "tiffin_auto_*.db")))
    latest_auto_backup = None
    if auto_backups:
        latest_auto_backup = os.path.basename(auto_backups[-1]).replace("tiffin_auto_", "").replace(".db", "")
    return render_template("settings.html", settings=current, account_name=account_name_of(current), upi_ok=valid_upi_id(current.get("upi_id") or ""), latest_invoice_id=(latest_inv["id"] if latest_inv else None),
                            auto_backup_count=len(auto_backups), latest_auto_backup=latest_auto_backup)


@app.route("/settings/backup")
def backup_db():
    if not os.path.exists(DB_PATH):
        abort(404)
    filename = f"shudh_rasoi_backup_{today_date().isoformat()}.db"
    fd, tmp_name = tempfile.mkstemp(suffix=".db", dir=os.path.dirname(os.path.abspath(DB_PATH)))
    os.close(fd)
    try:
        consistent_db_copy(DB_PATH, tmp_name)  # includes everything not yet merged from the write-ahead log
        with open(tmp_name, "rb") as fh:
            payload = io.BytesIO(fh.read())
    finally:
        try:
            os.remove(tmp_name)
        except OSError:
            pass
    return send_file(payload, as_attachment=True, download_name=filename,
                      mimetype="application/octet-stream")


@app.route("/settings/restore", methods=["POST"])
def restore_db():
    if not request.form.get("confirm"):
        flash("You must confirm before restoring — this replaces all current data.", "error")
        return redirect(url_for("settings_page"))
    file = request.files.get("backup_file")
    if not file or file.filename == "":
        flash("Choose a backup .db file to restore.", "error")
        return redirect(url_for("settings_page"))
    data_dir = os.path.dirname(os.path.abspath(DB_PATH))
    tmp_path = os.path.join(data_dir, "_restore_tmp.db")
    file.save(tmp_path)
    try:
        test_conn = sqlite3.connect(tmp_path)
        try:
            integrity = test_conn.execute("PRAGMA integrity_check").fetchone()[0]
            tables = [r[0] for r in test_conn.execute(
                "SELECT name FROM sqlite_master WHERE type='table'"
            ).fetchall()]
        finally:
            test_conn.close()
        if integrity != "ok":
            raise ValueError("The file appears to be damaged.")
        if "customers" not in tables:
            raise ValueError("Not a valid Shudh Rasoi backup file.")
    except Exception as e:
        try:
            os.remove(tmp_path)
        except OSError:
            pass
        flash(f"Restore failed — {e}", "error")
        return redirect(url_for("settings_page"))

    global _last_automation_check
    backup_current = os.path.join(data_dir, f"tiffin_before_restore_{datetime.now().strftime('%Y%m%d%H%M%S')}.db")
    if os.path.exists(DB_PATH):
        consistent_db_copy(DB_PATH, backup_current)   # safety copy, including un-merged WAL data
    for suffix in ("-wal", "-shm"):                    # a stale WAL must never be replayed onto the restored file
        try:
            os.remove(DB_PATH + suffix)
        except OSError:
            pass
    saved_users = []
    try:
        _uc = sqlite3.connect(DB_PATH)
        _uc.row_factory = sqlite3.Row
        saved_users = [dict(r) for r in _uc.execute("SELECT * FROM users").fetchall()]
        _uc.close()
    except sqlite3.Error:
        saved_users = []
    os.replace(tmp_path, DB_PATH)
    init_db()                                          # upgrades older backups (rate history, snapshots)
    if saved_users:                                    # keep the current sign-in account, whatever the backup held
        _rc = sqlite3.connect(DB_PATH)
        try:
            _rc.execute("DELETE FROM users")
            for u in saved_users:
                cols = list(u.keys())
                _rc.execute(f"INSERT INTO users ({', '.join(cols)}) VALUES ({', '.join('?' * len(cols))})",
                            [u[c] for c in cols])
            _rc.commit()
        finally:
            _rc.close()
    _last_automation_check = None
    flash("Data restored successfully. Your previous database was kept as a safety copy next to it.", "success")
    return redirect(url_for("dashboard"))


# ---------------------------------------------------------------------------
# Routes: Month history (permanent monthly snapshots)
# ---------------------------------------------------------------------------
@app.route("/history")
def history():
    conn = get_db()
    months = history_months(conn)
    conn.close()
    return render_template("history.html", months=months, current_ym=current_ym())


@app.route("/history/<ym>")
def history_month(ym):
    ym_ok = valid_ym(ym)
    if not ym_ok:
        abort(404)
    conn = get_db()
    snap = get_month_snapshot(conn, ym_ok)
    data = snap["data"] if snap else build_month_data(conn, ym_ok)
    versions = [dict(r) for r in conn.execute(
        "SELECT version, status, closed_by, closed_at, superseded_at FROM month_snapshots WHERE ym=? ORDER BY version DESC",
        (ym_ok,)).fetchall()]
    end = month_bounds(ym_ok)[1]
    cur_ym = current_ym()
    first_ym = first_data_month(conn) or cur_ym
    conn.close()
    return render_template(
        "history_month.html", ym=ym_ok, data=data, snap=snap, versions=versions,
        can_close=(snap is None and end < today_date()),
        prev_ym=(shift_month(ym_ok, -1) if ym_ok > first_ym and ym_ok > "2000-01" else None), next_ym=(shift_month(ym_ok, 1) if ym_ok < cur_ym else None),
        is_current=(ym_ok == cur_ym), is_future=(ym_ok > cur_ym),
    )


@app.route("/history/<ym>/close", methods=["POST"])
def history_close(ym):
    ym_ok = valid_ym(ym)
    if not ym_ok:
        abort(404)
    conn = get_db()
    try:
        version = close_month(conn, ym_ok, closed_by="manual")
        if version:
            flash(f"{month_label(ym_ok)} is now closed. Its figures are saved permanently and can no longer change.", "success")
        else:
            flash(f"{month_label(ym_ok)} is already closed.", "error")
    except ValueError as e:
        flash(str(e), "error")
    conn.close()
    return redirect(url_for("history_month", ym=ym_ok))


@app.route("/history/<ym>/reopen", methods=["POST"])
def history_reopen(ym):
    ym_ok = valid_ym(ym)
    if not ym_ok:
        abort(404)
    if not request.form.get("confirm"):
        flash("Tick the confirmation box to reopen a closed month.", "error")
        return redirect(url_for("history_month", ym=ym_ok))
    conn = get_db()
    if reopen_month(conn, ym_ok):
        flash(f"{month_label(ym_ok)} is reopened for corrections. The previous snapshot is kept on file; "
              f"close the month again when you're done.", "success")
    else:
        flash(f"{month_label(ym_ok)} isn't closed.", "error")
    conn.close()
    return redirect(url_for("history_month", ym=ym_ok))


@app.route("/history/<ym>/export.json")
def history_export(ym):
    ym_ok = valid_ym(ym)
    if not ym_ok:
        abort(404)
    conn = get_db()
    snap = get_month_snapshot(conn, ym_ok)
    data = snap["data"] if snap else build_month_data(conn, ym_ok)
    conn.close()
    payload = io.BytesIO(json.dumps(data, indent=2, ensure_ascii=False).encode("utf-8"))
    return send_file(payload, as_attachment=True, download_name=f"shudh_rasoi_{ym_ok}_snapshot.json",
                     mimetype="application/json")


# ---------------------------------------------------------------------------
# Excel export
# ---------------------------------------------------------------------------
INR_FMT = '"₹"#,##0'
GREEN_HEX = "1F3B2C"
BAND_HEX = "F4F8F5"
BORDER_HEX = "BFBFBF"
_thin = Side(style="thin", color=BORDER_HEX)
_box = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)


def _xh(cell):
    cell.font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
    cell.fill = PatternFill("solid", fgColor=GREEN_HEX)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = _box


def _xc(cell, number_format=None, align="left"):
    cell.font = Font(name="Arial", size=10)
    cell.alignment = Alignment(horizontal=align, vertical="center")
    cell.border = _box
    if number_format:
        cell.number_format = number_format


def _xtable(ws, headers, rows, col_formats=None, widths=None, title=None):
    r = 1
    if title:
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))
        tc = ws.cell(row=1, column=1, value=title)
        tc.font = Font(name="Arial", size=13, bold=True, color="FFFFFF")
        tc.fill = PatternFill("solid", fgColor=GREEN_HEX)
        tc.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 24
        r = 2
    for i, h in enumerate(headers, start=1):
        _xh(ws.cell(row=r, column=i, value=h))
    ws.row_dimensions[r].height = 22
    if widths:
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(i)].width = w
    for ridx, row in enumerate(rows, start=r + 1):
        for cidx, val in enumerate(row, start=1):
            fmt = col_formats.get(cidx) if col_formats else None
            align = "right" if fmt else "left"
            cell = ws.cell(row=ridx, column=cidx, value=val)
            _xc(cell, number_format=fmt, align=align)
            if (ridx - r) % 2 == 0:
                cell.fill = PatternFill("solid", fgColor=BAND_HEX)
    ws.freeze_panes = f"A{r+1}"


@app.route("/export/excel")
def export_excel():
    conn = get_db()
    settings = get_settings()
    kpis = get_kpis(conn)
    dues_rows = get_all_dues(conn)

    customers_raw = conn.execute(
        "SELECT c.*, p.name AS pg_name FROM customers c LEFT JOIN pgs p ON p.id=c.pg_id ORDER BY c.name"
    ).fetchall()
    pgs_raw = conn.execute("SELECT * FROM pgs ORDER BY name").fetchall()
    payments_raw = conn.execute(
        """SELECT r.date, c.name AS customer_name, r.amount, r.payment_mode, r.notes, i.invoice_number
           FROM revenue r JOIN customers c ON c.id=r.customer_id
           LEFT JOIN invoices i ON i.id=r.invoice_id ORDER BY r.date DESC"""
    ).fetchall()
    extras_raw = conn.execute(
        """SELECT e.date, COALESCE(c.name, e.guest_name, 'Walk-in') AS who, e.description,
                  e.amount, e.payment_status, e.payment_mode
           FROM extra_orders e LEFT JOIN customers c ON c.id=e.customer_id ORDER BY e.date DESC"""
    ).fetchall()
    expenses_raw = conn.execute(
        "SELECT date, category, description, amount, payment_mode FROM expenses WHERE expense_type='Business' ORDER BY date DESC"
    ).fetchall()
    personal_expenses_raw = conn.execute(
        "SELECT date, category, description, amount, payment_mode FROM expenses WHERE expense_type='Personal' ORDER BY date DESC"
    ).fetchall()
    savings_raw = conn.execute(
        "SELECT date, entry_type, amount, description, payment_mode FROM savings ORDER BY date DESC"
    ).fetchall()
    invoices_raw = conn.execute(
        """SELECT i.invoice_number, c.name AS customer_name, i.period_start, i.period_end,
                  i.subscription_amount, i.extra_amount, i.total_amount, i.due_date
           FROM invoices i JOIN customers c ON c.id=i.customer_id ORDER BY i.generated_date DESC"""
    ).fetchall()
    monthly_rows = get_monthly_summary(conn, months_back=36)
    conn.close()

    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"
    ws.sheet_view.showGridLines = False
    ws.merge_cells("A1:B1")
    t = ws["A1"]
    t.value = f"{settings.get('business_name','Shudh Rasoi').upper()} — SNAPSHOT EXPORT"
    t.font = Font(name="Arial", size=14, bold=True, color="FFFFFF")
    t.fill = PatternFill("solid", fgColor=GREEN_HEX)
    t.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 26
    ws["A2"] = f"Exported on {datetime.now().strftime('%d-%b-%Y %H:%M')}"
    ws["A2"].font = Font(name="Arial", size=9, italic=True, color="666666")
    ws.column_dimensions["A"].width = 32
    ws.column_dimensions["B"].width = 20
    summary_data = [
        ("Total Revenue (all time)", kpis["total_revenue"], INR_FMT),
        ("Total Business Expense (all time)", kpis["total_expense"], INR_FMT),
        ("Net Profit — Business (all time)", kpis["net_profit"], INR_FMT),
        ("Profit Margin %", kpis["margin"] / 100, "0.0%"),
        ("Active Customers", kpis["active_customers"], "0"),
        ("Revenue This Month", kpis["month_revenue"], INR_FMT),
        ("Business Expense This Month", kpis["month_expense"], INR_FMT),
        ("Billed This Month", kpis["month_billed"], INR_FMT),
        ("Total Outstanding Dues", kpis["pending_total"], INR_FMT),
        ("Total Personal Expense (all time)", kpis["total_personal_expense"], INR_FMT),
        ("Total Savings Set Aside (all time)", kpis["savings_total_deposits"], INR_FMT),
        ("Currently in FD (principal only)", kpis["savings_fd_balance"], INR_FMT),
        ("Remaining Available Funds (all time)", kpis["remaining_funds"], INR_FMT),
    ]
    r = 4
    for label, val, fmt in summary_data:
        lc = ws.cell(row=r, column=1, value=label)
        lc.font = Font(name="Arial", size=10, bold=True)
        lc.border = _box
        vc = ws.cell(row=r, column=2, value=val)
        vc.border = _box
        vc.number_format = fmt
        vc.alignment = Alignment(horizontal="right")
        r += 1

    ws2 = wb.create_sheet("PGs")
    ws2.sheet_view.showGridLines = False
    _xtable(ws2, ["PG Name", "Address", "Contact Person", "Contact Number", "Notes"],
            [(p["name"], p["address"] or "", p["contact_person"] or "", p["contact_number"] or "", p["notes"] or "") for p in pgs_raw],
            widths=[22, 28, 18, 16, 26])

    ws3 = wb.create_sheet("Customers")
    ws3.sheet_view.showGridLines = False
    headers = ["Name", "Category", "PG", "Area", "Contact", "Plan", "Monthly Rate",
               "Join Date", "Status", "Outstanding"]
    rows = []
    dues_map = {d["id"]: d["outstanding"] for d in dues_rows}
    for c in customers_raw:
        rows.append((c["name"], c["category"], c["pg_name"] or "", c["area"] or "", c["contact"] or "",
                     plan_label(c), c["monthly_rate"], c["join_date"] or "", c["status"],
                     dues_map.get(c["id"], 0)))
    _xtable(ws3, headers, rows, col_formats={7: INR_FMT, 10: INR_FMT},
            widths=[22, 11, 18, 20, 14, 18, 13, 12, 10, 13])

    ws4 = wb.create_sheet("Payments")
    ws4.sheet_view.showGridLines = False
    _xtable(ws4, ["Date", "Customer", "Amount", "Mode", "Invoice #", "Notes"],
            [(p["date"], p["customer_name"], p["amount"], p["payment_mode"], p["invoice_number"] or "", p["notes"] or "") for p in payments_raw],
            col_formats={3: INR_FMT}, widths=[13, 22, 13, 14, 16, 24])

    ws5 = wb.create_sheet("Extra Orders")
    ws5.sheet_view.showGridLines = False
    _xtable(ws5, ["Date", "Customer/Guest", "Description", "Amount", "Status", "Mode"],
            [(e["date"], e["who"], e["description"] or "", e["amount"], e["payment_status"], e["payment_mode"] or "") for e in extras_raw],
            col_formats={4: INR_FMT}, widths=[13, 20, 24, 13, 11, 14])

    ws6 = wb.create_sheet("Business Expenses")
    ws6.sheet_view.showGridLines = False
    _xtable(ws6, ["Date", "Category", "Description", "Amount", "Mode"],
            [(e["date"], e["category"], e["description"] or "", e["amount"], e["payment_mode"] or "") for e in expenses_raw],
            col_formats={4: INR_FMT}, widths=[13, 24, 26, 13, 14])

    ws6b = wb.create_sheet("Personal Expenses")
    ws6b.sheet_view.showGridLines = False
    _xtable(ws6b, ["Date", "Category", "Description", "Amount", "Mode"],
            [(e["date"], e["category"], e["description"] or "", e["amount"], e["payment_mode"] or "") for e in personal_expenses_raw],
            col_formats={4: INR_FMT}, widths=[13, 24, 26, 13, 14])

    ws6c = wb.create_sheet("Savings & FD")
    ws6c.sheet_view.showGridLines = False
    _xtable(ws6c, ["Date", "Type", "Amount", "Notes", "Mode"],
            [(s["date"], s["entry_type"], s["amount"], s["description"] or "", s["payment_mode"] or "") for s in savings_raw],
            col_formats={3: INR_FMT}, widths=[13, 14, 13, 26, 14])

    ws7 = wb.create_sheet("Invoices")
    ws7.sheet_view.showGridLines = False
    _xtable(ws7, ["Invoice #", "Customer", "Period Start", "Period End", "Subscription", "Extras", "Total", "Due Date"],
            [(i["invoice_number"], i["customer_name"], i["period_start"], i["period_end"],
              i["subscription_amount"], i["extra_amount"], i["total_amount"], i["due_date"] or "") for i in invoices_raw],
            col_formats={5: INR_FMT, 6: INR_FMT, 7: INR_FMT}, widths=[15, 20, 13, 13, 13, 11, 13, 12])

    ws8 = wb.create_sheet("Dues")
    ws8.sheet_view.showGridLines = False
    _xtable(ws8, ["Customer", "Category", "Status", "Total Billed", "Outstanding"],
            [(d["name"], d["category"], d["status"], d["total_billed"], d["outstanding"]) for d in dues_rows],
            col_formats={4: INR_FMT, 5: INR_FMT}, widths=[22, 12, 11, 14, 14])

    ws9 = wb.create_sheet("Monthly Summary")
    ws9.sheet_view.showGridLines = False
    _xtable(ws9, ["Month", "Revenue", "Expense", "Net Profit", "Margin %"],
            [(m["label"], m["revenue"], m["expense"], m["profit"], m["margin"] / 100) for m in monthly_rows],
            col_formats={2: INR_FMT, 3: INR_FMT, 4: INR_FMT, 5: "0.0%"}, widths=[13, 15, 15, 15, 12])

    conn = get_db()  # (the connection used above was already closed; request teardown closes this one)
    ws10 = wb.create_sheet("Rate History")
    ws10.sheet_view.showGridLines = False
    rate_rows = conn.execute(
        """SELECT c.name, r.* FROM customer_rates r JOIN customers c ON c.id=r.customer_id
           ORDER BY c.name, r.effective_month""").fetchall()
    _xtable(ws10, ["Customer", "Effective From", "Plan", "Monthly Rate", "Breakfast Rate", "Lunch Rate", "Dinner Rate", "Note"],
            [(r["name"], r["effective_month"], plan_label(r) if customer_meals(r) else "No meals",
              effective_monthly_rate(r), r["rate_breakfast"], r["rate_lunch"], r["rate_dinner"], r["note"] or "")
             for r in rate_rows],
            col_formats={4: INR_FMT, 5: INR_FMT, 6: INR_FMT, 7: INR_FMT}, widths=[22, 14, 22, 14, 14, 12, 12, 26])

    ws11 = wb.create_sheet("Invoice Calculations")
    ws11.sheet_view.showGridLines = False
    snap_rows = conn.execute(
        """SELECT s.*, c.name AS customer_name FROM invoice_snapshots s
           JOIN customers c ON c.id=s.customer_id WHERE s.status='active' ORDER BY s.ym DESC, c.name""").fetchall()
    _xtable(ws11, ["Invoice #", "Customer", "Month", "Monthly Rate", "Rate From", "Daily Rate", "Billable Days",
                   "Total Meals", "Tiffins Taken", "Skipped", "On Leave", "Calculated", "Final Subscription", "Extras", "Total"],
            [(r["invoice_number"], r["customer_name"], r["ym"], r["rate_monthly"], r["rate_effective_from"] or "",
              r["daily_rate"], r["billable_days"], r["scheduled_meals"], r["delivered_meals"], r["skipped_meals"],
              r["leave_meals"], r["calculated_amount"], r["final_subscription_amount"], round(r["final_total"] - r["final_subscription_amount"], 2), r["final_total"])
             for r in snap_rows],
            col_formats={4: INR_FMT, 6: '"₹"#,##0.00', 12: INR_FMT, 13: INR_FMT, 14: INR_FMT, 15: INR_FMT},
            widths=[15, 20, 10, 13, 11, 11, 10, 11, 12, 9, 9, 13, 16, 10, 12])

    ws12 = wb.create_sheet("Month Snapshots")
    ws12.sheet_view.showGridLines = False
    _xtable(ws12, ["Month", "State", "Billed", "Revenue", "Expense", "Net Profit", "Closed On"],
            [(m["label"], m["state"].title(), m["billed"], m["revenue"], m["expense"], m["profit"], m["closed_at"] or "")
             for m in history_months(conn)],
            col_formats={3: INR_FMT, 4: INR_FMT, 5: INR_FMT, 6: INR_FMT}, widths=[13, 11, 14, 14, 14, 14, 20])

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    filename = f"shudh_rasoi_export_{today_date().isoformat()}.xlsx"
    return send_file(
        buf, as_attachment=True, download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    init_db()
    _db_ready = True
    host = os.environ.get("TIFFIN_HOST", "127.0.0.1")
    port = safe_int(os.environ.get("TIFFIN_PORT"), default=5000)
    debug = os.environ.get("TIFFIN_DEBUG") == "1"
    print("\n  Shudh Rasoi Business Manager is running.")
    print(f"  Open http://{'127.0.0.1' if host in ('0.0.0.0', '') else host}:{port} in your browser.\n")
    if debug:
        app.run(host=host, port=port, debug=True)
    else:
        try:
            from waitress import serve          # production-grade server (pip install waitress)
            serve(app, host=host, port=port, threads=4)
        except ImportError:
            app.run(host=host, port=port, debug=False)
