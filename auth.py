"""Authentication for Shudh Rasoi.

* One administrator account stored in the `users` table of the application database.
* Passwords are hashed with Werkzeug (scrypt by default) — never stored in plain text.
* First run seeds `admin` / `admin` (flagged "must change"); the owner changes them in Account.
* Every route except /login, /static and /healthz requires a signed-in session.
* All POST requests must carry a CSRF token (injected into every <form> automatically).
* Repeated failed sign-ins are slowed down by a temporary lockout.
"""
import hmac
import re
import secrets
import threading
import time
from datetime import datetime, timedelta
from urllib.parse import urlparse

from flask import (
    abort, flash, g, jsonify, redirect, render_template, request, session, url_for
)
from werkzeug.security import check_password_hash, generate_password_hash

DEFAULT_USERNAME = "admin"
DEFAULT_PASSWORD = "admin"
USERNAME_RE = re.compile(r"^[A-Za-z0-9._-]{3,32}$")
MIN_PASSWORD_LEN = 8
COMMON_PASSWORDS = {
    "admin", "admin123", "password", "password1", "12345678", "123456789",
    "qwerty123", "11111111", "iloveyou", "welcome1", "letmein1",
}
MAX_FAILURES = 5          # failed attempts before a temporary lockout
LOCKOUT_SECONDS = 300     # 5 minutes
FAILURE_WINDOW = 900      # failures older than this are forgotten

PUBLIC_ENDPOINTS = {"login", "static", "healthz"}

_failures = {}            # (ip, username) -> [count, first_ts, locked_until]
_failures_lock = threading.Lock()
_DUMMY_HASH = generate_password_hash("not-the-password-" + secrets.token_hex(8))


# ---------------------------------------------------------------------------
# Schema + seeding
# ---------------------------------------------------------------------------
def ensure_auth_schema(conn):
    """Create the users table and seed the first-run administrator. Idempotent and additive."""
    conn.execute(
        """CREATE TABLE IF NOT EXISTS users (
               id INTEGER PRIMARY KEY AUTOINCREMENT,
               username TEXT NOT NULL UNIQUE COLLATE NOCASE,
               password_hash TEXT NOT NULL,
               must_change INTEGER NOT NULL DEFAULT 0,
               onboarded INTEGER NOT NULL DEFAULT 0,
               pw_version INTEGER NOT NULL DEFAULT 1,
               created_at TEXT NOT NULL,
               updated_at TEXT,
               last_login_at TEXT
           )""")
    if conn.execute("SELECT COUNT(*) FROM users").fetchone()[0] == 0:
        conn.execute(
            "INSERT INTO users (username, password_hash, must_change, created_at) VALUES (?,?,1,?)",
            (DEFAULT_USERNAME, generate_password_hash(DEFAULT_PASSWORD),
             datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
    conn.commit()


def _now():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def csrf_token():
    tok = session.get("csrf")
    if not tok:
        tok = secrets.token_urlsafe(32)
        session["csrf"] = tok
    return tok


def _safe_next(target):
    """Only allow same-site relative redirects after login (prevents open redirects)."""
    if not target:
        return None
    p = urlparse(target)
    if p.scheme or p.netloc or not target.startswith("/") or target.startswith("//") or "\\" in target:
        return None
    if target.startswith("/login") or target.startswith("/logout"):
        return None
    return target


def validate_new_password(pw, username):
    if len(pw) < MIN_PASSWORD_LEN:
        return f"The new password must be at least {MIN_PASSWORD_LEN} characters long."
    if len(pw) > 128:
        return "The new password is too long (128 characters maximum)."
    if pw.lower() in COMMON_PASSWORDS:
        return "That password is too easy to guess. Pick something less common."
    if username and pw.lower() == username.lower():
        return "The password can't be the same as the username."
    if pw.strip() != pw:
        return "The password can't start or end with a space."
    return None


def _client_key(username):
    return (request.remote_addr or "?", (username or "").strip().lower()[:64])


def _lock_remaining(key):
    with _failures_lock:
        rec = _failures.get(key)
        if rec and rec[2] > time.time():
            return int(rec[2] - time.time()) + 1
    return 0


def _record_failure(key):
    now = time.time()
    with _failures_lock:
        rec = _failures.get(key)
        if not rec or now - rec[1] > FAILURE_WINDOW:
            rec = [0, now, 0]
        rec[0] += 1
        if rec[0] >= MAX_FAILURES:
            rec[2] = now + LOCKOUT_SECONDS
            rec[0] = 0
            rec[1] = now
        _failures[key] = rec


def _clear_failures(key):
    with _failures_lock:
        _failures.pop(key, None)


def reset_throttle():
    """Test helper."""
    with _failures_lock:
        _failures.clear()


# ---------------------------------------------------------------------------
# Wiring
# ---------------------------------------------------------------------------
def init_auth(app, get_db, ensure_ready):
    """Register the guard, security headers, CSRF handling and the auth routes on `app`.
    Must be called before any other before_request hook is registered so that nothing
    runs for an unauthenticated visitor."""
    import os
    app.config.setdefault("SESSION_COOKIE_HTTPONLY", True)
    app.config.setdefault("SESSION_COOKIE_SAMESITE", "Lax")
    app.config.setdefault("SESSION_COOKIE_NAME", "shudh_rasoi_session")
    app.config["SESSION_COOKIE_SECURE"] = os.environ.get("TIFFIN_HTTPS") == "1"
    app.config["PERMANENT_SESSION_LIFETIME"] = timedelta(hours=12)

    def load_user():
        uid = session.get("uid")
        if not uid:
            return None
        conn = get_db()
        row = conn.execute("SELECT * FROM users WHERE id=?", (uid,)).fetchone()
        if not row or row["pw_version"] != session.get("pwv"):
            return None
        return row

    @app.before_request
    def _auth_guard():
        ensure_ready()
        ep = request.endpoint
        if ep in PUBLIC_ENDPOINTS:
            return None
        if request.path == "/favicon.ico":
            return None
        user = load_user()
        if user is None:
            session.pop("uid", None)
            wants_json = request.is_json or request.path.endswith("/toggle") \
                or request.headers.get("X-Requested-With") == "fetch"
            if wants_json:
                return jsonify({"ok": False, "error": "Your session has expired. Please sign in again."}), 401
            nxt = request.full_path.rstrip("?") if request.method == "GET" else None
            return redirect(url_for("login", next=_safe_next(nxt) or None))
        g.user = user
        if request.method == "POST":
            sent = request.form.get("csrf_token") or request.headers.get("X-CSRF-Token") or ""
            if not hmac.compare_digest(sent, session.get("csrf", "")):
                if request.is_json or request.headers.get("X-Requested-With") == "fetch":
                    return jsonify({"ok": False, "error": "Security token expired. Reload the page."}), 403
                flash("Your session security token expired. Please try again.", "error")
                return redirect(request.referrer or url_for("dashboard"))
        return None

    @app.context_processor
    def _auth_context():
        user = g.get("user")
        return {
            "csrf_token": csrf_token,
            "current_user": user["username"] if user is not None else None,
            "must_change_credentials": bool(user["must_change"]) if user is not None else False,
            "needs_onboarding": (user is not None and not user["onboarded"]),
        }

    @app.after_request
    def _security_headers(resp):
        resp.headers.setdefault("X-Content-Type-Options", "nosniff")
        resp.headers.setdefault("X-Frame-Options", "DENY")
        resp.headers.setdefault("Referrer-Policy", "same-origin")
        resp.headers.setdefault(
            "Content-Security-Policy",
            "default-src 'self'; img-src 'self' data: blob:; style-src 'self' 'unsafe-inline'; "
            "script-src 'self' 'unsafe-inline'; font-src 'self' data:; frame-ancestors 'none'; "
            "form-action 'self'; base-uri 'self'; object-src 'none'")
        if not request.path.startswith("/static/"):
            resp.headers["Cache-Control"] = "no-store"
        # Put the CSRF token into every POST form, so no template needs to remember it.
        if resp.mimetype == "text/html" and not resp.direct_passthrough and resp.status_code < 500:
            body = resp.get_data(as_text=True)
            if "<form" in body:
                tok = session.get("csrf") or csrf_token()
                field = f'<input type="hidden" name="csrf_token" value="{tok}">'
                body = re.sub(
                    r'(<form\b[^>]*\bmethod\s*=\s*["\']?post["\']?[^>]*>)',
                    lambda m: m.group(1) + field, body, flags=re.I)
                resp.set_data(body)
        return resp

    # ----- routes ---------------------------------------------------------
    @app.route("/healthz")
    def healthz():
        return jsonify({"status": "ok"})

    @app.route("/login", methods=["GET", "POST"])
    def login():
        if g.get("user") is None and session.get("uid"):
            session.pop("uid", None)
        conn = get_db()
        if session.get("uid") and load_user() is not None:
            return redirect(_safe_next(request.args.get("next")) or url_for("dashboard"))
        default_hint = bool(conn.execute(
            "SELECT 1 FROM users WHERE must_change=1 LIMIT 1").fetchone())
        error = None
        username = ""
        if request.method == "POST":
            sent = request.form.get("csrf_token") or ""
            if not hmac.compare_digest(sent, session.get("csrf", "")):
                error = "Your sign-in page expired. Please try again."
            else:
                username = (request.form.get("username") or "").strip()
                password = request.form.get("password") or ""
                key = _client_key(username)
                wait = _lock_remaining(key)
                if wait:
                    mins = max(1, round(wait / 60))
                    error = f"Too many failed attempts. Please wait about {mins} minute{'s' if mins != 1 else ''} and try again."
                else:
                    row = conn.execute(
                        "SELECT * FROM users WHERE username = ? COLLATE NOCASE", (username,)).fetchone()
                    ok = check_password_hash(row["password_hash"] if row else _DUMMY_HASH, password)
                    if row and ok:
                        _clear_failures(key)
                        old_csrf_free = True  # noqa: F841  (session is rebuilt below)
                        session.clear()
                        session.permanent = True
                        session["uid"] = row["id"]
                        session["pwv"] = row["pw_version"]
                        csrf_token()
                        conn.execute("UPDATE users SET last_login_at=? WHERE id=?", (_now(), row["id"]))
                        conn.commit()
                        return redirect(_safe_next(request.form.get("next") or request.args.get("next"))
                                        or url_for("dashboard"))
                    _record_failure(key)
                    error = "Incorrect username or password."
        csrf_token()
        return render_template("login.html", error=error, username=username,
                               next=_safe_next(request.args.get("next") or request.form.get("next")) or "",
                               default_hint=default_hint)

    @app.route("/logout", methods=["POST"])
    def logout():
        session.clear()
        flash("You have been signed out.", "success")
        return redirect(url_for("login"))

    @app.route("/account", methods=["GET"])
    def account():
        return render_template("account.html")

    @app.route("/account/credentials", methods=["POST"])
    def account_credentials():
        conn = get_db()
        user = g.user
        current = request.form.get("current_password") or ""
        new_username = (request.form.get("new_username") or "").strip()
        new_password = request.form.get("new_password") or ""
        confirm = request.form.get("confirm_password") or ""

        def fail(msg):
            flash(msg, "error")
            return redirect(url_for("account"))

        if not check_password_hash(user["password_hash"], current):
            return fail("Your current password is incorrect, so nothing was changed.")
        username_changed = bool(new_username) and new_username.lower() != user["username"].lower()
        # allow changing only the capitalisation of the username too
        if new_username and new_username != user["username"]:
            username_changed = True
        if not username_changed and not new_password:
            return fail("Enter a new username, a new password, or both.")
        final_username = new_username if username_changed else user["username"]
        if username_changed:
            if not USERNAME_RE.match(new_username):
                return fail("Usernames must be 3–32 characters: letters, numbers, dot, dash or underscore.")
            clash = conn.execute(
                "SELECT 1 FROM users WHERE username=? COLLATE NOCASE AND id<>?",
                (new_username, user["id"])).fetchone()
            if clash:
                return fail("That username is already taken.")
        new_hash = user["password_hash"]
        if new_password:
            if new_password != confirm:
                return fail("The new password and its confirmation don't match.")
            problem = validate_new_password(new_password, final_username)
            if problem:
                return fail(problem)
            if check_password_hash(user["password_hash"], new_password):
                return fail("The new password must be different from your current one.")
            new_hash = generate_password_hash(new_password)
        elif user["must_change"] and final_username.lower() == DEFAULT_USERNAME and not new_password:
            return fail("You're still using the default password. Please also choose a new password.")
        conn.execute(
            "UPDATE users SET username=?, password_hash=?, must_change=0, pw_version=pw_version+1, updated_at=? WHERE id=?",
            (final_username, new_hash, _now(), user["id"]))
        conn.commit()
        fresh = conn.execute("SELECT * FROM users WHERE id=?", (user["id"],)).fetchone()
        session.clear()                      # new session after a credential change
        session.permanent = True
        session["uid"] = fresh["id"]
        session["pwv"] = fresh["pw_version"]
        csrf_token()
        flash("Your sign-in details were updated. Use them next time you sign in.", "success")
        return redirect(url_for("account"))

    @app.route("/onboarding/complete", methods=["POST"])
    def onboarding_complete():
        conn = get_db()
        conn.execute("UPDATE users SET onboarded=1 WHERE id=?", (g.user["id"],))
        conn.commit()
        return jsonify({"ok": True})

    @app.route("/onboarding/restart", methods=["POST"])
    def onboarding_restart():
        conn = get_db()
        conn.execute("UPDATE users SET onboarded=0 WHERE id=?", (g.user["id"],))
        conn.commit()
        return redirect(url_for("dashboard"))
