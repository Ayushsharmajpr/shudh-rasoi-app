"""Tests for rate history by effective month, permanent monthly snapshots, month locking,
history dashboards and invoice summaries.

Run from the project folder:   python -m unittest discover -s tests -v
"""
import hashlib
import io
import json
import os
import shutil
import sqlite3
import subprocess
import sys
import tempfile
import unittest
from datetime import date

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import app as shop  # noqa: E402

REAL_BACKUP = os.environ.get("TIFFIN_TEST_BACKUP", "")


class Base(unittest.TestCase):
    """Fresh database per test, with a controllable 'today'."""

    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        shop.DB_PATH = os.path.join(self.tmp, "t.db")
        shop.BACKUP_DIR = os.path.join(self.tmp, "backups")
        self.today = date(2026, 6, 15)
        shop.today_date = lambda: self.today
        shop._last_automation_check = None
        shop.init_db()
        shop._db_ready = True
        conn = shop.get_db()
        # start from an empty business (init_db seeds demo data on a brand-new install)
        for t in ("revenue", "expenses", "extra_orders", "daily_log", "invoices", "customer_rates",
                  "customer_pauses", "customers", "pgs", "savings", "personal_income", "recurring_expenses"):
            conn.execute(f"DELETE FROM {t}")
        conn.commit()
        conn.close()
        shop.app.test_client_class = __import__('authclient').AuthClient
        self.client = shop.app.test_client()

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    # ---- helpers -------------------------------------------------------------
    def set_today(self, y, m, d):
        self.today = date(y, m, d)
        shop._last_automation_check = None

    def db(self):
        conn = sqlite3.connect(shop.DB_PATH)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA foreign_keys = ON")
        return conn

    def add_customer(self, name, join, rate=0.0, meals=("lunch", "dinner"), per_meal=None, status="Active"):
        pm = per_meal or {}
        conn = self.db()
        cur = conn.execute(
            """INSERT INTO customers (name, type, category, area, contact, plan, monthly_rate, join_date, status,
                   meal_breakfast, meal_lunch, meal_dinner, rate_breakfast, rate_lunch, rate_dinner)
               VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
            (name, "Regular", "Regular", "Sector 5", "9000000000", " + ".join(m.title() for m in meals), rate, join, status,
             int("breakfast" in meals), int("lunch" in meals), int("dinner" in meals),
             pm.get("breakfast", 0), pm.get("lunch", 0), pm.get("dinner", 0)))
        cid = cur.lastrowid
        shop.seed_rate_history(conn, cid)
        conn.commit()
        conn.close()
        return cid

    def customer(self, cid):
        conn = self.db()
        r = conn.execute("SELECT * FROM customers WHERE id=?", (cid,)).fetchone()
        conn.close()
        return r

    def edit_customer(self, cid, effective_month=None, **overrides):
        c = self.customer(cid)
        form = {"name": c["name"], "category": c["category"], "area": c["area"] or "", "contact": c["contact"] or "",
                "join_date": c["join_date"], "monthly_rate": str(c["monthly_rate"] or 0),
                "rate_breakfast": str(c["rate_breakfast"] or 0), "rate_lunch": str(c["rate_lunch"] or 0),
                "rate_dinner": str(c["rate_dinner"] or 0), "notes": ""}
        for m in ("breakfast", "lunch", "dinner"):
            if c[f"meal_{m}"]:
                form[f"meal_{m}"] = "on"
        for k, v in overrides.items():
            if v is None:
                form.pop(k, None)
            else:
                form[k] = v
        if effective_month:
            form["effective_month"] = effective_month
        return self.client.post(f"/customers/{cid}/edit", data=form, follow_redirects=True)

    def log(self, cid, d, meal, status):
        conn = self.db()
        conn.execute("INSERT INTO daily_log (customer_id, date, meal, status) VALUES (?,?,?,?)", (cid, d, meal, status))
        conn.commit()
        conn.close()

    def generate(self, ym):
        conn = shop.get_db()
        created, skipped = shop.generate_invoices_for_month(conn, ym)
        conn.close()
        return created

    def invoice_for(self, cid, ym):
        conn = self.db()
        r = conn.execute("SELECT * FROM invoices WHERE customer_id=? AND substr(period_start,1,7)=?", (cid, ym)).fetchone()
        conn.close()
        return r

    def snapshot_blob(self, cid, ym):
        conn = self.db()
        r = conn.execute("SELECT snapshot_json FROM invoice_snapshots WHERE customer_id=? AND ym=? AND status='active'",
                         (cid, ym)).fetchone()
        conn.close()
        return r["snapshot_json"] if r else None

    def set_setting(self, key, value):
        conn = self.db()
        conn.execute("INSERT INTO settings (key,value) VALUES (?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value", (key, value))
        conn.commit()
        conn.close()


# =====================================================================================
class RateHistoryTests(Base):

    def test_multiple_rate_changes_resolve_by_effective_month(self):
        """100 -> 200 from Oct, 200 -> 300 from Dec: each month uses the rate in force that month."""
        cid = self.add_customer("Asha", "2026-06-10", rate=100)
        conn = shop.get_db()
        plan = {"monthly_rate": 200, "rate_breakfast": 0, "rate_lunch": 0, "rate_dinner": 0,
                "meal_breakfast": 0, "meal_lunch": 1, "meal_dinner": 1}
        shop.set_plan_from_month(conn, cid, "2026-10", plan)
        plan["monthly_rate"] = 300
        shop.set_plan_from_month(conn, cid, "2026-12", plan)
        conn.commit()
        got = {ym: shop.effective_monthly_rate(shop.plan_for_month(conn, cid, ym))
               for ym in ["2026-05", "2026-06", "2026-07", "2026-09", "2026-10", "2026-11", "2026-12", "2027-01"]}
        conn.close()
        self.assertEqual(got, {"2026-05": 100, "2026-06": 100, "2026-07": 100, "2026-09": 100,
                               "2026-10": 200, "2026-11": 200, "2026-12": 300, "2027-01": 300})

    def test_past_invoices_never_change_when_rate_changes_later(self):
        cid = self.add_customer("Asha", "2026-08-01", rate=100)

        # August and September are billed at 100 (invoices generated after each month ends)
        self.set_today(2026, 9, 1)
        self.generate("2026-08")
        # rate goes 100 -> 200 from October, edited during September
        self.set_today(2026, 9, 28)
        r = self.edit_customer(cid, effective_month="2026-10", monthly_rate="200")
        self.assertIn(b"applies from Oct 2026", r.data)
        self.assertEqual(self.customer(cid)["monthly_rate"], 100, "today's rate is still 100 until October")

        self.set_today(2026, 10, 1)
        self.generate("2026-09")
        aug_blob = self.snapshot_blob(cid, "2026-08")
        sep_blob = self.snapshot_blob(cid, "2026-09")
        self.assertEqual(self.invoice_for(cid, "2026-08")["total_amount"], 100)
        self.assertEqual(self.invoice_for(cid, "2026-09")["total_amount"], 100)

        # October onward: 200
        self.set_today(2026, 11, 1)
        self.generate("2026-10")
        self.assertEqual(self.invoice_for(cid, "2026-10")["total_amount"], 200)
        self.client.get("/customers")   # first page load of the day refreshes the displayed current rate
        self.assertEqual(self.customer(cid)["monthly_rate"], 200, "customer list mirrors the rate now in force")

        # second change: 200 -> 300 from December
        self.set_today(2026, 11, 5)
        self.edit_customer(cid, effective_month="2026-12", monthly_rate="300")
        self.set_today(2026, 12, 1)
        self.generate("2026-11")
        self.set_today(2027, 1, 1)
        self.generate("2026-12")
        amounts = {ym: self.invoice_for(cid, ym)["total_amount"] for ym in
                   ["2026-08", "2026-09", "2026-10", "2026-11", "2026-12"]}
        self.assertEqual(amounts, {"2026-08": 100, "2026-09": 100, "2026-10": 200, "2026-11": 200, "2026-12": 300})

        # nothing issued earlier was touched by any later change
        self.assertEqual(self.snapshot_blob(cid, "2026-08"), aug_blob)
        self.assertEqual(self.snapshot_blob(cid, "2026-09"), sep_blob)

        # each snapshot remembers the rate it used and where that rate came from
        conn = self.db()
        rows = {r["ym"]: (r["rate_monthly"], r["rate_effective_from"]) for r in
                conn.execute("SELECT ym, rate_monthly, rate_effective_from FROM invoice_snapshots WHERE status='active'")}
        conn.close()
        self.assertEqual(rows["2026-09"], (100, "2026-08"))
        self.assertEqual(rows["2026-10"], (200, "2026-10"))
        self.assertEqual(rows["2026-12"], (300, "2026-12"))

    def test_rate_change_cannot_reach_back_into_an_invoiced_month(self):
        cid = self.add_customer("Asha", "2026-08-01", rate=100)
        self.set_today(2026, 10, 1)
        self.generate("2026-09")
        r = self.edit_customer(cid, effective_month="2026-09", monthly_rate="999")
        self.assertIn(b"can only be changed from", r.data)
        conn = self.db()
        n = conn.execute("SELECT COUNT(*) FROM customer_rates WHERE customer_id=?", (cid,)).fetchone()[0]
        conn.close()
        self.assertEqual(n, 1, "no rate row was added")
        # and the database itself refuses, whatever code path is used
        conn = self.db()
        with self.assertRaises(sqlite3.IntegrityError) as ctx:
            conn.execute("""INSERT INTO customer_rates (customer_id, effective_month, monthly_rate, created_at)
                            VALUES (?,?,?,?)""", (cid, "2026-09", 999, "x"))
        self.assertIn("RATE_LOCKED", str(ctx.exception))
        conn.close()

    def test_unchanged_form_does_not_touch_scheduled_changes(self):
        cid = self.add_customer("Asha", "2026-06-01", rate=100)
        self.edit_customer(cid, effective_month="2026-09", monthly_rate="250")
        self.edit_customer(cid, effective_month="2026-07", notes="just a note")  # no plan change
        conn = self.db()
        rows = [(r["effective_month"], r["monthly_rate"]) for r in
                conn.execute("SELECT * FROM customer_rates WHERE customer_id=? ORDER BY effective_month", (cid,))]
        conn.close()
        self.assertEqual(rows, [("2026-06", 100), ("2026-09", 250)])

    def test_scheduled_change_can_be_removed_but_not_a_past_one(self):
        cid = self.add_customer("Asha", "2026-05-01", rate=100)
        self.edit_customer(cid, effective_month="2026-09", monthly_rate="250")
        conn = self.db()
        sched = conn.execute("SELECT id FROM customer_rates WHERE effective_month='2026-09'").fetchone()["id"]
        first = conn.execute("SELECT id FROM customer_rates WHERE effective_month='2026-05'").fetchone()["id"]
        conn.close()
        r = self.client.post(f"/customers/{cid}/rates/{first}/delete", follow_redirects=True)
        self.assertIn(b"Only a scheduled", r.data)
        self.client.post(f"/customers/{cid}/rates/{sched}/delete", follow_redirects=True)
        conn = self.db()
        left = [r["effective_month"] for r in conn.execute("SELECT * FROM customer_rates WHERE customer_id=?", (cid,))]
        conn.close()
        self.assertEqual(left, ["2026-05"])

    def test_meal_plan_history_is_versioned_too(self):
        """Dropping dinner from November must not rewrite October's tiffin counts."""
        cid = self.add_customer("Meera", "2026-09-01", rate=3000)
        self.set_today(2026, 10, 20)
        self.edit_customer(cid, effective_month="2026-11", meal_dinner=None)
        conn = shop.get_db()
        oct_rows = {r["name"]: r for r in shop.get_monthly_tiffin_summary(conn, "2026-10")}
        conn.close()
        self.assertEqual(oct_rows["Meera"]["meals"], ["lunch", "dinner"])

        self.set_today(2026, 11, 20)
        conn = shop.get_db()
        nov_rows = {r["name"]: r for r in shop.get_monthly_tiffin_summary(conn, "2026-11")}
        oct_again = {r["name"]: r for r in shop.get_monthly_tiffin_summary(conn, "2026-10")}
        conn.close()
        self.assertEqual(nov_rows["Meera"]["meals"], ["lunch"])
        self.assertEqual(oct_again["Meera"]["meals"], ["lunch", "dinner"])
        self.assertEqual(oct_again["Meera"]["delivered"], 62)  # 31 days x 2 meals, still intact

    def test_per_meal_rates_override_flat_and_are_versioned(self):
        cid = self.add_customer("Kabir", "2026-09-01", rate=0, per_meal={"lunch": 1200, "dinner": 1500})
        self.set_today(2026, 10, 25)
        self.edit_customer(cid, effective_month="2026-11", rate_lunch="1400")
        conn = shop.get_db()
        sep = shop.compute_billing_calc(conn, conn.execute("SELECT * FROM customers WHERE id=?", (cid,)).fetchone(), "2026-09")
        nov = shop.compute_billing_calc(conn, conn.execute("SELECT * FROM customers WHERE id=?", (cid,)).fetchone(), "2026-11")
        conn.close()
        self.assertEqual((sep["rate"]["mode"], sep["rate"]["monthly"], sep["calculated_amount"]), ("per_meal", 2700, 2700))
        self.assertEqual((nov["rate"]["monthly"], nov["calculated_amount"]), (2900, 2900))


# =====================================================================================
class BillingCalculationTests(Base):

    def test_prorated_join_and_leave_days_use_that_months_rate(self):
        # Nov has 30 days. Joins 11 Nov -> 20 service days. 2 full leave days -> 18 billable days.
        cid = self.add_customer("Dev", "2026-11-11", rate=3000)
        self.log(cid, "2026-11-15", "lunch", "Leave")
        self.log(cid, "2026-11-15", "dinner", "Leave")
        self.log(cid, "2026-11-20", "lunch", "Leave")      # a single leave meal still credits that day
        self.log(cid, "2026-11-22", "dinner", "Skipped")   # skipped meals do NOT reduce the bill
        self.log(cid, "2026-11-23", "lunch", "Skipped")
        self.set_today(2026, 12, 1)
        conn = shop.get_db()
        calc = shop.compute_billing_calc(conn, conn.execute("SELECT * FROM customers WHERE id=?", (cid,)).fetchone(), "2026-11")
        conn.close()
        self.assertEqual(calc["active_days"], 20)
        self.assertEqual(calc["leave_days"], 2)
        self.assertEqual(calc["billable_days"], 18)
        self.assertEqual(calc["daily_rate"], 100)
        self.assertEqual(calc["calculated_amount"], 1800)
        self.assertEqual(calc["scheduled_meals"], 40)
        self.assertEqual(calc["skipped_meals"], 2)
        self.assertEqual(calc["leave_meals"], 3)
        self.assertEqual(calc["delivered_meals"], 40 - 2 - 3)
        self.assertEqual(calc["delivered_meals"] + calc["skipped_meals"] + calc["leave_meals"], calc["scheduled_meals"])

    def test_no_billing_before_join_month_and_zero_rate_gives_zero(self):
        cid = self.add_customer("Late", "2026-12-05", rate=3100)
        free = self.add_customer("Free", "2026-01-01", rate=0)
        conn = shop.get_db()
        late = shop.compute_billing_calc(conn, conn.execute("SELECT * FROM customers WHERE id=?", (cid,)).fetchone(), "2026-11")
        zero = shop.compute_billing_calc(conn, conn.execute("SELECT * FROM customers WHERE id=?", (free,)).fetchone(), "2026-11")
        conn.close()
        self.assertEqual((late["active_days"], late["calculated_amount"]), (0, 0))
        self.assertEqual(zero["calculated_amount"], 0)

    def test_invoice_extras_only_unpaid_and_never_double_billed(self):
        cid = self.add_customer("Eli", "2026-09-01", rate=3000)
        conn = self.db()
        conn.execute("INSERT INTO extra_orders (date, customer_id, description, amount, payment_status, payment_mode) VALUES (?,?,?,?,?,?)",
                     ("2026-09-10", cid, "paid cash thali", 120, "Paid", "Cash"))
        conn.execute("INSERT INTO extra_orders (date, customer_id, description, amount, payment_status) VALUES (?,?,?,?,?)",
                     ("2026-09-12", cid, "unpaid sweet box", 200, "Unpaid"))
        conn.commit()
        conn.close()
        self.set_today(2026, 10, 1)
        self.generate("2026-09")
        inv = self.invoice_for(cid, "2026-09")
        self.assertEqual(inv["extra_amount"], 200, "the extra already paid in cash is not billed again")
        self.assertEqual(inv["total_amount"], 3200)
        # an invoiced extra must be settled via its invoice, not marked paid separately
        conn = self.db()
        eid = conn.execute("SELECT id FROM extra_orders WHERE payment_status='Unpaid'").fetchone()["id"]
        conn.close()
        r = self.client.post(f"/extra-orders/{eid}/mark-paid", data={"payment_mode": "Cash"}, follow_redirects=True)
        self.assertIn(b"already on an invoice", r.data)

    def test_edit_amount_adjusts_final_but_keeps_original_calculation(self):
        cid = self.add_customer("Fay", "2026-09-01", rate=3000)
        self.set_today(2026, 10, 1)
        self.generate("2026-09")
        iid = self.invoice_for(cid, "2026-09")["id"]
        self.client.post(f"/invoices/{iid}/edit-amount", data={"subscription_amount": "2500"})
        conn = self.db()
        s = conn.execute("SELECT * FROM invoice_snapshots WHERE invoice_id=?", (iid,)).fetchone()
        conn.close()
        self.assertEqual((s["calculated_amount"], s["final_subscription_amount"], s["final_total"]), (3000, 2500, 2500))
        page = self.client.get(f"/invoices/{iid}").data.decode()
        self.assertIn("adjusted from", page.lower())

    def test_deleting_invoice_voids_snapshot_and_never_reuses_the_number(self):
        cid = self.add_customer("Gus", "2026-09-01", rate=3000)
        self.set_today(2026, 10, 1)
        self.generate("2026-09")
        first = self.invoice_for(cid, "2026-09")
        self.client.post(f"/invoices/{first['id']}/delete")
        conn = self.db()
        voided = conn.execute("SELECT status FROM invoice_snapshots WHERE invoice_id=?", (first["id"],)).fetchone()
        conn.close()
        self.assertEqual(voided["status"], "voided")
        self.generate("2026-09")
        second = self.invoice_for(cid, "2026-09")
        self.assertNotEqual(first["invoice_number"], second["invoice_number"])

    def test_pause_is_remembered_after_resume(self):
        cid = self.add_customer("Hana", "2026-09-01", rate=3000)
        self.set_today(2026, 10, 10)
        self.client.post(f"/customers/{cid}/status", data={"status": "Paused", "pause_from": "2026-10-10", "pause_to": "2026-10-14"})
        self.set_today(2026, 10, 20)
        self.client.post(f"/customers/{cid}/status", data={"status": "Active"})
        self.assertIsNone(self.customer(cid)["pause_from"])
        self.set_today(2026, 11, 1)
        self.generate("2026-10")
        inv = self.invoice_for(cid, "2026-10")
        # Oct has 31 days; 5 paused days are credited -> 26 billable days
        self.assertAlmostEqual(inv["subscription_amount"], round(3000 / 31 * 26, 2), places=2)
        conn = self.db()
        s = conn.execute("SELECT leave_days, billable_days FROM invoice_snapshots WHERE invoice_id=?", (inv["id"],)).fetchone()
        conn.close()
        self.assertEqual((s["leave_days"], s["billable_days"]), (5, 26))


# =====================================================================================
class SnapshotAndHistoryTests(Base):

    def build_september(self):
        cid = self.add_customer("Asha", "2026-08-01", rate=3000)
        self.log(cid, "2026-09-03", "lunch", "Skipped")
        self.log(cid, "2026-09-04", "dinner", "Skipped")
        self.log(cid, "2026-09-10", "lunch", "Leave")
        self.log(cid, "2026-09-10", "dinner", "Leave")
        conn = self.db()
        conn.execute("INSERT INTO expenses (date, category, description, amount, payment_mode, expense_type) VALUES (?,?,?,?,?,?)",
                     ("2026-09-05", "Gas Cylinder", "lpg", 500, "Cash", "Business"))
        conn.execute("INSERT INTO expenses (date, category, description, amount, payment_mode, expense_type) VALUES (?,?,?,?,?,?)",
                     ("2026-09-06", "Miscellaneous", "snacks", 80, "UPI", "Personal"))
        conn.execute("INSERT INTO revenue (date, customer_id, amount, payment_mode, notes) VALUES (?,?,?,?,?)",
                     ("2026-09-15", cid, 1000, "UPI", "advance"))
        conn.commit()
        conn.close()
        self.set_today(2026, 10, 1)
        self.generate("2026-09")
        return cid

    def test_month_snapshot_captures_everything(self):
        cid = self.build_september()
        self.set_today(2026, 10, 10)
        conn = shop.get_db()
        self.assertEqual(shop.close_month(conn, "2026-09", "manual"), 1)
        snap = shop.get_month_snapshot(conn, "2026-09")
        conn.close()
        d = snap["data"]
        k = d["kpis"]
        self.assertEqual((k["billed"], k["revenue"], k["business_expense"], k["profit"], k["personal_expense"]),
                         (2900, 1000, 500, 500, 80))
        row = d["customers"][0]
        self.assertEqual(row["calc"]["rate"]["monthly"], 3000)
        self.assertEqual((row["calc"]["scheduled_meals"], row["calc"]["delivered_meals"],
                          row["calc"]["skipped_meals"], row["calc"]["leave_meals"]), (60, 56, 2, 2))
        self.assertEqual((row["invoice_total"], row["received"]), (2900, 1000))
        e = d["entries"]
        self.assertEqual(len(e["payments"]), 1)
        self.assertEqual(len(e["business_expenses"]), 1)
        self.assertEqual(len(e["personal_expenses"]), 1)
        self.assertEqual(len(e["tiffin_records"]), 4)
        self.assertEqual(e["invoices"][0]["final_total"], 2900)

    def test_closed_month_is_frozen_against_later_rates_and_entries(self):
        cid = self.build_september()
        self.set_today(2026, 10, 10)
        conn = shop.get_db()
        shop.close_month(conn, "2026-09", "manual")
        conn.close()
        conn = self.db()
        before = conn.execute("SELECT snapshot_json FROM month_snapshots WHERE ym='2026-09'").fetchone()[0]
        conn.close()
        page_before = self.client.get("/history/2026-09").data.decode()

        # future edits: new rate from November, new October entries, new customer, new payment
        self.edit_customer(cid, effective_month="2026-11", monthly_rate="9999")
        conn = self.db()
        conn.execute("INSERT INTO expenses (date, category, description, amount, payment_mode, expense_type) VALUES (?,?,?,?,?,?)",
                     ("2026-10-11", "Gas Cylinder", "lpg", 777, "Cash", "Business"))
        conn.execute("INSERT INTO revenue (date, customer_id, amount, payment_mode, notes) VALUES (?,?,?,?,?)",
                     ("2026-10-12", cid, 1900, "UPI", "pays september"))
        conn.commit()
        conn.close()
        self.add_customer("Newcomer", "2026-10-01", rate=5000)

        conn = self.db()
        after = conn.execute("SELECT snapshot_json FROM month_snapshots WHERE ym='2026-09'").fetchone()[0]
        conn.close()
        self.assertEqual(hashlib.sha256(before.encode()).hexdigest(), hashlib.sha256(after.encode()).hexdigest())
        page_after = self.client.get("/history/2026-09").data.decode()
        strip = lambda h: h.split("<main", 1)[-1]
        self.assertEqual(strip(page_before), strip(page_after))
        self.assertIn("Saved snapshot", page_after)
        self.assertIn("2,900", page_after)

    def test_database_blocks_changes_inside_closed_months(self):
        cid = self.build_september()
        self.set_today(2026, 10, 10)
        conn = shop.get_db()
        shop.close_month(conn, "2026-09", "manual")
        conn.close()
        conn = self.db()
        stmts = [
            ("INSERT INTO revenue (date, customer_id, amount, payment_mode) VALUES ('2026-09-20', ?, 5, 'Cash')", (cid,)),
            ("INSERT INTO expenses (date, category, amount, payment_mode, expense_type) VALUES ('2026-09-20','x',5,'Cash','Business')", ()),
            ("DELETE FROM expenses WHERE date='2026-09-05'", ()),
            ("UPDATE revenue SET amount=1 WHERE date='2026-09-15'", ()),
            ("INSERT INTO daily_log (customer_id, date, meal, status) VALUES (?, '2026-09-20', 'lunch', 'Skipped')", (cid,)),
            ("INSERT INTO extra_orders (date, description, amount, payment_status) VALUES ('2026-09-20','x',5,'Unpaid')", ()),
            ("UPDATE invoices SET total_amount=1", ()),
            ("DELETE FROM invoices", ()),
        ]
        for sql, params in stmts:
            with self.assertRaises(sqlite3.IntegrityError, msg=sql) as ctx:
                conn.execute(sql, params)
            self.assertIn("MONTH_LOCKED", str(ctx.exception), sql)
        # October is still open
        conn.execute("INSERT INTO expenses (date, category, amount, payment_mode, expense_type) VALUES ('2026-10-20','x',5,'Cash','Business')")
        # saved snapshots can never be edited or deleted
        for sql in ["UPDATE month_snapshots SET revenue=0", "UPDATE month_snapshots SET snapshot_json='{}'",
                    "DELETE FROM month_snapshots", "UPDATE invoice_snapshots SET rate_monthly=1",
                    "UPDATE invoice_snapshots SET snapshot_json='{}'", "DELETE FROM invoice_snapshots"]:
            with self.assertRaises(sqlite3.IntegrityError, msg=sql) as ctx:
                conn.execute(sql)
            self.assertIn("SNAPSHOT_IMMUTABLE", str(ctx.exception), sql)
        conn.close()

    def test_routes_report_locked_months_politely(self):
        cid = self.build_september()
        self.set_today(2026, 10, 10)
        conn = shop.get_db()
        shop.close_month(conn, "2026-09", "manual")
        conn.close()
        # tiffin toggle -> JSON 409 with an explanation
        r = self.client.post("/daily/toggle", json={"customer_id": cid, "date": "2026-09-20", "meal": "lunch", "status": "Skipped"})
        self.assertEqual(r.status_code, 409)
        self.assertIn("closed", r.get_json()["error"])
        # a payment dated inside the closed month -> friendly flash, nothing saved
        r = self.client.post("/payments", data={"customer_id": cid, "amount": "100", "date": "2026-09-25", "payment_mode": "Cash"},
                             follow_redirects=True)
        self.assertIn(b"closed month", r.data)
        conn = self.db()
        self.assertEqual(conn.execute("SELECT COUNT(*) FROM revenue").fetchone()[0], 1)
        conn.close()
        # invoices of a closed month are locked in the UI
        iid = self.invoice_for(cid, "2026-09")["id"]
        page = self.client.get(f"/invoices/{iid}").data.decode()
        self.assertIn("is closed, so this invoice is locked", page)
        self.assertNotIn("Delete invoice", page)
        # generating invoices for a closed month is refused
        r = self.client.post("/billing", data={"month": "2026-09"}, follow_redirects=True)
        self.assertIn(b"is closed", r.data)

    def test_payment_for_old_invoice_dated_today_is_allowed(self):
        cid = self.build_september()
        self.set_today(2026, 10, 10)
        conn = shop.get_db()
        shop.close_month(conn, "2026-09", "manual")
        conn.close()
        iid = self.invoice_for(cid, "2026-09")["id"]
        r = self.client.post(f"/invoices/{iid}/pay", data={"amount": "1900", "payment_mode": "UPI", "date": "2026-10-10"},
                             follow_redirects=True)
        self.assertIn(b"Payment recorded", r.data)
        conn = self.db()
        paid = conn.execute("SELECT SUM(amount) FROM revenue WHERE invoice_id=?", (iid,)).fetchone()[0]
        conn.close()
        self.assertEqual(paid, 1900)

    def test_reopen_edit_and_close_again_keeps_every_version(self):
        self.build_september()
        self.set_today(2026, 10, 10)
        conn = shop.get_db()
        shop.close_month(conn, "2026-09", "manual")
        conn.close()
        r = self.client.post("/history/2026-09/reopen", data={}, follow_redirects=True)
        self.assertIn(b"Tick the confirmation", r.data)
        self.client.post("/history/2026-09/reopen", data={"confirm": "1"})
        conn = self.db()
        conn.execute("INSERT INTO expenses (date, category, amount, payment_mode, expense_type) VALUES ('2026-09-28','Gas Cylinder',250,'Cash','Business')")
        conn.commit()
        conn.close()
        # a reopened month is never auto-closed again
        conn = shop.get_db()
        self.assertEqual(shop.auto_close_elapsed_months(conn), [])
        conn.close()
        self.client.post("/history/2026-09/close")
        conn = self.db()
        versions = [(r["version"], r["status"]) for r in conn.execute("SELECT version, status FROM month_snapshots WHERE ym='2026-09' ORDER BY version")]
        v1 = json.loads(conn.execute("SELECT snapshot_json FROM month_snapshots WHERE ym='2026-09' AND version=1").fetchone()[0])
        v2 = json.loads(conn.execute("SELECT snapshot_json FROM month_snapshots WHERE ym='2026-09' AND version=2").fetchone()[0])
        conn.close()
        self.assertEqual(versions, [(1, "superseded"), (2, "active")])
        self.assertEqual(v1["kpis"]["business_expense"], 500)
        self.assertEqual(v2["kpis"]["business_expense"], 750)

    def test_current_month_cannot_be_closed_and_open_month_is_live(self):
        self.build_september()
        self.set_setting("auto_close_months", "0")
        self.set_today(2026, 10, 10)
        conn = shop.get_db()
        with self.assertRaises(ValueError):
            shop.close_month(conn, "2026-10")
        conn.close()
        page = self.client.get("/history/2026-10").data.decode()
        self.assertIn("Live view", page)
        self.assertNotIn("Close this month now", page)   # still in progress
        self.assertIn("Close this month now", self.client.get("/history/2026-09").data.decode())

    def test_auto_close_waits_for_grace_period_and_closes_in_order(self):
        self.build_september()
        self.set_today(2026, 10, 5)               # 5 days after month end: still inside the grace window
        conn = shop.get_db()
        shop.auto_close_elapsed_months(conn)
        conn.close()
        conn = self.db()
        closed = [r["ym"] for r in conn.execute("SELECT ym FROM month_snapshots WHERE status='active' ORDER BY ym")]
        conn.close()
        self.assertNotIn("2026-09", closed)
        self.set_today(2026, 10, 6)
        conn = shop.get_db()
        shop.auto_close_elapsed_months(conn)
        conn.close()
        conn = self.db()
        row = conn.execute("SELECT closed_by FROM month_snapshots WHERE ym='2026-09' AND status='active'").fetchone()
        conn.close()
        self.assertEqual(row["closed_by"], "auto")

    def test_auto_close_can_be_switched_off(self):
        self.build_september()
        self.set_setting("auto_close_months", "0")
        self.set_today(2026, 11, 20)
        conn = shop.get_db()
        self.assertEqual(shop.auto_close_elapsed_months(conn), [])
        conn.close()

    def test_every_history_month_page_renders(self):
        self.build_september()
        self.set_today(2026, 10, 10)
        conn = shop.get_db()
        shop.close_month(conn, "2026-09", "manual")
        conn.close()
        for url in ["/history", "/history/2026-09", "/history/2026-10", "/history/2026-08", "/history/2027-02",
                    "/history/2026-09/export.json", "/monthly/2026-09", "/monthly", "/daily?date=2026-09-12&month=2026-09"]:
            self.assertEqual(self.client.get(url).status_code, 200, url)
        self.assertEqual(self.client.get("/history/not-a-month").status_code, 404)
        data = json.loads(self.client.get("/history/2026-09/export.json").data)
        self.assertEqual(data["kpis"]["billed"], 2900)

    def test_closed_month_tracker_matches_what_was_recorded(self):
        cid = self.build_september()
        self.set_today(2026, 10, 10)
        conn = shop.get_db()
        live = {r["name"]: (r["delivered"], r["skipped"], r["leave"]) for r in shop.get_monthly_tiffin_summary(conn, "2026-09")}
        shop.close_month(conn, "2026-09", "manual")
        # change the plan afterwards; the closed month's tracker must not move
        self.edit_customer(cid, effective_month="2026-11", meal_dinner=None)
        frozen = {r["name"]: (r["delivered"], r["skipped"], r["leave"]) for r in shop.get_monthly_tiffin_summary(conn, "2026-09")}
        conn.close()
        self.assertEqual(live, frozen)
        self.assertEqual(frozen["Asha"], (56, 2, 2))


# =====================================================================================
class InvoiceSummaryTests(Base):

    def pdf_text(self, iid):
        r = self.client.get(f"/invoices/{iid}/pdf")
        self.assertEqual(r.status_code, 200)
        path = os.path.join(self.tmp, "inv.pdf")
        with open(path, "wb") as fh:
            fh.write(r.data)
        out = subprocess.run(["pdftotext", "-layout", path, "-"], capture_output=True, text=True)
        return out.stdout

    def test_invoice_pdf_and_page_show_the_full_monthly_summary(self):
        cid = self.add_customer("Ila", "2026-09-01", rate=3000)
        self.log(cid, "2026-09-03", "lunch", "Skipped")
        self.log(cid, "2026-09-04", "dinner", "Skipped")
        self.log(cid, "2026-09-10", "lunch", "Leave")
        self.log(cid, "2026-09-10", "dinner", "Leave")
        self.set_today(2026, 10, 1)
        self.generate("2026-09")
        iid = self.invoice_for(cid, "2026-09")["id"]

        text = self.pdf_text(iid)
        compact = " ".join(text.split())
        for needle in ["Monthly summary", "Total meals", "in plan", "taken", "skipped", "on leave",
                       "Lunch + Dinner", "3,000.00 per month", "3,000.00 ÷ 30 days = ₹100.00 per day",
                       "100.00 × 29 billable days = ₹2,900.00", "Leave days credited 1 day",
                       "Skipped tiffins are not deducted"]:
            self.assertIn(needle, compact, needle)
        # the four headline numbers: 60 meals in plan, 56 taken, 2 skipped, 2 on leave
        nums = [t for t in text.split() if t.isdigit()]
        for n in ("60", "56"):
            self.assertIn(n, nums)

        page = self.client.get(f"/invoices/{iid}").data.decode()
        for needle in ["Monthly summary", "Tiffins taken", "29 of 30", "Rate in effect from Sep 2026"]:
            self.assertIn(needle, page)

    def test_old_invoice_still_shows_its_own_rate_after_rate_changes(self):
        cid = self.add_customer("Jai", "2026-08-01", rate=100)
        self.set_today(2026, 9, 1)
        self.generate("2026-08")
        iid = self.invoice_for(cid, "2026-08")["id"]
        text_before = " ".join(self.pdf_text(iid).split())
        self.set_today(2026, 9, 10)
        self.edit_customer(cid, effective_month="2026-10", monthly_rate="200")
        self.set_today(2026, 12, 1)
        self.edit_customer(cid, effective_month="2027-01", monthly_rate="400", meal_dinner=None)
        text_after = " ".join(self.pdf_text(iid).split())
        self.assertEqual(text_before, text_after)
        self.assertIn("100.00 per month", text_after)
        self.assertIn("Lunch + Dinner", text_after)

    def test_invoice_issued_mid_month_is_flagged_and_unchanged_by_later_records(self):
        cid = self.add_customer("Kia", "2026-09-01", rate=3000)
        self.set_today(2026, 9, 7)
        self.generate("2026-09")                        # issued before month end
        iid = self.invoice_for(cid, "2026-09")["id"]
        self.assertIn("before the month ended", " ".join(self.pdf_text(iid).split()))
        self.log(cid, "2026-09-20", "lunch", "Leave")   # recorded after the invoice was issued
        self.log(cid, "2026-09-20", "dinner", "Leave")
        self.assertEqual(self.invoice_for(cid, "2026-09")["total_amount"], 3000)
        page = self.client.get(f"/invoices/{iid}").data.decode()
        self.assertIn("changed after the invoice was generated", page)

    def test_special_characters_in_names_do_not_break_the_pdf(self):
        cid = self.add_customer("R&D <Mess> Co", "2026-09-01", rate=3000)
        self.set_today(2026, 10, 1)
        self.generate("2026-09")
        iid = self.invoice_for(cid, "2026-09")["id"]
        self.assertIn("R&D", self.pdf_text(iid))


# =====================================================================================
class RobustnessTests(Base):

    def test_backup_download_includes_uncheckpointed_wal_data(self):
        cid = self.add_customer("Lia", "2026-06-01", rate=100)
        conn = sqlite3.connect(shop.DB_PATH)
        conn.execute("INSERT INTO expenses (date, category, amount, payment_mode, expense_type) VALUES ('2026-06-14','Gas Cylinder',123,'Cash','Business')")
        conn.commit()
        conn.close()
        r = self.client.get("/settings/backup")
        self.assertEqual(r.status_code, 200)
        path = os.path.join(self.tmp, "dl.db")
        with open(path, "wb") as fh:
            fh.write(r.data)
        check = sqlite3.connect(path)
        self.assertEqual(check.execute("SELECT amount FROM expenses WHERE date='2026-06-14'").fetchone()[0], 123)
        self.assertEqual(check.execute("PRAGMA integrity_check").fetchone()[0], "ok")
        check.close()

    def test_restore_upgrades_an_old_backup_and_keeps_a_safety_copy(self):
        cid = self.add_customer("Mo", "2026-06-01", rate=100)
        old = os.path.join(self.tmp, "old.db")
        c = sqlite3.connect(old)
        c.executescript("""
            CREATE TABLE customers (id INTEGER PRIMARY KEY, name TEXT, type TEXT, area TEXT, contact TEXT, plan TEXT,
                monthly_rate REAL, join_date TEXT, status TEXT, category TEXT, pg_id INTEGER, meal_breakfast INTEGER DEFAULT 0,
                meal_lunch INTEGER DEFAULT 1, meal_dinner INTEGER DEFAULT 0, rate_breakfast REAL DEFAULT 0, rate_lunch REAL DEFAULT 0,
                rate_dinner REAL DEFAULT 0, pause_from TEXT, pause_to TEXT, notes TEXT);
            INSERT INTO customers (id, name, type, plan, monthly_rate, join_date, status, category) VALUES
                (1, 'Old Customer', 'Regular', 'Lunch', 2000, '2026-05-02', 'Active', 'Regular');
        """)
        c.commit()
        c.close()
        with open(old, "rb") as fh:
            r = self.client.post("/settings/restore", data={"confirm": "1", "backup_file": (fh, "old.db")},
                                 content_type="multipart/form-data", follow_redirects=True)
        self.assertIn(b"restored successfully", r.data)
        conn = self.db()
        self.assertEqual(conn.execute("SELECT name FROM customers").fetchall()[0]["name"], "Old Customer")
        self.assertEqual(conn.execute("SELECT effective_month, monthly_rate FROM customer_rates").fetchone()[:], ("2026-05", 2000))
        conn.close()
        self.assertTrue([f for f in os.listdir(self.tmp) if f.startswith("tiffin_before_restore_")])

    def test_garbage_restore_is_rejected_and_data_untouched(self):
        self.add_customer("Nia", "2026-06-01", rate=100)
        r = self.client.post("/settings/restore", data={"confirm": "1", "backup_file": (io.BytesIO(b"not a database"), "x.db")},
                             content_type="multipart/form-data", follow_redirects=True)
        self.assertIn(b"Restore failed", r.data)
        conn = self.db()
        self.assertEqual(conn.execute("SELECT COUNT(*) FROM customers").fetchone()[0], 1)
        conn.close()

    def test_init_db_is_idempotent(self):
        cid = self.add_customer("Oz", "2026-06-01", rate=100)
        conn = self.db()
        before = [tuple(r) for r in conn.execute("SELECT * FROM customer_rates")]
        conn.close()
        for _ in range(3):
            shop.init_db()
        conn = self.db()
        after = [tuple(r) for r in conn.execute("SELECT * FROM customer_rates")]
        conn.close()
        self.assertEqual(before, after)

    def test_settings_forms_do_not_clobber_each_others_fields(self):
        self.set_setting("auto_close_months", "0")
        self.client.post("/settings", data={"opening_balance_cash": "500"})   # the opening-balances form
        conn = self.db()
        row = conn.execute("SELECT value FROM settings WHERE key='auto_close_months'").fetchone()
        conn.close()
        self.assertEqual(row["value"], "0")

    def test_deleting_a_customer_with_no_history_works(self):
        cid = self.add_customer("Pip", "2026-06-01", rate=100)
        r = self.client.post(f"/customers/{cid}/delete", follow_redirects=True)
        self.assertIn(b"Customer removed", r.data)

    def test_every_main_page_renders(self):
        cid = self.add_customer("Qin", "2026-06-01", rate=100)
        for url in ["/", "/customers", f"/customers/{cid}", f"/customers/{cid}/edit", "/daily", "/extra-orders",
                    "/billing", "/payments", "/dues", "/expenses", "/monthly", "/reports", "/settings", "/savings",
                    "/personal-income", "/accounts", "/pgs", "/history", "/export/excel"]:
            self.assertEqual(self.client.get(url).status_code, 200, url)


# =====================================================================================
@unittest.skipUnless(REAL_BACKUP and os.path.exists(REAL_BACKUP), "set TIFFIN_TEST_BACKUP to a real backup to run this")
class RealBackupMigrationTests(unittest.TestCase):
    """Runs the upgrade against a real database from the live app."""

    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        shop.DB_PATH = os.path.join(self.tmp, "real.db")
        shop.BACKUP_DIR = os.path.join(self.tmp, "b")
        shutil.copy(REAL_BACKUP, shop.DB_PATH)
        self.today = date(2026, 10, 7)
        shop.today_date = lambda: self.today
        shop._last_automation_check = None
        shop.init_db()
        shop._db_ready = True

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def test_existing_invoices_keep_their_exact_amounts(self):
        conn = sqlite3.connect(shop.DB_PATH)
        conn.row_factory = sqlite3.Row
        rows = conn.execute("""SELECT i.invoice_number, i.subscription_amount, i.total_amount,
                                      s.calculated_amount, s.final_subscription_amount, s.final_total, s.rate_monthly
                               FROM invoices i JOIN invoice_snapshots s ON s.invoice_id=i.id AND s.status='active'""").fetchall()
        n_inv = conn.execute("SELECT COUNT(*) FROM invoices").fetchone()[0]
        conn.close()
        self.assertEqual(len(rows), n_inv)
        for r in rows:
            self.assertEqual(r["final_subscription_amount"], r["subscription_amount"], r["invoice_number"])
            self.assertEqual(r["final_total"], r["total_amount"], r["invoice_number"])
            self.assertAlmostEqual(r["calculated_amount"], r["subscription_amount"], places=2, msg=r["invoice_number"])

    def test_past_months_are_closed_and_every_page_works(self):
        client = shop.app.test_client()
        conn = sqlite3.connect(shop.DB_PATH)
        closed = [r[0] for r in conn.execute("SELECT ym FROM month_snapshots WHERE status='active' ORDER BY ym")]
        conn.close()
        self.assertTrue(closed)
        for ym in closed:
            self.assertEqual(client.get(f"/history/{ym}").status_code, 200)
        for url in ["/", "/customers", "/daily", "/billing", "/monthly", "/reports", "/history", "/export/excel"]:
            self.assertEqual(client.get(url).status_code, 200, url)


if __name__ == "__main__":
    unittest.main(verbosity=2)
