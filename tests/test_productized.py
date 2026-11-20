"""
Tests for the client-ready parts: config.env checks, plain-English errors,
the daily schedule, and the release folder.

    python -m unittest discover tests -v
"""

import io
import logging
import os
import sys
import tempfile
import unittest
import zipfile
from contextlib import redirect_stdout
from pathlib import Path
from types import SimpleNamespace
from unittest import mock

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import build_release            # noqa: E402
import invoice_reminder as ir   # noqa: E402
import make_sample_invoices     # noqa: E402

logging.getLogger("invoice_reminder").propagate = False
logging.getLogger("invoice_reminder").addHandler(logging.NullHandler())

KEYS = ["SHOP_NAME", "SHOP_PHONE", "SHOP_EMAIL", "SIGNED_BY", "MPESA_PAYBILL", "MPESA_TILL", "BANK_DETAILS",
        "INVOICE_FILE", "OUTBOX_FOLDER", "CHECK_TIME", "EMAIL_ME_THE_LIST", "GMAIL_ADDRESS",
        "GMAIL_APP_PASSWORD", "NOTIFY_TO"]
GOOD = {"SHOP_NAME": "Mlimani Hardware", "SHOP_PHONE": "0712 345 678", "SHOP_EMAIL": "accounts@mlimani.co.ke",
        "MPESA_PAYBILL": "522522", "CHECK_TIME": "07:30", "EMAIL_ME_THE_LIST": "no"}


class ConfigTest(unittest.TestCase):
    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp())
        self.addCleanup(ir.load_settings)                       # restore real settings after the test
        env = mock.patch.dict(os.environ, {})
        env.start()
        self.addCleanup(env.stop)
        for key in KEYS:                                         # start clean, whatever is in your .env
            os.environ.pop(key, None)
        make_sample_invoices.build(self.tmp / "invoices.xlsx", [])

    def configure(self, values=None, text=None):
        """Write a config.env like a client would, and load it the way the tool does."""
        path = self.tmp / "config.env"
        if text is None:
            values = dict(values or {})
            values.setdefault("INVOICE_FILE", str(self.tmp / "invoices.xlsx"))
            text = "\n".join(f"{k}={v}" for k, v in values.items())
        path.write_text(text, encoding="utf-8")
        ir.load_settings(path)
        return ir.check_config()


class TestConfigCheck(ConfigTest):
    def test_untouched_template_lists_what_to_fill_in(self):
        template = (ROOT / "config.env.example").read_text(encoding="utf-8")
        problems, _ = self.configure(text=template.replace("INVOICE_FILE=invoices.xlsx",
                                                           f"INVOICE_FILE={self.tmp / 'invoices.xlsx'}"))
        joined = " ".join(problems)
        self.assertIn("SHOP_NAME still says 'Your Shop Name'", joined)
        self.assertIn("SHOP_PHONE still says", joined)
        self.assertNotIn("GMAIL", joined)          # email is off in the template, so Gmail lines don't matter

    def test_filled_in_config_passes(self):
        problems, notes = self.configure(GOOD)
        self.assertEqual(problems, [])
        self.assertEqual(notes, [])

    def test_typical_typos_are_named_with_their_line(self):
        problems, _ = self.configure({**GOOD, "MPESA_PAYBILL": "522 522A", "CHECK_TIME": "7.30am",
                                      "SHOP_PHONE": "0712", "SHOP_EMAIL": "accounts.mlimani"})
        joined = " ".join(problems)
        for expected in ("MPESA_PAYBILL '522 522A' should be 5 to 7 digits",
                         "CHECK_TIME '7.30am' should be a 24-hour time",
                         "SHOP_PHONE '0712' doesn't look like a phone number",
                         "SHOP_EMAIL 'accounts.mlimani' isn't a valid email"):
            self.assertIn(expected, joined)

    def test_single_digit_hour_is_accepted(self):
        problems, _ = self.configure({**GOOD, "CHECK_TIME": "7:30"})
        self.assertEqual(problems, [])
        self.assertEqual(ir.CHECK_TIME, "07:30")

    def test_missing_sheet_message_points_at_invoice_file(self):
        problems, _ = self.configure({**GOOD, "INVOICE_FILE": "C:/Sales/daily/invoices.xlsx"})
        self.assertTrue(any("Could not find your invoice sheet" in p and "INVOICE_FILE in config.env" in p
                            for p in problems))

    def test_relative_and_absolute_paths(self):
        self.configure({**GOOD, "INVOICE_FILE": "invoices.xlsx", "OUTBOX_FOLDER": "Outbox"})
        self.assertEqual(ir.DEFAULT_SHEET, ir.HERE / "invoices.xlsx")      # next to the tool
        self.configure({**GOOD, "OUTBOX_FOLDER": str(self.tmp / "Reminders")})
        self.assertEqual(ir.DEFAULT_OUTBOX, self.tmp / "Reminders")

    def test_no_payment_method_is_a_note_not_an_error(self):
        problems, notes = self.configure({k: v for k, v in GOOD.items() if k != "MPESA_PAYBILL"})
        self.assertEqual(problems, [])
        self.assertIn("call us for payment details", notes[0])

    def test_email_on_needs_real_gmail_and_app_password(self):
        problems, _ = self.configure({**GOOD, "EMAIL_ME_THE_LIST": "yes", "GMAIL_ADDRESS": "yourshop@gmail.com",
                                      "GMAIL_APP_PASSWORD": "mypassword123"})
        joined = " ".join(problems)
        self.assertIn("GMAIL_ADDRESS still says", joined)
        self.assertIn("16-letter App Password", joined)
        problems, _ = self.configure({**GOOD, "EMAIL_ME_THE_LIST": "yes", "GMAIL_ADDRESS": "mlimani@gmail.com",
                                      "GMAIL_APP_PASSWORD": "abcd efgh ijkl mnop", "NOTIFY_TO": "a@b.co.ke, c@d.com"})
        self.assertEqual(problems, [])


class TestFriendlyErrors(ConfigTest):
    def run_main(self, *args):
        out = io.StringIO()
        with mock.patch.object(sys, "argv", ["invoice_reminder.py", *args]), redirect_stdout(out):
            with self.assertRaises(SystemExit) as exit_:
                ir.main()
        return exit_.exception.code, out.getvalue()

    def test_bad_config_gives_a_numbered_list_not_a_traceback(self):
        self.configure({**GOOD, "MPESA_PAYBILL": "abc", "CHECK_TIME": "25:00"})
        code, out = self.run_main()
        self.assertEqual(code, 1)
        self.assertIn("Please fix these in config.env, then try again:", out)
        self.assertIn("1. MPESA_PAYBILL", out)
        self.assertIn("2. CHECK_TIME", out)
        self.assertNotIn("Traceback", out)

    def test_check_reports_ok(self):
        self.configure(GOOD)
        code, out = self.run_main("--check")
        self.assertEqual(code, 0)
        self.assertIn("Settings OK: Mlimani Hardware", out)

    def test_unexpected_crash_shows_a_calm_message(self):
        out = io.StringIO()
        with mock.patch.object(ir, "get_data", side_effect=KeyError("boom")), redirect_stdout(out):
            ok = ir.run_automation(self.tmp / "invoices.xlsx", self.tmp / "Outbox")
        self.assertFalse(ok)
        self.assertIn("Something unexpected went wrong. Nothing was lost", out.getvalue())
        self.assertNotIn("KeyError", out.getvalue())

    def test_email_login_failure_wording(self):
        import smtplib
        self.configure({**GOOD, "GMAIL_ADDRESS": "m@gmail.com", "GMAIL_APP_PASSWORD": "abcdefghijklmnop"})
        bad_login = smtplib.SMTPAuthenticationError(535, b"Username and Password not accepted")
        with mock.patch.object(smtplib, "SMTP_SSL", side_effect=bad_login):
            with self.assertRaisesRegex(ir.ReminderError, "Email login failed. Double-check GMAIL_ADDRESS"):
                ir.notify_owner("summary", ir.date.today(), {"new": [], "updated": [], "files": []}, self.tmp)


class TestSchedule(ConfigTest):
    def fake_run(self, code=0):
        calls = []

        def run(cmd, **kwargs):
            calls.append(cmd)
            return SimpleNamespace(returncode=code, stdout="", stderr="ERROR: Access is denied.")
        return calls, run

    def test_installs_at_check_time_from_config(self):
        self.configure({**GOOD, "CHECK_TIME": "18:00"})
        calls, run = self.fake_run()
        with mock.patch.object(ir, "ON_WINDOWS", True), redirect_stdout(io.StringIO()):
            self.assertTrue(ir.install_schedule(run))
        create = calls[0]
        self.assertEqual(create[:4], ["schtasks", "/Create", "/TN", "Invoice Reminder"])
        self.assertEqual(create[create.index("/ST") + 1], "18:00")
        self.assertIn("--scheduled", create[create.index("/TR") + 1])
        self.assertIn("-AllowStartIfOnBatteries", calls[1][-1])        # laptops on battery still run it

    def test_exe_schedules_itself(self):
        with mock.patch.object(ir, "FROZEN", True), \
                mock.patch.object(sys, "executable", r"C:\Shop\InvoiceReminder\InvoiceReminder.exe"):
            self.assertEqual(ir.schedule_command(), r'"C:\Shop\InvoiceReminder\InvoiceReminder.exe" --scheduled')

    def test_access_denied_suggests_run_as_administrator(self):
        self.configure(GOOD)
        calls, run = self.fake_run(code=1)
        out = io.StringIO()
        with mock.patch.object(ir, "ON_WINDOWS", True), redirect_stdout(out):
            self.assertFalse(ir.install_schedule(run))
        self.assertIn("Run as administrator", out.getvalue())
        self.assertEqual(len(calls), 1)


class TestRelease(unittest.TestCase):
    def test_client_folder_has_everything_and_nothing_private(self):
        tmp = Path(tempfile.mkdtemp())
        with mock.patch.object(build_release, "RELEASE", tmp), \
                mock.patch.object(build_release, "FOLDER", tmp / "InvoiceReminder"):
            zip_path = build_release.package(exe=None, sample=False)
        folder = tmp / "InvoiceReminder"
        names = {p.name for p in folder.iterdir()}
        self.assertEqual(names, {"invoice_reminder.py", "requirements.txt", "config.env", "invoices.xlsx",
                                 "setup.bat", "run.bat", "remove_schedule.bat", "READ ME FIRST.txt"})
        self.assertEqual((folder / "config.env").read_text(encoding="utf-8"),
                         (ROOT / "config.env.example").read_text(encoding="utf-8"))   # template, never your settings
        from openpyxl import load_workbook
        rows = list(load_workbook(folder / "invoices.xlsx").active.iter_rows(values_only=True))
        self.assertEqual(rows[0][:2], ("Invoice No", "Client Name"))
        self.assertTrue(all(v is None for row in rows[1:] for v in row))              # blank for the client
        with zipfile.ZipFile(zip_path) as z:
            self.assertIn("InvoiceReminder/READ ME FIRST.txt", z.namelist())
            self.assertFalse(any(n.endswith((".env", "reminder_history.csv", "automation.log"))
                                 for n in z.namelist() if not n.endswith("config.env")))

    def test_read_me_first_setup_is_three_steps(self):
        text = (ROOT / "docs" / "READ_ME_FIRST.txt").read_text(encoding="utf-8")
        self.assertEqual([s for s in ("STEP 1.", "STEP 2.", "STEP 3.", "STEP 4.") if s in text],
                         ["STEP 1.", "STEP 2.", "STEP 3."])


if __name__ == "__main__":
    unittest.main()
