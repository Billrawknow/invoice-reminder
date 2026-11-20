"""
Tests for the invoice reminder (Week 5 test table and more).

    python -m unittest discover tests -v

Each test works in its own temporary folder, so your real Outbox,
reminder_history.csv and automation.log are never touched.
"""

import csv
import logging
import smtplib
import socket
import sys
import tempfile
import unittest
from datetime import date
from decimal import Decimal
from pathlib import Path
from unittest import mock

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import invoice_reminder as ir          # noqa: E402
import make_sample_invoices as sample  # noqa: E402
from openpyxl import Workbook          # noqa: E402

TODAY = date(2026, 11, 19)

# Keep test runs out of the real automation.log
logging.getLogger("invoice_reminder").propagate = False
logging.getLogger("invoice_reminder").addHandler(logging.NullHandler())


class ReminderTestCase(unittest.TestCase):
    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp())
        self.outbox = self.tmp / "Outbox"
        patcher = mock.patch.object(ir, "HISTORY_FILE", self.tmp / "reminder_history.csv")
        patcher.start()
        self.addCleanup(patcher.stop)
        self.sheet = self.make_sheet(sample.HEADER, sample.ROWS)

    def make_sheet(self, header, rows, name="invoices.xlsx"):
        book = Workbook()
        ws = book.active
        ws.title = "Invoices"
        if header:
            ws.append(header)
        for r in rows:
            ws.append(r)
        path = self.tmp / name
        book.save(path)
        return path

    def run_tool(self, sheet=None, today=TODAY, notify=False):
        """Run once; return (ok, log lines). Log lines are captured, not written to automation.log."""
        with mock.patch("builtins.print"), self.assertLogs("invoice_reminder", level="INFO") as logs:
            ok = ir.run_automation(sheet or self.sheet, self.outbox, today, notify)
        return ok, logs.output

    def drafts(self):
        return sorted(p.name for p in self.outbox.glob("*.txt") if not p.name.startswith("_"))


class TestClientData(ReminderTestCase):
    """MUST PASS: run with client data, check the output is correct."""

    def test_finds_the_right_overdue_invoices(self):
        invoices, skipped = ir.get_data(self.sheet)
        overdue = ir.process(invoices, TODAY)
        self.assertEqual(len(invoices), 12)
        self.assertEqual([s[1] for s in skipped], ["Tea Hotel Renovations"])
        self.assertEqual([i.invoice_no for i in overdue], [
            "INV-2026-0088", "INV-2026-0101", "INV-2026-0106", "INV-2026-0110",
            "INV-2026-0104", "INV-2026-0109", "INV-2026-0107"])
        self.assertEqual(sum(i.balance for i in overdue), Decimal("394050"))

    def test_paid_due_today_and_future_invoices_are_left_alone(self):
        self.run_tool()
        names = " ".join(self.drafts())
        for inv in ("0111", "0112", "0114", "0115", "0116"):   # paid, paid, due today, future, future
            self.assertNotIn(inv, names)

    def test_stage_gets_firmer_with_age(self):
        self.run_tool()
        stages = {n.split("_")[0]: n.split("_")[1] for n in self.drafts()}
        self.assertEqual(stages["INV-2026-0107"], "first")    # 4 days
        self.assertEqual(stages["INV-2026-0109"], "first")    # 7 days
        self.assertEqual(stages["INV-2026-0104"], "second")   # 14 days
        self.assertEqual(stages["INV-2026-0088"], "final")    # 50 days

    def test_letter_uses_balance_not_full_amount(self):
        self.run_tool()
        text = next(self.outbox.glob("INV-2026-0104_*")).read_text(encoding="utf-8")
        self.assertIn("Balance due:    KES 66,800", text)
        self.assertIn("KES 60,000 you have already paid", text)
        self.assertIn("Attn: The Bursar", text)

    def test_messy_cells_still_parse(self):
        # Row for INV-2026-0110 has "KES 9,800" as text and 05/11/2026 as a text date
        invoices, _ = ir.get_data(self.sheet)
        rono = next(i for i in invoices if i.invoice_no == "INV-2026-0110")
        self.assertEqual(rono.amount, Decimal("9800"))
        self.assertEqual(rono.due, date(2026, 11, 5))

    def test_every_sms_fits_one_message(self):
        invoices, _ = ir.get_data(self.sheet)
        with mock.patch.object(ir, "MPESA_PAYBILL", "123456"):
            for invoice in ir.process(invoices, TODAY):
                self.assertLessEqual(len(ir.draft_sms(invoice)), 160, invoice.invoice_no)

    def test_csv_export_works_too(self):
        path = self.tmp / "invoices.csv"
        with open(path, "w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            w.writerow(["invoice no", "client name", "invoice amount", "due date", "paid"])
            w.writerow(["QK-0001", "Kiptoo Builders Ltd", "48500.00", "2026-10-10", "no"])
        ok, _ = self.run_tool(path)
        self.assertTrue(ok)
        self.assertEqual(len(self.drafts()), 1)


class TestEmptyData(ReminderTestCase):
    """MUST PASS: run with empty data, handle it gracefully."""

    def test_header_only(self):
        ok, logs = self.run_tool(self.make_sheet(sample.HEADER, [], "empty.xlsx"))
        self.assertTrue(ok)
        self.assertEqual(self.drafts(), [])
        self.assertIn("0 invoices read", logs[-1])

    def test_completely_blank_workbook(self):
        ok, _ = self.run_tool(self.make_sheet(None, [], "blank.xlsx"))
        self.assertTrue(ok)
        self.assertEqual(self.drafts(), [])

    def test_missing_file_is_logged_not_crashed(self):
        ok, logs = self.run_tool(self.tmp / "nope.xlsx")
        self.assertFalse(ok)
        self.assertIn("FAILED", logs[-1])
        self.assertIn("Could not find your invoice sheet", logs[-1])
        self.assertIn("INVOICE_FILE in config.env", logs[-1])

    def test_missing_column_explains_what_to_fix(self):
        ok, logs = self.run_tool(self.make_sheet(["Client Name", "Invoice Amount", "Paid"],
                                                 [["X", 100, "no"]], "bad.xlsx"))
        self.assertFalse(ok)
        self.assertIn("due date", logs[-1])

    def test_bad_paid_value_skips_row(self):
        sheet = self.make_sheet(["Client Name", "Invoice Amount", "Due Date", "Paid"],
                                [["A", 100, date(2026, 11, 1), "maybe"], ["B", 200, date(2026, 11, 1), "no"]])
        ok, logs = self.run_tool(sheet)
        self.assertTrue(ok)
        self.assertEqual(len(self.drafts()), 1)
        self.assertTrue(any("Row 2 skipped" in line for line in logs))


class TestNoInternet(ReminderTestCase):
    """MUST PASS: no internet. The error is logged and the drafts are still saved."""

    def test_offline_notify_logs_error_and_keeps_drafts(self):
        offline = socket.gaierror(11001, "getaddrinfo failed")
        with mock.patch.object(ir, "setting", side_effect=lambda n, d="": "x@gmail.com" if "GMAIL" in n else d), \
                mock.patch.object(smtplib, "SMTP_SSL", side_effect=offline):
            ok, logs = self.run_tool(notify=True)
        self.assertFalse(ok)
        self.assertIn("FAILED", logs[-1])
        self.assertIn("No internet?", logs[-1])
        self.assertEqual(len(self.drafts()), 7)

    def test_notify_without_email_settings_is_skipped(self):
        with mock.patch.object(ir, "setting", return_value=""):
            ok, logs = self.run_tool(notify=True)
        self.assertTrue(ok)
        self.assertIn("owner skipped", logs[-1])


class TestNoDuplicates(ReminderTestCase):
    """CHECK THIS: run 3 times in a row, nothing is duplicated."""

    def test_three_runs_same_day(self):
        _, first = self.run_tool()
        _, second = self.run_tool()
        _, third = self.run_tool()
        self.assertEqual(len(self.drafts()), 7)
        self.assertIn("7 new drafts", first[-1])
        self.assertIn("0 new drafts, 7 already drafted", second[-1])
        self.assertIn("0 new drafts, 7 already drafted", third[-1])
        with open(ir.HISTORY_FILE, encoding="utf-8") as f:
            self.assertEqual(len(list(csv.DictReader(f))), 7)

    def test_sent_drafts_are_not_redrafted(self):
        self.run_tool()
        sent = self.outbox / "Sent"
        sent.mkdir()
        for p in self.outbox.glob("INV-*.txt"):
            p.rename(sent / p.name)
        self.run_tool()
        self.assertEqual(self.drafts(), [])

    def test_next_day_only_new_stages_get_drafts(self):
        self.run_tool()
        # 20 Nov: AIC Church moves from 7 to 8 days (second reminder) and
        # Bett Electricals, due 19 Nov, becomes 1 day overdue. Everyone else: unchanged.
        _, logs = self.run_tool(today=date(2026, 11, 20))
        self.assertIn("2 new drafts, 6 already drafted", logs[-1])
        self.assertTrue(any(n.startswith("INV-2026-0114_first") for n in self.drafts()))
        self.assertTrue(any(n.startswith("INV-2026-0109_second") for n in self.drafts()))
        text = next(self.outbox.glob("INV-2026-0109_second*")).read_text(encoding="utf-8")
        self.assertIn("Pastor Kirui", text)

    def test_final_notice_mentions_earlier_reminders_only_if_true(self):
        self.run_tool()
        first_time = next(self.outbox.glob("INV-2026-0088_final*")).read_text(encoding="utf-8")
        self.assertNotIn("sent earlier reminders", first_time)
        self.run_tool(today=date(2026, 12, 20))   # St. Mary's reaches 45 days
        later = next(self.outbox.glob("INV-2026-0104_final*")).read_text(encoding="utf-8")
        self.assertIn("sent earlier reminders", later)


class TestPaymentsChange(ReminderTestCase):
    def test_paid_invoice_draft_moves_to_cancelled(self):
        self.run_tool()
        rows = [r[:] for r in sample.ROWS]
        rows[1][-1] = "yes"   # Kiptoo Builders paid
        self.run_tool(self.make_sheet(sample.HEADER, rows))
        self.assertFalse(any(n.startswith("INV-2026-0101") for n in self.drafts()))
        self.assertTrue((self.outbox / "Cancelled" / "INV-2026-0101_final_Kiptoo_Builders_Ltd.txt").exists())

    def test_part_payment_updates_unsent_draft(self):
        self.run_tool()
        rows = [r[:] for r in sample.ROWS]
        rows[1][7] = 20000    # Kiptoo pays 20,000 of 48,500
        _, logs = self.run_tool(self.make_sheet(sample.HEADER, rows))
        self.assertIn("1 new drafts", logs[-1])
        text = (self.outbox / "INV-2026-0101_final_Kiptoo_Builders_Ltd.txt").read_text(encoding="utf-8")
        self.assertIn("Balance due:    KES 28,500", text)
        self.assertEqual(len(self.drafts()), 7)


class TestParsing(unittest.TestCase):
    def test_dates(self):
        for text in ("2026-11-05", "05/11/2026", "05-11-2026", "5 Nov 2026", "05.11.2026"):
            self.assertEqual(ir.parse_date(text), date(2026, 11, 5), text)
        with self.assertRaises(ValueError):
            ir.parse_date("next week")

    def test_money(self):
        self.assertEqual(ir.parse_money("KES 12,500"), Decimal("12500"))
        self.assertEqual(ir.parse_money("Ksh 1,200.50"), Decimal("1200.50"))
        self.assertEqual(ir.parse_money("7,350/="), Decimal("7350"))
        self.assertEqual(ir.parse_money(None), Decimal("0"))

    def test_paid(self):
        self.assertTrue(ir.parse_paid("Yes"))
        self.assertFalse(ir.parse_paid(" no "))
        self.assertFalse(ir.parse_paid(None))
        with self.assertRaises(ValueError):
            ir.parse_paid("maybe")


if __name__ == "__main__":
    unittest.main()
