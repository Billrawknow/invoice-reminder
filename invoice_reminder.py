"""
Invoice Reminder: drafts polite follow-ups for overdue invoices.

Built for a hardware shop that sells on credit to contractors, schools,
churches and fundis. It reads the shop's invoice sheet (Excel or CSV) and, for
every unpaid invoice past its due date, writes a ready-to-send reminder to
Outbox/ as a .txt file: an email version and a short SMS/WhatsApp version.

    python invoice_reminder.py                    # check invoices.xlsx once
    python invoice_reminder.py --file march.xlsx  # another sheet (or a .csv)
    python invoice_reminder.py --today 2026-11-19 # pretend it's another day (testing)
    python invoice_reminder.py --notify           # also email the owner today's chase list
    python invoice_reminder.py --daily 07:30      # keep running, check every day at 07:30

Reminders get firmer as an invoice gets older:
    1-7 days overdue    friendly reminder
    8-30 days overdue   second reminder (offers a payment plan)
    31+ days overdue    final reminder (asks for payment within 7 days)

Running it again never duplicates a reminder. reminder_history.csv records
every draft, so an invoice gets one draft per stage, and a new one only when it
moves to a firmer stage or its balance changes.
"""

import argparse
import csv
import logging
import os
import re
import shutil
import smtplib
import ssl
import sys
import time
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from decimal import Decimal, InvalidOperation
from email.message import EmailMessage
from pathlib import Path

from dotenv import load_dotenv

HERE = Path(__file__).resolve().parent
load_dotenv(HERE / ".env")

LOG_FILE = HERE / "automation.log"
HISTORY_FILE = HERE / "reminder_history.csv"
DEFAULT_SHEET = HERE / "invoices.xlsx"
DEFAULT_OUTBOX = HERE / "Outbox"

logging.basicConfig(filename=LOG_FILE, level=logging.INFO, encoding="utf-8",
                    format="%(asctime)s | %(levelname)s | %(message)s")
log = logging.getLogger("invoice_reminder")


def setting(name, default=""):
    return os.getenv(name, default).strip()


SHOP_NAME = setting("SHOP_NAME", "Mlimani Hardware & Building Supplies")
SHOP_PHONE = setting("SHOP_PHONE", "0700 000 000")
SHOP_EMAIL = setting("SHOP_EMAIL", "accounts@example.com")
SIGNED_BY = setting("SIGNED_BY", "Accounts Office")
MPESA_PAYBILL = setting("MPESA_PAYBILL", "")
MPESA_TILL = setting("MPESA_TILL", "")
BANK_DETAILS = setting("BANK_DETAILS", "")


class ReminderError(Exception):
    """A problem the shop can fix (missing file, wrong columns, bad email login)."""


# ------------------------------------------------------------------ columns

# The sheet can use any of these header names (case and spacing don't matter).
COLUMNS = {
    "client":      ["client name", "client", "customer", "customer name", "name"],
    "amount":      ["invoice amount", "amount", "total", "invoice total"],
    "due":         ["due date", "due", "date due"],
    "paid":        ["paid", "paid (yes/no)", "paid?", "status"],
    # optional
    "invoice_no":  ["invoice no", "invoice number", "invoice", "invoice #", "inv no", "number"],
    "email":       ["email", "client email", "email address"],
    "phone":       ["phone", "phone number", "mobile", "tel"],
    "amount_paid": ["amount paid", "paid amount", "part paid", "deposit"],
    "items":       ["items", "description", "goods", "details"],
    "contact":     ["contact person", "contact", "attention", "attn"],
}
REQUIRED = ["client", "amount", "due", "paid"]

YES = {"yes", "y", "paid", "true", "1", "done", "cleared"}
NO = {"no", "n", "unpaid", "false", "0", "not paid", "pending", "partial", "part paid", ""}

STAGES = [  # (from this many days overdue, stage key, label)
    (31, "final", "Final reminder"),
    (8, "second", "Second reminder"),
    (1, "first", "Friendly reminder"),
]


@dataclass
class Invoice:
    row: int
    invoice_no: str
    client: str
    amount: Decimal
    amount_paid: Decimal
    due: date
    paid: bool
    email: str = ""
    phone: str = ""
    items: str = ""
    contact: str = ""

    @property
    def balance(self):
        return self.amount - self.amount_paid

    def days_overdue(self, today):
        return (today - self.due).days

    def stage(self, today):
        days = self.days_overdue(today)
        for start, key, label in STAGES:
            if days >= start:
                return key, label
        return None, None


# ------------------------------------------------------------------ parsing

def norm(text):
    return re.sub(r"\s+", " ", str(text or "")).strip().lower()


def parse_money(value):
    if value is None or value == "":
        return Decimal("0")
    if isinstance(value, (int, float, Decimal)):
        return Decimal(str(value))
    cleaned = re.sub(r"(?i)kes|ksh|sh|/=|,|\s", "", str(value))
    try:
        return Decimal(cleaned)
    except InvalidOperation:
        raise ValueError(f"'{value}' is not an amount")


DATE_FORMATS = ["%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y", "%d %b %Y", "%d %B %Y", "%d/%m/%y"]


def parse_date(value):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = str(value or "").strip()
    if not text:
        raise ValueError("due date is empty")
    text = text.split(" ")[0] if re.match(r"\d{4}-\d{2}-\d{2} ", text) else text
    for fmt in DATE_FORMATS:   # Kenyan sheets write 05/11/2026 for 5 November
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    raise ValueError(f"'{value}' is not a date (use 2026-11-05 or 05/11/2026)")


def parse_paid(value):
    text = norm(value)
    if text in YES:
        return True
    if text in NO:
        return False
    raise ValueError(f"paid column says '{value}' (use yes or no)")


# ------------------------------------------------------------------ 1. get data

def read_rows(path):
    """Return the sheet as a list of rows (lists of cell values)."""
    if not path.exists():
        raise ReminderError(f"Can't find {path.name}. Put the invoice sheet next to the script "
                            f"or pass --file path\\to\\sheet.xlsx")
    try:
        if path.suffix.lower() == ".csv":
            with open(path, newline="", encoding="utf-8-sig") as f:
                return [row for row in csv.reader(f)]
        if path.suffix.lower() in (".xlsx", ".xlsm"):
            from openpyxl import load_workbook
            book = load_workbook(path, read_only=True, data_only=True)
            sheet = book["Invoices"] if "Invoices" in book.sheetnames else book.worksheets[0]
            rows = [list(r) for r in sheet.iter_rows(values_only=True)]
            book.close()
            return rows
    except PermissionError:
        raise ReminderError(f"{path.name} is locked. Close it in Excel and run again.")
    except ReminderError:
        raise
    except Exception as err:
        raise ReminderError(f"Couldn't open {path.name} ({type(err).__name__}). Is it a real Excel file?")
    raise ReminderError(f"{path.name}: use an .xlsx or .csv file.")


def find_columns(header):
    names = [norm(h) for h in header]
    found = {}
    for key, aliases in COLUMNS.items():
        for alias in aliases:
            if alias in names:
                found[key] = names.index(alias)
                break
    missing = [k for k in REQUIRED if k not in found]
    if missing:
        wanted = ", ".join(COLUMNS[k][0] for k in missing)
        raise ReminderError(f"The sheet is missing these columns: {wanted}. "
                            f"Row 1 must have: client name, invoice amount, due date, paid")
    return found


def get_data(path):
    """Read the sheet and return (invoices, skipped_rows)."""
    rows = read_rows(path)
    rows = [r for r in rows if any(c not in (None, "") for c in r)]   # drop blank lines
    if not rows:
        return [], []
    cols = find_columns(rows[0])

    def cell(row, key):
        i = cols.get(key)
        return row[i] if i is not None and i < len(row) else None

    invoices, skipped = [], []
    for n, row in enumerate(rows[1:], start=2):
        client = str(cell(row, "client") or "").strip()
        try:
            if not client:
                raise ValueError("client name is empty")
            amount = parse_money(cell(row, "amount"))
            invoice = Invoice(
                row=n,
                invoice_no=str(cell(row, "invoice_no") or "").strip() or f"ROW{n}",
                client=client,
                amount=amount,
                amount_paid=parse_money(cell(row, "amount_paid")),
                due=parse_date(cell(row, "due")),
                paid=parse_paid(cell(row, "paid")),
                email=str(cell(row, "email") or "").strip(),
                phone=str(cell(row, "phone") or "").strip(),
                items=str(cell(row, "items") or "").strip(),
                contact=str(cell(row, "contact") or "").strip(),
            )
        except ValueError as err:
            skipped.append((n, client or "?", str(err)))
            log.warning(f"Row {n} skipped ({client or 'no name'}): {err}")
            continue
        invoices.append(invoice)
    return invoices, skipped


# ------------------------------------------------------------------ 2. process

def kes(amount):
    amount = Decimal(amount)
    return f"KES {amount:,.0f}" if amount == amount.to_integral() else f"KES {amount:,.2f}"


def nice_date(d):
    return f"{d:%A} {d.day} {d:%B %Y}"


def short_date(d):
    return f"{d.day} {d:%b}"


def payment_lines(invoice):
    lines = []
    if MPESA_PAYBILL:
        lines.append(f"  M-Pesa Paybill {MPESA_PAYBILL}, Account {invoice.invoice_no}")
    if MPESA_TILL:
        lines.append(f"  M-Pesa Buy Goods Till {MPESA_TILL}")
    if BANK_DETAILS:
        lines.append(f"  Bank: {BANK_DETAILS}")
    return lines or [f"  Call us on {SHOP_PHONE} for payment details."]


ORG_WORDS = ("ltd", "limited", "school", "church", "company", "works", "academy", "college",
             "county", "enterprises", "agencies", "contractors", "co-op", "sacco", "hotel",
             "electricals", "hospital", "&", " co.", "group", "services", "ministry")


def is_organisation(client):
    return any(w in f" {client.lower()}" for w in ORG_WORDS)


def is_role(contact):
    return contact.lower().startswith("the ")


def letter_header(invoice):
    """Opening lines: 'Attn: The Bursar' for roles, 'Dear Mr. Kiptoo' for people."""
    if invoice.contact and is_role(invoice.contact):
        return [f"Attn: {invoice.contact}, {invoice.client}", "", "Dear Sir/Madam,"]
    if invoice.contact:
        return [f"Dear {invoice.contact},"]
    return ["Dear Sir/Madam,"] if is_organisation(invoice.client) else [f"Dear {invoice.client},"]


def sms_name(invoice):
    if invoice.contact:
        return "" if is_role(invoice.contact) else invoice.contact
    return "" if is_organisation(invoice.client) else invoice.client.split()[0]


def draft_email(invoice, today, reminded_before=False):
    key, label = invoice.stage(today)
    days = invoice.days_overdue(today)
    day_word = "day" if days == 1 else "days"
    subject = f"{label}: Invoice {invoice.invoice_no}, {kes(invoice.balance)} due {short_date(invoice.due)}"

    part_paid = ""
    if invoice.amount_paid > 0:
        part_paid = f"Thank you for the {kes(invoice.amount_paid)} you have already paid."

    if key == "first":
        opening = (f"I hope you are well. This is a friendly reminder that invoice {invoice.invoice_no} "
                   f"was due on {nice_date(invoice.due)} and is now {days} {day_word} overdue.")
        middle = ("If you have already paid, thank you, and please ignore this message. "
                  "Kindly send us the M-Pesa or bank reference so we can update our records.")
    elif key == "second":
        opening = (f"We have not yet received payment for invoice {invoice.invoice_no}, which was due "
                   f"on {nice_date(invoice.due)} and is now {days} {day_word} overdue.")
        middle = ("If something on the invoice is not right, or you would like to pay in instalments, "
                  f"please call me on {SHOP_PHONE} and we will agree a plan that works for you. "
                  "If you have already paid, please send us the reference so we can update our records.")
    else:
        pay_by = today + timedelta(days=7)
        earlier = ("We have sent earlier reminders and have not yet received payment."
                   if reminded_before else "We have not yet received payment.")
        opening = (f"Invoice {invoice.invoice_no} was due on {nice_date(invoice.due)} and is now "
                   f"{days} days overdue. {earlier}")
        middle = (f"Please settle the balance by {nice_date(pay_by)}, or call me on {SHOP_PHONE} before "
                  "then to agree a payment plan. Until the account is up to date we may have to pause "
                  "further supplies on credit. We value your business and would like to keep serving you.")

    lines = letter_header(invoice) + ["", opening, ""]
    if invoice.items:
        lines.append(f"  Goods supplied: {invoice.items}")
    lines += [f"  Invoice amount: {kes(invoice.amount)}",
              f"  Balance due:    {kes(invoice.balance)}", ""]
    if part_paid:
        lines += [part_paid.strip(), ""]
    lines += ["You can pay by:"] + payment_lines(invoice) + ["", middle, "",
              "Thank you for buying from us.", "", "Kind regards,", SIGNED_BY, SHOP_NAME,
              f"{SHOP_PHONE} | {SHOP_EMAIL}"]
    return subject, "\n".join(lines)


def draft_sms(invoice):
    """Short version for SMS or WhatsApp. Aims for one SMS (160 characters)."""
    name = sms_name(invoice)
    hi = f"Hi {name}, " if name else ""
    shop = SHOP_NAME.split("&")[0].strip()
    if MPESA_PAYBILL:
        pay = f"Pay via Paybill {MPESA_PAYBILL} Acc {invoice.invoice_no}."
    elif MPESA_TILL:
        pay = f"Pay via Till {MPESA_TILL}."
    else:
        pay = ""
    options = [
        f"{shop}: {hi}invoice {invoice.invoice_no} balance {kes(invoice.balance)} was due "
        f"{short_date(invoice.due)}. {pay} Already paid? Pls send ref. Call {SHOP_PHONE}",
        f"{shop}: {hi}inv {invoice.invoice_no} bal {kes(invoice.balance)} due "
        f"{short_date(invoice.due)}. {pay} Paid? Send ref. {SHOP_PHONE}",
        f"{shop}: inv {invoice.invoice_no} bal {kes(invoice.balance)} due {short_date(invoice.due)}. "
        f"{pay} Call {SHOP_PHONE}",
    ]
    options = [re.sub(r"\s+", " ", o).strip() for o in options]
    return next((o for o in options if len(o) <= 160), options[-1])


def process(invoices, today):
    """Pick the overdue invoices, firmest first."""
    overdue = [i for i in invoices if not i.paid and i.balance > 0 and i.days_overdue(today) >= 1]
    return sorted(overdue, key=lambda i: (-i.days_overdue(today), i.client))


# ------------------------------------------------------------------ 3. deliver

def safe_name(text):
    return re.sub(r"[^A-Za-z0-9-]+", "_", text).strip("_")[:60]


def draft_filename(invoice, stage):
    return f"{safe_name(invoice.invoice_no)}_{stage}_{safe_name(invoice.client)}.txt"


def load_history():
    if not HISTORY_FILE.exists():
        return {}
    with open(HISTORY_FILE, newline="", encoding="utf-8") as f:
        return {(r["invoice_no"], r["stage"]): r for r in csv.DictReader(f)}


def save_history(entries):
    new_file = not HISTORY_FILE.exists()
    with open(HISTORY_FILE, "a", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=["drafted_on", "invoice_no", "client", "stage", "balance", "file"])
        if new_file:
            writer.writeheader()
        writer.writerows(entries)


def write_draft(path, invoice, today, reminded_before=False):
    subject, body = draft_email(invoice, today, reminded_before)
    sms = draft_sms(invoice)
    to = invoice.email or "(no email on file - send by SMS/WhatsApp below)"
    text = "\n".join([
        f"TO:       {to}",
        f"PHONE:    {invoice.phone or '(no phone on file)'}",
        f"SUBJECT:  {subject}",
        "=" * 72, "", body, "", "=" * 72,
        f"SMS / WHATSAPP VERSION ({len(sms)} characters)", "-" * 72, sms, "",
    ])
    path.write_text(text, encoding="utf-8")


def deliver(overdue, invoices, today, outbox):
    """Write a draft for each overdue invoice not already drafted at this stage."""
    outbox.mkdir(exist_ok=True)
    history = load_history()
    new_entries, results = [], {"new": [], "updated": [], "already": [], "cancelled": [], "files": []}

    for invoice in overdue:
        stage, _ = invoice.stage(today)
        name = draft_filename(invoice, stage)
        path = outbox / name
        done = history.get((invoice.invoice_no, stage))
        if done and Decimal(done["balance"]) == invoice.balance:
            results["already"].append(invoice)
            continue
        if done and not path.exists():
            # The earlier draft was already sent (moved out of Outbox) but the
            # balance has changed since, so draft an updated one.
            name = name.replace(".txt", f"_updated_{today:%Y%m%d}.txt")
            path = outbox / name
        reminded_before = any(k[0] == invoice.invoice_no for k in history)
        write_draft(path, invoice, today, reminded_before)
        results["updated" if done else "new"].append(invoice)
        results["files"].append(name)
        new_entries.append({"drafted_on": today.isoformat(), "invoice_no": invoice.invoice_no,
                            "client": invoice.client, "stage": stage, "balance": str(invoice.balance),
                            "file": name})

    # Invoices now marked paid: pull any unsent draft out of the Outbox so
    # nobody chases a customer who has already paid.
    settled = {safe_name(i.invoice_no) for i in invoices if i.paid or i.balance <= 0}
    for path in outbox.glob("*.txt"):
        if path.name.startswith("_"):
            continue
        if any(path.name.startswith(f"{inv}_") for inv in settled):
            cancelled = outbox / "Cancelled"
            cancelled.mkdir(exist_ok=True)
            shutil.move(str(path), str(cancelled / path.name))
            results["cancelled"].append(path.name)

    if new_entries:
        save_history(new_entries)
    return results


def write_chase_list(overdue, results, skipped, today, outbox):
    """One summary file for the owner: who owes what, oldest first."""
    total = sum((i.balance for i in overdue), Decimal("0"))
    lines = [f"{SHOP_NAME}: overdue invoices on {nice_date(today)}", "",
             f"{len(overdue)} overdue invoice(s), {kes(total)} outstanding.", ""]
    if overdue:
        lines.append(f"{'Days':>5}  {'Invoice':<15} {'Client':<34} {'Balance':>14}  Stage")
        lines.append("-" * 84)
        for i in overdue:
            _, label = i.stage(today)
            lines.append(f"{i.days_overdue(today):>5}  {i.invoice_no:<15} {i.client[:34]:<34} "
                         f"{kes(i.balance):>14}  {label}")
        lines.append("")
    new = results["new"] + results["updated"]
    lines.append(f"New drafts in Outbox today: {len(new)}")
    lines += [f"  - {name}" for name in results["files"]]
    if results["already"]:
        lines.append(f"Already drafted earlier (not repeated): {len(results['already'])}")
    if results["cancelled"]:
        lines.append(f"Moved to Outbox/Cancelled because the invoice is now paid: {len(results['cancelled'])}")
        lines += [f"  - {n}" for n in results["cancelled"]]
    if skipped:
        lines += ["", "ROWS TO FIX IN THE SHEET (skipped):"]
        lines += [f"  - Row {n} ({client}): {why}" for n, client, why in skipped]
    lines += ["", "How to use: open each new draft, send the email or copy the SMS into WhatsApp,",
              "then move the file into Outbox/Sent so you know it has gone."]
    text = "\n".join(lines) + "\n"
    (outbox / "_chase_list.txt").write_text(text, encoding="utf-8")
    return text, total


def notify_owner(summary, today, results, outbox):
    """Email the chase list to the owner, with today's new drafts attached."""
    sender = setting("GMAIL_ADDRESS")
    password = setting("GMAIL_APP_PASSWORD").replace(" ", "")
    to = setting("NOTIFY_TO") or sender
    if not sender or not password:
        log.info("Notify skipped: GMAIL_ADDRESS / GMAIL_APP_PASSWORD not set in .env")
        return "skipped (email not set up)"
    msg = EmailMessage()
    new = results["new"] + results["updated"]
    msg["Subject"] = f"Payment chase list {today:%d %b}: {len(new)} new reminder(s) ready"
    msg["From"] = f"{SHOP_NAME} <{sender}>"
    msg["To"] = to
    msg.set_content(summary)
    for name in results["files"]:
        path = outbox / name
        if path.exists():
            msg.add_attachment(path.read_bytes(), maintype="text", subtype="plain", filename=path.name)
    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=ssl.create_default_context(), timeout=30) as s:
            s.login(sender, password)
            s.send_message(msg)
    except smtplib.SMTPAuthenticationError:
        raise ReminderError("Gmail rejected the login. GMAIL_APP_PASSWORD must be a 16-character App Password.")
    except (smtplib.SMTPException, OSError) as err:
        raise ReminderError(f"Couldn't email the chase list ({type(err).__name__}). "
                            f"No internet? The drafts are still saved in Outbox.")
    return f"emailed to {to}"


# ------------------------------------------------------------------ run

def run_automation(sheet, outbox, today=None, notify=False, trigger="manual"):
    today = today or date.today()
    print(f"Starting at {datetime.now():%Y-%m-%d %H:%M:%S} (checking invoices as of {today})")
    try:
        invoices, skipped = get_data(sheet)                     # 1. fetch
        overdue = process(invoices, today)                      # 2. process
        results = deliver(overdue, invoices, today, outbox)     # 3. deliver
        summary, total = write_chase_list(overdue, results, skipped, today, outbox)
        print("\n" + summary)
        note = ""
        if notify:
            note = " | owner " + notify_owner(summary, today, results, outbox)
        made = len(results["new"]) + len(results["updated"])
        log.info(f"OK - {trigger} | {len(invoices)} invoices read, {len(overdue)} overdue ({kes(total)}), "
                 f"{made} new drafts, {len(results['already'])} already drafted, "
                 f"{len(results['cancelled'])} cancelled, rows skipped: {len(skipped)}{note}")   # 4. log
        if not invoices and not skipped:
            print("The sheet has no invoices yet. Nothing to do.")
        return True
    except ReminderError as err:
        print(f"PROBLEM: {err}")
        log.error(f"FAILED - {trigger} | {err}")
        return False
    except Exception as err:   # never let a surprise crash go unrecorded
        print(f"UNEXPECTED PROBLEM: {type(err).__name__}: {err}")
        log.exception(f"FAILED - {trigger} | unexpected {type(err).__name__}: {err}")
        return False


def main():
    p = argparse.ArgumentParser(description="Draft reminders for overdue invoices.")
    p.add_argument("--file", type=Path, default=DEFAULT_SHEET, help="invoice sheet (.xlsx or .csv)")
    p.add_argument("--outbox", type=Path, default=DEFAULT_OUTBOX, help="folder for the drafts")
    p.add_argument("--today", type=date.fromisoformat, help="pretend today is YYYY-MM-DD (testing)")
    p.add_argument("--notify", action="store_true", help="email the owner the chase list")
    p.add_argument("--daily", metavar="HH:MM", help="keep running and check every day at this time")
    p.add_argument("--scheduled", action="store_true", help=argparse.SUPPRESS)
    args = p.parse_args()
    trigger = "scheduled" if args.scheduled else "manual"

    if not args.daily:
        ok = run_automation(args.file, args.outbox, args.today, args.notify, trigger)
        sys.exit(0 if ok else 1)

    import schedule   # only needed for --daily
    try:
        schedule.every().day.at(args.daily).do(
            run_automation, args.file, args.outbox, None, args.notify, "daily-loop")
    except schedule.ScheduleValueError:
        sys.exit(f"--daily needs a time like 07:30, not '{args.daily}'")
    print(f"Running. Invoices will be checked every day at {args.daily}. Leave this window open "
          f"(Ctrl+C to stop).")
    log.info(f"Daily loop started, checking at {args.daily}")
    try:
        while True:
            schedule.run_pending()
            time.sleep(30)
    except KeyboardInterrupt:
        log.info("Daily loop stopped by user")
        print("Stopped.")


if __name__ == "__main__":
    main()
