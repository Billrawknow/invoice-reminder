# Invoice Reminder for Hardware Shops

This tool reads a hardware shop's credit-sales sheet (Excel or CSV) and finds every **unpaid invoice past its due date**. For each one it drafts a polite follow-up and saves it as a `.txt` file in `Outbox/`. Each file holds an email and a short SMS/WhatsApp version (160 characters or less) that includes the M-Pesa Paybill and the invoice number as the account.

Built for AI Masterclass 2:
- **Week 5:** Automation Tool for a Business Problem.
- **Week 6, Day 28:** Productizing. It is now a client deliverable that needs no Python on the shop's computer.

**Client:** a hardware shop that sells on credit to contractors, schools, churches and fundis.
**Service description and prices:** [SERVICE.md](SERVICE.md) (KES 12,000 setup + KES 2,500/month).
**The client's setup guide:** [docs/READ_ME_FIRST.txt](docs/READ_ME_FIRST.txt)

## Day 28: from script to client product

| Lesson step | What was done |
|---|---|
| 1. Pick the strongest tool | This one: a real client type, a price, 23 tests, and a path into QuoteKit |
| 2. No hardcoded values | Everything a shop changes is in **`config.env`**: shop details, Paybill/Till/bank, `INVOICE_FILE`, `OUTBOX_FOLDER`, `CHECK_TIME`, and `EMAIL_ME_THE_LIST` with the Gmail settings. Relative paths sit next to the tool, and full paths such as `C:/Users/Shop/Documents/invoices.xlsx` work too |
| 3. Setup in 3 steps or fewer | Unzip, double-click `setup.bat` (fill in Notepad), then double-click `run.bat`. `setup.bat` checks the settings and reopens Notepad until they're right, then schedules the daily run at `CHECK_TIME` |
| 4. Validation and clear errors | `--check` validates every setting **before** anything runs and gives one numbered list naming the exact `config.env` line. A traceback never reaches the client; unexpected errors go to `automation.log` |
| 5. Service description | [SERVICE.md](SERVICE.md): what it does, who it's for, what it replaces, deliverables, KES prices |
| Extra: no Python for the client | `build_release.py` makes `InvoiceReminder.exe` (PyInstaller) and zips the client folder |

### Error messages before and after

| Situation | Before (Week 5 or raw Python) | After (what the client reads) |
|---|---|---|
| Sheet not where config says | `Can't find invoices.xlsx. Put the invoice sheet next to the script or pass --file` | `Could not find your invoice sheet (C:\Sales\invoices.xlsx). Check that INVOICE_FILE in config.env is correct.` |
| Settings not filled in | Letters went out signed "Your Shop Name" | `1. SHOP_NAME still says 'Your Shop Name'. Put in your real details.` |
| Typo in Paybill | `MPESA_PAYBILL=52252B` printed in every letter | `MPESA_PAYBILL '52252B' should be 5 to 7 digits only (no spaces or letters).` |
| Wrong time format | Task Scheduler error | `CHECK_TIME '7.30am' should be a 24-hour time like 07:30 or 18:00.` (`7:30` is accepted as `07:30`) |
| Gmail login rejected | `SMTPAuthenticationError: (535, b'Username and Password not accepted')` | `Email login failed. Double-check GMAIL_ADDRESS and the 16-letter GMAIL_APP_PASSWORD in config.env. The drafts are still saved in Outbox.` |
| Sheet open in Excel | `PermissionError: [Errno 13]` | `invoices.xlsx is locked. Close it in Excel and run again.` |
| Scheduling blocked | `ERROR: Access is denied.` | `Windows wouldn't create the daily task. Right-click setup.bat and choose 'Run as administrator'` |
| Anything unexpected | A Python traceback | `Something unexpected went wrong. Nothing was lost. Please send the file automation.log to your support contact.` (The traceback goes in the log for you) |

### Build the client folder

```bash
pip install pyinstaller
python build_release.py             # on Windows: builds InvoiceReminder.exe, then the folder and the zip
python build_release.py --sample    # same, with the 13 sample invoices for a demo
python build_release.py --no-exe    # ships the .py instead (client needs Python)
```

```
release/InvoiceReminder/            ->  release/InvoiceReminder_v1.0.zip
  InvoiceReminder.exe    the tool, nothing to install
  config.env             settings template with placeholders the check refuses to run with
  invoices.xlsx          blank sheet with the right columns and a yes/no dropdown
  setup.bat              edit settings -> check -> schedule daily
  run.bat                check now, open the Outbox
  remove_schedule.bat    turn the daily check off
  READ ME FIRST.txt      3-step setup plus what every message means
```

Your own `config.env`, invoices, Outbox, log and history are never copied. The build refuses to package them.

## The business problem

Hardware shops in Kenya give credit all the time: a contractor takes 60 bags of cement on account, or a school buys roofing sheets on an LPO. Then someone has to go through the invoice book every week, work out who is late, and type each reminder by hand. When the shop is busy, that doesn't happen, and old invoices are simply forgotten.

## What it does

```
invoices.xlsx  ->  1. read every row     ->  2. pick unpaid + overdue  ->  3. draft reminders  ->  Outbox/
                   (skip bad rows,          (firmest stage first)        (email + SMS,            _chase_list.txt
                    flag them)                                            no duplicates)          automation.log
```

| Days overdue | Stage | Tone |
|---|---|---|
| 1 to 7 | `first` | Friendly: "If you have already paid, thank you, please send the reference" |
| 8 to 30 | `second` | Offers instalments: "call me and we will agree a plan" |
| 31 or more | `final` | Asks for payment within 7 days, or credit supplies may be paused |

It also handles the things a real shop runs into:
- **Part payments:** the letter thanks the customer for what they've paid and chases only the balance.
- **Paid invoices:** an unsent draft is moved to `Outbox/Cancelled`, so no one chases a customer who has paid.
- **Messy cells:** `KES 9,800`, `7,350/=`, `05/11/2026` (Kenyan day/month order) and `Yes`/`no` all parse. A row it can't read is skipped and listed under "Rows to fix" in the chase list, never guessed.
- **Addressing:** a contact of "The Bursar" gives "Attn: The Bursar" followed by "Dear Sir/Madam", "Mr. Kiptoo" gives "Dear Mr. Kiptoo", and a person with no contact gets their own name.
- **No duplicates:** `reminder_history.csv` records each draft, so an invoice gets **one draft per stage**. A new draft is only made when the invoice moves to a firmer stage or its balance changes.

Example output from the sample sheet: [examples/](examples/)

```
Mlimani Hardware & Building Supplies: overdue invoices on Thursday 19 November 2026

7 overdue invoice(s), KES 394,050 outstanding.

 Days  Invoice         Client                                    Balance  Stage
------------------------------------------------------------------------------------
   50  INV-2026-0088   Kericho Water & Sanitation Co.        KES 214,000  Final reminder
   40  INV-2026-0101   Kiptoo Builders Ltd                    KES 48,500  Final reminder
   20  INV-2026-0106   Chebet Plumbing Works                  KES 15,600  Second reminder
   14  INV-2026-0110   Rono & Sons Contractors                 KES 9,800  Second reminder
   14  INV-2026-0104   St. Mary's Girls Secondary School      KES 66,800  Second reminder
    7  INV-2026-0109   AIC Church Kapsoit                     KES 32,000  Friendly reminder
    4  INV-2026-0107   Joseph Langat                           KES 7,350  Friendly reminder

ROWS TO FIX IN THE SHEET (skipped):
  - Row 14 (Tea Hotel Renovations): due date is empty
```

## Set it up on the shop's computer (3 steps)

1. **Unzip `InvoiceReminder_v1.0.zip`** somewhere you'll keep it, such as Documents.
2. **Double-click `setup.bat`.** Notepad opens `config.env`: type the shop name, phone and Paybill, then save and close. If anything is wrong, it says which line to fix and opens Notepad again. Once the settings are right, it schedules the daily check at `CHECK_TIME`. It runs on battery too, and catches up if the computer was off.
3. **Put the shop's invoices in `invoices.xlsx`**, then double-click `run.bat`. The Outbox folder opens with today's reminders.

To turn off the daily check, double-click `remove_schedule.bat`. To set it up from this repo instead of the release folder, run `setup.bat` here: it installs the requirements with Python first.

## The invoice sheet

Row 1 holds the headers. Their order doesn't matter, and neither does capitalisation.

| Column | Required | Notes |
|---|---|---|
| Client Name | yes | |
| Invoice Amount | yes | `48500`, `48,500` or `KES 48,500` |
| Due Date | yes | An Excel date, `2026-11-05` or `05/11/2026` |
| Paid | yes | `yes` or `no` (the sample sheet has a dropdown) |
| Invoice No | no | Strongly recommended: it is the M-Pesa account number in the reminder |
| Contact Person | no | `The Bursar` or `Mr. Kiptoo` |
| Email, Phone | no | With no email, the file says "send by SMS/WhatsApp" |
| Items | no | Shown as "Goods supplied" |
| Amount Paid | no | For part payments |

If the workbook has a sheet named **Invoices** it is used, otherwise the first sheet. Running `python make_sample_invoices.py` rebuilds the 13-row sample.

## Running it by hand

```bash
pip install -r requirements.txt
cp config.env.example config.env               # then fill it in
python invoice_reminder.py --check              # validate config.env, change nothing
python invoice_reminder.py                      # check the sheet once (add --open to open the Outbox)
python invoice_reminder.py --install-schedule   # daily at CHECK_TIME via Task Scheduler (prints a cron line on Linux/Mac)
python invoice_reminder.py --remove-schedule
python invoice_reminder.py --file march.csv     # another sheet, or a CSV export
python invoice_reminder.py --today 2026-11-19   # pretend it's another day (testing)
python invoice_reminder.py --notify             # email the owner this time (or set EMAIL_ME_THE_LIST=yes)
python invoice_reminder.py --daily 07:30        # stay running, check every day at 07:30 (schedule library)
```

The script has the four parts every production automation needs:

```python
invoices, skipped = get_data(sheet)                     # 1. fetch
overdue = process(invoices, today)                      # 2. process
results = deliver(overdue, invoices, today, outbox)     # 3. deliver
log.info(f"OK - {len(invoices)} invoices read, ...")    # 4. log
```

**`setup.bat` vs `--daily`:** the lesson's `schedule` loop (`--daily`) only works while a terminal window stays open. Task Scheduler (`setup.bat`) survives restarts and runs without a window, so that's what a client gets.

## Tests (Week 5 test table)

```bash
python -m unittest discover tests -v      # 40 tests, all in temp folders (your Outbox and config are untouched)
```

| Test | Result | How |
|---|---|---|
| Run with client data, check output is correct | **PASS** | 7 of 12 invoices overdue, KES 394,050. Paid, due-today and future invoices are left alone. The right stage per age, the balance used instead of the full amount, and every SMS is 160 characters or less |
| Run with empty data: crash or graceful? | **PASS** | A header-only sheet and a blank workbook both log `OK - 0 invoices read`. A missing file or missing column gives a plain-English `FAILED` log line, with no crash |
| Disconnect internet: is the error logged? | **PASS** | `--notify` with no connection logs `FAILED ... (gaierror). No internet? The drafts are still saved in Outbox.` Drafting itself needs no internet |
| Run it 3 times in a row: duplicates? | **PASS** | Run 1 makes 7 new drafts. Runs 2 and 3 make 0 new drafts and record "7 already drafted". Sent drafts moved to `Outbox/Sent` are not redrafted |
| Run with the client's actual data | **CHECK** | The sample includes a text amount, a dd/mm text date and a missing due date. Ask the shop for a real export before go-live |

Three runs in a row on 19 Nov (`automation.log`):

```
2026-11-19 07:34:29,213 | INFO | OK - manual | 12 invoices read, 7 overdue (KES 394,050), 7 new drafts, 0 already drafted, 0 cancelled, rows skipped: 1
2026-11-19 07:34:29,350 | INFO | OK - manual | 12 invoices read, 7 overdue (KES 394,050), 0 new drafts, 7 already drafted, 0 cancelled, rows skipped: 1
2026-11-19 07:34:29,485 | INFO | OK - manual | 12 invoices read, 7 overdue (KES 394,050), 0 new drafts, 7 already drafted, 0 cancelled, rows skipped: 1
```

Internet off, with `--notify`:

```
2026-11-19 07:32:16,088 | ERROR | FAILED - manual | Couldn't email the chase list (gaierror). No internet? The drafts are still saved in Outbox.
```

## Pricing

The full client-facing version is in [SERVICE.md](SERVICE.md). Price it on what it saves, not on how long it took to build.

| Package | Price | Includes |
|---|---|---|
| Setup | **KES 12,000** | Install with the shop's details and real invoices, 30 minutes of training, 1 month of support |
| Setup + 3 months | **KES 18,000** | The above, plus fixes and one change request a month for 3 months |
| Care plan | **KES 2,500/mo** | Weekly log check, keep it running through Windows updates, improvements each quarter |

**Justification:**
- A shop with about 40 credit invoices a month spends around 3 hours a week checking the book and typing reminders. That is 12 hours a month, and at KES 500/hour of the owner's time it comes to about **KES 6,000 a month**. Setup pays for itself in about 2 months on time alone.
- The bigger saving is cash. Customers who get a reminder on day 1, not day 40, pay sooner. Collecting **one** forgotten KES 15,000 invoice covers the whole setup fee.

## Next: making this part of QuoteKit

The columns were chosen to match what QuoteKit already keeps for each invoice: number, client, contact, items, total, amount paid, due date and status. That gives a clear path from this script into the product:

1. **Now:** export unpaid invoices from QuoteKit as CSV. `python invoice_reminder.py --file quotekit_export.csv` works with no changes (this is tested).
2. **In QuoteKit:** the same rules become a nightly `pg-boss` job in `apps/worker`:
   - query `status = unpaid AND due_date < today`;
   - work out the stage with the same 1/8/31-day bands;
   - insert into a `payment_reminders` table with a **unique (invoice_id, stage, balance_cents)** constraint, which is the same no-duplicates rule as `reminder_history.csv`, enforced by PostgreSQL.
3. **Automatic stop:** QuoteKit already records payments and M-Pesa references, so a payment cancels pending reminders directly. That replaces the `Outbox/Cancelled` step.
4. **Delivery:** the email goes through QuoteKit's existing mailer, and the SMS through Africa's Talking. The owner can approve each reminder or turn on auto-send for the friendly stage.
5. **Kenya-specific:** a government or school client that deducted **2% withholding VAT** leaves a small "balance" that isn't a debt. QuoteKit should match it to the WHVAT certificate instead of chasing it.

Money stays as integer cents in QuoteKit. This script uses `Decimal` for the same reason: float rounding should never reach a customer's letter.

## Files

| File | What it is |
|---|---|
| `invoice_reminder.py` | The tool: config check, read, find overdue invoices, draft, log, notify, schedule |
| `config.env.example` | The client's settings template, with plain-English notes on every line |
| `setup.bat`, `run.bat`, `remove_schedule.bat` | Double-click files. They use `InvoiceReminder.exe` if it's there, otherwise Python |
| `build_release.py` | Builds the `.exe` and the client folder/zip |
| `SERVICE.md` | Service description: what it does, who it's for, what it replaces, prices, deliverables |
| `docs/READ_ME_FIRST.txt` | The client's 3-step setup guide and message explanations (shipped as `READ ME FIRST.txt`) |
| `make_sample_invoices.py` | Builds the sample `invoices.xlsx`, or a blank client template with `--blank` |
| `invoices.xlsx` | Sample data (all names, emails and phone numbers are made up) |
| `tests/test_invoice_reminder.py` | 23 tests covering the Week 5 test table |
| `tests/test_productized.py` | 17 tests for config checks, friendly errors, scheduling and the release folder |
| `examples/` | Sample drafts and the chase list |

`config.env`, `.env`, `Outbox/`, `reminder_history.csv`, `automation.log`, `build/`, `dist/` and `release/` are git-ignored. On a client's machine they hold passwords and real customer data.

## Stack

Python 3.10+, openpyxl, python-dotenv, schedule, unittest, PyInstaller, and Windows Task Scheduler.
