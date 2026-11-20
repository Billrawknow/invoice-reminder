INVOICE REMINDER - READ ME FIRST
================================

Every morning this tool looks at your invoice sheet, finds every customer
who has not paid on time, and writes a polite reminder for each one.
Each reminder has an email version and a short SMS/WhatsApp version, with
your Paybill and the invoice number already filled in.

You read each reminder and send it yourself. Nothing is sent without you.


SETTING IT UP (3 steps, about 10 minutes)
-----------------------------------------

STEP 1.  Unzip this folder and put it somewhere you will keep it,
         for example in Documents.

STEP 2.  Double-click  setup.bat
         Notepad opens your settings. Type your shop name, phone number
         and Paybill, then SAVE (Ctrl+S) and CLOSE Notepad.
         If something is wrong, setup tells you exactly which line to fix
         and opens Notepad again.

STEP 3.  Put your credit sales in  invoices.xlsx  (one row per invoice),
         then double-click  run.bat
         The Outbox folder opens with today's reminders.

That is it. From now on it checks by itself every morning at the time in
your settings (7:30 am unless you changed it).


EVERY DAY
---------

1. Open the Outbox folder. Start with  _chase_list.txt : it shows who owes
   you, how much, and for how long.
2. Open each reminder and send it, by email or by copying the short
   message into WhatsApp.
3. Move each reminder you have sent into  Outbox\Sent

When a customer pays, write  yes  in the Paid column of invoices.xlsx.
If they pay part, write the amount in Amount Paid.


THE INVOICE SHEET
-----------------

Columns you must fill:   Client Name, Invoice Amount, Due Date, Paid (yes/no)
Columns that help:       Invoice No, Contact Person, Email, Phone, Items,
                         Amount Paid

Dates can be written 05/11/2026 (day/month/year). Amounts can be written
48500, 48,500 or KES 48,500.


IF YOU SEE A MESSAGE
--------------------

"Please fix these in config.env"
    Open config.env (right-click > Open with > Notepad), fix the lines
    listed, save, and try again.

"Could not find your invoice sheet"
    The INVOICE_FILE line in config.env does not point to your sheet.
    Check the name and the folder.

"invoices.xlsx is locked. Close it in Excel"
    Close the sheet in Excel, then double-click run.bat again.

"ROWS TO FIX IN THE SHEET" (at the bottom of _chase_list.txt)
    A row is missing something, such as a due date. It is skipped, not
    guessed. Fix that row and it will be included next time.

"Email login failed"
    Only if you turned on EMAIL_ME_THE_LIST. Check the Gmail address and
    the 16-letter App Password in config.env. Your reminders are still in
    the Outbox.

"Something unexpected went wrong"
    Nothing is lost. Send the file  automation.log  from this folder to
    your support contact.


WHAT YOU NEVER HAVE TO WORRY ABOUT
----------------------------------

- No customer gets the same reminder twice. A new one is written only
  when an invoice gets more overdue (after 8 days, then after 31 days)
  or the balance changes.
- If a customer pays before you send their reminder, it is moved to
  Outbox\Cancelled so you do not chase someone who has paid.
- To stop the daily check, double-click  remove_schedule.bat
  To start it again, double-click  setup.bat


SUPPORT
-------

Rawknow Company Ltd  |  Bildad Ronoh  |  bildadronoh@gmail.com
