# Invoice Reminder

**Automated payment follow-ups for Kenyan hardware shops and suppliers who sell on credit**

### KES 12,000 setup + KES 2,500/month (optional care plan)

---

## What it does

Every morning, it reads your credit-sales sheet and finds every customer who hasn't paid on time. For each one it writes a polite reminder that's ready to send. Each reminder comes as an email and as a short SMS/WhatsApp message, with your Paybill and the invoice number already in it.

Reminders get firmer as an invoice gets older:
- **Friendly** at 1 day late.
- **Offers a payment plan** at 8 days.
- **Asks for payment within 7 days** after 30 days.

You also get a one-page **chase list**: who owes you, how much, and for how long, with the oldest debts first.

You read and send every reminder yourself; nothing goes out without you. Nobody is reminded twice for the same thing. If a customer pays before you send their reminder, it's pulled out automatically.

## Who it's for

- Hardware shops, building-materials suppliers and wholesalers with **20 to 200 credit invoices a month**.
- Typical customers on credit are contractors, fundis, schools, churches and county offices.
- Invoices are kept in Excel, or anything that saves as Excel or CSV, including QuoteKit exports.
- The shop has a Windows computer. **Python and technical skills are not needed.**

## What it replaces

- **About 3 hours a week** of going through the invoice book, working out who is late, and typing reminders one by one.
- **Forgotten invoices.** When the shop is busy, chasing stops, and an invoice that is 60 days late is far harder to collect than one that is 6 days late.

At KES 500 an hour of the owner's time, 3 hours a week is **about KES 6,000 a month**. The setup fee is covered in about two months on time alone. Collecting **one** forgotten KES 15,000 invoice covers it on day one.

## What you get

| | |
|---|---|
| **The tool** | `InvoiceReminder.exe` with nothing to install, plus the files below |
| **Your settings file** | `config.env`: your shop name, Paybill, sheet location and check time, all changed in Notepad |
| **Invoice sheet** | `invoices.xlsx` with the right columns and a yes/no Paid dropdown, or we connect your existing sheet |
| **Setup guide** | `READ ME FIRST.txt`: 3 steps, plus what every message means |
| **Installation** | We install it on your computer with your details and your real invoices, and check the first reminders with you |
| **Training** | 30 minutes covering the daily routine: open, send, move to Sent |
| **Support** | 1 month included: fixes, plus help with any message you don't understand |

## Prices

| Package | Price | Includes |
|---|---|---|
| **Setup** | **KES 12,000** once | Everything above, with 1 month of support |
| **Setup + 3 months** | **KES 18,000** once | Setup, plus 3 months of support and one change a month (e.g. your own letter wording or a new column) |
| **Care plan** | **KES 2,500/month** | We check it weekly, keep it working through Windows updates, and add a small improvement each quarter. Cancel any time |

Changes outside these packages, such as SMS sending through Africa's Talking or a different sheet layout, are quoted before any work starts.

## What we need from you

- Your shop name, phone, email, and M-Pesa Paybill or Till number (and bank details if you take transfers).
- Your current invoice list, or a sample of 10 rows so we can check it reads correctly.
- 1 hour at the shop computer for installation and training.

## Coming next

The same reminder rules are being built into **QuoteKit**, our quotation, invoicing and eTIMS app. There, reminders start automatically when an invoice is issued and stop the moment a payment is recorded. Invoice Reminder clients can move across with their history.

---

**Rawknow Company Ltd** · Bildad Ronoh · bildadronoh@gmail.com
