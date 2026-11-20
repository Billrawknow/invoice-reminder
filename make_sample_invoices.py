"""
Creates invoices.xlsx: a sample credit-sales sheet for a hardware shop.

The columns match what QuoteKit keeps for each invoice, so a QuoteKit export
can replace this sheet later without changing the reminder script.

    python make_sample_invoices.py                  # invoices.xlsx with 13 sample invoices
    python make_sample_invoices.py --blank out.xlsx # empty template for a client (headers + dropdown)
"""

import argparse

from datetime import date
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

HEADER = ["Invoice No", "Client Name", "Contact Person", "Email", "Phone", "Items",
          "Invoice Amount", "Amount Paid", "Due Date", "Paid"]

# Due dates are set around 19 Nov 2026, the day the tool was tested.
ROWS = [
    ["INV-2026-0088", "Kericho Water & Sanitation Co.", "The Procurement Officer", "procurement@kewasco.example.com", "0711 000 101",
     "40 x 4-inch PVC pipes, 12 gate valves", 214000, 0, date(2026, 9, 30), "no"],
    ["INV-2026-0101", "Kiptoo Builders Ltd", "Mr. Kiptoo", "accounts@kiptoobuilders.example.com", "0722 000 102",
     "60 bags cement, 30 x Y12 bars", 48500, 0, date(2026, 10, 10), "no"],
    ["INV-2026-0104", "St. Mary's Girls Secondary School", "The Bursar", "bursar@stmarys.example.com", "0733 000 103",
     "Roofing sheets (gauge 30) x 45, ridges", 126800, 60000, date(2026, 11, 5), "no"],
    ["INV-2026-0106", "Chebet Plumbing Works", "Ms. Chebet", "chebetplumbing@example.com", "0712 000 104",
     "PPR pipes and fittings", 15600, 0, date(2026, 10, 30), "no"],
    ["INV-2026-0107", "Joseph Langat", "", "", "0798 000 105",
     "2 tins gloss paint, brushes, thinner", 7350, 0, date(2026, 11, 15), "no"],
    ["INV-2026-0109", "AIC Church Kapsoit", "Pastor Kirui", "aickapsoit@example.com", "0720 000 106",
     "Wheelbarrows x 2, 10 bags cement", 32000, 0, date(2026, 11, 12), "no"],
    ["INV-2026-0110", "Rono & Sons Contractors", "", "ronosons@example.com", "0725 000 107",
     "Binding wire, nails, timber 2x4", "KES 9,800", None, "05/11/2026", "No"],
    ["INV-2026-0111", "Mercy Wanjiku", "", "mercyw@example.com", "0701 000 108",
     "Door locks x 4, hinges", 4200, 0, date(2026, 11, 1), "yes"],
    ["INV-2026-0112", "Kapkatet Boys High School", "The Bursar", "bursar@kapkatet.example.com", "0734 000 109",
     "Desks hardware: bolts, screws, varnish", 18900, 18900, date(2026, 11, 8), "yes"],
    ["INV-2026-0114", "Bett Electricals", "Mr. Bett", "bettelectricals@example.com", "0716 000 110",
     "Cable 2.5mm x 3 rolls, sockets", 22750, 0, date(2026, 11, 19), "no"],
    ["INV-2026-0115", "Sigowet Farmers Co-op", "The Manager", "sigowetcoop@example.com", "0727 000 111",
     "Chain-link fence x 6 rolls, posts", 67400, 20000, date(2026, 11, 26), "no"],
    ["INV-2026-0116", "Daniel Mutai", "", "", "0790 000 112",
     "Plumbing tools set", 5600, 0, date(2026, 12, 3), "no"],
    ["INV-2026-0117", "Tea Hotel Renovations", "Mr. Otieno", "maintenance@teahotel.example.com", "0721 000 113",
     "Tiles 40 boxes, grout", 88300, 0, None, "no"],   # due date missing: skipped and flagged
]


def build(out, rows):
    book = Workbook()
    sheet = book.active
    sheet.title = "Invoices"
    sheet.append(HEADER)
    for row in rows:
        sheet.append(row)

    navy = PatternFill("solid", fgColor="1F4E79")
    for cell in sheet[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = navy
        cell.alignment = Alignment(vertical="center")
    widths = [15, 34, 22, 36, 14, 40, 15, 13, 12, 7]
    for col, width in zip("ABCDEFGHIJ", widths):
        sheet.column_dimensions[col].width = width
    for row in sheet.iter_rows(min_row=2, max_row=max(len(rows) + 1, 500)):
        row[6].number_format = row[7].number_format = "#,##0"
        row[8].number_format = "dd/mm/yyyy"
    sheet.freeze_panes = "A2"

    yes_no = DataValidation(type="list", formula1='"yes,no"', allow_blank=False)
    sheet.add_data_validation(yes_no)
    yes_no.add("J2:J1000")

    book.save(out)
    return out


def main():
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--blank", metavar="FILE", type=Path, help="write an empty client template instead")
    args = p.parse_args()
    if args.blank:
        build(args.blank, [])
        print(f"Saved blank template {args.blank}")
    else:
        out = build(Path(__file__).resolve().parent / "invoices.xlsx", ROWS)
        print(f"Saved {out.name} with {len(ROWS)} invoices.")


if __name__ == "__main__":
    main()
