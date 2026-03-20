"""
Transaction Management Sheet Generator
Generates a polished Excel workbook (TransactionManagement.xlsx) with:
  - Sheet 1: Transactions  (data entry with dropdowns, formulas, formatting)
  - Sheet 2: Dashboard     (live summary using COUNTIF/SUMIF formulas)
  - Sheet 3: Instructions  (usage guide)
"""

from openpyxl import Workbook
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side,
    GradientFill
)
from openpyxl.styles.numbers import FORMAT_DATE_DDMMYY
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.formatting.rule import CellIsRule, FormulaRule
import datetime

OUTPUT_FILE = "TransactionManagement.xlsx"
DATA_ROWS   = 500   # rows pre-formatted for data entry

# ─── Colour palette ───────────────────────────────────────────────────────────
C_HEADER_BG  = "1E3A5F"   # dark navy
C_HEADER_FG  = "FFFFFF"
C_ALT_ROW    = "F0F5FF"   # very light blue for alternating rows
C_PAID       = "C6EFCE"   # green fill  (payment paid)
C_PAID_FT    = "276221"
C_PENDING    = "FFEB9C"   # orange fill (payment pending)
C_PENDING_FT = "9C6500"
C_DUE        = "FFC7CE"   # red fill    (payment due)
C_DUE_FT     = "9C0006"
C_DEL_OK     = "DDEBF7"   # blue fill   (delivered)
C_DEL_OK_FT  = "1F4E79"
C_DEL_PEN    = "E2EFDA"   # light green (delivery pending)

# ─── Helper: thin border ──────────────────────────────────────────────────────
def thin_border():
    s = Side(style='thin', color='D9D9D9')
    return Border(left=s, right=s, top=s, bottom=s)

def header_border():
    s = Side(style='medium', color='1E3A5F')
    return Border(left=s, right=s, top=s, bottom=s)

# ─── Helper: apply fill ───────────────────────────────────────────────────────
def fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)


# ══════════════════════════════════════════════════════════════════════════════
#  SHEET 1 — TRANSACTIONS
# ══════════════════════════════════════════════════════════════════════════════
def build_transactions_sheet(ws):
    ws.title = "Transactions"

    # ── Column definitions ────────────────────────────────────────────────────
    columns = [
        ("A", "Transaction ID",       14),
        ("B", "Date",                 13),
        ("C", "Type",                 10),
        ("D", "Buyer Company",        24),
        ("E", "Seller Company",       24),
        ("F", "Product / Item",       30),
        ("G", "Quantity",             11),
        ("H", "Unit Price (₹)",       16),
        ("I", "Total Amount (₹)",     18),
        ("J", "Payment Status",       16),
        ("K", "Amount Paid (₹)",      16),
        ("L", "Balance Due (₹)",      16),
        ("M", "Payment Due Date",     17),
        ("N", "Delivery Status",      16),
        ("O", "Expected Delivery",    18),
        ("P", "Notes",                35),
    ]

    # ── Write & style headers ─────────────────────────────────────────────────
    for col_letter, title, width in columns:
        col_idx = ord(col_letter) - ord('A') + 1
        cell = ws.cell(row=1, column=col_idx, value=title)
        cell.font      = Font(bold=True, color=C_HEADER_FG, size=11)
        cell.fill      = fill(C_HEADER_BG)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border    = header_border()
        ws.column_dimensions[col_letter].width = width

    ws.row_dimensions[1].height = 36

    # ── Freeze header row ─────────────────────────────────────────────────────
    ws.freeze_panes = "A2"

    # ── Auto-filter on all columns ────────────────────────────────────────────
    ws.auto_filter.ref = f"A1:P1"

    # ── Date number format ────────────────────────────────────────────────────
    DATE_FMT = "DD/MM/YYYY"
    CURR_FMT = '#,##0.00'

    # ── Data rows ─────────────────────────────────────────────────────────────
    for r in range(2, DATA_ROWS + 2):
        # Alternating row background
        row_fill = fill(C_ALT_ROW) if r % 2 == 0 else fill("FFFFFF")

        for col_idx in range(1, 17):
            cell = ws.cell(row=r, column=col_idx)
            cell.fill   = row_fill
            cell.border = thin_border()
            cell.alignment = Alignment(vertical="center")

        # Col A — Transaction ID placeholder hint
        ws.cell(row=r, column=1).alignment = Alignment(horizontal="center", vertical="center")

        # Col B — Date format
        ws.cell(row=r, column=2).number_format = DATE_FMT

        # Col G — Quantity: number
        ws.cell(row=r, column=7).alignment = Alignment(horizontal="center", vertical="center")

        # Col H — Unit Price
        ws.cell(row=r, column=8).number_format = CURR_FMT

        # Col I — Total Amount = Qty * Unit Price
        ws.cell(row=r, column=9).value = f"=IF(G{r}*H{r}=0,\"\",G{r}*H{r})"
        ws.cell(row=r, column=9).number_format = CURR_FMT
        ws.cell(row=r, column=9).font = Font(bold=True)

        # Col K — Amount Paid
        ws.cell(row=r, column=11).number_format = CURR_FMT

        # Col L — Balance Due = Total - Paid
        ws.cell(row=r, column=12).value = f"=IF(I{r}=\"\",\"\",I{r}-IF(K{r}=\"\",0,K{r}))"
        ws.cell(row=r, column=12).number_format = CURR_FMT
        ws.cell(row=r, column=12).font = Font(color="9C0006")

        # Col M — Payment Due Date
        ws.cell(row=r, column=13).number_format = DATE_FMT

        # Col O — Expected Delivery Date
        ws.cell(row=r, column=15).number_format = DATE_FMT

        # Col P — Notes: wrap text
        ws.cell(row=r, column=16).alignment = Alignment(wrap_text=True, vertical="top")

    # ── Data Validation dropdowns ─────────────────────────────────────────────
    last_row = DATA_ROWS + 1

    dv_type = DataValidation(
        type="list", formula1='"Buy,Sell"',
        showDropDown=False, showErrorMessage=True,
        errorTitle="Invalid", error="Choose Buy or Sell"
    )
    dv_type.sqref = f"C2:C{last_row}"
    ws.add_data_validation(dv_type)

    dv_pay = DataValidation(
        type="list", formula1='"Paid,Pending,Due"',
        showDropDown=False, showErrorMessage=True,
        errorTitle="Invalid", error="Choose Paid, Pending or Due"
    )
    dv_pay.sqref = f"J2:J{last_row}"
    ws.add_data_validation(dv_pay)

    dv_del = DataValidation(
        type="list", formula1='"Delivered,Pending"',
        showDropDown=False, showErrorMessage=True,
        errorTitle="Invalid", error="Choose Delivered or Pending"
    )
    dv_del.sqref = f"N2:N{last_row}"
    ws.add_data_validation(dv_del)

    # ── Conditional formatting ────────────────────────────────────────────────
    range_pay = f"J2:J{DATA_ROWS + 1}"
    range_del = f"N2:N{DATA_ROWS + 1}"
    full_row  = f"A2:P{DATA_ROWS + 1}"

    # Payment Status — Paid (green)
    ws.conditional_formatting.add(range_pay, CellIsRule(
        operator="equal", formula=['"Paid"'],
        fill=fill(C_PAID), font=Font(bold=True, color=C_PAID_FT)
    ))
    # Payment Status — Pending (yellow)
    ws.conditional_formatting.add(range_pay, CellIsRule(
        operator="equal", formula=['"Pending"'],
        fill=fill(C_PENDING), font=Font(bold=True, color=C_PENDING_FT)
    ))
    # Payment Status — Due (red)
    ws.conditional_formatting.add(range_pay, CellIsRule(
        operator="equal", formula=['"Due"'],
        fill=fill(C_DUE), font=Font(bold=True, color=C_DUE_FT)
    ))

    # Delivery — Delivered (blue)
    ws.conditional_formatting.add(range_del, CellIsRule(
        operator="equal", formula=['"Delivered"'],
        fill=fill(C_DEL_OK), font=Font(bold=True, color=C_DEL_OK_FT)
    ))
    # Delivery — Pending (light green)
    ws.conditional_formatting.add(range_del, CellIsRule(
        operator="equal", formula=['"Pending"'],
        fill=fill(C_DEL_PEN), font=Font(bold=True, color="375623")
    ))

    # Overdue rows: Due Date is in the past AND Payment != Paid → highlight full row
    today_serial = (datetime.date.today() - datetime.date(1900, 1, 1)).days + 2
    ws.conditional_formatting.add(full_row, FormulaRule(
        formula=[f'AND($M2<>"", $M2<TODAY(), $J2<>"Paid")'],
        fill=fill("FFE0E0"),
        font=Font(color="9C0006")
    ))

    # ── Sample data ────────────────────────────────────────────────────────────
    today = datetime.date.today()
    d = lambda days: today + datetime.timedelta(days=days)

    # Columns: TxnID, Date, Type, Buyer, Seller, Product, Qty, UnitPrice,
    #          [Total-skip], PayStatus, PaidAmt, [BalDue-skip],
    #          DueDate, DelStatus, ExpDelivery, Notes
    samples = [
        ("TXN-0001", d(-45), "Buy",  "ABC Enterprises",      "XYZ Traders",          "Office Chairs",         10, 2500,   "Paid",    25000, d(-15), "Delivered", d(-30), "Bulk order for new office wing"),
        ("TXN-0002", d(-30), "Sell", "Sunrise Retail Co.",   "ABC Enterprises",      "LED Strip Lights",      50,  350,   "Paid",    17500, d(-10), "Delivered", d(-20), "Monthly supply contract"),
        ("TXN-0003", d(-20), "Buy",  "ABC Enterprises",      "Metro Supplies Ltd.",  "A4 Paper (Box)",       100,  480,   "Paid",    48000, d(-5),  "Delivered", d(-12), "Stationery restock Q1"),
        ("TXN-0004", d(-18), "Sell", "Greenleaf Corp.",      "ABC Enterprises",      "Laptop Bags",           30,  750,   "Paid",    22500, d(-8),  "Delivered", d(-15), "Corporate gift order"),
        ("TXN-0005", d(-15), "Buy",  "ABC Enterprises",      "TechZone Pvt. Ltd.",   "USB-C Hubs",            20, 1200,   "Pending",     0, d(15),  "Pending",   d(10),  "Awaiting delivery confirmation"),
        ("TXN-0006", d(-12), "Sell", "BlueStar Industries",  "ABC Enterprises",      "Printer Cartridges",    60,  320,   "Pending",     0, d(18),  "Delivered", d(-5),  "Invoice sent; awaiting payment"),
        ("TXN-0007", d(-10), "Buy",  "ABC Enterprises",      "Global Furniture Co.", "Standing Desks",         8, 8500,   "Pending", 34000, d(20),  "Pending",   d(25),  "Partial advance paid; balance on delivery"),
        ("TXN-0008", d(-8),  "Sell", "Horizon Traders",      "ABC Enterprises",      "Wireless Keyboards",    25,  950,   "Pending",     0, d(12),  "Pending",   d(8),   "Order confirmed; production in progress"),
        ("TXN-0009", d(-35), "Buy",  "ABC Enterprises",      "Speedy Logistics",     "Packaging Material",   200,  150,   "Due",         0, d(-10), "Delivered", d(-28), "Payment overdue — follow up with accounts"),
        ("TXN-0010", d(-25), "Sell", "Redwood Distributors", "ABC Enterprises",      "Steel Shelving Units",  15, 3200,   "Due",     16000, d(-5),  "Delivered", d(-18), "Partial payment received; balance due"),
        ("TXN-0011", d(-22), "Buy",  "ABC Enterprises",      "Alpha Electronics",    "HDMI Cables (pack)",    80,  220,   "Due",         0, d(-2),  "Delivered", d(-20), "Urgent: payment deadline crossed"),
        ("TXN-0012", d(-5),  "Sell", "Pinnacle Group",       "ABC Enterprises",      "Office Sofas",           5, 12000,  "Pending",     0, d(30),  "Pending",   d(45),  "Custom order; 45-day delivery window"),
        ("TXN-0013", d(-3),  "Buy",  "ABC Enterprises",      "SwiftPrint Solutions", "Business Cards (500)",   4,  800,   "Paid",     3200, d(0),   "Pending",   d(5),   "Branding refresh batch"),
        ("TXN-0014", d(-2),  "Sell", "Nexus Retail Ltd.",    "ABC Enterprises",      "Extension Cords",      100,  180,   "Pending",     0, d(21),  "Pending",   d(14),  "New client — net-30 payment terms"),
        ("TXN-0015", d(-1),  "Buy",  "ABC Enterprises",      "CloudServe India",     "Annual SaaS License",    1, 85000,  "Paid",    85000, d(0),   "Delivered", d(-1),  "Annual ERP renewal — auto-renews next year"),
        ("TXN-0016", d(0),   "Sell", "Eastgate Wholesalers", "ABC Enterprises",      "Projector Screens",     10, 4500,   "Pending",     0, d(15),  "Pending",   d(20),  "New bulk deal — first transaction"),
    ]

    SKIP_COLS = {9, 12}   # Total Amount and Balance Due are formulas
    for data_row_idx, row_data in enumerate(samples, start=2):
        (txn_id, date, txn_type, buyer, seller, product,
         qty, unit_price, pay_status, paid_amt,
         due_date, del_status, exp_delivery, notes) = row_data

        values = [txn_id, date, txn_type, buyer, seller, product,
                  qty, unit_price, None,        # col 9 = formula
                  pay_status, paid_amt, None,   # col 12 = formula
                  due_date, del_status, exp_delivery, notes]

        for col_idx, val in enumerate(values, start=1):
            if col_idx in SKIP_COLS:
                continue
            ws.cell(row=data_row_idx, column=col_idx).value = val


# ══════════════════════════════════════════════════════════════════════════════
#  SHEET 2 — DASHBOARD
# ══════════════════════════════════════════════════════════════════════════════
def build_dashboard_sheet(ws):
    ws.title = "Dashboard"
    ws.sheet_view.showGridLines = False

    # Title
    ws.merge_cells("B2:F2")
    title = ws["B2"]
    title.value     = "Transaction Management — Dashboard"
    title.font      = Font(bold=True, size=16, color=C_HEADER_FG)
    title.fill      = fill(C_HEADER_BG)
    title.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 40

    ws.merge_cells("B3:F3")
    sub = ws["B3"]
    sub.value     = f"Auto-updated from Transactions sheet  |  Generated: {datetime.date.today().strftime('%d %b %Y')}"
    sub.font      = Font(italic=True, size=10, color="666666")
    sub.alignment = Alignment(horizontal="center", vertical="center")

    # ── Section helper ─────────────────────────────────────────────────────────
    def section_header(row, text, bg="2563EB"):
        ws.merge_cells(f"B{row}:F{row}")
        cell = ws[f"B{row}"]
        cell.value     = text
        cell.font      = Font(bold=True, color="FFFFFF", size=11)
        cell.fill      = fill(bg)
        cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws.row_dimensions[row].height = 26

    def kv_row(row, label, formula, val_fmt=None, label_bg="EEF2FF", val_bg="FFFFFF"):
        lc = ws.cell(row=row, column=2, value=label)
        lc.font      = Font(bold=True, size=11)
        lc.fill      = fill(label_bg)
        lc.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        lc.border    = thin_border()
        ws.row_dimensions[row].height = 24

        ws.merge_cells(f"C{row}:F{row}")
        vc = ws.cell(row=row, column=3)
        vc.value     = formula
        vc.font      = Font(bold=True, size=12, color="1E3A5F")
        vc.fill      = fill(val_bg)
        vc.alignment = Alignment(horizontal="center", vertical="center")
        vc.border    = thin_border()
        if val_fmt:
            vc.number_format = val_fmt

    # Column widths
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 30
    for c in ["C", "D", "E", "F"]:
        ws.column_dimensions[c].width = 14

    # ── Overview ───────────────────────────────────────────────────────────────
    section_header(5, "📋  OVERVIEW")
    kv_row(6,  "Total Transactions",
           "=COUNTA(Transactions!A2:A501)-COUNTIF(Transactions!A2:A501,\"\")")
    kv_row(7,  "Total Transaction Value (₹)",
           "=IFERROR(SUM(Transactions!I2:I501),0)", '#,##0.00')
    kv_row(8,  "Buy Transactions",
           "=COUNTIF(Transactions!C2:C501,\"Buy\")")
    kv_row(9,  "Sell Transactions",
           "=COUNTIF(Transactions!C2:C501,\"Sell\")")

    # ── Payment ────────────────────────────────────────────────────────────────
    section_header(11, "💳  PAYMENT STATUS", "276221")
    kv_row(12, "Paid — Count",
           "=COUNTIF(Transactions!J2:J501,\"Paid\")", None, "D4EDDA")
    kv_row(13, "Paid — Total Amount (₹)",
           "=IFERROR(SUMIF(Transactions!J2:J501,\"Paid\",Transactions!I2:I501),0)",
           '#,##0.00', "D4EDDA")
    kv_row(14, "Pending — Count",
           "=COUNTIF(Transactions!J2:J501,\"Pending\")", None, "FFF3CD")
    kv_row(15, "Pending — Total Amount (₹)",
           "=IFERROR(SUMIF(Transactions!J2:J501,\"Pending\",Transactions!I2:I501),0)",
           '#,##0.00', "FFF3CD")
    kv_row(16, "Due / Overdue — Count",
           "=COUNTIF(Transactions!J2:J501,\"Due\")", None, "F8D7DA")
    kv_row(17, "Due / Overdue — Total Amount (₹)",
           "=IFERROR(SUMIF(Transactions!J2:J501,\"Due\",Transactions!I2:I501),0)",
           '#,##0.00', "F8D7DA")
    kv_row(18, "Total Balance Due (₹)",
           "=IFERROR(SUM(Transactions!L2:L501),0)", '#,##0.00', "F8D7DA")

    # ── Delivery ───────────────────────────────────────────────────────────────
    section_header(20, "🚚  DELIVERY STATUS", "7C3AED")
    kv_row(21, "Delivered — Count",
           "=COUNTIF(Transactions!N2:N501,\"Delivered\")", None, "DDEBF7")
    kv_row(22, "Delivery Pending — Count",
           "=COUNTIF(Transactions!N2:N501,\"Pending\")", None, "E2EFDA")

    # ── Footer ─────────────────────────────────────────────────────────────────
    ws.merge_cells("B24:F24")
    note = ws["B24"]
    note.value     = "All values are automatically calculated from the Transactions sheet."
    note.font      = Font(italic=True, size=9, color="999999")
    note.alignment = Alignment(horizontal="center")


# ══════════════════════════════════════════════════════════════════════════════
#  SHEET 3 — INSTRUCTIONS
# ══════════════════════════════════════════════════════════════════════════════
def build_instructions_sheet(ws):
    ws.title = "Instructions"
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 24
    ws.column_dimensions["C"].width = 60

    ws.merge_cells("B1:C1")
    title = ws["B1"]
    title.value     = "How to Use — Transaction Management Sheet"
    title.font      = Font(bold=True, size=15, color=C_HEADER_FG)
    title.fill      = fill(C_HEADER_BG)
    title.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 36

    instructions = [
        ("GETTING STARTED", None),
        ("Step 1", "Go to the 'Transactions' sheet to enter data."),
        ("Step 2", "Fill in each column from left to right for every transaction."),
        ("Step 3", "The 'Total Amount' and 'Balance Due' columns calculate automatically — do not edit them."),
        ("Step 4", "Check the 'Dashboard' sheet for a live summary of all transactions."),

        ("COLUMN GUIDE", None),
        ("Transaction ID", "Enter a unique ID for each transaction (e.g., TXN-0001, TXN-0002, ...)."),
        ("Date",           "The date the transaction was made (DD/MM/YYYY)."),
        ("Type",           "Select 'Buy' if your company is purchasing, 'Sell' if selling. Choose from dropdown."),
        ("Buyer Company",  "Name of the company that is buying goods/services."),
        ("Seller Company", "Name of the company that is selling goods/services."),
        ("Product / Item", "Name of the product or service. For multiple items, separate with commas."),
        ("Quantity",       "Number of units."),
        ("Unit Price (₹)", "Price per unit in Indian Rupees."),
        ("Total Amount",   "AUTO-CALCULATED: Quantity × Unit Price. Do not edit."),
        ("Payment Status", "Select from dropdown: Paid / Pending / Due."),
        ("Amount Paid",    "How much has already been paid (partial or full)."),
        ("Balance Due",    "AUTO-CALCULATED: Total Amount - Amount Paid. Do not edit."),
        ("Payment Due Date","The deadline by which payment must be made."),
        ("Delivery Status","Select from dropdown: Delivered / Pending."),
        ("Expected Delivery","Expected date for product/service delivery."),
        ("Notes",          "Any additional remarks, terms, or conditions."),

        ("COLOUR CODES", None),
        ("Green (Payment)", "Payment has been received — Paid."),
        ("Yellow/Orange",   "Payment is Pending — action needed."),
        ("Red",             "Payment is Due or Overdue — urgent."),
        ("Red row (full)",  "Entire row turns light red when Payment Due Date has passed and payment is not Paid."),
        ("Blue (Delivery)", "Item has been Delivered successfully."),

        ("TIPS", None),
        ("Filtering",      "Use the dropdown arrows in the header row of Transactions to filter by any column."),
        ("Sorting",        "Click any header to sort transactions by that column."),
        ("Adding rows",    "Simply type in the next empty row — all formatting and formulas apply automatically."),
        ("Dashboard",      "The Dashboard sheet updates automatically as you add or edit transactions."),
    ]

    row = 3
    for label, detail in instructions:
        if detail is None:
            # Section header
            ws.merge_cells(f"B{row}:C{row}")
            cell = ws[f"B{row}"]
            cell.value     = label
            cell.font      = Font(bold=True, color="FFFFFF", size=11)
            cell.fill      = fill("2563EB")
            cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
            ws.row_dimensions[row].height = 26
        else:
            lc = ws.cell(row=row, column=2, value=label)
            lc.font      = Font(bold=True, size=10)
            lc.fill      = fill("EEF2FF")
            lc.alignment = Alignment(vertical="top", wrap_text=True)
            lc.border    = thin_border()
            ws.row_dimensions[row].height = 20

            dc = ws.cell(row=row, column=3, value=detail)
            dc.font      = Font(size=10)
            dc.alignment = Alignment(wrap_text=True, vertical="top")
            dc.border    = thin_border()

        row += 1


# ══════════════════════════════════════════════════════════════════════════════
#  MAIN
# ══════════════════════════════════════════════════════════════════════════════
def main():
    wb = Workbook()
    ws_txn  = wb.active
    ws_dash = wb.create_sheet()
    ws_inst = wb.create_sheet()

    build_transactions_sheet(ws_txn)
    build_dashboard_sheet(ws_dash)
    build_instructions_sheet(ws_inst)

    # Sheet tab colors
    ws_txn.sheet_properties.tabColor  = "2563EB"
    ws_dash.sheet_properties.tabColor = "16A34A"
    ws_inst.sheet_properties.tabColor = "7C3AED"

    # Start on Transactions sheet
    wb.active = ws_txn

    wb.save(OUTPUT_FILE)
    print(f"✅  Created: {OUTPUT_FILE}")
    print(f"   • Sheet 1: Transactions  ({DATA_ROWS} rows pre-formatted)")
    print(f"   • Sheet 2: Dashboard     (live summary formulas)")
    print(f"   • Sheet 3: Instructions  (usage guide)")


if __name__ == "__main__":
    main()
