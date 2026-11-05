"""Builds admin_excel_toolkit.xlsx. All customers and invoices are made up (fixed random seed)."""
import random
from datetime import date, timedelta

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter as col
from openpyxl.worksheet.datavalidation import DataValidation

random.seed(2006)
FONT = "Arial"
HEAD = PatternFill("solid", fgColor="1F3864")
INPUT = PatternFill("solid", fgColor="FFF2CC")
RED = PatternFill("solid", bgColor="F8CBAD")
AMBER = PatternFill("solid", bgColor="FFE699")
GREEN = PatternFill("solid", bgColor="C6E0B4")


def header(ws, row, names, widths):
    for i, name in enumerate(names, 1):
        c = ws.cell(row=row, column=i, value=name)
        c.font = Font(name=FONT, bold=True, color="FFFFFF")
        c.fill = HEAD
        c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        ws.column_dimensions[col(i)].width = widths[i - 1]
    ws.row_dimensions[row].height = 32


wb = Workbook()
g = wb.active
g.title = "Guide"
g.column_dimensions["A"].width = 110
for i, t in enumerate([
    "Admin Excel Toolkit", "",
    "Portfolio project. All customers, emails and invoices are made up.", "",
    "Clean Up: a messy contact list on the left, cleaned by formulas on the right (TRIM, PROPER, LOWER, COUNTIF).",
    "Customers: the customer list that invoices look up.",
    "Invoices: yellow cells are typed in. INDEX/MATCH finds the customer; VAT, total, due date and status are formulas.",
    "Report: unpaid invoices by age and monthly totals (SUMIFS). Change the report date to see the figures move.",
    "", "Assumptions are on the Report sheet: report date, VAT rate and payment terms.",
    "VAT of 23% is the Irish standard rate, used here as an example.",
], 1):
    g.cell(row=i, column=1, value=t)
g["A1"].font = Font(name=FONT, bold=True, size=14)

# ---------- Clean Up ----------
cu = wb.create_sheet("Clean Up")
header(cu, 1, ["Raw Name", "Raw Email", "Raw Phone", "Clean Name", "Clean Email", "Email OK?", "Clean Phone", "Duplicate?"],
       [26, 30, 16, 24, 30, 9, 15, 11])
first = ["aoife", "CONOR", "Niamh", "sean", "Maria", "PIOTR", "fatima", "James", "ANNA", "liam", "Grace", "omar"]
last = ["byrne", "O'Neill", "KELLY", "murphy", "Silva", "NOWAK", "khan", "Walsh", "RYAN", "doyle", "Nolan", "Ali"]
raw = []
for f, l in zip(first, last):
    name = random.choice(["{} {}", "  {} {}", "{}  {}", "{} {}  "]).format(f, l)
    mail = f"{f.lower()}.{l.lower().replace(chr(39), '')}@example.ie"
    mail = random.choice([mail, mail.upper(), f"  {mail}", mail])
    # 555-style digits so that no real person's number is used.
    num = f"555{random.randint(1000, 9999)}"
    phone = random.choice([f"087 {num}", f"087-{num}", f"087{num}"])
    raw.append((name, mail, phone))
raw.insert(5, (raw[1][0].upper(), raw[1][1], raw[1][2]))  # a duplicate
raw.insert(9, ("Tom Hayes", "tom.hayes.example.ie", "086 5550100"))  # a bad email
for r, row in enumerate(raw, 2):
    for c, v in enumerate(row, 1):
        cu.cell(row=r, column=c, value=v)
    cu[f"D{r}"] = f"=PROPER(TRIM(A{r}))"
    cu[f"E{r}"] = f"=LOWER(TRIM(B{r}))"
    cu[f"F{r}"] = f'=IF(ISNUMBER(FIND("@",E{r})),"Yes","Check")'
    cu[f"G{r}"] = f'=SUBSTITUTE(SUBSTITUTE(C{r}," ",""),"-","")'
    cu[f"H{r}"] = f'=IF(COUNTIF($E$2:E{r},E{r})>1,"Duplicate","")'
CU = len(raw) + 1
cu.conditional_formatting.add(f"F2:F{CU}", CellIsRule(operator="equal", formula=['"Check"'], fill=RED))
cu.conditional_formatting.add(f"H2:H{CU}", CellIsRule(operator="equal", formula=['"Duplicate"'], fill=AMBER))
cu.freeze_panes = "A2"

# ---------- Customers ----------
cs = wb.create_sheet("Customers")
header(cs, 1, ["Customer ID", "Customer", "County"], [12, 30, 12])
names = [("Liffey Office Supplies", "Dublin"), ("Harbour Cafe Group", "Dublin"), ("Greenway Garden Centre", "Wicklow"),
         ("Shannon Print Works", "Clare"), ("Northside Hardware", "Dublin"), ("Atlantic Bike Shop", "Galway"),
         ("Kilkenny Crafts", "Kilkenny"), ("Boyne Valley Foods", "Meath"), ("Cork Tile Centre", "Cork"),
         ("Phoenix Pharmacy", "Dublin")]
for r, (n, co) in enumerate(names, 2):
    cs[f"A{r}"] = f"C{100 + r - 1}"
    cs[f"B{r}"] = n
    cs[f"C{r}"] = co
CS = len(names) + 1

# ---------- Invoices ----------
REPORT_DATE, VAT, TERMS = "Report!$B$3", "Report!$B$4", "Report!$B$5"
iv = wb.create_sheet("Invoices")
header(iv, 1, ["Invoice No", "Customer ID", "Customer", "Invoice Date", "Net (EUR)", "VAT (EUR)", "Total (EUR)",
               "Due Date", "Paid Date", "Status", "Days Overdue"], [11, 10, 28, 12, 11, 10, 11, 12, 12, 10, 9])
start = date(2026, 4, 1)
report = date(2026, 10, 31)
inv = sorted(start + timedelta(days=random.randrange((report - start).days + 1)) for _ in range(90))
for r, d in enumerate(inv, 2):
    iv[f"A{r}"] = f"INV-{2000 + r - 2}"
    iv[f"B{r}"] = f"C{random.randint(101, 100 + len(names))}"
    iv[f"C{r}"] = f'=IFERROR(INDEX(Customers!$B$2:$B${CS},MATCH(B{r},Customers!$A$2:$A${CS},0)),"Unknown ID")'
    iv[f"D{r}"] = d
    iv[f"E{r}"] = round(random.uniform(45, 2400), 2)
    iv[f"F{r}"] = f"=ROUND(E{r}*{VAT},2)"
    iv[f"G{r}"] = f"=E{r}+F{r}"
    iv[f"H{r}"] = f"=D{r}+{TERMS}"
    pay_in = random.randint(5, 55)
    if random.random() < 0.8 and d + timedelta(days=pay_in) <= report:
        iv[f"I{r}"] = d + timedelta(days=pay_in)
    iv[f"J{r}"] = f'=IF(I{r}<>"","Paid",IF(H{r}<{REPORT_DATE},"Overdue","Open"))'
    iv[f"K{r}"] = f'=IF(J{r}="Overdue",{REPORT_DATE}-H{r},0)'
    for c in "DHI":
        iv[f"{c}{r}"].number_format = "yyyy-mm-dd"
    for c in "EFG":
        iv[f"{c}{r}"].number_format = "#,##0.00"
    for c in "BDEI":
        iv[f"{c}{r}"].fill = INPUT
IN = len(inv) + 1
dv = DataValidation(type="list", formula1=f"Customers!$A$2:$A${CS}", allow_blank=False,
                    showErrorMessage=True, errorTitle="Unknown customer", error="Pick a Customer ID from the list.")
iv.add_data_validation(dv)
dv.add(f"B2:B{IN}")
iv.freeze_panes = "B2"
iv.auto_filter.ref = f"A1:K{IN}"
iv.conditional_formatting.add(f"J2:J{IN}", CellIsRule(operator="equal", formula=['"Overdue"'], fill=RED))
iv.conditional_formatting.add(f"J2:J{IN}", CellIsRule(operator="equal", formula=['"Open"'], fill=AMBER))
iv.conditional_formatting.add(f"J2:J{IN}", CellIsRule(operator="equal", formula=['"Paid"'], fill=GREEN))

# ---------- Report ----------
rp = wb.create_sheet("Report")
rp.column_dimensions["A"].width = 30
for c in "BCD":
    rp.column_dimensions[c].width = 15
rp["A1"] = "Invoice Report (sample data)"
rp["A1"].font = Font(name=FONT, bold=True, size=14)
for r, (label, v, fmt) in enumerate([("Report date", report, "yyyy-mm-dd"), ("VAT rate", 0.23, "0%"),
                                     ("Payment terms (days)", 30, "0")], 3):
    rp[f"A{r}"] = label
    rp[f"B{r}"] = v
    rp[f"B{r}"].number_format = fmt
    rp[f"B{r}"].fill = INPUT
    rp[f"B{r}"].font = Font(name=FONT, color="0000FF")


def block_header(row, labels):
    for i, n in enumerate(labels, 1):
        c = rp.cell(row=row, column=i, value=n)
        c.font = Font(name=FONT, bold=True, color="FFFFFF")
        c.fill = HEAD


G, J, K, D = (f"Invoices!${x}$2:${x}${IN}" for x in "GJKD")
block_header(7, ["Unpaid by age", "Invoices", "Total (EUR)"])
buckets = [("Not yet due", '"Open"', None, None), ("1-30 days overdue", '"Overdue"', 1, 30),
           ("31-60 days overdue", '"Overdue"', 31, 60), ("61+ days overdue", '"Overdue"', 61, 100000)]
for r, (label, status, lo, hi) in enumerate(buckets, 8):
    rp[f"A{r}"] = label
    if lo is None:
        rp[f"B{r}"] = f"=COUNTIFS({J},{status})"
        rp[f"C{r}"] = f"=SUMIFS({G},{J},{status})"
    else:
        rp[f"B{r}"] = f'=COUNTIFS({J},{status},{K},">={lo}",{K},"<={hi}")'
        rp[f"C{r}"] = f'=SUMIFS({G},{J},{status},{K},">={lo}",{K},"<={hi}")'
    rp[f"C{r}"].number_format = "#,##0.00"
rp["A12"] = "Total unpaid"
rp["B12"] = "=SUM(B8:B11)"
rp["C12"] = "=SUM(C8:C11)"
rp["C12"].number_format = "#,##0.00"
for c in ("A12", "B12", "C12"):
    rp[c].font = Font(name=FONT, bold=True)
block_header(14, ["Month", "Invoices", "Invoiced (EUR)"])
m = date(2026, 4, 1)
for r in range(15, 22):
    rp[f"A{r}"] = m
    rp[f"A{r}"].number_format = "mmm yyyy"
    rp[f"A{r}"].alignment = Alignment(horizontal="left")
    nxt = date(m.year + (m.month == 12), m.month % 12 + 1, 1)
    rp[f"B{r}"] = f'=COUNTIFS({D},">="&A{r},{D},"<"&DATE(YEAR(A{r}),MONTH(A{r})+1,1))'
    rp[f"C{r}"] = f'=SUMIFS({G},{D},">="&A{r},{D},"<"&DATE(YEAR(A{r}),MONTH(A{r})+1,1))'
    rp[f"C{r}"].number_format = "#,##0.00"
    m = nxt
rp["A22"] = "Total"
rp["B22"] = "=SUM(B15:B21)"
rp["C22"] = "=SUM(C15:C21)"
rp["C22"].number_format = "#,##0.00"
for c in ("A22", "B22", "C22"):
    rp[c].font = Font(name=FONT, bold=True)
ch = BarChart()
ch.title = "Invoiced per month (EUR)"
ch.legend = None
ch.add_data(Reference(rp, min_col=3, min_row=14, max_row=21), titles_from_data=True)
ch.set_categories(Reference(rp, min_col=1, min_row=15, max_row=21))
ch.series[0].graphicalProperties.solidFill = "1F3864"
ch.height, ch.width = 8, 15
rp.add_chart(ch, "E3")

for ws in wb.worksheets:
    for row in ws.iter_rows():
        for c in row:
            if c.value is not None and c.font.name != FONT:
                c.font = Font(name=FONT, bold=c.font.bold, color=c.font.color)
wb.save("admin_excel_toolkit.xlsx")
print("invoices", len(inv), "contacts", len(raw))
