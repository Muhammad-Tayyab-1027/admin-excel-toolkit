# Admin Excel Toolkit

An Excel workbook covering everyday office tasks: cleaning a messy contact list, tracking invoices and reporting on what is unpaid. I built it to practise the skills asked for in Irish data entry and junior admin jobs.

**All customers, emails and invoices are made up.**

[Open the workbook](admin_excel_toolkit.xlsx)

## What it does

**Clean Up**
- Fixes spacing and capital letters in names
- Standardises email addresses and phone numbers
- Flags duplicates and emails with no @ sign

**Invoices**
- Tracks 90 sample invoices
- Looks up the customer from a Customer ID, with a drop-down to prevent typing errors
- Calculates VAT, total, due date, status and days overdue

**Report**
- Groups unpaid invoices by age: not yet due, 1-30, 31-60 and 61+ days overdue
- Totals invoices per month, with a chart
- Updates when you change the report date, VAT rate or payment terms

## Excel skills used

`INDEX` + `MATCH`, `IFERROR`, `SUMIFS`, `COUNTIFS`, `COUNTIF`, `TRIM`, `PROPER`, `LOWER`, `SUBSTITUTE`, `FIND`, `ROUND`, data validation, conditional formatting, filters, frozen panes and charts.

## Try it

1. Change the **Report date** on the Report sheet and watch invoices become overdue.
2. Type a Paid Date for an unpaid invoice on the Invoices sheet.
3. Add a messy name to the **Clean Up** sheet by copying a row down.

## Rebuilding the file

`build_workbook.py` creates the workbook with Python and `openpyxl`. Install `openpyxl`, then run:

```
python build_workbook.py
```

Built with AI assistance (Claude Code).
