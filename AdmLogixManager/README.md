# Office Manager (Adm Logix Rental Equipment)

A simple Windows desktop program to replace the daily Excel work:

| Screen | What it does |
|---|---|
| **Dashboard** | Bank balance, cash in hand, today's expense, month income/expense, salary due, client receivable, 6-month chart, expense-by-category chart, latest entries, reminders (Iqama expiry, unpaid salaries, old unpaid invoices) |
| **Daily Book** | Day-by-day income & expense for **Bank** or **Cash**. Opening (B/F) → Income → Expense → Closing is calculated automatically. One tick adds *Transfer Fee + VAT* lines. Descriptions you typed before are suggested and the category is filled for you. Export **Day PDF** (same layout as the old sheet), **Day Excel**, or the **Month Workbook** (one sheet per day, like `02-Bank Expense-2026.xlsx`). |
| **Salaries** | Generate the monthly salary sheet from active employees, add allowance / overtime / deduction / advance, record full or partial payments (each payment goes into the Daily Book automatically), *Pay All Remaining*, Salary Sheet PDF/Excel, Payslip PDF. |
| **Employees** | ID No, Iqama No, name, basic salary, status (Active / On Leave / Removed / Final Exit), Iqama expiry, phone… plus your own custom fields. Payment history per employee. Nominal Roll PDF/Excel. |
| **Clients & Invoices** | ID-wise company list (current / previous clients), invoices per client, record receipts (goes into the Daily Book as income), Statement of Account PDF/Excel with *amount in words* (like the Spectra sheet). |
| **Reports** | Any date range, filter by type / category / account / text. Totals plus a category breakdown. PDF/Excel. |
| **Settings & Import** | Company name and report titles, opening balances, default transfer fee and VAT %, **custom fields** (add new fields to Employees, Clients or Daily entries any time), categories, **Import Excel**, Backup / Restore. |

Data you enter once (employees, clients, categories, settings) is stored once and reused everywhere.

## Install and run (Windows)

1. Install Python 3.10 or newer from https://www.python.org/downloads/. During setup, tick **"Add python.exe to PATH"**.
2. Double-click **`run.bat`**. The first time, it installs two small libraries (`openpyxl`, `reportlab`) and then opens the program.

To get a normal `.exe` that needs no Python, double-click **`build_exe.bat`**. Then use `dist\OfficeManager\OfficeManager.exe`. You can copy that whole folder to any PC.

## First use: import your existing files

Go to **Settings & Import → Import Excel File…** and import, in this order:

1. `NOMINAL_ROLL_OF_EMPLOYEES_SEPTEMBER.xlsx` (sheet **UPDATED BY AHSAN**, salary column **by ahsan**)
2. `ID_WISE_COMPANIES_LIST.xlsx` (sheet **ID WISE CPMPANIES**, then again with sheet **Client List (2)** to fill Current/Previous)
3. `02-Bank_Expense-2026.xlsx` (all day sheets at once; the opening balance is set from the first sheet)
4. Any client statement, such as `Spectra_Division_Contracting_Company.xlsx`

The program detects the file type itself. The bank workbook import checks every day's balance against the old sheet. With the current file, all 168 days match, and the balance after 10-Oct-2026 is **22,968.08**, the same as the printed PDF.

## Where things are saved

* `data\office.db`: **all your data in one file**. Copy it to move to another computer.
* `exports\`: every PDF/Excel you generate (it opens automatically).
* `backups\`: an automatic backup each time you close the program (the last 30 are kept). Use **Restore Backup…** to go back.

## Tips

* Amount boxes accept sums: typing `400+120` saves 520.
* Dates can be typed as `14-Oct-2026`, `14/10/2026` or picked with the 📅 button.
* In the Daily Book, click a row to edit it, press **New** for a fresh entry, and press **Ctrl+S** to save.
* Don't delete employees who left. Set their status to *Removed* / *Final Exit* so their history stays.
