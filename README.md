# Expense Coder: how to use it

A free tool that turns bank statement CSVs and scanned bills into an Excel sheet of coded expenses, plus files you can import into Sage or QuickBooks.

Everything it uses is free and open-source: Python, pandas, openpyxl, and RapidOCR to read the scanned bills. Nothing is sent to the internet. The OCR model downloads once, the first time you run it.

## One-time install

```
py -m pip install --user -r requirements.txt
```

## Try it with the sample data

The `sample/` folder holds a made-up trial balance, two bank CSVs and five bill images. Copy them into `input/`, then run setup and run:

```
xcopy /E /I sample input
py expense_coder.py setup
py expense_coder.py run
```

**Client data stays on your PC.** The `.gitignore` blocks `input/`, `output/`, `rules.xlsx` and any CSV or Excel file outside `sample/`, so real data is never uploaded to GitHub.

## Folders

```
input/trial_balance.csv     <- your trial balance (CSV or XLSX)
input/bank/                 <- bank CSVs. Start each file name with the bank's account code,
                               e.g. 1200_barclays_oct.csv, 1210_hsbc_oct.csv
input/bills/                <- scanned bills (JPG, PNG, PDF)
rules.xlsx                  <- the tool's "brain" (you can edit it)
output/coded_expenses.xlsx  <- the result
output/sage50_import.csv    <- ready for Sage 50 import (bank payments)
output/journal_import.csv   <- double-entry journal for Sage Accounting / QuickBooks
```

## Monthly steps

1. **First time per client:** run `py expense_coder.py setup`. It reads the trial balance and creates `rules.xlsx`. Open it and check:
   - The **Type** column (Expense / Bank / Other).
   - The **Keywords** for each expense account. These are pre-filled, and you can add your own.
2. Put the bank CSVs and bill scans in the folders, then run `py expense_coder.py run`.
3. Open `output/coded_expenses.xlsx`:
   - **Columns A–E:** Date | Expense code | Bank code | Narrative | Amount
   - **Yellow rows:** low confidence, so check the code.
   - **Red rows:** a bill with no matching bank payment (paid in cash or by card?).
4. Fix any wrong code in column B, save, then run `py expense_coder.py learn`. The tool remembers that vendor, so it gets it right next month.
5. Import `output/sage50_import.csv` (Sage 50) or `output/journal_import.csv` (Sage Accounting or QuickBooks).

## How it decides the code

1. **Vendor rules** (rules.xlsx > Vendors): your past corrections. Confidence: High.
2. **Keywords** found in the bank description and the bill text. Confidence: Medium, or Low if two accounts tie.
3. **No match:** the row goes to the suspense code (9999) for you to choose. Confidence: Low.

## Check before importing

- **Column headings:** download your software's own import template and line the headings up. Sage 50 uses audit-trail import with Type BP. QuickBooks journals need account *names* unless account numbers are switched on.
- **Tax code:** the Sage file uses tax code T9 (no VAT) for every row. Change `DefaultTaxCode` in rules.xlsx > Settings, or edit the rows that carry VAT.
- **Bank feed:** don't import these if the same bank feed is also live in the software, or every payment will be counted twice.
