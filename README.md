# Monthly Commission Payment Tracker

Monthly version of the commissions reconciliation pipeline (see [commissions-x-sales](https://github.com/edgar0garcs38/commissions-x-sales)). It checks which insurance policies were paid each month and matches them against the issuance/renewal base.

## What it does
1. **Paid status:** parses currency values in EU and US formats (`1.234,56`, `(123)` for negatives, `$`) and classifies each payment
2. **Cancellation filter:** removes policies that appear in the cancellation report
3. **Issuance/renewal matching:** normalizes policy numbers and matches the renewals base against the main dataset
4. **Global summary:** aggregates payments

Output is one Excel workbook with the sheets `Paid Status`, `Match Results` and `Payment Global Summary`.

## Tech stack
Python · pandas · regex

## Run it
```bash
pip install -r requirements.txt
python main.py
```
Expects `main.xlsx`, `cancelreport.xlsx` and `Renovaciones.xlsx` in the same folder.

> **Note:** The input files contain confidential client and company data, so they are not included in this repository. Only the source code is published.
