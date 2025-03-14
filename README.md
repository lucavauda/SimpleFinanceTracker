# SimpleFinanceTracker
A Python-based financial analysis tool that transforms my bank account statements into Excel reports with matplotlib charts.

## Features

- Imports bank transaction data in European format (CSV/Excel)
- Combines debit and credit entries into a single ledger
- Categorizes and analyzes spending patterns
- Generates Excel reports with:
  - Transaction details and summaries
  - Monthly balance trends
  - Category-based expense analysis
  - Income source breakdown
- Visualizes financial data with matplotlib charts
- Calculates financial metrics and statistics

## Why

I was tired to have my messy account statement and wanted to have a simple python script to help me handle it.
Claude Sonnet helped me develop this. 

## How to use it

Pretty simple: 
First activate the virtual environment `finance_env\Scripts\activate`

Then load the bank account statement in the same folder, change the name of the file accordingly and call the `python3 ./script_finance.py` script and check the Excel spreadsheet at the end.

It works for my statement but I don't think it will work for yours.
I made this for me. :)
