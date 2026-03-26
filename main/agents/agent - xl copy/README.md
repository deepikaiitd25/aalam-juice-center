# Excel Generation Agent (A2A Compatible)

An A2A-compatible agent that autonomously generates structured `.xlsx` spreadsheets based on natural language plans.

## Features
- Parses structural plans to create workbooks and sheets.
- Generates tabular data with labeled columns.
- Inserts computed fields and summary rows using `pandas` and `openpyxl`.

## Setup

1. Install dependencies:
```bash
pip install -e .