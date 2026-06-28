# Credit Card Statement Tools

A desktop GUI application for processing Capital One and Chase credit card statements. Built with Python and tkinter — runs locally, no browser or server required, no data leaves your machine.

---

## What it does

Managing foreign currency transactions and custom budget categories across multiple credit cards means a lot of manual copy-paste work. This tool eliminates that by parsing your PDF statements and downloaded CSVs directly, letting you preview everything before exporting a clean, correctly ordered CSV you can paste straight into your master spreadsheet.

---

## Features

### Capital One Statement → Foreign CSV
Parses a Capital One Venture X PDF statement and extracts every transaction charged in a non-USD currency. Captures the original foreign amount, currency code, and exchange rate as printed on the statement. Handles transactions that split across page boundaries. Rows with no foreign exchange data are included with `-` placeholders so column order never shifts when merging sheets.

### Chase → Foreign CSV
Same as above for Chase Sapphire statements, which use a different PDF format (`NEW ZEALAND DOLLAR` / `EXCHG RATE` pattern). USD transactions are included as purple-highlighted rows with dash placeholders so nothing is omitted. Also handles transactions whose exchange rate line appears on the next page.

### Capital One Remap Categories
Takes the raw CSV export from Capital One's website and adds a `My Category` column mapped from Capital One's default categories to your own personal labels. The mapping table is fully editable in the UI — change any label, click Recalculate, and the table updates instantly without re-uploading.

### Chase Remap Categories
Same as above for Chase CSV exports, with Chase-specific categories (Bills & Utilities, Food & Drink, Groceries, Shopping, Travel, etc.).

### Sort by Master CSV
After parsing a PDF, click **Sort by Master CSV** and select your downloaded bank CSV. The rows reorder to exactly match your master spreadsheet's order — so you can export and paste without any manual resorting. Matching uses date + amount with description similarity as a tiebreaker for duplicate entries on the same day.

---

## Supported Currencies

| Currency | Symbol | Card |
|----------|--------|------|
| NZD | NZ$ | Chase, Capital One |
| EUR | € | Both |
| GBP | £ | Chase |
| DKK | kr | Capital One |
| ISK | kr | Capital One |
| CZK | Kč | Capital One |
| TRY | ₺ | Capital One |
| CHF / CHW | Fr | Capital One |
| CAD | CA$ | Chase |
| AUD | A$ | Chase |
| JPY | ¥ | Chase |

---

## Requirements

- Python 3.8+
- [pdfplumber](https://github.com/jsvine/pdfplumber)

```bash
pip install pdfplumber
```

---

## Usage

### Windows
Double-click `launch.bat`, or run from terminal:

```bash
python capital_one_gui.py
```

### Mac / Linux

```bash
python3 capital_one_gui.py
```

---

## Workflow

### Foreign currency extraction (PDF tabs)

1. Click **Upload PDF Statement** and select your Capital One or Chase PDF
2. Review transactions in the preview table — filter by currency or search by merchant
3. Optionally click **Sort by Master CSV** to align row order with your spreadsheet
4. Click **Export CSV**

### Category remapping (Remap tabs)

1. Click **Upload CSV** and select your raw export from capitalone.com or chase.com
2. Edit any category label directly in the mapping panel, then click **Recalculate**
3. Green rows are matched, amber rows have unrecognized categories
4. Click **Export Remapped CSV**

---

## Category Mappings

### Capital One

| Capital One | Your Label |
|-------------|------------|
| Car Rental | Travel Fund |
| Dining | Food |
| Entertainment | Savings |
| Health Care | Savings |
| Insurance | Medical Insurance |
| Internet | Utilities |
| Merchandise | Remainder |
| Other Travel | Travel Fund |
| Phone/Cable | Utilities |
| Lodging | Travel Fund |
| Gas/Automotive | Utilities |
| Other Services | Travel Fund |
| Payment/Credit | Savings |
| Professional Services | Utilities |

### Chase

| Chase | Your Label |
|-------|------------|
| Bills & Utilities | Utilities |
| Food & Drink | Food |
| Gas | Utilities |
| Groceries | Food |
| Home | Savings |
| Shopping | Remainder |
| Travel | Travel Fund |

All mappings are editable at runtime — changes apply immediately without restarting.
