# Amazon Orders & Returns PDF Generator

This Python script automates the process of combining Amazon order summaries with transaction and return data into nicely formatted PDFs. It also exports a CSV of any missing summary PDFs for manual handling.

---

## Features

- Combine **Amazon order summary PDFs** with a **transaction table**.
- Append **returns table** below transactions (two rows per return: info + title).
- Clickable **[LINK TO TRANSACTIONS]** for each order.
- Tables formatted for **Excel-like appearance** with consistent column widths.
- Sort transactions and returns from **oldest to newest**.
- Export CSV of **missing summary PDFs**.
- Test mode: process only a subset of orders for troubleshooting.
- Auto-removes temporary PDFs after merging.

---

## Setup

1. **Clone repository**:
```bash
git clone <repo_url>
cd <repo_folder>
```

2. **Install dependencies**:
```bash
pip install pandas reportlab PyPDF2
```

3. **Update paths in `generate_pdfs.py`** if your files are in different directories:
```python
CSV_PATH = "./reconciliation/..."
SUMMARY_DIR = "./order_summary"
RETURNS_CSV = "./returns/..."
OUTPUT_DIR = "./output"
```

4. **Optional settings** at the top of the script:
```python
TEST_MODE = True      # True = process only first MAX_ORDERS orders
MAX_ORDERS = 5        # Number of orders to process in test mode
```

---

## Usage

Run the script:

```bash
python generate_pdfs.py
```

- The script will process each order in the transactions CSV:
  1. Look for the corresponding summary PDF in `order_summary/`.
  2. Generate a transaction table PDF.
  3. Append any returns for that order below the transactions table.
  4. Merge with the original summary PDF into a single combined PDF.
  5. Output files are saved to `output/`.

- If a summary PDF is **missing**, the order ID will be exported to `output/missing_summary_pdfs.csv`.

---

## Data Handling

### Transactions Table

- Columns included (configurable via `COLUMNS_TO_INCLUDE`):
  ```
  Transaction Date, Payment Reference ID, Transaction Type, Currency, Payment Amount,
  Account Group, Card, Account User, Order Date, Order ID, PO Number, Order Status
  ```
- `Card` = concatenation of `Payment Instrument Type` + `Payment Identifier`.
- Sorted **oldest → newest**.
- Link to Amazon transactions page included at top: `[LINK TO TRANSACTIONS]`

### Returns Table

- Columns included (configurable via `RETURN_INFO_COLUMNS`):
  ```
  Order Date, Order Id, Account Group, PO Number, Account User,
  Return Date, Return Reason, Return Quantity, ASIN
  ```
- Two rows per return:
  1. Info row (all columns except Title)
  2. Title row spanning full table width
- Sorted **oldest → newest**.
- Table styling matches transactions table for consistent appearance.

