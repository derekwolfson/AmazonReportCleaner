import os
import pandas as pd
from reportlab.lib.pagesizes import LETTER, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from PyPDF2 import PdfReader, PdfWriter

# ================== SETTINGS ==================
TEST_MODE = False         # True = only process first MAX_ORDERS orders
MAX_ORDERS = 5           # Number of orders to process in test mode

CSV_PATH = "./reconciliation/reconciliation_from_20250101_to_20251210_20251210_0902.csv"
SUMMARY_DIR = "./order_summary"
RETURNS_CSV = "./returns/returns_from_20250101_to_20251210_20251210_0954.csv"
OUTPUT_DIR = "./output"
MISSING_CSV_PATH = os.path.join(OUTPUT_DIR, "missing_summary_pdfs.csv")

COLUMNS_TO_INCLUDE = [
    "Transaction Date",
    "Payment Reference ID",
    "Transaction Type",
    "Currency",
    "Payment Amount",
    "Account Group",
    "Card",  # concatenated Payment Instrument Type + Payment Identifier
    "Account User",
    "Order Date",
    "Order ID",
    "PO Number",
    "Order Status"
]

URL_TEMPLATE = "https://www.amazon.com/cpe/yourpayments/transactions?transactionTag={order_id}&ref_=ppx_od_dt_b_yt"

RETURN_INFO_COLUMNS = [
    "Order Date", "Order Id", "Account Group", "PO Number", "Account User",
    "Return Date", "Return Reason", "Return Quantity", "ASIN"
]

os.makedirs(OUTPUT_DIR, exist_ok=True)

# ---------------- PDF Helpers ----------------
def find_summary_pdf(order_id):
    for filename in os.listdir(SUMMARY_DIR):
        if order_id in filename:
            return os.path.join(SUMMARY_DIR, filename)
    return None

def make_transaction_page(order_id, rows, returns_df, output_pdf):
    rows = rows.copy()
    rows['Card'] = rows['Payment Instrument Type'].fillna('') + ' ' + rows['Payment Identifier'].fillna('')

    styles = getSampleStyleSheet()
    header_style = ParagraphStyle(name='header', fontSize=6, leading=6, alignment=0)

    # ---------------- Transactions Table ----------------
    data_trans = [[Paragraph(col, header_style) for col in COLUMNS_TO_INCLUDE]]
    for _, r in rows.iterrows():
        data_trans.append([r.get(col, '') for col in COLUMNS_TO_INCLUDE])

    page_width, page_height = landscape(LETTER)
    margin = 40
    usable_width = page_width - 2 * margin

    # Column widths for transactions
    num_cols_trans = len(COLUMNS_TO_INCLUDE)
    base_width_trans = usable_width / num_cols_trans
    col_widths_trans = [base_width_trans] * num_cols_trans

    # Adjust specific columns
    idx_transaction_date = COLUMNS_TO_INCLUDE.index("Transaction Date")
    idx_payment_ref = COLUMNS_TO_INCLUDE.index("Payment Reference ID")
    idx_order_id = COLUMNS_TO_INCLUDE.index("Order ID")
    idx_order_status = COLUMNS_TO_INCLUDE.index("Order Status")
    col_widths_trans[idx_payment_ref] += 36
    col_widths_trans[idx_order_id] += 36
    col_widths_trans[idx_transaction_date] -= 18
    col_widths_trans[idx_order_status] -= 18

    pdf = SimpleDocTemplate(output_pdf, pagesize=landscape(LETTER))
    story = []

    # Title + link
    story.append(Paragraph(f"Transactions for Order ID {order_id}", styles["Heading2"]))
    story.append(Spacer(1, 4))
    url = URL_TEMPLATE.format(order_id=order_id)
    story.append(Paragraph(f'<link href="{url}">[LINK TO TRANSACTIONS]</link>',
                            ParagraphStyle('link', fontSize=6, textColor=colors.blue, leftIndent=0)))
    story.append(Spacer(1, 6))

    table_trans = Table(data_trans, colWidths=col_widths_trans)
    table_trans.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.grey),
        ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
        ('ALIGN', (0,0), (-1,-1), 'LEFT'),
        ('VALIGN', (0,0), (-1,0), 'TOP'),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTNAME', (0,1), (-1,-1), 'Helvetica'),
        ('FONTSIZE', (0,0), (-1,-1), 6),
        ('TOPPADDING', (0,0), (-1,-1), 1),
        ('BOTTOMPADDING', (0,0), (-1,-1), 1),
        ('LEFTPADDING', (0,0), (-1,-1), 2),
        ('RIGHTPADDING', (0,0), (-1,-1), 2),
        ('GRID', (0,0), (-1,-1), 0.25, colors.black),
    ]))
    story.append(table_trans)
    story.append(Spacer(1, 6))

    # ---------------- Returns Table ----------------
    order_returns = returns_df[returns_df['Order Id'] == order_id]
    if not order_returns.empty:
        story.append(Paragraph(f"Returns for Order ID {order_id}", styles["Heading2"]))
        story.append(Spacer(1, 4))

        # Build data: header row + two rows per return
        data_returns = [[Paragraph(col, header_style) for col in RETURN_INFO_COLUMNS]]
        for _, r in order_returns.iterrows():
            info_row = [r.get(col, '') for col in RETURN_INFO_COLUMNS]
            data_returns.append(info_row)
            title_row = [r.get("Title", '')] + [''] * (len(RETURN_INFO_COLUMNS)-1)
            data_returns.append(title_row)

        # Scale returns columns to match transactions table total width
        total_width_trans = sum(col_widths_trans)
        col_widths_returns = [total_width_trans / len(RETURN_INFO_COLUMNS)] * len(RETURN_INFO_COLUMNS)

        table_returns = Table(data_returns, colWidths=col_widths_returns)
        style_returns = TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.grey),
            ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
            ('ALIGN', (0,0), (-1,-1), 'LEFT'),
            ('VALIGN', (0,0), (-1,0), 'TOP'),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('FONTNAME', (0,1), (-1,-1), 'Helvetica'),
            ('FONTSIZE', (0,0), (-1,-1), 6),
            ('TOPPADDING', (0,0), (-1,-1), 1),
            ('BOTTOMPADDING', (0,0), (-1,-1), 1),
            ('LEFTPADDING', (0,0), (-1,-1), 2),
            ('RIGHTPADDING', (0,0), (-1,-1), 2),
            ('GRID', (0,0), (-1,-1), 0.25, colors.black),
        ])
        # Span title rows
        for i in range(2, len(data_returns), 2):
            style_returns.add('SPAN', (0,i), (-1,i))
            style_returns.add('FONTNAME', (0,i), (-1,i), 'Helvetica-Bold')
        table_returns.setStyle(style_returns)
        story.append(table_returns)
        story.append(Spacer(1, 6))

    pdf.build(story)

def append_pdfs(original, extra_page, output):
    writer = PdfWriter()
    orig_reader = PdfReader(original)
    extra_reader = PdfReader(extra_page)
    for page in orig_reader.pages:
        writer.add_page(page)
    for page in extra_reader.pages:
        writer.add_page(page)
    with open(output, "wb") as f:
        writer.write(f)

# ---------------- Main ----------------
def main():
    df = pd.read_csv(CSV_PATH, dtype=str).fillna("")
    returns_df = pd.read_csv(RETURNS_CSV, dtype=str).fillna("")

    # ---------------- Sort by date ----------------
    df['Transaction Date'] = pd.to_datetime(df['Transaction Date'], errors='coerce')
    df = df.sort_values('Transaction Date')
    df['Transaction Date'] = df['Transaction Date'].dt.strftime('%Y-%m-%d')

    returns_df['Return Date'] = pd.to_datetime(returns_df['Return Date'], errors='coerce')
    returns_df = returns_df.sort_values('Return Date')
    returns_df['Return Date'] = returns_df['Return Date'].dt.strftime('%Y-%m-%d')

    groups = df.groupby("Order ID")

    missing_orders = []

    for i, (order_id, rows) in enumerate(groups):
        if TEST_MODE and i >= MAX_ORDERS:
            break

        print(f"Processing Order ID {order_id}...")
        summary_pdf = find_summary_pdf(order_id)
        if not summary_pdf:
            print(f"  !! No summary PDF found for {order_id}")
            missing_orders.append(order_id)
            continue

        temp_pdf = os.path.join(OUTPUT_DIR, f"{order_id}_transactions.pdf")
        make_transaction_page(order_id, rows, returns_df, temp_pdf)

        out_pdf = os.path.join(OUTPUT_DIR, f"Amazon_{order_id}_combined.pdf")
        append_pdfs(summary_pdf, temp_pdf, out_pdf)

        os.remove(temp_pdf)
        print(f"  ✔ Output written: {out_pdf}")

    if missing_orders:
        pd.DataFrame({"Order ID": missing_orders}).to_csv(MISSING_CSV_PATH, index=False)
        print(f"\n⚠ Missing summary PDFs exported to: {MISSING_CSV_PATH}")
    else:
        print("\nAll summary PDFs found.")

if __name__ == "__main__":
    main()
