"""Where the raw exports live, and how each source maps onto one common schema."""
from pathlib import Path

ROOT = Path(__file__).parent
RAW_DIR = ROOT / "data" / "raw"
PRODUCTS_FILE = ROOT / "data" / "products.csv"
OUTPUT_DIR = ROOT / "output"

# Timestamps with a timezone are converted to the business's local time before taking the date.
REPORT_TZ = "America/New_York"

# Every source is mapped onto these columns.
COLUMNS = ["order_id", "order_date", "sku", "qty", "unit_price", "email", "status"]

# Files of one source are read in name order. When the same order line appears in two
# exports, the later file wins, because it has the latest status.
SOURCES = {
    "Shopify": {
        "pattern": "shopify_orders_*.csv",
        "read": {"encoding": "utf-8-sig"},  # Shopify CSVs start with a byte-order mark
        "columns": {
            "Name": "order_id",
            "Created at": "order_date",
            "Lineitem sku": "sku",
            "Lineitem quantity": "qty",
            "Lineitem price": "unit_price",
            "Email": "email",
            "Financial Status": "status",
        },
        # Shopify fills order-level fields on the first line item of each order only.
        "fill_down": ["order_date", "email", "status"],
        "date_format": "%Y-%m-%d %H:%M:%S %z",
        "exclude_status": ["refunded", "voided"],
    },
    "Amazon": {
        "pattern": "amazon_orders_*.txt",
        "read": {"sep": "\t"},  # Amazon order reports are tab-separated
        "columns": {
            "amazon-order-id": "order_id",
            "purchase-date": "order_date",
            "sku": "sku",
            "quantity": "qty",
            "item-price": "unit_price",
            "buyer-email": "email",
            "order-status": "status",
        },
        "date_format": "ISO8601",
        "price_is_line_total": True,  # Amazon's item-price is quantity x unit price
        "exclude_status": ["cancelled"],
    },
    "Store": {
        "pattern": "store_sales_*.xlsx",
        "read": {"skiprows": 2},  # a title row and a blank row sit above the header
        "columns": {
            "Order No": "order_id",
            "Date": "order_date",
            "SKU": "sku",
            "Qty": "qty",
            "Unit Price": "unit_price",
            "Customer Email": "email",
        },
        "date_format": "%d/%m/%Y",  # day first: 03/04/2026 is 3 April, not 4 March
        "exclude_status": [],
    },
}
