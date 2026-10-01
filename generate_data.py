"""Generate messy sample exports that copy real Shopify, Amazon and till formats.

The data is synthetic (fixed random seed), so the repository contains no real customer data.
"""
import csv
import random
from datetime import datetime, timezone
from zoneinfo import ZoneInfo

from openpyxl import Workbook

import config

# sku, product, category, price, how often it sells
PRODUCTS = [
    ("MUG-001", "Café Mug 350ml", "Kitchen", 14.00, 10),
    ("MUG-002", "Travel Mug", "Kitchen", 24.00, 8),
    ("POT-020", "Glass Teapot", "Kitchen", 32.00, 5),
    ("ESP-090", "Espresso Machine", "Kitchen", 1249.00, 0.15),
    ("TEA-010", "Green Tea 100g", "Tea", 9.50, 12),
    ("TEA-011", "Earl Grey 100g", "Tea", 8.50, 12),
    ("TEA-012", "Matcha 30g", "Tea", 19.00, 7),
    ("BAG-030", "Canvas Tote", "Accessories", 18.00, 6),
    ("NB-040", "Notebook A5", "Stationery", 7.50, 9),
    ("PEN-041", "Gel Pen 3-pack", "Stationery", 4.50, 9),
    ("CND-050", "Soy Candle", "Home", 16.00, 7),
    ("PLT-060", "Ceramic Planter", "Home", 22.00, 5),
    ("GFT-070", "Gift Box", "Gifts", 45.00, 4),
]
NY = ZoneInfo(config.REPORT_TZ)
START, END = datetime(2026, 1, 1), datetime(2026, 6, 30, 21, 0)


def order_time(rng):
    return START + (END - START) * rng.random() ** 0.8  # sales grow over the half-year


def order_lines(rng):
    """One to three different products, mostly quantity 1."""
    n = rng.choice([1, 1, 1, 2, 2, 3])
    picked = {}
    while len(picked) < n:
        p = rng.choices(PRODUCTS, weights=[p[4] for p in PRODUCTS])[0]
        picked[p[0]] = (p, rng.choices([1, 2, 3], weights=[75, 20, 5])[0])
    return list(picked.values())


def email(rng):
    e = f"customer{rng.randint(1, 400)}@example.com"
    r = rng.random()
    return e.upper() if r < 0.07 else f" {e.title()}" if r < 0.15 else e


def shopify(rng):
    orders = sorted(
        (dict(time=order_time(rng).replace(tzinfo=NY), email=email(rng), lines=order_lines(rng),
              status=rng.choices(["paid", "refunded"], weights=[95, 5])[0]) for _ in range(420)),
        key=lambda o: o["time"])
    files = {"shopify_orders_2026-01_to_03.csv": [], "shopify_orders_2026-03_to_06.csv": []}
    for n, o in enumerate(orders):
        o["id"] = f"#{1001 + n}"
        month = o["time"].month
        if month <= 3:
            status = o["status"]
            if month == 3:  # the first export ran on 31 March, before some payments and refunds
                status = "paid" if status == "refunded" else rng.choices(["paid", "pending"], [85, 15])[0]
            files["shopify_orders_2026-01_to_03.csv"].append((o, status))
        if month >= 3:
            files["shopify_orders_2026-03_to_06.csv"].append((o, o["status"]))

    typo_orders = {"#1060", "#1333"}
    header = ["Name", "Email", "Financial Status", "Created at", "Currency", "Lineitem quantity",
              "Lineitem name", "Lineitem price", "Lineitem sku"]
    for name, rows in files.items():
        with open(config.RAW_DIR / name, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow(header)
            for o, status in rows:
                for i, (p, qty) in enumerate(o["lines"]):
                    sku = "MUG001" if o["id"] in typo_orders and i == 0 else p[0]
                    if rng.random() < 0.04:
                        sku = sku.lower() + " "
                    first = i == 0  # order-level fields only on the first line item
                    w.writerow([o["id"], o["email"] if first else "", status if first else "",
                                o["time"].strftime("%Y-%m-%d %H:%M:%S %z") if first else "",
                                "USD", qty, p[1], f"{p[3]:.2f}", sku])


def amazon(rng):
    rows = []
    for _ in range(330):
        t = order_time(rng).replace(tzinfo=NY).astimezone(timezone.utc)
        oid = f"{rng.randint(111, 114)}-{rng.randint(10**6, 10**7 - 1)}-{rng.randint(10**6, 10**7 - 1)}"
        status = rng.choices(["Shipped", "Cancelled", "Pending"], weights=[92, 5, 3])[0]
        buyer = f"{rng.getrandbits(40):x}@marketplace.amazon.com"
        for p, qty in order_lines(rng):
            rows.append({"amazon-order-id": oid, "purchase-date": t.isoformat(timespec="seconds"),
                         "order-status": status, "product-name": p[1], "sku": p[0], "quantity": qty,
                         "currency": "USD", "item-price": f"{p[3] * qty:.2f}", "buyer-email": buyer})
    rows.sort(key=lambda r: r["purchase-date"])
    for i in sorted(rng.sample(range(len(rows)), 6), reverse=True):
        rows.insert(i + 1, dict(rows[i]))  # part of the report was pasted in twice
    rows[rng.randrange(len(rows))]["item-price"] = ""

    with open(config.RAW_DIR / "amazon_orders_2026-H1.txt", "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0]), delimiter="\t")
        w.writeheader()
        w.writerows(rows)


def store(rng):
    wb = Workbook()
    ws = wb.active
    ws.title = "Sales"
    ws.append(["Downtown Store - Sales Log, Jan-Jun 2026"])
    ws.append([])
    ws.append(["Order No", "Date", "Product", "SKU", "Qty", "Unit Price", "Customer Email", "Staff"])
    days = sorted(order_time(rng).date() for _ in range(260))
    units = 0
    for n, d in enumerate(days, start=1):
        date_cell = d.strftime("%d/%m/%Y") if rng.random() < 0.8 else datetime(d.year, d.month, d.day)
        buyer = email(rng) if rng.random() < 0.6 else None  # walk-in customers leave no email
        staff = rng.choice(["Ana", "Ben", "Chloe"])
        for p, qty in order_lines(rng):
            sku = f" {p[0].lower()}" if rng.random() < 0.25 else p[0]
            price = p[3] if rng.random() < 0.5 else f"${p[3]:,.2f}"
            ws.append([f"S-{n:04d}", date_cell, p[1], sku, qty if rng.random() < 0.85 else f"{qty} ",
                       price, buyer, staff])
            units += qty
        if n % 65 == 0:  # returns are logged in the same sheet with a negative quantity
            p = rng.choice(PRODUCTS[:-1])
            ws.append([f"R-{n // 65:04d}", date_cell, p[1], p[0], -1, p[3], None, staff])
    ws.append([])
    ws.append([None, None, "TOTAL", None, units])
    wb.save(config.RAW_DIR / "store_sales_2026-H1.xlsx")


def products():
    with open(config.PRODUCTS_FILE, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["sku", "product", "category", "list_price"])
        w.writerows(p[:4] for p in PRODUCTS)


if __name__ == "__main__":
    rng = random.Random(42)
    config.RAW_DIR.mkdir(parents=True, exist_ok=True)
    products()
    shopify(rng)
    amazon(rng)
    store(rng)
    print(f"Sample exports written to {config.RAW_DIR}")
