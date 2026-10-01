"""Excel report: summary with a chart, products, the clean data, and a data quality sheet."""
import pandas as pd

from clean import REASONS


def write_report(clean, removed, rows_read, path):
    month = clean["order_date"].dt.strftime("%Y-%m")
    by_month = clean.pivot_table(index=month, columns="source", values="revenue", aggfunc="sum", fill_value=0)
    by_month["Total"] = by_month.sum(axis=1)
    by_month.index.name = "Month"

    products = (clean.groupby(["product", "category"], as_index=False)
                .agg(units=("qty", "sum"), revenue=("revenue", "sum"))
                .sort_values("revenue", ascending=False))
    products["share"] = products["revenue"] / products["revenue"].sum()

    quality = quality_table(clean, removed, rows_read)
    orders = clean[["source", "order_id"]].drop_duplicates()

    with pd.ExcelWriter(path, engine="xlsxwriter", datetime_format="yyyy-mm-dd") as writer:
        book = writer.book
        money = book.add_format({"num_format": "$#,##0.00"})
        pct = book.add_format({"num_format": "0.0%"})
        head = book.add_format({"bold": True, "bg_color": "#DDEBF7", "bottom": 1})
        title = book.add_format({"bold": True, "font_size": 16})
        label = book.add_format({"font_color": "#666666"})
        kpi = book.add_format({"bold": True, "font_size": 14, "align": "left"})
        kpi_money = book.add_format({"bold": True, "font_size": 14, "align": "left", "num_format": "$#,##0"})

        def table(df, sheet, startrow=0, index=False):
            df.to_excel(writer, sheet_name=sheet, startrow=startrow, index=index)
            ws = writer.sheets[sheet]
            labels = ([df.index.name] if index else []) + list(df.columns)
            for col, text in enumerate(labels):
                ws.write(startrow, col, text, head)
            return ws

        # Summary
        ws = table(by_month, "Summary", startrow=7, index=True)
        first, last = clean["order_date"].min(), clean["order_date"].max()
        ws.write("A1", f"Sales report, {first:%b %Y} to {last:%b %Y}", title)
        ws.write("A2", f"Sources: {', '.join(sorted(clean['source'].unique()))}. "
                       f"{len(removed)} of {sum(rows_read.values())} rows left out, see Data Quality.", label)
        revenue = clean["revenue"].sum()
        for col, (name, value, fmt) in enumerate([
            ("Revenue", revenue, kpi_money),
            ("Orders", len(orders), kpi),
            ("Units", clean["qty"].sum(), kpi),
            ("Avg order value", revenue / len(orders), kpi_money),
        ]):
            ws.write(3, col, name, label)
            ws.write(4, col, value, fmt)
        ws.set_column(0, len(by_month.columns), 15, money)

        chart = book.add_chart({"type": "column", "subtype": "stacked"})
        n = len(by_month)
        for i in range(1, len(by_month.columns)):  # every source column, not Total
            chart.add_series({"name": ["Summary", 7, i],
                              "categories": ["Summary", 8, 0, 7 + n, 0],
                              "values": ["Summary", 8, i, 7 + n, i]})
        chart.set_title({"name": "Revenue by month and channel"})
        chart.set_y_axis({"num_format": "$#,##0", "major_gridlines": {"visible": True}})
        chart.set_legend({"position": "bottom"})
        chart.set_size({"width": 600, "height": 320})
        ws.insert_chart(9 + n, 0, chart)

        # Products
        ws = table(products, "Products")
        ws.set_column("A:A", 22)
        ws.set_column("B:C", 12)
        ws.set_column("D:D", 14, money)
        ws.set_column("E:E", 10, pct)

        # Clean data
        ws = table(clean, "Clean Data")
        ws.autofilter(0, 0, len(clean), len(clean.columns) - 1)
        ws.freeze_panes(1, 0)
        ws.set_column("A:G", 13)
        ws.set_column("E:E", 18)
        ws.set_column("H:I", 12, money)
        ws.set_column("J:K", 32)

        # Data quality: rows read, removed per reason, kept. Then every removed row.
        ws = table(quality, "Data Quality", index=True)
        ws.set_column("A:A", 52)
        ws.set_column("B:E", 11)
        ws.set_landscape()
        ws.fit_to_pages(1, 0)

        ws = table(removed[["reason", "source", "source_file", "source_row", "order_id", "order_date",
                            "sku", "qty", "unit_price", "email", "status"]], "Removed Rows")
        ws.autofilter(0, 0, len(removed), 10)
        ws.freeze_panes(1, 0)
        ws.set_column("A:A", 52)
        ws.set_column("B:B", 10)
        ws.set_column("C:C", 32)
        ws.set_column("D:I", 13)
        ws.set_column("J:J", 32)


def quality_table(clean, removed, rows_read):
    """Rows read, rows removed per reason, and rows kept. Each column adds up."""
    counts = removed.groupby(["reason", "source"]).size().unstack(fill_value=0)
    counts = counts.reindex(index=[r for r in REASONS if r in counts.index])
    table = pd.concat([
        pd.DataFrame(rows_read, index=["Rows read"]),
        -counts,
        clean.groupby("source").size().to_frame("Rows in report").T,
    ]).fillna(0).astype(int)
    table["Total"] = table.sum(axis=1)
    table.index.name = "Step"
    return table
