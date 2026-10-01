"""Clean, merge and de-duplicate every sales export in a folder, then build the Excel report."""
import argparse
import sys
from pathlib import Path

import pandas as pd

import config
from clean import load_source, process, standardize
from report import write_report


def run(input_dir, output_dir, products_file):
    frames, rows_read = [], {}
    for name, cfg in config.SOURCES.items():
        for path in sorted(input_dir.glob(cfg["pattern"])):
            df = standardize(load_source(name, cfg, path), cfg)
            frames.append(df)
            rows_read[name] = rows_read.get(name, 0) + len(df)
            print(f"  read {path.name}: {len(df)} rows")
    if not frames:
        sys.exit(f"No input files found in {input_dir}")

    products = pd.read_csv(products_file, dtype={"sku": "str"})
    clean, removed = process(frames, products)
    # Every row read is either in the report or listed with a reason. Nothing disappears silently.
    if sum(rows_read.values()) != len(clean) + len(removed):
        raise RuntimeError("Row counts don't add up: some rows went missing")

    output_dir.mkdir(parents=True, exist_ok=True)
    clean.to_csv(output_dir / "clean_orders.csv", index=False)
    write_report(clean, removed, rows_read, output_dir / "sales_report.xlsx")
    return clean, removed, rows_read


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", type=Path, default=config.RAW_DIR, help="folder with the exports")
    parser.add_argument("--output", type=Path, default=config.OUTPUT_DIR)
    parser.add_argument("--products", type=Path, default=config.PRODUCTS_FILE)
    args = parser.parse_args(argv)

    clean, removed, rows_read = run(args.input, args.output, args.products)
    print(f"{sum(rows_read.values())} rows read, {len(clean)} in the report, {len(removed)} left out:")
    for reason, n in removed["reason"].value_counts().items():
        print(f"  {n:4d}  {reason}")
    print(f"Revenue ${clean['revenue'].sum():,.2f}. Report: {args.output / 'sales_report.xlsx'}")


if __name__ == "__main__":
    main()
