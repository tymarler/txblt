#!/usr/bin/env python3
"""Correlate Square item export CSV to sim rental summary table."""

from __future__ import annotations

import argparse
import csv
from collections import defaultdict
from decimal import Decimal
from pathlib import Path

TIERS = [
    ("Large Bay", 45, 1),
    ("Large Bay", 90, 2),
    ("Large Bay", 135, 3),
    ("Large Bay", 180, 4),
    ("Private Bay", 55, 1),
    ("Private Bay", 110, 2),
    ("Standard Bay (Right Hand Only)", 35, 1),
    ("Standard Bay (Right Hand Only)", 70, 2),
]
TIER_HOURS = {(item, price): hours for item, price, hours in TIERS}

# Split-bill line totals that roll up to a full session price.
SPLIT_SNAPS = {
    ("Large Bay", 11.25): ("Large Bay", 45),
    ("Standard Bay (Right Hand Only)", 17.5): ("Standard Bay (Right Hand Only)", 35),
    ("Standard Bay (Right Hand Only)", 52.5): ("Standard Bay (Right Hand Only)", 35),
}


def money(value: str) -> Decimal:
    return Decimal(value.replace("$", "").replace(",", ""))


def normalize_item_and_price(item: str, total: Decimal) -> tuple[str, int | None]:
    if item == "Standard Bay":
        item = "Standard Bay (Right Hand Only)"

    total_f = float(total)
    if item == "Large Bay" and total_f == 110:
        item = "Private Bay"

    known_prices = [price for tier_item, price, _ in TIERS if tier_item == item]
    if total_f in known_prices:
        return item, int(total_f)

    snap = SPLIT_SNAPS.get((item, total_f))
    if snap:
        return snap

    return item, None


def summarize(input_path: Path) -> tuple[list[dict[str, object]], list[dict[str, object]]]:
    rows = []
    with input_path.open(newline="", encoding="utf-8") as handle:
        for row in csv.DictReader(handle):
            if row["Category"] == "Sim Rental":
                rows.append(row)

    by_transaction: dict[str, dict[str, object]] = defaultdict(
        lambda: {"item": None, "total": Decimal(0), "qty": Decimal(0)}
    )
    for row in rows:
        txn_id = row["Transaction ID"]
        entry = by_transaction[txn_id]
        entry["item"] = row["Item"]
        entry["total"] += money(row["Gross Sales"])
        entry["qty"] = max(entry["qty"], Decimal(row["Qty"]))

    summary: dict[tuple[str, int], dict[str, object]] = defaultdict(
        lambda: {"quantity": Decimal(0), "total": Decimal(0), "total_hours": 0}
    )
    unmatched = []
    for txn_id, entry in by_transaction.items():
        item, price = normalize_item_and_price(entry["item"], entry["total"])
        if price is None:
            unmatched.append(
                {
                    "transaction_id": txn_id,
                    "item": entry["item"],
                    "gross_sales": float(entry["total"]),
                    "qty": float(entry["qty"]),
                }
            )
            continue

        key = (item, price)
        summary[key]["quantity"] += entry["qty"]
        summary[key]["total"] += entry["total"]
        summary[key]["total_hours"] += TIER_HOURS[key] * float(entry["qty"])

    table_rows = []
    for item, price, _ in TIERS:
        data = summary.get((item, price), {"quantity": 0, "total": 0, "total_hours": 0})
        table_rows.append(
            {
                "Type": item,
                "Cost": price,
                "Quantity": int(data["quantity"]),
                "Total": float(data["total"]),
                "Total Hours": int(data["total_hours"]),
            }
        )

    total_quantity = sum(row["Quantity"] for row in table_rows)
    total_amount = sum(row["Total"] for row in table_rows)
    total_hours = sum(row["Total Hours"] for row in table_rows)
    table_rows.append(
        {
            "Type": "Total:",
            "Cost": "",
            "Quantity": total_quantity,
            "Total": total_amount,
            "Total Hours": total_hours,
        }
    )
    return table_rows, unmatched


def write_csv(rows: list[dict[str, object]], output_path: Path) -> None:
    with output_path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=["Type", "Cost", "Quantity", "Total", "Total Hours"])
        writer.writeheader()
        for row in rows:
            writer.writerow(row)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input_csv", type=Path, help="Square items export CSV")
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=Path("output/sim-rental-summary.csv"),
        help="Output summary CSV",
    )
    args = parser.parse_args()

    rows, unmatched = summarize(args.input_csv)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    write_csv(rows, args.output)

    print(f"Wrote summary to {args.output}")
    if unmatched:
        print(f"Warning: {len(unmatched)} transaction(s) could not be mapped to a price tier.")


if __name__ == "__main__":
    main()
