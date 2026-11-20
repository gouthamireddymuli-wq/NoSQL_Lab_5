"""Lab 5.1 Try It Yourself – Extended Sales Analytics Pipeline

Extra pipelines built on the same orders collection:
  1. Revenue & order-count breakdown by REGION
  2. Best customer per region ($group with $first after sorting)
  3. $facet: run multiple aggregations in a single pass
"""

import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from config.connection import get_db, reset_collection, banner
from rich.table import Table
from rich.console import Console

console = Console()

ORDERS = [
    {"order_id": "ORD-101", "customer": "Arjun",  "status": "completed", "total": 2500, "items": 3,  "date": "2024-01-15", "region": "North"},
    {"order_id": "ORD-102", "customer": "Sneha",  "status": "completed", "total": 1800, "items": 1,  "date": "2024-01-20", "region": "South"},
    {"order_id": "ORD-103", "customer": "Arjun",  "status": "completed", "total": 3200, "items": 5,  "date": "2024-02-01", "region": "North"},
    {"order_id": "ORD-104", "customer": "Kavya",  "status": "cancelled", "total":  900, "items": 2,  "date": "2024-02-05", "region": "East"},
    {"order_id": "ORD-105", "customer": "Rahul",  "status": "completed", "total": 5600, "items": 8,  "date": "2024-02-10", "region": "West"},
    {"order_id": "ORD-106", "customer": "Arjun",  "status": "completed", "total": 1200, "items": 2,  "date": "2024-02-15", "region": "North"},
    {"order_id": "ORD-107", "customer": "Sneha",  "status": "completed", "total": 4100, "items": 6,  "date": "2024-03-01", "region": "South"},
    {"order_id": "ORD-108", "customer": "Kavya",  "status": "completed", "total": 2900, "items": 4,  "date": "2024-03-05", "region": "East"},
    {"order_id": "ORD-109", "customer": "Rahul",  "status": "pending",   "total": 7800, "items": 10, "date": "2024-03-10", "region": "West"},
    {"order_id": "ORD-110", "customer": "Divya",  "status": "completed", "total": 1500, "items": 1,  "date": "2024-03-15", "region": "North"},
]


def main():
    banner("Lab 5.1 Try It Yourself: Extended Analytics")
    col = reset_collection("nosql_labs", "orders_try")
    col.insert_many(ORDERS)
    print(f"[OK] Inserted {len(ORDERS)} orders.\n")

    # ── PIPELINE 1: Revenue by Region (completed orders only) ─────────────────
    print("=== 1. Revenue by Region (completed orders only) ===")
    pipeline_region = [
        {"$match": {"status": "completed"}},
        {"$group": {
            "_id":           "$region",
            "total_revenue": {"$sum": "$total"},
            "order_count":   {"$sum": 1},
            "avg_order":     {"$avg": "$total"},
            "total_items":   {"$sum": "$items"},
        }},
        {"$sort": {"total_revenue": -1}},
        {"$project": {
            "region":        "$_id",
            "total_revenue": 1,
            "order_count":   1,
            "avg_order":     {"$round": ["$avg_order", 2]},
            "total_items":   1,
            "_id":           0,
        }},
    ]
    tbl = Table(title="Revenue by Region")
    tbl.add_column("Region", style="cyan")
    tbl.add_column("Revenue",     style="green", justify="right")
    tbl.add_column("Orders",      justify="right")
    tbl.add_column("Avg Order",   justify="right")
    tbl.add_column("Items Sold",  justify="right")
    for r in col.aggregate(pipeline_region):
        tbl.add_row(r["region"], f"${r['total_revenue']:,.0f}",
                    str(r["order_count"]), f"${r['avg_order']:,.2f}",
                    str(r["total_items"]))
    console.print(tbl)

    # ── PIPELINE 2: Top Customer per Region ──────────────────────────────────
    print("\n=== 2. Top Customer per Region (highest single-order total) ===")
    pipeline_top_region = [
        {"$match": {"status": "completed"}},
        {"$sort": {"total": -1}},
        {"$group": {
            "_id":         "$region",
            "top_customer": {"$first": "$customer"},
            "top_total":    {"$first": "$total"},
            "top_order":    {"$first": "$order_id"},
        }},
        {"$sort": {"top_total": -1}},
    ]
    for r in col.aggregate(pipeline_top_region):
        print(f"  {r['_id']:<8} -> {r['top_customer']:<10}  "
              f"(best order {r['top_order']}: ${r['top_total']:,})")

    # ── PIPELINE 3: $facet – multi-dimensional analytics in one pass ─────────
    print("\n=== 3. $facet: Status Summary + Region Summary (single pass) ===")
    pipeline_facet = [
        {"$facet": {
            "by_status": [
                {"$group": {"_id": "$status", "count": {"$sum": 1},
                            "revenue": {"$sum": "$total"}}},
                {"$sort": {"revenue": -1}},
            ],
            "by_region": [
                {"$match": {"status": "completed"}},
                {"$group": {"_id": "$region", "revenue": {"$sum": "$total"},
                            "orders": {"$sum": 1}}},
                {"$sort": {"revenue": -1}},
            ],
            "grand_total": [
                {"$match": {"status": "completed"}},
                {"$group": {"_id": None, "total": {"$sum": "$total"},
                            "count": {"$sum": 1}}},
            ],
        }}
    ]
    result = list(col.aggregate(pipeline_facet))[0]

    print("\n  [By Status]")
    for s in result["by_status"]:
        print(f"    {s['_id']:<12} {s['count']} orders  ${s['revenue']:>8,}")

    print("\n  [By Region – completed only]")
    for s in result["by_region"]:
        print(f"    {s['_id']:<8} {s['orders']} orders  ${s['revenue']:>8,}")

    gt = result["grand_total"][0]
    print(f"\n  Grand Total (completed): ${gt['total']:,} across {gt['count']} orders")

    banner("Lab 5.1 Try It Yourself Complete")


if __name__ == "__main__":
    main()
