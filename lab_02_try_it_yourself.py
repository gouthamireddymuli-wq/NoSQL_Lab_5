"""Lab 5.2 Try It Yourself – Extended $lookup Joins

Extra pipelines on the same three collections:
  1. Customer spending summary – how much each customer has spent in total
     (lookup orders from customers side, then $unwind + $group)
  2. Orders fully enriched with BOTH customer AND all product details
     in a single aggregation pipeline
  3. High-value customers – customers whose total spend exceeds a threshold,
     with their order history listed
"""

import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from config.connection import get_db, reset_collection, banner

ORDERS = [
    {"order_id": "ORD-201", "customer_id": "C001", "product_ids": ["P001", "P002"], "total": 350},
    {"order_id": "ORD-202", "customer_id": "C002", "product_ids": ["P003"],          "total": 150},
    {"order_id": "ORD-203", "customer_id": "C003", "product_ids": ["P001", "P004", "P005"], "total": 680},
    {"order_id": "ORD-204", "customer_id": "C001", "product_ids": [],                "total":   0},
]

CUSTOMERS = [
    {"customer_id": "C001", "name": "Arjun", "tier": "Gold",     "city": "Bangalore"},
    {"customer_id": "C002", "name": "Sneha", "tier": "Silver",   "city": "Mumbai"},
    {"customer_id": "C003", "name": "Rahul", "tier": "Platinum", "city": "Delhi"},
]

PRODUCTS = [
    {"product_id": "P001", "name": "Widget A", "category": "Electronics",  "price": 200},
    {"product_id": "P002", "name": "Gadget B", "category": "Accessories",  "price": 150},
    {"product_id": "P003", "name": "Widget C", "category": "Electronics",  "price": 150},
    {"product_id": "P004", "name": "Gadget D", "category": "Tools",        "price": 280},
    {"product_id": "P005", "name": "Widget E", "category": "Consumables",  "price": 200},
]


def main():
    banner("Lab 5.2 Try It Yourself: Extended $lookup")
    orders_col    = reset_collection("nosql_labs", "orders_join")
    customers_col = reset_collection("nosql_labs", "customers_join")
    products_col  = reset_collection("nosql_labs", "products_join")
    orders_col.insert_many(ORDERS)
    customers_col.insert_many(CUSTOMERS)
    products_col.insert_many(PRODUCTS)
    print(f"[OK] Seeded orders ({len(ORDERS)}), customers ({len(CUSTOMERS)}), "
          f"products ({len(PRODUCTS)}).\n")

    # ── PIPELINE 1: Customer Spending Summary (lookup orders → group) ─────────
    print("=== 1. Customer Spending Summary ===")
    pipeline_spend = [
        {"$lookup": {
            "from":         "orders_join",
            "localField":   "customer_id",
            "foreignField": "customer_id",
            "as":           "orders",
        }},
        {"$project": {
            "customer_id": 1,
            "name":        1,
            "tier":        1,
            "city":        1,
            "order_count": {"$size": "$orders"},
            "total_spent": {"$sum":  "$orders.total"},
        }},
        {"$sort": {"total_spent": -1}},
    ]
    for r in customers_col.aggregate(pipeline_spend):
        print(f"  {r['name']:<8} [{r['tier']:<8}] "
              f"{r['order_count']} orders  total=${r['total_spent']:>5,}  ({r['city']})")

    # ── PIPELINE 2: Fully-enriched orders (customer + all product details) ────
    print("\n=== 2. Fully Enriched Orders (customer + products in one pipeline) ===")
    pipeline_full = [
        # Join customer
        {"$lookup": {
            "from":         "customers_join",
            "localField":   "customer_id",
            "foreignField": "customer_id",
            "as":           "customer",
        }},
        {"$unwind": "$customer"},
        # Join products (array lookup)
        {"$lookup": {
            "from":         "products_join",
            "localField":   "product_ids",
            "foreignField": "product_id",
            "as":           "products",
        }},
        {"$project": {
            "_id":           0,
            "order_id":      1,
            "total":         1,
            "customer_name": "$customer.name",
            "customer_tier": "$customer.tier",
            "product_names": "$products.name",
            "product_count": {"$size": "$products"},
        }},
    ]
    for r in orders_col.aggregate(pipeline_full):
        plist = ", ".join(r["product_names"]) if r["product_names"] else "(none)"
        print(f"  {r['order_id']}  {r['customer_name']:<6} [{r['customer_tier']}]"
              f"  ${r['total']}  |  {r['product_count']} product(s): {plist}")

    # ── PIPELINE 3: High-Value Customers (total spend > threshold) ───────────
    THRESHOLD = 300
    print(f"\n=== 3. High-Value Customers (total spend > ${THRESHOLD}) ===")
    pipeline_hv = [
        {"$lookup": {
            "from":         "orders_join",
            "localField":   "customer_id",
            "foreignField": "customer_id",
            "as":           "orders",
        }},
        {"$addFields": {
            "total_spent": {"$sum": "$orders.total"},
            "order_ids":   "$orders.order_id",
        }},
        {"$match": {"total_spent": {"$gt": THRESHOLD}}},
        {"$project": {
            "_id":        0,
            "name":       1,
            "tier":       1,
            "total_spent":1,
            "order_ids":  1,
        }},
        {"$sort": {"total_spent": -1}},
    ]
    for r in customers_col.aggregate(pipeline_hv):
        print(f"  {r['name']:<8} [{r['tier']:<8}]  "
              f"total=${r['total_spent']}  orders={r['order_ids']}")

    banner("Lab 5.2 Try It Yourself Complete")


if __name__ == "__main__":
    main()
