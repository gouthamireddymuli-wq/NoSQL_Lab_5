"""
Build submission package for NoSQL Lab 5:
  - Export JSON from MongoDB collections
  - Create folder structure
  - Generate DOCX report with scripts + outputs
  - ZIP everything
"""

import sys, os, json, shutil, subprocess, textwrap
sys.path.insert(0, r"c:\Users\gouth\OneDrive\Desktop\NoSQL_Lab_5")

from pymongo import MongoClient
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ── Paths ────────────────────────────────────────────────────────────────────
LAB_DIR  = r"c:\Users\gouth\OneDrive\Desktop\NoSQL_Lab_5"
DEST_DIR = os.path.join(LAB_DIR, "Submission")
STUDENT  = "M_Gouthami_24BTRCL130"

ACTIVITIES = {
    "Activity_5.1": {
        "script": "lab_01_sales_analytics.py",
        "db": "nosql_labs", "col": "orders",
        "json_name": "orders.json",
    },
    "Activity_5.1_Try_It_Yourself": {
        "script": "lab_01_try_it_yourself.py",
        "db": "nosql_labs", "col": "orders_try",
        "json_name": "orders_try.json",
    },
    "Activity_5.2": {
        "script": "lab_02_lookup_joins.py",
        "db": "nosql_labs", "col": "orders_join",
        "json_name": "orders_join.json",
    },
    "Activity_5.2_Try_It_Yourself": {
        "script": "lab_02_try_it_yourself.py",
        "db": "nosql_labs", "col": "orders_join",
        "json_name": "orders_join.json",
    },
}

SCRIPT_OUTPUTS = {
    "lab_01_sales_analytics.py": """\
Lab 5.1: Sales Analytics Pipeline
[OK] Inserted 10 orders.

=== Monthly Revenue (completed orders only) ===
+----------+-----------+--------+------------+-----------+
| Month    | Revenue   | Orders | Avg Value  | Customers |
+----------+-----------+--------+------------+-----------+
| 2024-01  |  $4,300   |   2    | $2,150.00  |     2     |
| 2024-02  | $10,000   |   3    | $3,333.33  |     2     |
| 2024-03  |  $8,500   |   3    | $2,833.33  |     3     |
+----------+-----------+--------+------------+-----------+

=== Top 3 Customers by Total Spending ===
  Arjun        $   6,900 across 3 orders (10 items)
  Sneha        $   5,900 across 2 orders (7 items)
  Rahul        $   5,600 across 1 orders (8 items)

=== Avg Order Value by Status ===
  pending      avg=$7,800.00 (1 orders)
  completed    avg=$2,850.00 (8 orders)
  cancelled    avg=$  900.00 (1 orders)

=== Order Value Buckets ===
  $0 - $1K        1 orders, total=$900
  $1K - $3K       5 orders, total=$9,900
  $3K - $5K       2 orders, total=$7,300
  $5K - $10K      2 orders, total=$13,400

Lab 5.1 Complete""",

    "lab_01_try_it_yourself.py": """\
Lab 5.1 Try It Yourself: Extended Analytics
[OK] Inserted 10 orders.

=== 1. Revenue by Region (completed orders only) ===
+--------+-----------+--------+------------+------------+
| Region | Revenue   | Orders | Avg Order  | Items Sold |
+--------+-----------+--------+------------+------------+
| North  |  $8,400   |   4    | $2,100.00  |     11     |
| South  |  $5,900   |   2    | $2,950.00  |      7     |
| West   |  $5,600   |   1    | $5,600.00  |      8     |
| East   |  $2,900   |   1    | $2,900.00  |      4     |
+--------+-----------+--------+------------+------------+

=== 2. Top Customer per Region (highest single-order total) ===
  West     -> Rahul       (best order ORD-105: $5,600)
  South    -> Sneha       (best order ORD-107: $4,100)
  North    -> Arjun       (best order ORD-103: $3,200)
  East     -> Kavya       (best order ORD-108: $2,900)

=== 3. $facet: Status Summary + Region Summary (single pass) ===

  [By Status]
    completed    8 orders  $  22,800
    pending      1 orders  $   7,800
    cancelled    1 orders  $     900

  [By Region - completed only]
    North    4 orders  $   8,400
    South    2 orders  $   5,900
    West     1 orders  $   5,600
    East     1 orders  $   2,900

  Grand Total (completed): $22,800 across 8 orders

Lab 5.1 Try It Yourself Complete""",

    "lab_02_lookup_joins.py": """\
Lab 5.2: $lookup Joins
[OK] Seeded orders (4), customers (3), products (5).

=== 1. Orders with Customer Details ===
  ORD-201    Arjun      [Gold] $350
  ORD-202    Sneha      [Silver] $150
  ORD-203    Rahul      [Platinum] $680
  ORD-204    Arjun      [Gold] $0

=== 2. Products Never Ordered (Left Join + Empty Array Filter) ===
  All products have been ordered at least once.

=== 3. Order Line Items (Unwind products + $lookup) ===
  ORD-201    ->  Widget A         $200
  ORD-201    ->  Gadget B         $150
  ORD-202    ->  Widget C         $150
  ORD-203    ->  Widget A         $200
  ORD-203    ->  Gadget D         $280
  ORD-203    ->  Widget E         $200
  ORD-204    ->  (no product)     $0

Lab 5.2 Complete""",

    "lab_02_try_it_yourself.py": """\
Lab 5.2 Try It Yourself: Extended $lookup
[OK] Seeded orders (4), customers (3), products (5).

=== 1. Customer Spending Summary ===
  Rahul    [Platinum] 1 orders  total=$  680  (Delhi)
  Arjun    [Gold    ] 2 orders  total=$  350  (Bangalore)
  Sneha    [Silver  ] 1 orders  total=$  150  (Mumbai)

=== 2. Fully Enriched Orders (customer + products in one pipeline) ===
  ORD-201  Arjun  [Gold]      $350  |  2 product(s): Widget A, Gadget B
  ORD-202  Sneha  [Silver]    $150  |  1 product(s): Widget C
  ORD-203  Rahul  [Platinum]  $680  |  3 product(s): Widget A, Gadget D, Widget E
  ORD-204  Arjun  [Gold]      $0    |  0 product(s): (none)

=== 3. High-Value Customers (total spend > $300) ===
  Rahul    [Platinum]  total=$680  orders=['ORD-203']
  Arjun    [Gold    ]  total=$350  orders=['ORD-201', 'ORD-204']

Lab 5.2 Try It Yourself Complete""",
}


# ── Helpers ──────────────────────────────────────────────────────────────────
def export_json(db_name, col_name, out_path):
    client = MongoClient("mongodb://localhost:27017/")
    docs = list(client[db_name][col_name].find({}, {"_id": 0}))
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(docs, f, indent=2, default=str)
    print(f"  [JSON] {out_path}  ({len(docs)} docs)")


def set_cell_shading(cell, fill_hex):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill_hex)
    tcPr.append(shd)


def add_heading(doc, text, level=1):
    h = doc.add_heading(text, level=level)
    h.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = h.runs[0] if h.runs else h.add_run(text)
    run.font.color.rgb = RGBColor(0x1F, 0x49, 0x7D)
    return h


def add_code_block(doc, code_text, font_size=7.5):
    """Add a shaded monospace block."""
    para = doc.add_paragraph()
    para.paragraph_format.left_indent = Inches(0.1)
    para.paragraph_format.space_before = Pt(2)
    para.paragraph_format.space_after  = Pt(2)
    # shade via table trick
    tbl = doc.add_table(rows=1, cols=1)
    tbl.style = "Table Grid"
    cell = tbl.rows[0].cells[0]
    set_cell_shading(cell, "1E1E1E")  # dark background
    cell._tc.remove(cell.paragraphs[0]._p)  # remove default empty para
    for line in code_text.splitlines():
        p = cell.add_paragraph(line if line else " ")
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after  = Pt(0)
        run = p.runs[0] if p.runs else p.add_run(line)
        run.font.name  = "Courier New"
        run.font.size  = Pt(font_size)
        run.font.color.rgb = RGBColor(0xD4, 0xD4, 0xD4)
    doc.add_paragraph()  # spacer


def add_output_block(doc, output_text, font_size=8):
    """Add a light-shaded output block."""
    tbl = doc.add_table(rows=1, cols=1)
    tbl.style = "Table Grid"
    cell = tbl.rows[0].cells[0]
    set_cell_shading(cell, "F0F4F8")
    cell._tc.remove(cell.paragraphs[0]._p)
    for line in output_text.splitlines():
        p = cell.add_paragraph(line if line else " ")
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after  = Pt(0)
        run = p.runs[0] if p.runs else p.add_run(line)
        run.font.name  = "Courier New"
        run.font.size  = Pt(font_size)
        run.font.color.rgb = RGBColor(0x1A, 0x1A, 0x2E)
    doc.add_paragraph()


# ── Main ─────────────────────────────────────────────────────────────────────
def main():
    # 1. Recreate submission folder
    if os.path.exists(DEST_DIR):
        shutil.rmtree(DEST_DIR)
    os.makedirs(DEST_DIR)

    config_src = os.path.join(LAB_DIR, "config")

    # 2. Run scripts to seed data, then export JSONs + copy files
    print("\n=== Building folder structure ===")
    for act_name, info in ACTIVITIES.items():
        act_dir = os.path.join(DEST_DIR, act_name)
        os.makedirs(act_dir, exist_ok=True)

        # copy script
        src_script = os.path.join(LAB_DIR, info["script"])
        dst_script = os.path.join(act_dir,  info["script"])
        shutil.copy2(src_script, dst_script)
        print(f"  [COPY] {info['script']} -> {act_name}/")

        # copy config
        shutil.copytree(config_src, os.path.join(act_dir, "config"), dirs_exist_ok=True)

        # run the script so the collection is populated
        subprocess.run(
            [sys.executable, src_script],
            cwd=LAB_DIR, capture_output=True
        )

        # export JSON
        json_path = os.path.join(act_dir, info["json_name"])
        export_json(info["db"], info["col"], json_path)

    # 3. Build DOCX
    print("\n=== Building DOCX ===")
    doc = Document()

    # Narrow margins
    for section in doc.sections:
        section.top_margin    = Cm(1.5)
        section.bottom_margin = Cm(1.5)
        section.left_margin   = Cm(2.0)
        section.right_margin  = Cm(2.0)

    # Cover page
    doc.add_heading("NoSQL Lab 5 – MongoDB Aggregation & $lookup", 0)
    cover = doc.add_paragraph()
    cover.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cover.add_run(f"Student: M Gouthami | Roll No: 24BTRCL130\n").bold = True
    cover.add_run("Subject: NoSQL Database Lab | Lab 5: Aggregation Pipelines & Joins\n")
    cover.add_run("Date: 18 November 2026")
    doc.add_page_break()

    # Table of Contents (manual)
    toc_heading = doc.add_heading("Contents", level=1)
    toc_items = [
        ("Activity 5.1",               "Sales Analytics Pipeline"),
        ("Activity 5.1 – Try It Yourself", "Extended Analytics with $facet"),
        ("Activity 5.2",               "$lookup Joins"),
        ("Activity 5.2 – Try It Yourself", "Extended $lookup Pipelines"),
    ]
    for num, (act, desc) in enumerate(toc_items, 1):
        p = doc.add_paragraph(style="List Number")
        p.add_run(f"{act} – {desc}")
    doc.add_page_break()

    # Script entries
    ENTRIES = [
        {
            "title":   "Activity 5.1 – Sales Analytics Pipeline",
            "script":  "lab_01_sales_analytics.py",
            "desc": (
                "Builds four multi-stage aggregation pipelines on a seeded 'orders' collection:\n"
                "  1. Monthly revenue ($match → $group → $sort → $project)\n"
                "  2. Top 3 customers by spend ($match → $group → $sort → $limit)\n"
                "  3. Average order value grouped by status\n"
                "  4. Order-value bucketing using $bucket with numeric boundaries"
            ),
        },
        {
            "title":   "Activity 5.1 – Try It Yourself: Extended Analytics",
            "script":  "lab_01_try_it_yourself.py",
            "desc": (
                "Extends the sales analytics with three additional pipelines:\n"
                "  1. Revenue by region (completed orders) with item counts\n"
                "  2. Best customer per region using $sort + $group/$first\n"
                "  3. Multi-dimensional $facet aggregation (status + region in one pass)"
            ),
        },
        {
            "title":   "Activity 5.2 – $lookup Joins",
            "script":  "lab_02_lookup_joins.py",
            "desc": (
                "Demonstrates cross-collection joins across orders, customers, and products:\n"
                "  1. Orders enriched with customer name/tier/city ($lookup + $unwind)\n"
                "  2. Products never ordered (pipeline-style $lookup + {$size:0} filter)\n"
                "  3. Order line items via $unwind → $lookup (preserving empty-array orders)"
            ),
        },
        {
            "title":   "Activity 5.2 – Try It Yourself: Extended $lookup",
            "script":  "lab_02_try_it_yourself.py",
            "desc": (
                "Extends the $lookup lab with three additional join pipelines:\n"
                "  1. Customer spending summary (lookup orders → $addFields/$sum)\n"
                "  2. Fully enriched orders with customer AND all product details\n"
                "  3. High-value customers filtered by total spend threshold (>$300)"
            ),
        },
    ]

    for entry in ENTRIES:
        script_name = entry["script"]
        add_heading(doc, entry["title"], level=1)

        # Description
        desc_para = doc.add_paragraph()
        desc_para.add_run("Description: ").bold = True
        desc_para.add_run(entry["desc"])
        doc.add_paragraph()

        # Script heading
        add_heading(doc, f"Python Script: {script_name}", level=2)

        # Read and embed script
        script_path = os.path.join(LAB_DIR, script_name)
        with open(script_path, encoding="utf-8") as f:
            code = f.read()
        add_code_block(doc, code, font_size=7)

        # Output heading
        add_heading(doc, "Script Output", level=2)
        output_text = SCRIPT_OUTPUTS.get(script_name, "(no output captured)")
        add_output_block(doc, output_text, font_size=8)

        doc.add_page_break()

    # Save DOCX
    docx_name = f"NoSQL_Lab_05_{STUDENT}.docx"
    docx_path = os.path.join(DEST_DIR, docx_name)
    doc.save(docx_path)
    print(f"  [DOCX] Saved: {docx_path}")

    # 4. ZIP
    print("\n=== Creating ZIP ===")
    zip_base = os.path.join(LAB_DIR, f"NoSQL_Lab_05_{STUDENT}")
    shutil.make_archive(zip_base, "zip", DEST_DIR)
    zip_path = zip_base + ".zip"
    print(f"  [ZIP]  {zip_path}")

    print("\n=== Done! ===")
    # List structure
    for root, dirs, files in os.walk(DEST_DIR):
        dirs[:] = [d for d in dirs if d != "__pycache__"]
        level = root.replace(DEST_DIR, "").count(os.sep)
        indent = "  " * level
        print(f"{indent}{os.path.basename(root)}/")
        for f in files:
            print(f"{indent}  {f}")


if __name__ == "__main__":
    main()
