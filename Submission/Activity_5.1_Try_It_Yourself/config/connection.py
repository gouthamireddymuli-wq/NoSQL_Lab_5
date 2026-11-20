"""Shared MongoDB connection helper for NoSQL Labs.
"""

import json
from pymongo import MongoClient
from rich.console import Console
from rich.panel import Panel
from rich.text import Text
from bson import json_util

console = Console()
DEFAULT_URI = "mongodb://localhost:27017/"


def get_client(uri=DEFAULT_URI):
    """Return a connected MongoClient."""
    return MongoClient(uri)


def get_db(db_name="nosql_labs", uri=DEFAULT_URI):
    """Return a database handle."""
    client = get_client(uri)
    return client[db_name]


def reset_collection(db_name, col_name, uri=DEFAULT_URI):
    """Drop the collection if it exists, returning a clean collection handle."""
    db = get_db(db_name, uri)
    db.drop_collection(col_name)
    return db[col_name]


def banner(title: str):
    """Print a clean visual banner."""
    console.print()
    panel = Panel(
        Text(title, justify="center", style="bold white on blue"),
        expand=False,
        border_style="cyan"
    )
    console.print(panel)
    console.print()


def print_json(data):
    """Print formatted JSON data."""
    if isinstance(data, str):
        console.print_json(data)
    else:
        json_str = json.dumps(data, indent=2, default=str)
        console.print_json(json_str)


def export_collection_to_json(db_name, col_name, output_path, uri=DEFAULT_URI):
    """Export all documents in a MongoDB collection to a pretty-printed JSON file."""
    db = get_db(db_name, uri)
    docs = list(db[col_name].find())
    json_data = json.loads(json_util.dumps(docs))
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(json_data, f, indent=2)
    print(f"[OK] Exported {len(docs)} documents to {output_path}")
    return output_path
