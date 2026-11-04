"""Create synthetic input and reference data; never load input rows into WMS."""
import argparse
import json
import sqlite3
from pathlib import Path

import pandas as pd

DEFAULT_OUTPUT = Path(__file__).resolve().parents[1] / "data"
CUSTOMER = "SYNTHETIC_CUSTOMER"
RULES = {
    "allowed_zones": ["AMBIENT", "CHILLED"],
    "allowed_location_types": ["PICK", "RESERVE"],
    "max_capacity_pallets": 20,
    "location_code_pattern": r"^LOC-\d{3}$",
    "required_fields": ["customer_id", "location_code", "zone", "location_type", "capacity_pallets"],
    "capacity_must_be_positive_integer": True,
    "location_code_must_be_unique": True,
}


def location_rows() -> list[dict]:
    rows = [dict(customer_id=CUSTOMER, location_code=f"LOC-{i:03d}",
                 zone="AMBIENT" if i % 2 else "CHILLED",
                 location_type="PICK" if i % 2 else "RESERVE",
                 capacity_pallets=10) for i in range(1, 21)]
    rows[12]["location_code"] = "LOC-001"  # duplicate within workbook
    rows[13]["location_code"] = None  # required field missing
    rows[14]["zone"] = "UNKNOWN"
    rows[15]["capacity_pallets"] = -1
    rows[16]["capacity_pallets"] = "many"
    rows[17]["location_type"] = "UNKNOWN"
    rows[18]["location_code"] = "bad code"
    rows[19]["capacity_pallets"] = 25
    return rows


def generate(output_dir: Path = DEFAULT_OUTPUT) -> tuple[Path, Path]:
    """Create new fixtures, refusing to overwrite either existing artifact."""
    output_dir.mkdir(parents=True, exist_ok=True)
    workbook = output_dir / "warehouse_locations.xlsx"
    database = output_dir / "synthetic_wms.sqlite3"
    if workbook.exists() or database.exists():
        raise FileExistsError("Fixture already exists; choose a new --output-dir to preserve local data.")
    pd.DataFrame(location_rows()).to_excel(workbook, index=False, sheet_name="locations")
    with sqlite3.connect(database) as connection:
        connection.executescript("""
            CREATE TABLE customers (customer_id TEXT PRIMARY KEY, name TEXT NOT NULL);
            CREATE TABLE locations (
                customer_id TEXT NOT NULL REFERENCES customers(customer_id),
                location_code TEXT NOT NULL,
                zone TEXT NOT NULL,
                location_type TEXT NOT NULL,
                capacity_pallets INTEGER NOT NULL CHECK(capacity_pallets > 0),
                PRIMARY KEY (customer_id, location_code)
            );
            CREATE TABLE customer_rules (
                customer_id TEXT NOT NULL REFERENCES customers(customer_id),
                rule_name TEXT NOT NULL,
                rule_value_json TEXT NOT NULL,
                PRIMARY KEY (customer_id, rule_name)
            );
        """)
        connection.execute("PRAGMA foreign_keys = ON")
        connection.execute("INSERT INTO customers VALUES (?, ?)", (CUSTOMER, "Synthetic Warehouse Customer"))
        connection.executemany("INSERT INTO locations VALUES (?, ?, ?, ?, ?)", [
            (CUSTOMER, "LOC-001", "AMBIENT", "PICK", 10),
            (CUSTOMER, "LOC-100", "CHILLED", "RESERVE", 20),
            (CUSTOMER, "LOC-101", "AMBIENT", "RESERVE", 15),
        ])
        connection.executemany("INSERT INTO customer_rules VALUES (?, ?, ?)",
                               [(CUSTOMER, key, json.dumps(value)) for key, value in RULES.items()])
    return workbook, database


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output-dir", type=Path, default=DEFAULT_OUTPUT)
    args = parser.parse_args()
    for path in generate(args.output_dir):
        print(path)


if __name__ == "__main__":
    main()
