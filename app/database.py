"""Read-only access to synthetic WMS reference data."""
import json
import math
import re
import sqlite3
from contextlib import closing
from pathlib import Path

DEFAULT_DATABASE = Path(__file__).resolve().parents[1] / "data/synthetic_wms.sqlite3"


class ValidationError(ValueError):
    def __init__(self, code: str, message: str, status: int = 400):
        super().__init__(message)
        self.code = code
        self.status = status


def connect_readonly(path: Path) -> sqlite3.Connection:
    return sqlite3.connect(path.resolve().as_uri() + "?mode=ro", uri=True)


def load_reference(path: Path, customer_id: str) -> tuple[dict, dict[str, set[str]]]:
    try:
        with closing(connect_readonly(path)) as connection:
            customers = {row[0]: set() for row in connection.execute("SELECT customer_id FROM customers")}
            if customer_id not in customers:
                raise ValidationError("UNKNOWN_CUSTOMER", "Requested customer does not exist.", 404)
            for customer, code in connection.execute("SELECT customer_id, location_code FROM locations"):
                customers[customer].add(code)
            rules = {name: json.loads(value) for name, value in connection.execute(
                "SELECT rule_name, rule_value_json FROM customer_rules WHERE customer_id = ?", (customer_id,))}
        required = {"required_fields", "allowed_zones", "allowed_location_types", "location_code_pattern",
                    "max_capacity_pallets", "location_code_must_be_unique", "capacity_must_be_positive_integer"}
        if not required <= rules.keys():
            raise ValueError("Incomplete customer rules")
        for key in ("required_fields", "allowed_zones", "allowed_location_types"):
            if not isinstance(rules[key], list) or not all(isinstance(v, str) for v in rules[key]):
                raise ValueError("Expected string lists")
        re.compile(rules["location_code_pattern"])
        limit = rules["max_capacity_pallets"]
        if isinstance(limit, bool) or not isinstance(limit, (int, float)) or not math.isfinite(limit) or limit <= 0:
            raise ValueError("Invalid capacity limit")
        for key in ("location_code_must_be_unique", "capacity_must_be_positive_integer"):
            if not isinstance(rules[key], bool):
                raise ValueError("Expected boolean rule")
        return rules, customers
    except (sqlite3.Error, ValueError, TypeError, re.error) as exc:
        if isinstance(exc, ValidationError):
            raise
        raise ValidationError("REFERENCE_DATA_ERROR", "WMS database or customer rules are unavailable or invalid.", 503) from exc
