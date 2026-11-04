"""Deterministic Excel validation shared by HTTP and CLI clients."""
from collections import Counter
from datetime import date, datetime, time
from io import BytesIO
from pathlib import Path
import math
import re
from zipfile import ZipFile

from openpyxl import load_workbook
import pandas as pd

from app.database import DEFAULT_DATABASE, ValidationError, load_reference

MAX_UPLOAD_BYTES = 5 * 1024 * 1024
MAX_EXPANDED_BYTES = 25 * 1024 * 1024
MAX_ROWS = 10_000
MAX_COLUMNS = 100
COLUMNS = ("customer_id", "location_code", "zone", "location_type", "capacity_pallets")


def read_workbook(content: bytes) -> tuple[list[dict], list[str]]:
    if len(content) > MAX_UPLOAD_BYTES:
        raise ValidationError("UPLOAD_TOO_LARGE", "Workbook exceeds the 5 MiB upload limit.", 413)
    try:
        with ZipFile(BytesIO(content)) as archive:
            if sum(item.file_size for item in archive.infolist()) > MAX_EXPANDED_BYTES:
                raise ValidationError("WORKBOOK_TOO_LARGE", "Expanded workbook exceeds 25 MiB.", 413)
        workbook = load_workbook(BytesIO(content), read_only=True, data_only=False)
        try:
            if "locations" not in workbook.sheetnames:
                raise ValidationError("MISSING_SHEET", "Workbook must contain a 'locations' sheet.")
            sheet = workbook["locations"]
            if (sheet.max_row or 0) > MAX_ROWS + 1 or (sheet.max_column or 0) > MAX_COLUMNS:
                raise ValidationError("WORKBOOK_TOO_LARGE", "Workbook exceeds 10,000 rows or 100 columns.", 413)
            iterator = sheet.iter_rows(values_only=True)
            headers = next(iterator, ())
            if any(not isinstance(h, str) for h in headers) or len(set(headers)) != len(headers):
                raise ValidationError("INVALID_HEADERS", "Headers must be unique, nonempty text values.")
            # openpyxl preserves Excel booleans; pandas' inference can otherwise turn them into integers.
            values = []
            for index, values_row in enumerate(iterator):
                if index >= MAX_ROWS or len(values_row) > MAX_COLUMNS:
                    raise ValidationError("WORKBOOK_TOO_LARGE", "Workbook exceeds 10,000 rows or 100 columns.", 413)
                values.append(values_row)
            frame = pd.DataFrame(values, columns=headers, dtype=object)
            records = frame.to_dict(orient="records")
            for row in records:
                for key, value in row.items():
                    if isinstance(value, (date, datetime, time)):
                        row[key] = value.isoformat()
            return records, list(headers)
        finally:
            workbook.close()
    except ValidationError:
        raise
    except Exception as exc:
        raise ValidationError("MALFORMED_WORKBOOK", "Upload is not a readable .xlsx workbook.") from exc


def missing(value) -> bool:
    return value is None or (isinstance(value, str) and not value.strip()) or (
        isinstance(value, float) and math.isnan(value))


def validate_workbook(content: bytes, customer_id: str, database: Path = DEFAULT_DATABASE) -> dict:
    records, headers = read_workbook(content)
    rules, customers = load_reference(database, customer_id)
    required = list(dict.fromkeys([*COLUMNS, *rules["required_fields"]]))
    absent = [field for field in required if field not in headers]
    if absent:
        raise ValidationError("MISSING_COLUMNS", "Missing required Excel columns: " + ", ".join(absent))
    duplicates = Counter((row["customer_id"], row["location_code"]) for row in records
                         if not missing(row["customer_id"]) and not missing(row["location_code"]))
    results = []
    counts = Counter()
    for row_number, row in enumerate(records, start=2):
        errors = []

        def add(code, message, field):
            value = row.get(field)
            errors.append(dict(code=code, message=message, field=field,
                               original_value=None if missing(value) and not isinstance(value, str) else value))
            counts[code] += 1

        for field in required:
            if missing(row.get(field)):
                add("REQUIRED_FIELD", "A nonempty value is required.", field)
        customer = row["customer_id"]
        code = row["location_code"]
        if not missing(customer):
            if customer not in customers:
                add("UNKNOWN_CUSTOMER", "Row customer does not exist in WMS.", "customer_id")
            if customer != customer_id:
                add("WRONG_CUSTOMER", "Row belongs to a different customer than requested.", "customer_id")
        if not missing(code):
            if not isinstance(code, str) or re.fullmatch(rules["location_code_pattern"], code) is None:
                add("INVALID_LOCATION_CODE", "Code does not match the customer's location pattern.", "location_code")
            if rules["location_code_must_be_unique"] and duplicates[(customer, code)] > 1:
                add("DUPLICATE_LOCATION", "Code occurs more than once for this customer in the workbook.", "location_code")
            if code in customers.get(customer, set()):
                add("LOCATION_EXISTS", "Location already exists for this customer in WMS.", "location_code")
        for field, rule, error in (("zone", "allowed_zones", "INVALID_ZONE"),
                                    ("location_type", "allowed_location_types", "INVALID_LOCATION_TYPE")):
            if not missing(row[field]) and row[field] not in rules[rule]:
                add(error, "Value is not allowed by the selected customer's rules.", field)
        capacity = row["capacity_pallets"]
        if not missing(capacity):
            # Do not coerce numeric text, booleans or formulas into a capacity.
            numeric = not isinstance(capacity, bool) and isinstance(capacity, (int, float)) and math.isfinite(capacity)
            if not numeric or capacity <= 0 or int(capacity) != capacity:
                add("INVALID_CAPACITY", "Capacity must be a positive integer Excel number.", "capacity_pallets")
            if numeric and capacity > rules["max_capacity_pallets"]:
                add("CAPACITY_EXCEEDED", "Capacity exceeds the customer's maximum.", "capacity_pallets")
        original = {key: None if missing(value) and not isinstance(value, str) else value for key, value in row.items()}
        results.append(dict(excel_row=row_number, status="INVALID" if errors else "VALID", values=original, exceptions=errors))
    invalid = sum(row["status"] == "INVALID" for row in results)
    return dict(total_records=len(results), valid_records=len(results) - invalid, invalid_records=invalid,
                error_counts=dict(sorted(counts.items())), rows=results)
