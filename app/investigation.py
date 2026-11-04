"""Read-only evidence gathering. Findings classify review work, never corrections."""
from collections import Counter, defaultdict
from contextlib import closing
from copy import deepcopy
from pathlib import Path
import sqlite3

from app.database import DEFAULT_DATABASE, ValidationError, connect_readonly, load_reference
from app.validation import COLUMNS, validate_workbook

LOCATION_FIELDS = ", ".join(COLUMNS)


def lookup_customer_rules(customer_id: str, database: Path = DEFAULT_DATABASE) -> dict:
    """Reuse Step 2's customer existence and rule-schema checks."""
    rules, _ = load_reference(database, customer_id)
    return rules


def _query_locations(database: Path, sql: str, parameters: tuple) -> list[dict]:
    try:
        with closing(connect_readonly(database)) as connection:
            connection.row_factory = sqlite3.Row
            return [dict(row) for row in connection.execute(sql, parameters)]
    except sqlite3.Error as exc:
        raise ValidationError("REFERENCE_DATA_ERROR", "WMS location evidence is unavailable.", 503) from exc


def lookup_location(customer_id: str, location_code: str, database: Path = DEFAULT_DATABASE) -> dict | None:
    """Find an exact customer/code pair; absence is evidence, not an error."""
    rows = _query_locations(database,
        f"SELECT {LOCATION_FIELDS} FROM locations WHERE customer_id = ? AND location_code = ?",
        (customer_id, location_code))
    return rows[0] if rows else None


def search_similar_locations(customer_id: str, location_code: str, database: Path = DEFAULT_DATABASE,
                             limit: int = 5) -> list[dict]:
    """Return at most 20 prefix candidates, ordered by exact code, excluding exact matches."""
    if isinstance(limit, bool) or not isinstance(limit, int) or not 1 <= limit <= 20:
        raise ValueError("Candidate limit must be an integer from 1 to 20.")
    if not isinstance(location_code, str) or len(location_code) < 3 or not location_code.strip():
        return []
    prefix = location_code[:3]
    rows = _query_locations(database,
        f"SELECT {LOCATION_FIELDS} FROM locations WHERE customer_id = ? "
        "AND substr(location_code, 1, 3) = ? AND location_code <> ? ORDER BY location_code LIMIT ?",
        (customer_id, prefix, location_code, limit))
    return [dict(location=row, matching_reason=f"Shares exact first three characters {prefix!r}.",
                 is_replacement=False) for row in rows]


def compare_record(uploaded: dict, existing: dict) -> dict:
    """Compare stored fields without converting or correcting either record."""
    matching, differing = [], []
    for field in COLUMNS:
        left, right = uploaded.get(field), existing.get(field)
        item = dict(field=field, uploaded_value=deepcopy(left), wms_value=deepcopy(right))
        same = left == right and (isinstance(left, bool) == isinstance(right, bool))
        (matching if same else differing).append(item)
    return dict(matching_fields=matching, differing_fields=differing)


def investigate_report(validation: dict, customer_id: str, database: Path = DEFAULT_DATABASE) -> dict:
    """Investigate invalid rows from an unmodified Step 2 validation report."""
    rules = lookup_customer_rules(customer_id, database)
    duplicate_rows = defaultdict(list)
    for row in validation["rows"]:
        values = row["values"]
        duplicate_rows[(values.get("customer_id"), values.get("location_code"))].append(row["excel_row"])
    investigations = []
    categories = Counter()
    for row in validation["rows"]:
        if row["status"] != "INVALID":
            continue
        values = row["values"]
        row_customer, code = values.get("customer_id"), values.get("location_code")
        # Never fetch another customer's WMS data because an uploaded row names it.
        scoped = row_customer == customer_id and isinstance(code, str) and bool(code.strip())
        existing = lookup_location(customer_id, code, database) if scoped else None
        candidates = search_similar_locations(customer_id, code, database) if scoped else []
        findings = []
        for exception in row["exceptions"]:
            error, field = exception["code"], exception["field"]
            category = "MANUAL_REVIEW"
            evidence = {}
            explanation = "No specialized investigation is available for this exception. Review the original validation evidence."
            if error == "LOCATION_EXISTS":
                category = "REVIEW_EXISTING_LOCATION"
                evidence = {"existing_location": existing}
                explanation = ("An exact customer/location match was retrieved from WMS." if existing else
                               "No scoped WMS match is currently available; the validation conflict could not be confirmed.")
            elif error == "DUPLICATE_LOCATION":
                category = "REVIEW_DUPLICATE"
                others = [number for number in duplicate_rows[(row_customer, code)] if number != row["excel_row"]]
                evidence = {"other_excel_rows": others}
                explanation = "Workbook rows share this exact customer and location code; no row is selected as authoritative."
            elif error in ("INVALID_ZONE", "INVALID_LOCATION_TYPE"):
                category = "REVIEW_BUSINESS_RULE"
                rule = "allowed_zones" if error == "INVALID_ZONE" else "allowed_location_types"
                evidence = {rule: rules[rule]}
                explanation = "The uploaded value is not in the requested customer's permitted values."
            elif error == "INVALID_LOCATION_CODE":
                category = "REQUEST_CUSTOMER_CLARIFICATION"
                evidence = {"location_code_pattern": rules["location_code_pattern"]}
                explanation = "The uploaded code does not match the requested customer's location-code pattern. Candidates are not replacements."
            elif error in ("INVALID_CAPACITY", "CAPACITY_EXCEEDED"):
                category = "REVIEW_BUSINESS_RULE"
                evidence = {"capacity_must_be_positive_integer": rules["capacity_must_be_positive_integer"],
                            "max_capacity_pallets": rules["max_capacity_pallets"]}
                explanation = "Capacity must be an Excel positive integer number and must not exceed the customer's limit."
            elif error == "REQUIRED_FIELD":
                category = "REQUEST_CUSTOMER_CLARIFICATION"
                evidence = {"missing_field": field, "original_value": values.get(field)}
                explanation = "A required field is missing or blank. No replacement value is inferred."
            elif error in ("UNKNOWN_CUSTOMER", "WRONG_CUSTOMER"):
                category = "REQUEST_CUSTOMER_CLARIFICATION"
                evidence = {"requested_customer_id": customer_id, "row_customer_id": row_customer}
                explanation = "Customer identity requires clarification; other customers' WMS records are not retrieved."
            findings.append(dict(exception_code=error, field=field, explanation=explanation,
                                 evidence=deepcopy(evidence), next_action_category=category))
            categories[category] += 1
        actions = {finding["next_action_category"] for finding in findings}
        # A row with multiple kinds of work requires manual review of all findings.
        action = next(iter(actions)) if len(actions) == 1 else "MANUAL_REVIEW"
        investigations.append(dict(excel_row=row["excel_row"], customer_id=row_customer, location_code=code,
            original_values=deepcopy(values), validation_exceptions=deepcopy(row["exceptions"]),
            findings=findings, wms_evidence=dict(lookup_performed=scoped, existing_location=existing,
                similar_locations=candidates), rules_customer_id=customer_id, business_rules=deepcopy(rules),
            comparison=compare_record(values, existing) if existing else None, next_action_category=action))
    return dict(investigated_records=len(investigations), findings_by_category=dict(sorted(categories.items())),
                records=investigations)


def investigate_workbook(content: bytes, customer_id: str, database: Path = DEFAULT_DATABASE) -> dict:
    """Validate with Step 2, then gather evidence for invalid records only."""
    validation = validate_workbook(content, customer_id, database)
    return dict(validation=validation, investigation=investigate_report(validation, customer_id, database))
