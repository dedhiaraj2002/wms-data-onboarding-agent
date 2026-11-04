"""Evidence, isolation and safety checks using temporary synthetic fixtures only."""
from copy import deepcopy
from io import BytesIO
import hashlib
import json
import sqlite3
import subprocess
import sys

from fastapi.testclient import TestClient
from openpyxl import Workbook
import pytest

from app.database import ValidationError
from app.investigation import (compare_record, investigate_report, investigate_workbook,
                              lookup_customer_rules, lookup_location, search_similar_locations)
from app.main import app, database_path
from app.validation import MAX_UPLOAD_BYTES, validate_workbook
from scripts.generate_sample_data import CUSTOMER, RULES, generate, location_rows


@pytest.fixture
def reference(tmp_path):
    return generate(tmp_path)


def excel(rows):
    book = Workbook()
    book.active.title = "locations"
    book.active.append(list(location_rows()[0]))
    for row in rows:
        book.active.append(list(row.values()))
    content = BytesIO()
    book.save(content)
    return content.getvalue()


@pytest.fixture
def client(reference):
    app.dependency_overrides[database_path] = lambda: reference[1]
    try:
        with TestClient(app) as session:
            yield session
    finally:
        app.dependency_overrides.clear()


def test_customer_rules_and_unknown_customer(reference):
    assert lookup_customer_rules(CUSTOMER, reference[1]) == RULES
    with pytest.raises(ValidationError) as error:
        lookup_customer_rules("UNKNOWN", reference[1])
    assert (error.value.code, error.value.status) == ("UNKNOWN_CUSTOMER", 404)


def test_exact_lookup_and_missing_records(reference):
    assert lookup_location(CUSTOMER, "LOC-001", reference[1]) == dict(
        customer_id=CUSTOMER, location_code="LOC-001", zone="AMBIENT", location_type="PICK", capacity_pallets=10)
    assert lookup_location(CUSTOMER, "LOC-999", reference[1]) is None
    assert lookup_location("UNKNOWN", "LOC-001", reference[1]) is None
    assert lookup_location(CUSTOMER, "' OR 1=1 --", reference[1]) is None


def test_customer_isolation(reference):
    # Test-only setup adds another tenant to a temporary DB, never the project fixture.
    with sqlite3.connect(reference[1]) as connection:
        connection.execute("INSERT INTO customers VALUES (?, ?)", ("OTHER", "Synthetic tenant"))
        connection.execute("INSERT INTO locations VALUES (?, ?, ?, ?, ?)", ("OTHER", "LOC-001", "PRIVATE", "PICK", 9))
    assert lookup_location("OTHER", "LOC-001", reference[1])["zone"] == "PRIVATE"
    assert lookup_location(CUSTOMER, "LOC-001", reference[1])["zone"] == "AMBIENT"
    candidates = search_similar_locations(CUSTOMER, "LOC-999", reference[1])
    assert all(item["location"]["customer_id"] == CUSTOMER for item in candidates)
    row = dict(location_rows()[0], customer_id="OTHER")
    result = investigate_workbook(excel([row]), CUSTOMER, reference[1])["investigation"]["records"][0]
    assert result["wms_evidence"] == dict(lookup_performed=False, existing_location=None, similar_locations=[])
    assert result["comparison"] is None
    assert "PRIVATE" not in json.dumps(result)


def test_similar_candidates_are_bounded_deterministic_and_explained(reference):
    candidates = search_similar_locations(CUSTOMER, "LOC-999", reference[1], limit=2)
    assert [item["location"]["location_code"] for item in candidates] == ["LOC-001", "LOC-100"]
    assert all(item["matching_reason"] == "Shares exact first three characters 'LOC'." for item in candidates)
    assert all(item["is_replacement"] is False for item in candidates)
    assert candidates == search_similar_locations(CUSTOMER, "LOC-999", reference[1], limit=2)
    assert all(item["location"]["location_code"] != "LOC-001"
               for item in search_similar_locations(CUSTOMER, "LOC-001", reference[1]))
    assert search_similar_locations(CUSTOMER, "BAD-999", reference[1]) == []
    assert search_similar_locations("UNKNOWN", "LOC-999", reference[1]) == []


@pytest.mark.parametrize("code", [None, "", "  ", "L", 123])
def test_similar_search_missing_or_short_code(reference, code):
    assert search_similar_locations(CUSTOMER, code, reference[1]) == []


@pytest.mark.parametrize("limit", [0, 21, True, "5"])
def test_similar_search_rejects_unsupported_limits(reference, limit):
    with pytest.raises(ValueError):
        search_similar_locations(CUSTOMER, "LOC-999", reference[1], limit=limit)


def test_comparison_preserves_originals(reference):
    existing = lookup_location(CUSTOMER, "LOC-001", reference[1])
    uploaded = dict(existing, capacity_pallets="10", zone="UNKNOWN")
    before = deepcopy((uploaded, existing))
    comparison = compare_record(uploaded, existing)
    assert [item["field"] for item in comparison["matching_fields"]] == ["customer_id", "location_code", "location_type"]
    assert comparison["differing_fields"] == [
        dict(field="zone", uploaded_value="UNKNOWN", wms_value="AMBIENT"),
        dict(field="capacity_pallets", uploaded_value="10", wms_value=10)]
    assert (uploaded, existing) == before
    assert compare_record(dict(existing, capacity_pallets=True), dict(existing, capacity_pallets=1))["differing_fields"]


def test_fixture_investigates_all_supported_exceptions_readonly(reference):
    source, db = reference
    digest = hashlib.sha256(db.read_bytes()).hexdigest()
    report = investigate_workbook(source.read_bytes(), CUSTOMER, db)
    assert report["validation"]["total_records"] == 20
    assert report["validation"]["invalid_records"] == 9
    investigation = report["investigation"]
    assert investigation["investigated_records"] == 9
    assert investigation["findings_by_category"] == {
        "REQUEST_CUSTOMER_CLARIFICATION": 2, "REVIEW_BUSINESS_RULE": 5,
        "REVIEW_DUPLICATE": 2, "REVIEW_EXISTING_LOCATION": 2}
    assert {finding["exception_code"] for row in investigation["records"] for finding in row["findings"]} == {
        "LOCATION_EXISTS", "DUPLICATE_LOCATION", "INVALID_ZONE", "INVALID_LOCATION_TYPE",
        "INVALID_LOCATION_CODE", "INVALID_CAPACITY", "CAPACITY_EXCEEDED", "REQUIRED_FIELD"}
    assert hashlib.sha256(db.read_bytes()).hexdigest() == digest
    json.dumps(report, allow_nan=False)


def test_duplicate_rows_existing_conflict_and_comparison(reference):
    report = investigate_workbook(reference[0].read_bytes(), CUSTOMER, reference[1])
    records = {row["excel_row"]: row for row in report["investigation"]["records"]}
    for current, other in [(2, 14), (14, 2)]:
        row = records[current]
        duplicate = next(item for item in row["findings"] if item["exception_code"] == "DUPLICATE_LOCATION")
        assert duplicate["evidence"] == {"other_excel_rows": [other]}
        assert row["wms_evidence"]["existing_location"]["location_code"] == "LOC-001"
        assert row["comparison"]["differing_fields"] == []
        assert row["next_action_category"] == "MANUAL_REVIEW"


@pytest.mark.parametrize("error,rule,value", [
    ("INVALID_ZONE", "allowed_zones", RULES["allowed_zones"]),
    ("INVALID_LOCATION_TYPE", "allowed_location_types", RULES["allowed_location_types"]),
    ("INVALID_LOCATION_CODE", "location_code_pattern", RULES["location_code_pattern"]),
    ("INVALID_CAPACITY", "capacity_must_be_positive_integer", True),
    ("CAPACITY_EXCEEDED", "max_capacity_pallets", 20),
    ("REQUIRED_FIELD", "missing_field", "location_code"),
])
def test_findings_contain_actual_rule_evidence(reference, error, rule, value):
    report = investigate_workbook(reference[0].read_bytes(), CUSTOMER, reference[1])
    finding = next(item for row in report["investigation"]["records"] for item in row["findings"] if item["exception_code"] == error)
    assert finding["evidence"][rule] == value
    if error == "REQUIRED_FIELD":
        assert finding["evidence"]["original_value"] is None


def test_no_report_mutation_and_missing_current_wms_evidence(reference):
    validation = validate_workbook(reference[0].read_bytes(), CUSTOMER, reference[1])
    # Simulate stale validation evidence without changing the reference database.
    validation["rows"][0]["values"]["location_code"] = "LOC-999"
    before = deepcopy(validation)
    result = investigate_report(validation, CUSTOMER, reference[1])["records"][0]
    assert result["wms_evidence"]["existing_location"] is None
    assert result["comparison"] is None
    finding = next(item for item in result["findings"] if item["exception_code"] == "LOCATION_EXISTS")
    assert "could not be confirmed" in finding["explanation"]
    assert validation == before
    assert result["original_values"] == validation["rows"][0]["values"]


def test_missing_customer_unknown_row_customer_and_unknown_exception(reference):
    rows = [dict(location_rows()[1], customer_id=None), dict(location_rows()[2], customer_id="UNKNOWN")]
    report = investigate_workbook(excel(rows), CUSTOMER, reference[1])
    assert report["investigation"]["investigated_records"] == 2
    assert all(not row["wms_evidence"]["lookup_performed"] for row in report["investigation"]["records"])
    validation = report["validation"]
    validation["rows"][0]["exceptions"].append(dict(code="FUTURE_ERROR", field="zone", message="Synthetic", original_value="AMBIENT"))
    result = investigate_report(validation, CUSTOMER, reference[1])
    assert result["records"][0]["findings"][-1]["next_action_category"] == "MANUAL_REVIEW"


def test_all_valid_no_investigations(reference):
    report = investigate_workbook(excel([location_rows()[1]]), CUSTOMER, reference[1])
    assert report["investigation"] == dict(investigated_records=0, findings_by_category={}, records=[])


def test_missing_database_not_created(tmp_path):
    missing = tmp_path / "missing.sqlite3"
    with pytest.raises(ValidationError) as error:
        lookup_location(CUSTOMER, "LOC-001", missing)
    assert error.value.code == "REFERENCE_DATA_ERROR"
    assert not missing.exists()


def test_api_and_existing_validation_match(client, reference):
    files = {"file": ("input.xlsx", reference[0].read_bytes())}
    response = client.post("/investigate", data={"customer_id": CUSTOMER}, files=files)
    assert response.status_code == 200
    report = response.json()
    assert report["investigation"]["investigated_records"] == 9
    assert report["validation"] == client.post("/validate", data={"customer_id": CUSTOMER}, files=files).json()
    assert client.get("/health").json() == {"status": "ok"}


@pytest.mark.parametrize("filename,content,customer,status,code", [
    ("input.csv", b"x", CUSTOMER, 400, "INVALID_FILE_FORMAT"),
    ("input.xlsx", b"x", CUSTOMER, 400, "MALFORMED_WORKBOOK"),
    ("input.xlsx", None, "UNKNOWN", 404, "UNKNOWN_CUSTOMER"),
    ("input.xlsx", b"x" * (MAX_UPLOAD_BYTES + 1), CUSTOMER, 413, "UPLOAD_TOO_LARGE"),
])
def test_api_errors(client, reference, filename, content, customer, status, code):
    response = client.post("/investigate", data={"customer_id": customer},
                           files={"file": (filename, reference[0].read_bytes() if content is None else content)})
    assert response.status_code == status
    assert response.json()["detail"]["code"] == code


def test_api_safety_protection(client):
    from zipfile import ZipFile
    content = BytesIO()
    with ZipFile(content, "w") as archive:
        archive.writestr("xl/vbaProject.bin", b"synthetic")
    response = client.post("/investigate", data={"customer_id": CUSTOMER}, files={"file": ("input.xlsx", content.getvalue())})
    assert response.status_code == 400
    assert response.json()["detail"]["code"] == "UNSAFE_WORKBOOK"
    assert client.post("/investigate").status_code == 422


def test_cli_json_and_preserves_existing_output(reference, tmp_path):
    output = tmp_path / "investigation.json"
    command = [sys.executable, "-m", "scripts.investigate_excel", str(reference[0]), "--customer-id", CUSTOMER,
               "--database", str(reference[1]), "--report", str(output)]
    run = subprocess.run(command, capture_output=True, text=True)
    assert run.returncode == 1
    assert "investigated_records: 9" in run.stdout
    assert "findings_by_category:" in run.stdout
    original = output.read_bytes()
    assert json.loads(original)["validation"]["total_records"] == 20
    assert subprocess.run(command, capture_output=True).returncode == 2
    assert output.read_bytes() == original


def test_duplicate_references_do_not_include_other_customers(reference):
    rows = [location_rows()[0], dict(location_rows()[0], customer_id="OTHER"), location_rows()[0]]
    result = investigate_workbook(excel(rows), CUSTOMER, reference[1])
    records = {row["excel_row"]: row for row in result["investigation"]["records"]}
    for current, other in [(2, 4), (4, 2)]:
        finding = next(item for item in records[current]["findings"] if item["exception_code"] == "DUPLICATE_LOCATION")
        assert finding["evidence"]["other_excel_rows"] == [other]
    assert not any(item["exception_code"] == "DUPLICATE_LOCATION" for item in records[3]["findings"])


def test_empty_workbook_and_cli_all_valid(reference, tmp_path):
    assert investigate_workbook(excel([]), CUSTOMER, reference[1])["investigation"]["records"] == []
    path = tmp_path / "valid.xlsx"
    path.write_bytes(excel([location_rows()[1]]))
    command = [sys.executable, "-m", "scripts.investigate_excel", str(path), "--customer-id", CUSTOMER,
               "--database", str(reference[1])]
    result = subprocess.run(command, capture_output=True, text=True)
    assert result.returncode == 0
    assert "investigated_records: 0" in result.stdout


def test_api_missing_database_returns_503(client, tmp_path):
    missing = tmp_path / "not-there.sqlite3"
    app.dependency_overrides[database_path] = lambda: missing
    response = client.post("/investigate", data={"customer_id": CUSTOMER}, files={"file": ("input.xlsx", excel([]))})
    assert response.status_code == 503
    assert response.json()["detail"]["code"] == "REFERENCE_DATA_ERROR"
    assert not missing.exists()
