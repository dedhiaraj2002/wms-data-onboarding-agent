from io import BytesIO
import hashlib
import json
import sqlite3
import subprocess
import sys

from fastapi.testclient import TestClient
from openpyxl import Workbook
import pytest

from app.database import ValidationError, connect_readonly
from app.main import app, database_path
from app.validation import MAX_UPLOAD_BYTES, validate_workbook
from scripts.generate_sample_data import CUSTOMER, generate


@pytest.fixture
def reference(tmp_path):
    return generate(tmp_path)


def workbook(rows, headers=None, sheet="locations"):
    book = Workbook()
    book.active.title = sheet
    book.active.append(headers if headers is not None else ["customer_id", "location_code", "zone", "location_type", "capacity_pallets"])
    for row in rows:
        book.active.append(row)
    output = BytesIO()
    book.save(output)
    return output.getvalue()


def valid(**changes):
    row = dict(customer_id=CUSTOMER, location_code="LOC-200", zone="AMBIENT", location_type="PICK", capacity_pallets=10)
    row.update(changes)
    return list(row.values())


def codes(row):
    return {error["code"] for error in row["exceptions"]}


def test_valid_and_read_only(reference):
    _, db = reference
    before = hashlib.sha256(db.read_bytes()).hexdigest()
    report = validate_workbook(workbook([valid()]), CUSTOMER, db)
    assert report["valid_records"] == 1
    assert report["rows"][0]["excel_row"] == 2
    assert hashlib.sha256(db.read_bytes()).hexdigest() == before
    connection = connect_readonly(db)
    try:
        with pytest.raises(sqlite3.OperationalError, match="readonly"):
            connection.execute("DELETE FROM locations")
    finally:
        connection.close()


def test_synthetic_all_defects(reference):
    source, db = reference
    report = validate_workbook(source.read_bytes(), CUSTOMER, db)
    assert (report["total_records"], report["valid_records"], report["invalid_records"]) == (20, 11, 9)
    expected = ["DUPLICATE_LOCATION", "REQUIRED_FIELD", "INVALID_ZONE", "INVALID_CAPACITY",
                "INVALID_CAPACITY", "INVALID_LOCATION_TYPE", "INVALID_LOCATION_CODE", "CAPACITY_EXCEEDED"]
    for row, error in zip(report["rows"][12:], expected):
        assert error in codes(row)
    for index in (0, 12):
        assert codes(report["rows"][index]) == {"DUPLICATE_LOCATION", "LOCATION_EXISTS"}
    assert report["rows"][12]["excel_row"] == 14
    assert report["error_counts"]["LOCATION_EXISTS"] == 2
    json.dumps(report, allow_nan=False)


@pytest.mark.parametrize("capacity", [-1, 0, 1.5, True, False, "many", "10"])
def test_invalid_capacity(reference, capacity):
    report = validate_workbook(workbook([valid(capacity_pallets=capacity)]), CUSTOMER, reference[1])
    assert "INVALID_CAPACITY" in codes(report["rows"][0])
    assert report["rows"][0]["values"]["capacity_pallets"] == capacity


@pytest.mark.parametrize("value", [None, "", "   "])
def test_required_values(reference, value):
    report = validate_workbook(workbook([valid(zone=value)]), CUSTOMER, reference[1])
    assert "REQUIRED_FIELD" in codes(report["rows"][0])


def test_multiple_errors(reference):
    report = validate_workbook(workbook([valid(location_code="bad", zone="bad", capacity_pallets=-1)]), CUSTOMER, reference[1])
    assert codes(report["rows"][0]) == {"INVALID_LOCATION_CODE", "INVALID_ZONE", "INVALID_CAPACITY"}


def test_customer_scoped_duplicates_and_unknown_customer(reference):
    rows = [valid(), valid(customer_id="OTHER"), valid()]
    report = validate_workbook(workbook(rows), CUSTOMER, reference[1])
    assert "DUPLICATE_LOCATION" in codes(report["rows"][0])
    assert "DUPLICATE_LOCATION" in codes(report["rows"][2])
    assert codes(report["rows"][1]) == {"UNKNOWN_CUSTOMER", "WRONG_CUSTOMER"}


def test_dynamic_customer_rules(reference):
    db = reference[1]
    with sqlite3.connect(db) as connection:
        connection.execute("INSERT INTO customers VALUES (?, ?)", ("OTHER", "Synthetic second customer"))
        connection.execute("UPDATE customer_rules SET rule_value_json = ? WHERE rule_name = ?", (json.dumps(["CUSTOM"]), "allowed_zones"))
        connection.execute("UPDATE customer_rules SET rule_value_json = ? WHERE rule_name = ?", (json.dumps("^CUSTOM-[0-9]+$"), "location_code_pattern"))
        connection.execute("UPDATE customer_rules SET rule_value_json = ? WHERE rule_name = ?", ("5", "max_capacity_pallets"))
    report = validate_workbook(workbook([valid(location_code="CUSTOM-1", zone="CUSTOM", capacity_pallets=5), valid(customer_id="OTHER")]), CUSTOMER, db)
    assert report["rows"][0]["status"] == "VALID"
    assert "WRONG_CUSTOMER" in codes(report["rows"][1])
    assert "UNKNOWN_CUSTOMER" not in codes(report["rows"][1])
    report = validate_workbook(workbook([valid()]), CUSTOMER, db)
    assert codes(report["rows"][0]) == {"INVALID_ZONE", "INVALID_LOCATION_CODE", "CAPACITY_EXCEEDED"}


@pytest.mark.parametrize("content,error", [
    (b"broken", "MALFORMED_WORKBOOK"),
    (workbook([], sheet="wrong"), "MISSING_SHEET"),
    (workbook([], headers=["customer_id"]), "MISSING_COLUMNS"),
    (workbook([], headers=["zone", "zone"]), "INVALID_HEADERS"),
    (b"x" * (MAX_UPLOAD_BYTES + 1), "UPLOAD_TOO_LARGE"),
])
def test_bad_workbooks(reference, content, error):
    with pytest.raises(ValidationError) as caught:
        validate_workbook(content, CUSTOMER, reference[1])
    assert caught.value.code == error


def test_unknown_requested_customer(reference):
    with pytest.raises(ValidationError) as caught:
        validate_workbook(workbook([valid()]), "' OR 1=1 --", reference[1])
    assert caught.value.code == "UNKNOWN_CUSTOMER"


def test_api(reference):
    app.dependency_overrides[database_path] = lambda: reference[1]
    try:
        with TestClient(app) as client:
            assert client.get("/health").json() == {"status": "ok"}
            response = client.post("/validate", data={"customer_id": CUSTOMER}, files={"file": ("input.xlsx", reference[0].read_bytes())})
            assert response.status_code == 200
            assert response.json()["invalid_records"] == 9
            for name, content, customer, status, code in [
                ("input.csv", b"x", CUSTOMER, 400, "INVALID_FILE_FORMAT"),
                ("input.xlsx", b"x", CUSTOMER, 400, "MALFORMED_WORKBOOK"),
                ("input.xlsx", workbook([], sheet="bad"), CUSTOMER, 400, "MISSING_SHEET"),
                ("input.xlsx", workbook([]), "UNKNOWN", 404, "UNKNOWN_CUSTOMER"),
                ("input.xlsx", b"x" * (MAX_UPLOAD_BYTES + 1), CUSTOMER, 413, "UPLOAD_TOO_LARGE"),
            ]:
                response = client.post("/validate", data={"customer_id": customer}, files={"file": (name, content)})
                assert response.status_code == status
                assert response.json()["detail"]["code"] == code
            assert client.post("/validate").status_code == 422
    finally:
        app.dependency_overrides.clear()


def test_cli(reference, tmp_path):
    report = tmp_path / "report.json"
    command = [sys.executable, "-m", "scripts.validate_excel", str(reference[0]), "--customer-id", CUSTOMER,
               "--database", str(reference[1]), "--report", str(report)]
    result = subprocess.run(command, capture_output=True, text=True)
    assert result.returncode == 1
    assert "invalid_records: 9" in result.stdout
    assert json.loads(report.read_text())["valid_records"] == 11
    assert subprocess.run(command, capture_output=True).returncode == 2


def test_blank_rows_preserve_excel_row_numbers(reference):
    report = validate_workbook(workbook([valid(), [None] * 5, valid(location_code="LOC-201")]), CUSTOMER, reference[1])
    assert report["total_records"] == 3
    assert report["rows"][1]["excel_row"] == 3
    assert len(report["rows"][1]["exceptions"]) == 5
    assert report["rows"][2]["excel_row"] == 4


def test_empty_workbook_with_headers(reference):
    report = validate_workbook(workbook([]), CUSTOMER, reference[1])
    assert report == dict(total_records=0, valid_records=0, invalid_records=0, error_counts={}, rows=[])


def test_formula_not_evaluated(reference):
    report = validate_workbook(workbook([valid(capacity_pallets="=5+5")]), CUSTOMER, reference[1])
    assert "INVALID_CAPACITY" in codes(report["rows"][0])
    assert report["rows"][0]["values"]["capacity_pallets"] == "=5+5"


def test_unavailable_database_is_not_created(tmp_path):
    db = tmp_path / "absent.sqlite3"
    with pytest.raises(ValidationError) as caught:
        validate_workbook(workbook([]), CUSTOMER, db)
    assert caught.value.code == "REFERENCE_DATA_ERROR"
    assert not db.exists()


def test_invalid_rules(reference):
    with sqlite3.connect(reference[1]) as connection:
        connection.execute("UPDATE customer_rules SET rule_value_json = ? WHERE rule_name = ?", ('"["', "location_code_pattern"))
    with pytest.raises(ValidationError) as caught:
        validate_workbook(workbook([]), CUSTOMER, reference[1])
    assert caught.value.code == "REFERENCE_DATA_ERROR"


def test_workbook_row_limit(reference):
    content = workbook([valid()] * 10_001)
    with pytest.raises(ValidationError) as caught:
        validate_workbook(content, CUSTOMER, reference[1])
    assert caught.value.code == "WORKBOOK_TOO_LARGE"


def test_expanded_archive_limit(reference):
    from zipfile import ZipFile, ZIP_DEFLATED
    stream = BytesIO()
    with ZipFile(stream, "w", compression=ZIP_DEFLATED) as archive:
        archive.writestr("large.xml", b"x" * (25 * 1024 * 1024 + 1))
    with pytest.raises(ValidationError) as caught:
        validate_workbook(stream.getvalue(), CUSTOMER, reference[1])
    assert caught.value.code == "WORKBOOK_TOO_LARGE"
