# WMS Data Onboarding Agent

Steps 1–2 provide a minimal API, reproducible synthetic warehouse fixtures, and a deterministic Excel validation engine.
Use Python 3.12 (SQLite is included in Python). No credentials or external services are required.

## Setup and run

From the repository root:

```bash
python3.12 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python -m scripts.generate_sample_data
python -m pytest -q
python -m uvicorn app.main:app --host 0.0.0.0 --port 8000
```

`GET /health` returns HTTP 200 with `{"status":"ok"}`. This is a liveness check.
Stop the development server with Ctrl+C.

## Project structure

- `app/main.py`: FastAPI application and health route.
- `scripts/generate_sample_data.py`: synthetic Excel and SQLite fixture generator.
- `data/`: generated local fixtures (ignored by Git).
- `tests/`: API and fixture tests using temporary directories.
- `requirements.txt`: pinned Python dependencies, including the test client.

## Synthetic fixtures

`data/warehouse_locations.xlsx` contains a `locations` sheet with exactly 20 records.
Columns are `customer_id`, `location_code`, `zone`, `location_type`, and
`capacity_pallets`. Rows 1–12 are valid in isolation; `LOC-001` already exists in the
reference WMS. Rows 13–20 intentionally contain, respectively: a duplicate code,
a missing code, an unknown zone, negative capacity, text capacity, an unknown type,
a malformed code, and capacity exceeding the customer limit. These fixtures exercise the Step 2 validation engine.

`data/synthetic_wms.sqlite3` contains `customers`, `locations`, and `customer_rules`.
One synthetic customer has three existing locations (`LOC-001`, `LOC-100`,
`LOC-101`). Rules store JSON values for allowed zones and types, required fields,
code format, uniqueness, and positive integer capacity with a maximum of 20 pallets.
Foreign keys are declared; future SQLite consumers must enable
`PRAGMA foreign_keys = ON` on every connection.

Generation refuses to overwrite either existing artifact. To generate another
fixture set safely:

```bash
python -m scripts.generate_sample_data --output-dir /tmp/wms-synthetic-example
```

Tests verify the health endpoint, workbook size and deliberate defects, database
integrity and reference contents, and preservation of existing files. The seed
database is only a synthetic reference fixture: no Excel rows are loaded into it.
AI agents, frontend, approval workflows, and SQL loading remain outside the current scope.

## Step 2: deterministic Excel validation

The API and CLI share `app/validation.py`. Reference data is loaded by
`app/database.py` using SQLite URI `mode=ro`; it cannot update or create a WMS
database. SQL parameters select the requested customer's rules. No LLM, automatic
corrections, approval workflow, or WMS loading is included.

### Platform setup

Linux/macOS (Python 3.12):

```bash
python3.12 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
# Only if synthetic fixtures do not already exist:
python -m scripts.generate_sample_data
python -m pytest -q
python -m uvicorn app.main:app --host 0.0.0.0 --port 8000
```

Windows PowerShell:

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
# Only if synthetic fixtures do not already exist:
.\.venv\Scripts\python.exe -m scripts.generate_sample_data
.\.venv\Scripts\python.exe -m pytest -q
.\.venv\Scripts\python.exe -m uvicorn app.main:app --port 8000
```

### Rules and input semantics

Every workbook must contain a `locations` sheet with the five Step 1 columns.
Extra columns are retained in the report. Duplicate or non-text headers are
rejected. Missing columns are workbook errors; empty cells and whitespace-only
values are row errors. Interior blank rows are validated and retain their actual
Excel row numbers. A header-only workbook returns zero records.

The requested customer must exist. Its JSON `customer_rules` provide required
fields, the regex pattern, allowed zones and types, maximum capacity, and the
uniqueness flag. Positive integer capacity is also a mandatory Step 2 invariant:
Excel boolean values, numeric text, formulas, zero, negative and fractional values
are rejected. Integer-valued numeric cells such as 10.0 are accepted.

Each row is checked against the **requested customer's rules**, even if its
customer ID differs; mismatched IDs are always flagged. Unknown row customers
receive both `UNKNOWN_CUSTOMER` and `WRONG_CUSTOMER`. Duplicate detection counts
exact `(customer_id, location_code)` pairs and flags every occurrence when the
customer's uniqueness rule is enabled. Existing WMS conflicts use each row's
customer and exact location code. No trimming, case conversion, correction, or
formula evaluation occurs. Dates are represented as ISO strings in JSON.

### API

`POST /validate` accepts multipart fields `customer_id` and `file`:

```bash
curl -X POST http://127.0.0.1:8000/validate \
  -F 'customer_id=SYNTHETIC_CUSTOMER' \
  -F 'file=@data/warehouse_locations.xlsx'
```

PowerShell can use `curl.exe` with the same `-F` arguments. The API returns HTTP
200 for a completed report, including reports with invalid rows. HTTP 400 means
invalid extension, malformed XLSX, missing `locations` sheet, or invalid/missing
headers. HTTP 404 means an unknown requested customer; 413 means a limit was
exceeded; 422 means missing form fields; 503 means unavailable or invalid reference
data. Errors have `detail.code` and `detail.message`.

Limits: 5 MiB uploaded file, 25 MiB expanded XLSX archive, 10,000 data rows and
100 columns. File size is checked before parsing; archive size is checked before
XML expansion. Uploads are handled as temporary multipart files and are closed
after validation. Parsing runs outside the async event loop.

### CLI

From the repository root, with the virtual environment active:

```bash
python -m scripts.validate_excel data/warehouse_locations.xlsx \
  --customer-id SYNTHETIC_CUSTOMER \
  --database data/synthetic_wms.sqlite3 \
  --report /tmp/wms-validation-report.json
```

PowerShell (using a new output filename):

```powershell
.\.venv\Scripts\python.exe -m scripts.validate_excel data/warehouse_locations.xlsx --customer-id SYNTHETIC_CUSTOMER --report "$env:TEMP\wms-validation-report.json"
```

`--report` is optional and refuses overwrites. Exit codes: 0 for all valid (or an
empty workbook), 1 for completed validation containing invalid records, 2 for
input/configuration/output errors. The CLI prints total, valid, invalid, and
exception counts.

### Example result

Example input: `SYNTHETIC_CUSTOMER, LOC-200, AMBIENT, PICK, 10` in the `locations`
sheet. The full JSON report for this one-row workbook is:

```json
{
  "total_records": 1,
  "valid_records": 1,
  "invalid_records": 0,
  "error_counts": {},
  "rows": [{
    "excel_row": 2,
    "status": "VALID",
    "values": {
      "customer_id": "SYNTHETIC_CUSTOMER",
      "location_code": "LOC-200",
      "zone": "AMBIENT",
      "location_type": "PICK",
      "capacity_pallets": 10
    },
    "exceptions": []
  }]
}
```

Each exception contains `code`, `message`, `field`, and `original_value`. Counts
count individual exceptions, so their sum can exceed the invalid row count.
The unchanged 20-row synthetic workbook yields **11 valid and 9 invalid records**:
both `LOC-001` rows have `DUPLICATE_LOCATION` and `LOCATION_EXISTS`; the other seven
intentional defects have their respective field errors. All errors on a record
are collected. Tests cover the fixture, edge cases, dynamic rules, database
read-only enforcement, API errors, CLI output, and Step 1 compatibility.
