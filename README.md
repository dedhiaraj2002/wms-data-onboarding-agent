# WMS Data Onboarding Agent

Steps 1–3 provide a minimal API, reproducible synthetic warehouse fixtures, deterministic Excel validation, and read-only WMS investigation tools.
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

### Workbook safety policy

Only plain `.xlsx` workbooks are accepted. Archive metadata and relationships are
inspected before openpyxl loads the workbook. Macro-enabled content, VBA/binary
parts, macro/dialog sheets, ActiveX/embedded objects, external-link parts, and
external relationships (including external hyperlinks) are rejected with HTTP
400 `UNSAFE_WORKBOOK`. Duplicate ZIP entries are also rejected. Existing upload,
expanded ZIP, row, and column limits still apply.

Excel error cells (such as `#N/A`), duration cells, nonfinite numbers, and other
unsupported cell values are rejected with HTTP 400 `UNSUPPORTED_CELL_VALUE`,
including the cell coordinate. Supported scalar values remain unchanged; dates
and times use ISO strings. Formulas remain unevaluated.

Customer rule schemas must contain exactly the supported rule names. Required
lists must be nonempty and contain unique, nonblank strings; required field names
must be supported. Patterns must be nonempty valid regex strings, capacity limits
must be positive finite numbers, and uniqueness must be boolean. Disabling the
mandatory positive-integer capacity setting is unsupported. Invalid or incomplete
configuration returns HTTP 503 `REFERENCE_DATA_ERROR`.

## Step 3: read-only WMS investigation

Validation answers **which rules were violated**. Investigation gathers evidence
that helps a person review those exceptions: the stored WMS location, comparisons,
permitted values, other duplicate Excel rows, and limited related-code candidates.
It never chooses a replacement, changes an input, approves a correction, or loads
anything into WMS. No LLM or agent framework is involved.

`app/investigation.py` provides these functions:

| Function | Purpose |
| --- | --- |
| `lookup_customer_rules` | Retrieve the selected customer's validated rule configuration using the Step 2 database layer |
| `lookup_location` | Parameterized exact lookup by both customer and location code; return `None` if absent |
| `search_similar_locations` | Customer-scoped, case-sensitive first-three-character prefix search; exclude the exact code |
| `compare_record` | Compare the five stored fields, preserving uploaded and WMS values in matching/differing field lists |
| `investigate_report` | Investigate invalid rows in an existing Step 2 validation report |
| `investigate_workbook` | Run Step 2 validation, then investigation; shared by API and CLI |

The standalone lookup tools require explicit customer IDs. The orchestrator uses
the requested customer's rules, just as validation does. If an uploaded row names
a different customer or lacks a customer/code, it does not look up that row's WMS
location or candidates. Duplicate references use exact `(customer_id, location_code)`
pairs, including every other involved Excel row number.

Similarity is deliberately simple: codes must share their exact first three
characters. Results are ordered by location code, default to five candidates,
and can be limited from 1 to 20 through the Python tool. Short, blank or non-text
codes yield no candidates. Each candidate includes its stored fields, matching
reason and `is_replacement: false`. Shared prefixes are weak evidence, not a claim
that any result is the intended location.

### API and CLI

Start the existing server from the repository root:

```bash
python -m uvicorn app.main:app --port 8000
```

Upload the same synthetic workbook:

```bash
curl -X POST http://127.0.0.1:8000/investigate \
  -F 'customer_id=SYNTHETIC_CUSTOMER' \
  -F 'file=@data/warehouse_locations.xlsx'
```

`POST /investigate` accepts the same multipart fields as `/validate`. Both routes
share upload handling, safety and size limits, and HTTP error mapping. The new
route calls the existing validator before gathering evidence. A completed report
returns HTTP 200 even when records are invalid; bad input and missing/invalid
reference data retain the Step 2 error codes.

For Linux/macOS, with the virtual environment activated:

```bash
python -m scripts.investigate_excel data/warehouse_locations.xlsx \
  --customer-id SYNTHETIC_CUSTOMER \
  --database data/synthetic_wms.sqlite3 \
  --report /tmp/wms-investigation.json
```

Windows PowerShell (no activation needed):

```powershell
.\.venv\Scripts\python.exe -m uvicorn app.main:app --port 8000
# In another terminal, from the repository root:
curl.exe -X POST http://127.0.0.1:8000/investigate -F "customer_id=SYNTHETIC_CUSTOMER" -F "file=@data/warehouse_locations.xlsx"
# Or run without a server:
.\.venv\Scripts\python.exe -m scripts.investigate_excel data/warehouse_locations.xlsx --customer-id SYNTHETIC_CUSTOMER --report "$env:TEMP\wms-investigation.json"
```

The optional JSON output must use a new path; existing files are never overwritten.
CLI exit codes match validation: 0 when no invalid records exist, 1 when a completed
report has invalid records, 2 on input/reference/output errors.

### Reports and review categories

The report has two top-level objects: `validation` (the unchanged Step 2 result)
and `investigation`. The latter includes `investigated_records`,
`findings_by_category`, and `records`. Each investigation record contains:

- `excel_row`, `customer_id`, `location_code`, and `original_values`.
- Original `validation_exceptions` and one evidence-based finding per exception.
- `wms_evidence`: whether a scoped lookup occurred, the exact stored location or
  `null`, and bounded similar candidates.
- `rules_customer_id`, `business_rules`, and a comparison when an exact record exists.
- `next_action_category`, a review classification only.

Finding categories are `REVIEW_EXISTING_LOCATION`, `REVIEW_DUPLICATE`,
`REQUEST_CUSTOMER_CLARIFICATION`, `REVIEW_BUSINESS_RULE`, and `MANUAL_REVIEW`.
When a row has several distinct categories, its row category is `MANUAL_REVIEW`;
all finding categories remain available. Unknown exception codes also receive
manual review. Counts are per finding, so one record can contribute more than one.

For the unchanged synthetic workbook, the summary is:

```json
{
  "total_records": 20,
  "invalid_records": 9,
  "investigated_records": 9,
  "findings_by_category": {
    "REQUEST_CUSTOMER_CLARIFICATION": 2,
    "REVIEW_BUSINESS_RULE": 5,
    "REVIEW_DUPLICATE": 2,
    "REVIEW_EXISTING_LOCATION": 2
  }
}
```

Example finding from Excel row 16 (`LOC-015`, uploaded zone `UNKNOWN`):

```json
{
  "exception_code": "INVALID_ZONE",
  "field": "zone",
  "explanation": "The uploaded value is not in the requested customer's permitted values.",
  "evidence": {"allowed_zones": ["AMBIENT", "CHILLED"]},
  "next_action_category": "REVIEW_BUSINESS_RULE"
}
```

For Excel rows 2 and 14, investigation retrieves the stored `LOC-001`, shows that
all five fields match, and reports the other duplicate row number. Neither record
is selected for retention or overwrite. Missing fields are reported without
invented values. If a validation conflict cannot be confirmed by a current scoped
lookup, the report says so and returns no comparison.

### Evidence limitations and tests

The synthetic reference database has only one customer and three stored locations.
Candidates and findings reflect only those records and that customer's rules.
An empty result does not prove that a location is invalid or identify a correction.
Validation and investigation use separate read-only connections, so changes by an
external writer between reads could make evidence differ; there is no cross-stage
snapshot guarantee. The tool reports absent evidence rather than fabricating it.
`investigate_report` expects a trusted report produced by Step 2, not arbitrary
external JSON; the HTTP route always generates that report itself.

Run `python -m pytest -q` for all Steps 1–3 tests. Investigation tests cover customer
isolation, deterministic candidate limits, exact lookups and missing records,
comparisons, every supported exception, duplicate row references, read-only
behavior, API protections, and CLI JSON output/overwrite protection. Test-only
extra tenants live in temporary databases; the existing synthetic fixtures are
unchanged. Step 4 AI reasoning is not implemented.
