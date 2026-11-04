# WMS Data Onboarding Agent

Step 1 provides a minimal API and reproducible, entirely synthetic warehouse fixtures.
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
a malformed code, and capacity exceeding the customer limit. These are fixture
examples for future validation work; there is no validation engine yet.

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
AI agents, frontend, approval workflows, and SQL loading are outside Step 1.
