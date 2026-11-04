"""Validate synthetic input without running the API server."""
import argparse
import json
from pathlib import Path

from app.database import DEFAULT_DATABASE, ValidationError
from app.validation import MAX_UPLOAD_BYTES, validate_workbook


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("workbook", type=Path)
    parser.add_argument("--customer-id", required=True)
    parser.add_argument("--database", type=Path, default=DEFAULT_DATABASE)
    parser.add_argument("--report", type=Path, help="Create a new JSON report; never overwrite an existing file")
    args = parser.parse_args()
    try:
        if args.workbook.suffix.lower() != ".xlsx":
            raise ValidationError("INVALID_FILE_FORMAT", "Input must be an .xlsx file.")
        with args.workbook.open("rb") as stream:
            report = validate_workbook(stream.read(MAX_UPLOAD_BYTES + 1), args.customer_id, args.database)
        if args.report:
            with args.report.open("x", encoding="utf-8") as stream:
                json.dump(report, stream, indent=2, allow_nan=False)
        for key in ("total_records", "valid_records", "invalid_records", "error_counts"):
            print(f"{key}: {report[key]}")
        return 1 if report["invalid_records"] else 0
    except (OSError, ValidationError) as exc:
        print(f"Validation failed: {exc}")
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
