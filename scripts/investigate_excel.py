"""Run validation and read-only investigation without an API server."""
import argparse
import json
from pathlib import Path

from app.database import DEFAULT_DATABASE, ValidationError
from app.investigation import investigate_workbook
from app.validation import MAX_UPLOAD_BYTES


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("workbook", type=Path)
    parser.add_argument("--customer-id", required=True)
    parser.add_argument("--database", type=Path, default=DEFAULT_DATABASE)
    parser.add_argument("--report", type=Path, help="Create a new JSON report; never overwrite")
    args = parser.parse_args()
    try:
        if args.workbook.suffix.lower() != ".xlsx":
            raise ValidationError("INVALID_FILE_FORMAT", "Input must be an .xlsx file.")
        with args.workbook.open("rb") as stream:
            report = investigate_workbook(stream.read(MAX_UPLOAD_BYTES + 1), args.customer_id, args.database)
        if args.report:
            with args.report.open("x", encoding="utf-8") as stream:
                json.dump(report, stream, indent=2, allow_nan=False)
        print(f"total_records: {report['validation']['total_records']}")
        print(f"invalid_records: {report['validation']['invalid_records']}")
        print(f"investigated_records: {report['investigation']['investigated_records']}")
        print(f"findings_by_category: {report['investigation']['findings_by_category']}")
        return 1 if report["validation"]["invalid_records"] else 0
    except (OSError, ValidationError) as exc:
        print(f"Investigation failed: {exc}")
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
