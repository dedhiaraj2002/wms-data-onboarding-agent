"""Allowlisted tool dispatch with backend-owned identity, records and evidence."""
from copy import deepcopy
from dataclasses import dataclass
from pathlib import Path

from pydantic import ValidationError as SchemaError

from app.agent_models import CustomerArgs, LocationArgs, SearchArgs
from app.database import ValidationError
from app.investigation import (compare_record, lookup_customer_rules, lookup_location,
                               search_similar_locations)

TOOL_MODELS = {
    "lookup_customer_rules": CustomerArgs,
    "lookup_location": LocationArgs,
    "search_similar_locations": SearchArgs,
    "compare_record": LocationArgs,
}
DESCRIPTIONS = {
    "lookup_customer_rules": "Read the authorized customer's rules.",
    "lookup_location": "Read an exact authorized customer/location record.",
    "search_similar_locations": "Read at most five prefix candidates; none is an approved replacement.",
    "compare_record": "Compare only the current uploaded row with its own exact WMS location. location_code must match the current row; neither record is supplied by the model.",
}


def tool_schemas() -> list[dict]:
    return [dict(type="function", function=dict(name=name, description=DESCRIPTIONS[name],
                parameters=model.model_json_schema(), strict=True)) for name, model in TOOL_MODELS.items()]


@dataclass(frozen=True)
class ToolCall:
    call_id: str
    name: str
    arguments: dict


class ToolExecutor:
    def __init__(self, customer_id: str, record: dict, database: Path, max_calls: int = 6):
        self.customer_id = customer_id
        self.record = deepcopy(record)
        self.database = database
        self.max_calls = max_calls
        self.trace = []
        self.evidence = {}
        self.attempts = 0

    def execute(self, call: ToolCall) -> dict:
        self.attempts += 1
        if self.attempts > self.max_calls:
            raise ValidationError("TOOL_LIMIT_EXCEEDED", "Agent exceeded the per-record tool-call limit.", 422)
        if call.name not in TOOL_MODELS:
            raise ValidationError("UNKNOWN_TOOL", "Agent requested a tool outside the allowlist.", 422)
        try:
            args = TOOL_MODELS[call.name].model_validate(call.arguments)
        except SchemaError as exc:
            raise ValidationError("INVALID_TOOL_ARGUMENTS", "Agent tool arguments failed schema validation.", 422) from exc
        if args.customer_id != self.customer_id:
            raise ValidationError("TOOL_SCOPE_VIOLATION", "Tool customer does not match the authorized customer.", 422)
        # An invalid row's customer is never used to widen authorization.
        if call.name != "lookup_customer_rules" and self.record["customer_id"] != self.customer_id:
            raise ValidationError("TOOL_SCOPE_VIOLATION", "Location tools are disabled for rows with mismatched customer identity.", 422)
        if call.name == "compare_record" and args.location_code != self.record["location_code"]:
            raise ValidationError("COMPARISON_TARGET_MISMATCH", "Comparison must use the current uploaded row's location code.", 422)
        if call.name == "lookup_customer_rules":
            result = lookup_customer_rules(self.customer_id, self.database)
        elif call.name == "lookup_location":
            result = lookup_location(self.customer_id, args.location_code, self.database)
        elif call.name == "search_similar_locations":
            result = search_similar_locations(self.customer_id, args.location_code, self.database, args.limit)
        else:
            existing = lookup_location(self.customer_id, self.record["location_code"], self.database)
            result = compare_record(self.record["original_values"], existing) if existing else None
        reference = f"row:{self.record['excel_row']}:tool:{len(self.trace) + 1}"
        self.evidence[reference] = deepcopy(result)
        self.trace.append(dict(excel_row=self.record["excel_row"], tool_name=call.name,
                               arguments=args.model_dump(), result_status="NOT_FOUND" if result is None else "SUCCESS",
                               evidence_reference=reference))
        return dict(evidence_reference=reference, result=deepcopy(result))
