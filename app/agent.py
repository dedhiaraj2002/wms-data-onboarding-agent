"""Bounded orchestration with deterministic checks on all tool calls and outputs."""
from copy import deepcopy
import json
from pathlib import Path

from pydantic import ValidationError as SchemaError

from app.agent_models import Recommendation, RecommendationResponse
from app.agent_providers import SYSTEM_PROMPT, make_provider
from app.agent_tools import ToolExecutor
from app.database import DEFAULT_DATABASE, ValidationError
from app.investigation import investigate_workbook

MAX_TOOL_CALLS = 6
MAX_AGENT_RECORDS = 50


def output_contract(record: dict) -> dict:
    """Keep factual fields fixed while allowing a separately labeled interpretation."""
    codes = [item["code"] for item in record["validation_exceptions"]]
    refs = [f"row:{record['excel_row']}:finding:{index}" for index in range(1, len(record["findings"]) + 1)]
    explanations = []
    for finding in record["findings"]:
        text = finding["explanation"]
        if finding["exception_code"] == "DUPLICATE_LOCATION":
            text += " Other duplicate Excel rows: " + json.dumps(finding["evidence"]["other_excel_rows"]) + "."
        explanations.append(text)
    conflicts = {"LOCATION_EXISTS", "DUPLICATE_LOCATION"} & set(codes)
    if conflicts:
        actions = [record["next_action_category"]]
    else:
        actions = ["REQUEST_CUSTOMER_CLARIFICATION", "MANUAL_REVIEW"]
    uncertainty = ["No change is approved or applied. Human review is required.",
                   "Evidence is limited to this synthetic workbook and the current WMS records."]
    if not record["wms_evidence"]["existing_location"]:
        uncertainty.append("No exact authorized WMS record is available; no replacement value can be established.")
    if conflicts:
        simulated = "Deterministic mock simulation: review the existing-location and duplicate evidence together; do not select a row to overwrite or retain automatically."
    else:
        simulated = "Deterministic mock simulation: ask the customer to clarify the invalid value rather than choosing a permitted value on their behalf."
    template = dict(excel_row=record["excel_row"], customer_id=record["customer_id"],
        location_code=record["location_code"], exception_codes=codes, suggested_next_action=actions[0],
        explanation="\n".join(explanations), evidence_references=refs, uncertainty=uncertainty,
        agent_rationale=dict(explanation=simulated,
                             uncertainty=["Simulated rationale only; no LLM inference was performed."]),
        requires_human_approval=True, proposed_changes=[])
    return dict(allowed_actions=actions, approved_explanation=template["explanation"], template=template)


def check_recommendation(output: dict, contract: dict, executor: ToolExecutor, mode: str = "mock") -> dict:
    if not isinstance(output, dict) or output.get("requires_human_approval") is not True:
        raise ValidationError("INVALID_MODEL_OUTPUT", "Human approval must be explicitly required.", 502)
    try:
        recommendation = Recommendation.model_validate(output).model_dump()
    except SchemaError as exc:
        raise ValidationError("INVALID_MODEL_OUTPUT", "Recommendation failed the strict output schema.", 502) from exc
    expected = contract["template"]
    for field in ("excel_row", "customer_id", "location_code", "exception_codes", "explanation", "uncertainty"):
        # JSON comparison preserves distinctions such as boolean True versus integer 1.
        if json.dumps(recommendation[field], sort_keys=True) != json.dumps(expected[field], sort_keys=True):
            raise ValidationError("INVALID_MODEL_OUTPUT", "Recommendation does not match backend-grounded evidence.", 502)
    refs = recommendation["evidence_references"]
    required = set(expected["evidence_references"])
    if (recommendation["suggested_next_action"] not in contract["allowed_actions"] or len(refs) != len(set(refs))
            or not required <= set(refs) or not set(refs) <= required | executor.evidence.keys()):
        raise ValidationError("INVALID_MODEL_OUTPUT", "Recommendation references unsupported actions or evidence.", 502)
    # Labels and verified facts are assembled only from backend-owned data.
    # Free-form rationale is intentionally NOT classified as factual evidence.
    recommendation["agent_rationale"].update(source="mock_simulation" if mode == "mock" else "openai_generated",
                                             verified=False, requires_human_review=True)
    record = executor.record
    recommendation["verified_facts"] = deepcopy(dict(
        validation_exceptions=record["validation_exceptions"],
        wms_evidence=record["wms_evidence"], business_rules=record["business_rules"],
        rules_customer_id=record["rules_customer_id"], comparison=record["comparison"],
        findings=record["findings"], tool_results=executor.evidence))
    return RecommendationResponse.model_validate(recommendation).model_dump()


def analyze_report(report: dict, customer_id: str, database: Path = DEFAULT_DATABASE,
                   mode: str = "mock", provider=None) -> dict:
    """Accept a trusted Step 3 report; the public API always builds it itself."""
    if mode not in ("mock", "openai"):
        raise ValidationError("INVALID_PROVIDER_MODE", "Provider mode must be mock or openai.", 422)
    records = report["investigation"]["records"]
    if len(records) > MAX_AGENT_RECORDS:
        raise ValidationError("AGENT_RECORD_LIMIT_EXCEEDED", "Analyze at most 50 invalid records per request.", 413)
    adapter = provider if provider is not None else make_provider(mode)
    recommendations, trace, tool_evidence = [], [], {}
    for record in records:
        contract = output_contract(record)
        context = dict(authorized_customer_id=customer_id, record=deepcopy(record), output_contract=contract)
        messages = [dict(role="system", content=SYSTEM_PROMPT),
                    dict(role="user", content=json.dumps(context, allow_nan=False))]
        executor = ToolExecutor(customer_id, record, database, MAX_TOOL_CALLS)
        used_ids = set()
        for _ in range(MAX_TOOL_CALLS + 1):
            # Do not let a provider implementation mutate backend-owned evidence or policy.
            reply = adapter.reply(deepcopy(messages), deepcopy(context))
            if reply.tool_calls:
                if reply.output is not None:
                    raise ValidationError("INVALID_MODEL_OUTPUT", "Tool calls and final output cannot be mixed.", 502)
                if executor.attempts + len(reply.tool_calls) > MAX_TOOL_CALLS:
                    raise ValidationError("TOOL_LIMIT_EXCEEDED", "Agent exceeded the per-record tool-call limit.", 422)
                for call in reply.tool_calls:
                    if not isinstance(call.call_id, str) or not call.call_id or call.call_id in used_ids:
                        raise ValidationError("INVALID_MODEL_OUTPUT", "Provider returned invalid tool-call identifiers.", 502)
                    used_ids.add(call.call_id)
                messages.append(dict(role="assistant", content=None, tool_calls=[
                    dict(id=call.call_id, type="function", function=dict(name=call.name, arguments=json.dumps(call.arguments)))
                    for call in reply.tool_calls]))
                for call in reply.tool_calls:
                    result = executor.execute(call)
                    messages.append(dict(role="tool", tool_call_id=call.call_id, content=json.dumps(result, allow_nan=False)))
            else:
                if not executor.trace:
                    raise ValidationError("INVALID_MODEL_OUTPUT", "Agent must execute a read-only tool before recommending.", 502)
                recommendations.append(check_recommendation(reply.output, contract, executor, mode))
                trace.extend(executor.trace)
                tool_evidence.update(executor.evidence)
                break
        else:
            raise ValidationError("TOOL_LIMIT_EXCEEDED", "Agent did not finish within the tool-call budget.", 422)
    return dict(**deepcopy(report), provider_mode=mode, recommendations=recommendations,
                tool_execution_trace=trace, tool_evidence=tool_evidence)


def analyze_workbook(content: bytes, customer_id: str, database: Path = DEFAULT_DATABASE, mode: str = "mock") -> dict:
    report = investigate_workbook(content, customer_id, database)
    return analyze_report(report, customer_id, database, mode)
