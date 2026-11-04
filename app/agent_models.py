"""Stable, strict schemas for recommendations and read-only tool arguments."""
from typing import Annotated, Literal

from pydantic import BaseModel, ConfigDict, Field, JsonValue


class StrictModel(BaseModel):
    model_config = ConfigDict(extra="forbid", strict=True)


class CustomerArgs(StrictModel):
    customer_id: str = Field(min_length=1, max_length=256)


class LocationArgs(CustomerArgs):
    location_code: str = Field(min_length=1, max_length=512)


class SearchArgs(LocationArgs):
    limit: int = Field(ge=1, le=5)


Action = Literal["REVIEW_EXISTING_LOCATION", "REVIEW_DUPLICATE", "REQUEST_CUSTOMER_CLARIFICATION",
                 "REVIEW_BUSINESS_RULE", "MANUAL_REVIEW"]
Scalar = str | int | float | bool | None


NonblankText = Annotated[str, Field(min_length=1, max_length=1500, pattern=r"\S")]


class AgentRationale(StrictModel):
    """Unverified interpretation, never a backend-verified fact or authorization."""
    explanation: NonblankText
    uncertainty: list[NonblankText] = Field(min_length=1, max_length=5)


class LabeledRationale(AgentRationale):
    source: Literal["mock_simulation", "openai_generated"]
    verified: Literal[False]
    requires_human_review: Literal[True]


class Recommendation(StrictModel):
    """Provider output: factual fields are checked; rationale is unverified text."""
    excel_row: int = Field(ge=2)
    customer_id: Scalar
    location_code: Scalar
    exception_codes: list[str]
    suggested_next_action: Action
    explanation: str = Field(min_length=1, max_length=12000)
    agent_rationale: AgentRationale = Field(description="Unverified model interpretation and uncertainty, never verified WMS facts or authorization.")
    evidence_references: list[str] = Field(min_length=1)
    uncertainty: list[str]
    requires_human_approval: Literal[True]
    # This MVP never proposes value changes. Human review is always required.
    proposed_changes: list[str] = Field(max_length=0)


class RecommendationResponse(Recommendation):
    """Backend-assembled response; provider cannot supply labels or verified facts."""
    agent_rationale: LabeledRationale = Field(description="Unverified interpretation with backend-owned provenance and mandatory human-review labels.")
    verified_facts: dict[str, JsonValue] = Field(description="Immutable backend evidence assembled from validation, investigation and executed tool results; never supplied by the model.")
