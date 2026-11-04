"""Stable, strict schemas for recommendations and read-only tool arguments."""
from typing import Literal

from pydantic import BaseModel, ConfigDict, Field


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


class Recommendation(StrictModel):
    excel_row: int = Field(ge=2)
    customer_id: Scalar
    location_code: Scalar
    exception_codes: list[str]
    suggested_next_action: Action
    explanation: str = Field(min_length=1, max_length=12000)
    evidence_references: list[str] = Field(min_length=1)
    uncertainty: list[str]
    requires_human_approval: Literal[True]
    # This MVP never proposes value changes. Human review is always required.
    proposed_changes: list[str] = Field(max_length=0)
