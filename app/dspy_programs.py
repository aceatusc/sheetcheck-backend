"""
dspy_programs.py — DSPy signatures and programs for SheetCheck.

Pydantic models define the exact shape of every input and output.
DSPy uses them to:
  - Inject a JSON schema into the prompt automatically
  - Parse and validate the LLM response back into typed objects
  - Give you `.model_dump()` for free instead of manual parse_json() calls

Programs return typed Pydantic objects. Callers (utils.py) call
.model_dump() or access fields directly — no raw string parsing needed.
"""

from __future__ import annotations

import logging
from functools import lru_cache
from typing import Any, Literal, Optional, Union

import dspy
from pydantic import BaseModel, Field

logger = logging.getLogger(__name__)


# ── DSPy LM configuration ─────────────────────────────────────────────────────

def configure_dspy(provider: str, model: str, api_key: str) -> None:
    lm_id = f"{provider}/{model}"
    try:
        lm = dspy.LM(lm_id, api_key=api_key, max_tokens=8096)
        dspy.configure(lm=lm)
        logger.debug("DSPy configured: %s", lm_id)
    except Exception as exc:
        logger.error("DSPy configure failed for %s: %s", lm_id, exc)
        raise


# ── Shared sub-models ─────────────────────────────────────────────────────────

class QAPair(BaseModel):
    q: str = Field(description="A 'Why ...?' design question about this step")
    a: str = Field(description="A concise answer explaining the design choice")


class SegmentParameter(BaseModel):
    label:   str                          = Field(description="Human-readable label shown in the UI")
    key:     str                          = Field(description="Variable name or literal value this maps to in the code")
    value:   Union[str, int, float]       = Field(description="Current value of this parameter")
    type:    Literal["number", "color", "select", "text"] = Field(description="UI control type")
    options: Optional[list[str]]          = Field(default=None, description="Choices for 'select' type only")


class Segment(BaseModel):
    id:               str                    = Field(description="Unique segment identifier, e.g. 'seg-1'")
    description:      str                    = Field(description="Short imperative label, e.g. 'Write header row'")
    sheet_context:    list[str]              = Field(description="Excel range addresses this segment touches, e.g. ['A1:E1']")
    explanation:      str                    = Field(description="1-2 sentences describing what the code does: inputs → outputs")
    predecessors:     list[str]              = Field(default_factory=list, description="IDs of segments this one depends on")
    qa_pairs:         list[QAPair]           = Field(default_factory=list, description="2-3 design Q&A pairs for this step")
    edit_suggestions: list[str]              = Field(default_factory=list, description="2-3 short prompts for edits the user might want")
    parameters:       list[SegmentParameter] = Field(default_factory=list, description="Tweakable constants hardcoded in the code")
    code:             str                    = Field(description="Office JS: await Excel.run(async (ctx) => { ... await ctx.sync(); });")
    undo_code:        str                    = Field(default="", description="Optional Office JS to reverse this segment")


# ── Output models ─────────────────────────────────────────────────────────────

class SegmentList(BaseModel):
    segments: list[Segment] = Field(description="Ordered list of code segments to execute")


class AskAnswer(BaseModel):
    answer:               str       = Field(description="Clear, concise answer in 1-3 sentences")
    follow_up_questions:  list[str] = Field(description="2 short suggested follow-up questions")


class RubricItem(BaseModel):
    id:      str  = Field(description="Requirement ID, e.g. 'h1' or 's2'")
    label:   str  = Field(description="Human-readable requirement text")
    checked: bool = Field(default=False, description="Whether this requirement is checked off")


class Rubric(BaseModel):
    hard_requirements: list[RubricItem] = Field(description="Must-have correctness criteria (1-2 items)")
    soft_requirements: list[RubricItem] = Field(description="Nice-to-have quality criteria (1-2 items)")


class VerifyResult(BaseModel):
    id:         str       = Field(description="Rubric item ID this result corresponds to")
    met:        bool      = Field(description="Whether the worksheet satisfies this requirement")
    reasoning:  str       = Field(description="One sentence explanation")
    references: list[str] = Field(description="Cell ranges that support the verdict, e.g. ['A1:E1']")


class VerifyResultList(BaseModel):
    results: list[VerifyResult] = Field(description="One entry per rubric item (hard and soft)")


# ── Worksheet context input model ─────────────────────────────────────────────

class WorksheetContext(BaseModel):
    """Structured representation of the current worksheet state passed from the add-in."""
    sheetData:    Optional[list[list[Any]]] = Field(default=None, description="2-D array of cell values")
    namedRanges:  Optional[dict[str, str]]  = Field(default=None, description="Named range → address mapping")
    activeCell:   Optional[str]             = Field(default=None, description="Currently selected cell address")
    sheetNames:   Optional[list[str]]       = Field(default=None, description="All sheet names in the workbook")


# ── Rubric hint input model ───────────────────────────────────────────────────

class RubricHint(BaseModel):
    hard_must_satisfy: list[str] = Field(default_factory=list)
    soft_nice_to_have: list[str] = Field(default_factory=list)


# ── Signatures ────────────────────────────────────────────────────────────────

class GenerateSegments(dspy.Signature):
    """
    Generate a sequence of Excel Office JS code segments that together
    accomplish the user's spreadsheet task. Each segment must be self-contained,
    independently executable, and include a clear explanation and Q&A pairs.
    """
    user_message: str              = dspy.InputField(desc="What the user wants to do in the spreadsheet")
    ws_context:   WorksheetContext = dspy.InputField(desc="Current worksheet state")
    rubric_hint:  RubricHint       = dspy.InputField(desc="Optional rubric requirements to satisfy (may be empty)")
    js_hint:      str              = dspy.InputField(desc="Known JS mistakes and fixes to avoid (may be empty string)")

    result: SegmentList = dspy.OutputField()


class EditSegments(dspy.Signature):
    """
    Modify the given segment based on user feedback, then regenerate all
    downstream segments so they remain consistent with the edit.
    The output must contain exactly 1 + len(remaining_segments) segments:
    the edited segment first, then the regenerated remainder in order.
    """
    user_message:       str              = dspy.InputField(desc="User's feedback describing the desired change")
    ws_context:         WorksheetContext = dspy.InputField(desc="Current worksheet state")
    original_segment:   Segment          = dspy.InputField(desc="The segment to edit")
    remaining_segments: list[Segment]    = dspy.InputField(desc="Segments that follow the edited one (may be empty)")
    js_hint:            str              = dspy.InputField(desc="Known JS mistakes and fixes to avoid (may be empty string)")

    result: SegmentList = dspy.OutputField()


class AnswerQuestion(dspy.Signature):
    """
    Answer a follow-up question about a specific step in a spreadsheet
    automation plan. Be concise and suggest natural follow-up questions.
    """
    user_message: str              = dspy.InputField(desc="The user's question")
    ws_context:   WorksheetContext = dspy.InputField(desc="Current worksheet state")
    current_step: Segment          = dspy.InputField(desc="The step the user is asking about")
    history:      list[dict]       = dspy.InputField(desc="Prior conversation turns (may be empty list)")

    result: AskAnswer = dspy.OutputField()


class ScaffoldRubric(dspy.Signature):
    """
    Generate a concise grading rubric for a spreadsheet task.
    Hard requirements are must-have correctness criteria.
    Soft requirements are nice-to-have quality criteria.
    Generate 1-2 of each.
    """
    user_message: str              = dspy.InputField(desc="Description of the spreadsheet task")
    ws_context:   WorksheetContext = dspy.InputField(desc="Current worksheet state")

    result: Rubric = dspy.OutputField()


class VerifyRubric(dspy.Signature):
    """
    Evaluate whether the current worksheet satisfies each rubric requirement.
    Be precise about cell references and give a one-sentence reasoning per item.
    """
    rubric:     Rubric             = dspy.InputField(desc="The rubric to evaluate against")
    ws_context: WorksheetContext   = dspy.InputField(desc="Current worksheet state")

    result: VerifyResultList = dspy.OutputField()


class ChatResponse(dspy.Signature):
    """Answer a general Excel / spreadsheet question helpfully and concisely."""
    user_message: str              = dspy.InputField(desc="User's question or request")
    ws_context:   WorksheetContext = dspy.InputField(desc="Current worksheet state")

    response: str = dspy.OutputField(desc="Helpful, concise answer — markdown OK")


# ── Programs ──────────────────────────────────────────────────────────────────

class SegmentProgram(dspy.Module):
    def __init__(self):
        self.predict = dspy.ChainOfThought(GenerateSegments)

    def forward(self, **kwargs) -> SegmentList:
        return self.predict(**kwargs).result


class EditProgram(dspy.Module):
    def __init__(self):
        self.predict = dspy.ChainOfThought(EditSegments)

    def forward(self, **kwargs) -> SegmentList:
        return self.predict(**kwargs).result


class AskProgram(dspy.Module):
    def __init__(self):
        self.predict = dspy.ChainOfThought(AnswerQuestion)

    def forward(self, **kwargs) -> AskAnswer:
        return self.predict(**kwargs).result


class RubricScaffoldProgram(dspy.Module):
    def __init__(self):
        self.predict = dspy.ChainOfThought(ScaffoldRubric)

    def forward(self, **kwargs) -> Rubric:
        return self.predict(**kwargs).result


class RubricVerifyProgram(dspy.Module):
    def __init__(self):
        self.predict = dspy.ChainOfThought(VerifyRubric)

    def forward(self, **kwargs) -> VerifyResultList:
        return self.predict(**kwargs).result


class ChatProgram(dspy.Module):
    def __init__(self):
        self.predict = dspy.ChainOfThought(ChatResponse)

    def forward(self, **kwargs) -> str:
        return self.predict(**kwargs).response


# ── Registry ──────────────────────────────────────────────────────────────────

_PROGRAMS: dict[str, type[dspy.Module]] = {
    "code":            SegmentProgram,
    "edit":            EditProgram,
    "ask":             AskProgram,
    "rubric_scaffold": RubricScaffoldProgram,
    "rubric_verify":   RubricVerifyProgram,
    "chat":            ChatProgram,
}


@lru_cache(maxsize=None)
def get_program(endpoint: str) -> dspy.Module:
    cls = _PROGRAMS.get(endpoint)
    if cls is None:
        raise ValueError(f"No DSPy program for endpoint '{endpoint}'")
    return cls()
