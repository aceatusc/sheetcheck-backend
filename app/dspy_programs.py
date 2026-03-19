"""
dspy_programs.py -- DSPy signatures and programs for SheetCheck.

Pydantic models mirror the exact wire format the Office add-in sends and
expects. Shape decisions are driven by the front-end JS:

  WorksheetContext  -- matches WorksheetContext.gather() output in worksheetContext.js
  StepSummary       -- matches the `step` field LLMClient.ask() sends (description+explanation only)
  AskAnswer         -- matches the { answer, follow_up_questions } shape stepNavigator expects
  Rubric            -- matches the rubric shape rubricManager uses throughout
  VerifyResultList  -- matches { results: [...] } that LLMClient.rubricVerify returns
"""

from __future__ import annotations

import logging
from functools import lru_cache
from typing import Any, Literal, Optional, Union

import dspy
from pydantic import BaseModel, ConfigDict, Field

logger = logging.getLogger(__name__)


# -- DSPy LM configuration ----------------------------------------------------

def configure_dspy(provider: str, model: str, api_key: str) -> None:
    lm_id = f"{provider}/{model}"
    try:
        lm = dspy.LM(lm_id, api_key=api_key, max_tokens=8096)
        dspy.configure(lm=lm)
        logger.debug("DSPy configured: %s", lm_id)
    except Exception as exc:
        logger.error("DSPy configure failed for %s: %s", lm_id, exc)
        raise


# -- Worksheet context ---------------------------------------------------------
#
# worksheetContext.js gather() returns:
#   {
#     selection:  { address, values, formulas } | null,
#     sheetData:  { usedRange: { address, values } } | null,
#     namedRanges: [{ name, value }] | null,
#     sheetNames:  string[] | null,
#   }
#
# extra="allow" so any new fields the add-in adds never break validation.

class UsedRange(BaseModel):
    model_config = ConfigDict(extra="allow")
    address: Optional[str]         = None
    values:  Optional[list[list[Any]]] = None


class Selection(BaseModel):
    model_config = ConfigDict(extra="allow")
    address:  Optional[str]             = None
    values:   Optional[list[list[Any]]] = None
    formulas: Optional[list[list[Any]]] = None


class SheetData(BaseModel):
    model_config = ConfigDict(extra="allow")
    usedRange: Optional[UsedRange] = None


class WorksheetContext(BaseModel):
    model_config = ConfigDict(extra="allow")
    selection:   Optional[Selection]        = None
    sheetData:   Optional[SheetData]        = None
    namedRanges: Optional[list[dict]]       = None
    sheetNames:  Optional[list[str]]        = None


# -- Segment models -----------------------------------------------------------

class QAPair(BaseModel):
    q: str = Field(description="A 'Why ...?' design question about this step")
    a: str = Field(description="A concise answer explaining the design choice")


class SegmentParameter(BaseModel):
    label:   str                                          = Field(description="Human-readable label shown in the UI")
    key:     str                                          = Field(description="Variable name or literal this maps to in the code")
    value:   Union[str, int, float]                       = Field(description="Current value of this parameter")
    type:    Literal["number", "color", "select", "text"] = Field(description="UI control type")
    options: Optional[list[str]]                          = Field(default=None, description="Choices for 'select' type only")


class Segment(BaseModel):
    id:               str                    = Field(description="Unique segment identifier, e.g. 'seg-1'")
    description:      str                    = Field(description="Short imperative label, e.g. 'Write header row'")
    sheet_context:    list[str]              = Field(description="Excel range addresses this segment touches")
    explanation:      str                    = Field(description="1-2 sentences: inputs to outputs")
    predecessors:     list[str]              = Field(default_factory=list)
    qa_pairs:         list[QAPair]           = Field(default_factory=list, description="2-3 design Q&A pairs")
    edit_suggestions: list[str]              = Field(default_factory=list, description="2-3 short edit prompts")
    parameters:       list[SegmentParameter] = Field(default_factory=list, description="Tweakable constants in the code")
    code:             str                    = Field(description="await Excel.run(async (ctx) => { ... await ctx.sync(); });")
    undo_code:        str                    = Field(default="")


class SegmentList(BaseModel):
    segments: list[Segment] = Field(description="Ordered list of code segments to execute")


# -- Ask -----------------------------------------------------------------------
#
# stepNavigator.js sends: { description, explanation } as the `step` field.
# We model that as StepSummary rather than a full Segment so the LLM isn't
# asked to reconstruct fields it was never given.

class StepSummary(BaseModel):
    model_config = ConfigDict(extra="allow")
    description: Optional[str] = None
    explanation: Optional[str] = None


class AskAnswer(BaseModel):
    answer:              str       = Field(description="Clear, concise answer in 1-3 sentences")
    follow_up_questions: list[str] = Field(description="2 short suggested follow-up questions")


# -- Rubric --------------------------------------------------------------------
#
# rubricManager.js uses: { hard_requirements: [{id,label,checked}], soft_requirements: [...] }
# rubricScaffold returns this shape directly (not nested).
# rubricVerify returns { results: [{id,met,reasoning,references}] }.

class RubricItem(BaseModel):
    id:      str  = Field(description="Requirement ID, e.g. 'h1' or 's2'")
    label:   str  = Field(description="Human-readable requirement text")
    checked: bool = Field(default=False)


class Rubric(BaseModel):
    hard_requirements: list[RubricItem] = Field(description="Must-have correctness criteria (1-2 items)")
    soft_requirements: list[RubricItem] = Field(description="Nice-to-have quality criteria (1-2 items)")


class VerifyResult(BaseModel):
    id:         str       = Field(description="Rubric item ID")
    met:        bool      = Field(description="Whether the worksheet satisfies this requirement")
    reasoning:  str       = Field(description="One sentence explanation")
    references: list[str] = Field(description="Supporting cell ranges, e.g. ['A1:E1']")


class VerifyResultList(BaseModel):
    results: list[VerifyResult] = Field(description="One entry per rubric item (hard and soft)")


# -- Rubric hint passed to /code ----------------------------------------------

class RubricHint(BaseModel):
    hard_must_satisfy: list[str] = Field(default_factory=list)
    soft_nice_to_have: list[str] = Field(default_factory=list)


# -- Signatures ----------------------------------------------------------------

class GenerateSegments(dspy.Signature):
    """
    Generate a sequence of Excel Office JS code segments that together
    accomplish the user's spreadsheet task. Each segment must be self-contained,
    independently executable, and include a clear explanation and Q&A pairs.
    """
    user_message: str              = dspy.InputField(desc="What the user wants to do in the spreadsheet")
    ws_context:   WorksheetContext = dspy.InputField(desc="Current worksheet state")
    rubric_hint:  RubricHint       = dspy.InputField(desc="Optional rubric requirements to satisfy (may be empty)")
    js_hint:      str              = dspy.InputField(desc="Known JS mistakes and fixes to avoid (may be empty)")

    result: SegmentList = dspy.OutputField()


class EditSegments(dspy.Signature):
    """
    Modify the given segment based on user feedback, then regenerate all
    downstream segments so they remain consistent with the edit.
    Output must contain exactly 1 + len(remaining_segments) segments:
    the edited segment first, then the regenerated remainder in order.
    """
    user_message:       str              = dspy.InputField(desc="User's feedback describing the desired change")
    ws_context:         WorksheetContext = dspy.InputField(desc="Current worksheet state")
    original_segment:   Segment          = dspy.InputField(desc="The segment to edit")
    remaining_segments: list[Segment]    = dspy.InputField(desc="Segments that follow the edited one (may be empty)")
    js_hint:            str              = dspy.InputField(desc="Known JS mistakes and fixes to avoid (may be empty)")

    result: SegmentList = dspy.OutputField()


class AnswerQuestion(dspy.Signature):
    """
    Answer a follow-up question about a specific step in a spreadsheet
    automation plan. Be concise and suggest natural follow-up questions.
    """
    user_message: str              = dspy.InputField(desc="The user's question")
    ws_context:   WorksheetContext = dspy.InputField(desc="Current worksheet state")
    current_step: StepSummary      = dspy.InputField(desc="Description and explanation of the step being asked about")
    history:      list[dict]       = dspy.InputField(desc="Prior conversation turns [{q, a}] (may be empty)")

    result: AskAnswer = dspy.OutputField()


class ScaffoldRubric(dspy.Signature):
    """
    Generate a concise grading rubric for a spreadsheet task.
    Hard requirements are must-have correctness criteria (1-2 items).
    Soft requirements are nice-to-have quality criteria (1-2 items).
    """
    user_message: str              = dspy.InputField(desc="Description of the spreadsheet task")
    ws_context:   WorksheetContext = dspy.InputField(desc="Current worksheet state")

    result: Rubric = dspy.OutputField()


class VerifyRubric(dspy.Signature):
    """
    Evaluate whether the current worksheet satisfies each rubric requirement.
    Be precise about cell references and give one-sentence reasoning per item.
    Cover every item in both hard_requirements and soft_requirements.
    """
    rubric:     Rubric           = dspy.InputField(desc="The rubric to evaluate against")
    ws_context: WorksheetContext = dspy.InputField(desc="Current worksheet state")

    result: VerifyResultList = dspy.OutputField()


class ChatResponse(dspy.Signature):
    """Answer a general Excel / spreadsheet question helpfully and concisely."""
    user_message: str              = dspy.InputField(desc="User's question or request")
    ws_context:   WorksheetContext = dspy.InputField(desc="Current worksheet state")

    response: str = dspy.OutputField(desc="Helpful, concise answer -- markdown OK")


# -- Programs ------------------------------------------------------------------

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


# -- Registry ------------------------------------------------------------------

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
