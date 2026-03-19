"""
dspy_programs.py -- DSPy signatures and programs for SheetCheck.

Threading note
--------------
Flask serves each request in its own thread. dspy.configure() is locked
to the thread that first calls it, so we must NOT call dspy.configure()
per-request. Instead, every call_program() passes the LM via dspy.context()
which is thread-local and safe to use from any thread.

Wire format
-----------
All Pydantic models mirror the exact shapes the Office add-in sends/expects,
derived from worksheetContext.js, stepNavigator.js, and rubricManager.js.
"""

from __future__ import annotations

import logging
from functools import lru_cache
from typing import Any, Literal, Optional, Union

import dspy
from pydantic import BaseModel, ConfigDict, Field

logger = logging.getLogger(__name__)


# -- LM factory (one LM instance per provider/model, cached) ------------------
#
# We deliberately do NOT call dspy.configure() globally. Each call_program()
# uses `with dspy.context(lm=...)` which is thread-local and safe.

@lru_cache(maxsize=None)
def get_lm(provider: str, model: str, api_key: str) -> dspy.LM:
    lm_id = f"{provider}/{model}"
    logger.debug("Creating LM: %s", lm_id)
    return dspy.LM(lm_id, api_key=api_key, max_tokens=8096)


# -- Worksheet context --------------------------------------------------------
#
# worksheetContext.js gather() returns:
#   {
#     selection:  { address, values, formulas } | null,
#     sheetData:  { usedRange: { address, values } } | null,
#     namedRanges: [{ name, value }] | null,
#     sheetNames:  string[] | null,
#   }

class UsedRange(BaseModel):
    model_config = ConfigDict(extra="allow")
    address: Optional[str]              = None
    values:  Optional[list[list[Any]]]  = None


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
    selection:   Optional[Selection]  = None
    sheetData:   Optional[SheetData]  = None
    namedRanges: Optional[list[dict]] = None
    sheetNames:  Optional[list[str]]  = None


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


# -- Ask ----------------------------------------------------------------------

class StepSummary(BaseModel):
    """Partial segment shape sent by stepNavigator._onAskSend()."""
    model_config = ConfigDict(extra="allow")
    description: Optional[str] = None
    explanation: Optional[str] = None


class AskAnswer(BaseModel):
    answer:              str       = Field(description="Clear, concise answer in 1-3 sentences")
    follow_up_questions: list[str] = Field(description="2 short suggested follow-up questions")


# -- Rubric -------------------------------------------------------------------

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


class RubricHint(BaseModel):
    hard_must_satisfy: list[str] = Field(default_factory=list)
    soft_nice_to_have: list[str] = Field(default_factory=list)


# -- Signatures ----------------------------------------------------------------

class GenerateSegments(dspy.Signature):
    """
    Generate a thorough sequence of Excel Office JS code segments that together
    fully accomplish the user's spreadsheet task.

    Decompose the task into as many fine-grained steps as make sense -- prefer
    more segments over fewer. Each distinct concern should be its own segment:
    writing data, applying formulas, formatting headers, formatting data rows,
    adding a totals row, colour-coding, auto-fitting columns, etc.
    A typical task should produce 5-10 segments; complex tasks may need more.

    Each segment must be:
    - Self-contained and independently executable
    - Scoped to a single coherent concern (not a catch-all "do everything" step)
    - Include a clear explanation and 2-3 Q&A pairs
    - Include all tweakable constants as parameters[]
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

    Apply the same decomposition principle as code generation: break work into
    as many fine-grained steps as make sense. Do not collapse remaining steps
    into fewer segments just because it is an edit -- preserve or increase
    granularity where appropriate.

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


# -- Programs -----------------------------------------------------------------

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


# -- Registry -----------------------------------------------------------------

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
