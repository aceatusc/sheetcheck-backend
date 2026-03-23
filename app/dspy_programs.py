"""dspy_programs.py -- DSPy signatures and programs for SheetCheck."""
from __future__ import annotations

import logging
from functools import lru_cache
from typing import Any, Literal, Optional, Union

import dspy
from pydantic import BaseModel, ConfigDict, Field, field_validator

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# LM factory
# ---------------------------------------------------------------------------

MAX_TOKENS: dict[str, int] = {
    "code":            64_000,
    "edit":            32_000,
    "ask":              2_000,
    "rubric_scaffold":  8_000,
    "rubric_verify":    8_000,
    "chat":             2_000,
}
_DEFAULT_MAX_TOKENS = 8_096


@lru_cache(maxsize=None)
def get_lm(provider: str, model: str, api_key: str, endpoint: str) -> dspy.LM:
    lm_id = f"{provider}/{model}"
    return dspy.LM(lm_id, api_key=api_key, max_tokens=MAX_TOKENS.get(endpoint, _DEFAULT_MAX_TOKENS))


# ---------------------------------------------------------------------------
# Worksheet context models
# ---------------------------------------------------------------------------

class UsedRange(BaseModel):
    model_config = ConfigDict(extra="allow")
    address: Optional[str]             = None
    values:  Optional[list[list[Any]]] = None


class Selection(BaseModel):
    model_config = ConfigDict(extra="allow")
    address:  Optional[str]             = None
    values:   Optional[list[list[Any]]] = None
    formulas: Optional[list[list[Any]]] = None


class CellStyle(BaseModel):
    model_config = ConfigDict(extra="allow")
    address:             Optional[str]   = None
    fillColor:           Optional[str]   = None
    fontColor:           Optional[str]   = None
    fontBold:            Optional[bool]  = None
    fontItalic:          Optional[bool]  = None
    fontSize:            Optional[float] = None
    numberFormat:        Optional[str]   = None
    horizontalAlignment: Optional[str]   = None


class ChartInfo(BaseModel):
    model_config = ConfigDict(extra="allow")
    name:      Optional[str] = None
    chartType: Optional[str] = None
    dataRange: Optional[str] = None
    title:     Optional[str] = None


class SheetInfo(BaseModel):
    model_config = ConfigDict(extra="allow")
    name:      str                       = None
    address:   Optional[str]             = None
    values:    Optional[list[list[Any]]] = None
    rowCount:  Optional[int]             = None
    colCount:  Optional[int]             = None
    truncated: Optional[bool]            = False
    shownRows: Optional[int]             = None
    error:     Optional[str]             = None


class SheetData(BaseModel):
    model_config = ConfigDict(extra="allow")
    usedRange: Optional[UsedRange]       = None
    allSheets: Optional[list[SheetInfo]] = Field(default=None,
        description="All worksheet tabs — use for cross-sheet refs (VLOOKUP sources, dropdowns, etc.)")
    styles:    Optional[list[CellStyle]] = None
    charts:    Optional[list[ChartInfo]] = None


class WorksheetContext(BaseModel):
    model_config = ConfigDict(extra="allow")
    selection:   Optional[Selection]  = None
    sheetData:   Optional[SheetData]  = None
    namedRanges: Optional[list[dict]] = None
    sheetNames:  Optional[list[str]]  = None


# ---------------------------------------------------------------------------
# Segment models
# ---------------------------------------------------------------------------

class QAPair(BaseModel):
    q: str = Field(description="'Why …?' question in sheet terms (cells/formulas/values), not code")
    a: str = Field(description="Concise answer — addresses, formula bar, visible output")


class SegmentParameter(BaseModel):
    label:   str                                           = Field(description="UI label")
    key:     str                                           = Field(description="Variable name in code")
    value:   Union[str, int, float]                        = Field(description="Current value")
    type:    Literal["number", "color", "select", "text"]  = Field(description="UI control type")
    options: Optional[list[str]]                           = Field(default=None, description="select options only")


# Office JS rules injected into the code field description so the LLM sees them
# exactly once, at the point of generation — not duplicated in every signature.
_OFFICE_JS_RULES = (
    "Office JS snippet. Rules: "
    "(1) await Excel.run(async (ctx) => { …; await ctx.sync(); }); "
    "(2) .values/.formulas/.numberFormat need a 2-D array matching range dims: "
    "[[v]] for 1x1, [[v1,v2]] for 1x2 row. Note: titleRange.merge() still requires "
    "the full original range width (e.g., 6 cols = [[v,'','','','','']]); "
    "(3) sheet.getRange() only — never range.getRange(); "
    "(4) no range.getRow() — use sheet.getRange('A1:Z1'); "
    "(5) no getLastCell().getEnd() — use sheet.getUsedRange()+load('rowCount') for last row; "
    "(6) dropdown: rng.dataValidation.rule={list:{inCellDropDown:true,source:'=Sheet!$A$1:$A$5'}} "
    "— source must start with '=', no .add(), no sheet.dataValidation; "
    "(7) no .conditionalFormatting — loop and set format per cell; "
    "(8) autofitColumns: range.getEntireColumn().format.autofitColumns() after writing data; "
    "(9) load() before ctx.sync(); never load and write same range in one sync block."
)


class OfficeJSCode(BaseModel):
    model_config = ConfigDict(extra="allow")
    code: str = Field(description=_OFFICE_JS_RULES)

    @classmethod
    def model_validate(cls, value, **kwargs):
        if isinstance(value, str):
            return cls(code=value)
        return super().model_validate(value, **kwargs)


class Segment(BaseModel):
    id:               str                    = Field(description="Unique ID e.g. 'seg-1'")
    description:      str                    = Field(description="Short imperative label")
    sheet_context:    list[str]              = Field(description="Range addresses touched")
    explanation:      str                    = Field(description="1-2 sentences: what changes and why")
    predecessors:     list[str]              = Field(default_factory=list)
    qa_pairs:         list[QAPair]           = Field(default_factory=list, description="2-3 Q&A pairs")
    edit_suggestions: list[str]              = Field(default_factory=list, description="2-3 short edit prompts")
    parameters:       list[SegmentParameter] = Field(default_factory=list, description="Tweakable constants")
    code:             OfficeJSCode           = Field(description="Office JS for this step")
    manual_steps:     str                    = Field(default="", description="Manual Excel UI steps if automation fails")

    @field_validator("code", mode="before")
    @classmethod
    def _coerce_code(cls, v):
        if isinstance(v, str):
            return {"code": v}
        return v

    def model_dump(self, **kwargs) -> dict:
        d = super().model_dump(**kwargs)
        if isinstance(d.get("code"), dict):
            d["code"] = d["code"].get("code", "")
        return d

    @property
    def code_str(self) -> str:
        return self.code.code if isinstance(self.code, OfficeJSCode) else str(self.code)


class SegmentList(BaseModel):
    segments: list[Segment] = Field(description="Ordered segments to execute")


# ---------------------------------------------------------------------------
# Ask models
# ---------------------------------------------------------------------------

class StepSummary(BaseModel):
    model_config = ConfigDict(extra="allow")
    description: Optional[str] = None
    explanation: Optional[str] = None


class AskAnswer(BaseModel):
    answer:              str       = Field(description="1-3 sentences in sheet terms — cells, formulas, visible values; no code")
    follow_up_questions: list[str] = Field(description="2 follow-up questions about what the user sees")


# ---------------------------------------------------------------------------
# Aspect / Verify models
# ---------------------------------------------------------------------------

class Aspect(BaseModel):
    id:    str = Field(description="Unique ID e.g. 'a1'")
    label: str = Field(description="What to verify (1-2 sentences)")


class AspectList(BaseModel):
    aspects: list[Aspect] = Field(description="3-6 aspects to review")


class VerifyResult(BaseModel):
    id:         str       = Field(description="Aspect ID")
    met:        bool      = Field(description="Whether the sheet satisfies this aspect")
    reasoning:  str       = Field(description="One sentence explanation")
    references: list[str] = Field(description="Supporting cell ranges e.g. ['A1:E1']")


class VerifyResultList(BaseModel):
    results: list[VerifyResult] = Field(description="One entry per aspect")


# Backward compat — /code rubric_hint
class RubricItem(BaseModel):
    id:      str  = Field(description="Aspect ID")
    label:   str  = Field(description="Aspect text")
    checked: bool = Field(default=False)


class Rubric(BaseModel):
    hard_requirements: list[RubricItem] = Field(default_factory=list)
    soft_requirements: list[RubricItem] = Field(default_factory=list)


class RubricHint(BaseModel):
    hard_must_satisfy: list[str] = Field(default_factory=list)
    soft_nice_to_have: list[str] = Field(default_factory=list)


# ---------------------------------------------------------------------------
# Signatures
# ---------------------------------------------------------------------------

class GenerateSegments(dspy.Signature):
    """
    Generate fine-grained Office JS segments that fully accomplish the user's task.
    One concern per segment — data, formulas, header style, row style, totals,
    colour-coding, dropdowns, autofit, etc. 5-15 segments is typical; more is better.
    Every tweakable constant goes in parameters[].
    """
    user_message: str              = dspy.InputField(desc="What the user wants")
    ws_context:   WorksheetContext = dspy.InputField(desc="Full workbook state")
    rubric_hint:  RubricHint       = dspy.InputField(desc="Requirements to satisfy")
    js_hint:      str              = dspy.InputField(desc="Recent JS errors to avoid")
    chat_history: list[str]        = dspy.InputField(desc="Conversation context")

    result: SegmentList = dspy.OutputField()


class EditSegments(dspy.Signature):
    """
    Apply user feedback to the segment; regenerate downstream segments for consistency.
    Output edited segment first, then remainder in order. Keep or increase granularity.
    """
    user_message:       str              = dspy.InputField(desc="Desired change")
    ws_context:         WorksheetContext = dspy.InputField(desc="Current workbook state")
    original_segment:   Segment          = dspy.InputField(desc="Segment to edit")
    remaining_segments: list[Segment]    = dspy.InputField(desc="Downstream segments")
    js_hint:            str              = dspy.InputField(desc="Recent JS errors to avoid")
    chat_history:       list[str]        = dspy.InputField(desc="Conversation context")

    result: SegmentList = dspy.OutputField()


class AnswerQuestion(dspy.Signature):
    """
    Answer a question about a spreadsheet step.
    The user sees their Excel sheet, not code — answer in sheet terms:
    cell addresses, formula bar, visible values, dropdown options. No JavaScript.
    """
    user_message: str              = dspy.InputField(desc="User's question")
    ws_context:   WorksheetContext = dspy.InputField(desc="Workbook state")
    current_step: StepSummary      = dspy.InputField(desc="Step being asked about")
    history:      list[dict]       = dspy.InputField(desc="Prior Q&A turns [{q, a}]")
    chat_history: list[str]        = dspy.InputField(desc="Conversation context")

    result: AskAnswer = dspy.OutputField()


class ScaffoldAspects(dspy.Signature):
    """
    Identify 3-6 aspects the user should verify after the task.
    Focus on non-obvious concerns: cross-sheet formulas, dropdown sources,
    data integrity, naming consistency. Use sheetData.allSheets for cross-sheet deps.
    """
    user_message: str              = dspy.InputField(desc="User's task description")
    ws_context:   WorksheetContext = dspy.InputField(desc="Full workbook state")
    chat_history: list[str]        = dspy.InputField(desc="Conversation context")

    result: AspectList = dspy.OutputField()


class VerifyAspects(dspy.Signature):
    """
    Evaluate whether the workbook satisfies each aspect.
    Check all sheets via sheetData.allSheets. Use chat_history for intent.
    One sentence per aspect; use sheet-qualified refs (e.g. Data!A2:B6).
    """
    aspects:      AspectList       = dspy.InputField(desc="Aspects to evaluate")
    ws_context:   WorksheetContext = dspy.InputField(desc="Full workbook state")
    chat_history: list[str]        = dspy.InputField(desc="Full conversation — what was asked and done")

    result: VerifyResultList = dspy.OutputField()


class ChatResponse(dspy.Signature):
    """Answer a general Excel / spreadsheet question helpfully and concisely."""
    user_message: str              = dspy.InputField(desc="User's question")
    ws_context:   WorksheetContext = dspy.InputField(desc="Workbook state")
    chat_history: list[str]        = dspy.InputField(desc="Conversation context")

    response: str = dspy.OutputField(desc="Helpful, concise answer — markdown OK")


# ---------------------------------------------------------------------------
# Programs
# ---------------------------------------------------------------------------

class SegmentProgram(dspy.Module):
    def __init__(self): self.predict = dspy.ChainOfThought(GenerateSegments)
    def forward(self, **kwargs) -> SegmentList: return self.predict(**kwargs).result

class EditProgram(dspy.Module):
    def __init__(self): self.predict = dspy.ChainOfThought(EditSegments)
    def forward(self, **kwargs) -> SegmentList: return self.predict(**kwargs).result

class AskProgram(dspy.Module):
    def __init__(self): self.predict = dspy.ChainOfThought(AnswerQuestion)
    def forward(self, **kwargs) -> AskAnswer: return self.predict(**kwargs).result

class AspectScaffoldProgram(dspy.Module):
    def __init__(self): self.predict = dspy.ChainOfThought(ScaffoldAspects)
    def forward(self, **kwargs) -> AspectList: return self.predict(**kwargs).result

class AspectVerifyProgram(dspy.Module):
    def __init__(self): self.predict = dspy.ChainOfThought(VerifyAspects)
    def forward(self, **kwargs) -> VerifyResultList: return self.predict(**kwargs).result

class ChatProgram(dspy.Module):
    def __init__(self): self.predict = dspy.ChainOfThought(ChatResponse)
    def forward(self, **kwargs) -> str: return self.predict(**kwargs).response


_PROGRAMS: dict[str, type[dspy.Module]] = {
    "code":            SegmentProgram,
    "edit":            EditProgram,
    "ask":             AskProgram,
    "rubric_scaffold": AspectScaffoldProgram,
    "rubric_verify":   AspectVerifyProgram,
    "chat":            ChatProgram,
}


@lru_cache(maxsize=None)
def get_program(endpoint: str) -> dspy.Module:
    cls = _PROGRAMS.get(endpoint)
    if cls is None:
        raise ValueError(f"No DSPy program for endpoint '{endpoint}'")
    return cls()
