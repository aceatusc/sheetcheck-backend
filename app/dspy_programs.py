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
from pydantic import BaseModel, ConfigDict, Field, field_validator

logger = logging.getLogger(__name__)


# -- LM factory (one LM instance per provider/model, cached) ------------------
#
# We deliberately do NOT call dspy.configure() globally. Each call_program()
# uses `with dspy.context(lm=...)` which is thread-local and safe.

# Per-endpoint max_tokens budgets.
# Set to the practical ceiling for each endpoint — segment generation needs
# headroom for 5-10 segments with code + qa_pairs + parameters each; Q&A and
# chat produce much shorter outputs and don't need a large budget.
# All values are well within each provider's hard limits:
#   Gemini 3 Flash / 3.1 variants: 64,000 output tokens
#   Mistral Small 2506 / Ministral: 131,072 output tokens
MAX_TOKENS: dict[str, int] = {
    "code":            64_000,   # 5-10 segments, each with code + qa + params
    "edit":            32_000,   # same shape as code
    "ask":              2_000,   # short answer + 2 follow-up questions
    "rubric_scaffold":  8_000,   # 2-4 rubric items
    "rubric_verify":    8_000,   # one reasoning line per rubric item
    "chat":             2_000,   # conversational answer
}
_DEFAULT_MAX_TOKENS = 8_096


@lru_cache(maxsize=None)
def get_lm(provider: str, model: str, api_key: str, endpoint: str) -> dspy.LM:
    lm_id = f"{provider}/{model}"
    max_tokens = MAX_TOKENS.get(endpoint, _DEFAULT_MAX_TOKENS)
    logger.debug("Creating LM: %s (max_tokens=%d)", lm_id, max_tokens)
    return dspy.LM(lm_id, api_key=api_key, max_tokens=max_tokens)


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


class CellStyle(BaseModel):
    """Styling snapshot for a single cell or range."""
    model_config = ConfigDict(extra="allow")
    address:             Optional[str]  = None
    fillColor:           Optional[str]  = None   # hex, e.g. "#1a1d27"
    fontColor:           Optional[str]  = None
    fontBold:            Optional[bool] = None
    fontItalic:          Optional[bool] = None
    fontSize:            Optional[float] = None
    numberFormat:        Optional[str]  = None   # e.g. "$#,##0", "0.0%"
    horizontalAlignment: Optional[str]  = None   # "Left" | "Center" | "Right"


class ChartInfo(BaseModel):
    """Summary of a chart object present on the sheet."""
    model_config = ConfigDict(extra="allow")
    name:       Optional[str]       = None   # chart object name
    chartType:  Optional[str]       = None   # "Line", "ColumnClustered", etc.
    dataRange:  Optional[str]       = None   # source data range address
    title:      Optional[str]       = None


class SheetInfo(BaseModel):
    """Data for a single worksheet (used in allSheets list)."""
    model_config = ConfigDict(extra="allow")
    name:      str                       = Field(description="Worksheet tab name")
    address:   Optional[str]             = None
    values:    Optional[list[list[Any]]] = None
    rowCount:  Optional[int]             = None
    colCount:  Optional[int]             = None
    truncated: Optional[bool]            = Field(default=False, description="True when the sheet was too large and only the first shownRows rows are included")
    shownRows: Optional[int]             = None
    error:     Optional[str]             = None


class SheetData(BaseModel):
    model_config = ConfigDict(extra="allow")
    usedRange: Optional[UsedRange]    = None
    allSheets: Optional[list[SheetInfo]] = Field(default=None,
        description=(
            "Data from ALL worksheets in the workbook. "
            "Use this to understand cross-sheet references (e.g. VLOOKUP into a Data sheet). "
            "The active sheet is always present in full; other sheets may be truncated."
        ))
    styles:    Optional[list[CellStyle]] = Field(default=None,
        description="Sampled cell styles — one entry per distinct formatted region")
    charts:    Optional[list[ChartInfo]] = Field(default=None,
        description="Charts present on the active sheet")


class WorksheetContext(BaseModel):
    model_config = ConfigDict(extra="allow")
    selection:   Optional[Selection]  = None
    sheetData:   Optional[SheetData]  = None
    namedRanges: Optional[list[dict]] = None
    sheetNames:  Optional[list[str]]  = None


# -- Segment models -----------------------------------------------------------

class QAPair(BaseModel):
    q: str = Field(description="A 'Why ...?' question about this step as the user sees it in the sheet (e.g. 'Why does column D show a dropdown?', 'Why does B3 contain a formula?') — never reference JavaScript or code")
    a: str = Field(description="A concise answer explaining the design choice in plain terms — describe cell addresses, formula bar contents, visible values; no code references")


class SegmentParameter(BaseModel):
    label:   str                                          = Field(description="Human-readable label shown in the UI")
    key:     str                                          = Field(description="Variable name or literal this maps to in the code")
    value:   Union[str, int, float]                       = Field(description="Current value of this parameter")
    type:    Literal["number", "color", "select", "text"] = Field(description="UI control type")
    options: Optional[list[str]]                          = Field(default=None, description="Choices for 'select' type only")


class OfficeJSCode(BaseModel):
    """
    A single self-contained Office JS snippet.

    MUST follow these rules exactly — violations cause runtime errors:

    STRUCTURE
    - Wrap in: await Excel.run(async (ctx) => { ... await ctx.sync(); });
    - Always use `async` in the Excel.run callback.

    RANGES & ARRAYS
    - .values / .numberFormat / .formulas ALWAYS take a 2-D array sized exactly to the range.
      Single cell: [[value]]. Column of N rows: [[v1],[v2],...]. Row of N cols: [[v1,v2,...]].
    - getRange address must exactly match the array dimensions (rows × cols). Count rows and cols
      before assigning — a mismatch causes "The number of rows or columns in the input array
      doesn't match the size or dimensions of the range."
    - NEVER read .values/.formulas without load() + await ctx.sync() first.
    - NEVER load() and write to the same range in one sync block.

    DATA VALIDATION (DROPDOWN LISTS)
    - To add a dropdown/list validation use the Ranges API, NOT the old DataValidation API:
        const rng = sheet.getRange("D2:D50");
        rng.dataValidation.rule = {
            list: { inCellDropDown: true, source: "=Data!$A$2:$A$6" }
        };
        await ctx.sync();
    - NEVER use sheet.dataValidation or range.dataValidation.add() — they do not exist.
    - The `source` string for a list rule that references another sheet MUST be a formula
      string starting with "=" e.g. "=Data!$A$2:$A$6". A plain address like "Data!A2:A6"
      causes "The argument is invalid or missing or has an incorrect format."

    RANGE NAVIGATION
    - getRange() is a Worksheet method ONLY. Call sheet.getRange('A1'), never range.getRange('A1').
    - Range.getLastCell() returns a Range — it has NO .getEnd() method. To get the last used row
      dynamically, load the usedRange rowCount and compute the address from that:
        const used = sheet.getUsedRange(); used.load('rowCount'); await ctx.sync();
        const lastRow = used.rowCount;  // then use getRange(`A2:A${lastRow}`)
    - Range.getRow() does not exist. Access rows via sheet.getRange('A1:Z1').
    - NEVER call range.getLastCell().getEnd(...) — getEnd is not a function on Range.

    FORMATTING
    - NEVER use .conditionalFormatting — not supported. Use explicit per-cell formatting in a loop.
    - Column autofit: range.getEntireColumn().format.autofitColumns() — NEVER .autofit().
    - Call autofitColumns() AFTER all data and formatting is written.
    """
    model_config = ConfigDict(extra="allow")

    code: str = Field(
        description=("Office JS snippet: await Excel.run(async (ctx) => { ... await ctx.sync(); }); ")
    )

    @classmethod
    def model_validate(cls, value, **kwargs):
        """Accept a plain string (from existing segment dicts) as well as a dict/object."""
        if isinstance(value, str):
            return cls(code=value)
        return super().model_validate(value, **kwargs)


class Segment(BaseModel):
    id:               str                    = Field(description="Unique segment identifier, e.g. 'seg-1'")
    description:      str                    = Field(description="Short imperative label, e.g. 'Write header row'")
    sheet_context:    list[str]              = Field(description="Excel range addresses this segment touches")
    explanation:      str                    = Field(description="1-2 sentences: inputs to outputs")
    predecessors:     list[str]              = Field(default_factory=list)
    qa_pairs:         list[QAPair]           = Field(default_factory=list, description="2-3 design Q&A pairs")
    edit_suggestions: list[str]             = Field(default_factory=list, description="2-3 short edit prompts")
    parameters:       list[SegmentParameter] = Field(default_factory=list, description="Tweakable constants in the code")
    code:             OfficeJSCode           = Field(description="The Office JS code for this step")
    undo_code:        str                    = Field(default="")
    manual_steps:     str                    = Field(default="", description="Step-by-step manual instructions for doing this step in the Excel UI.")

    @field_validator("code", mode="before")
    @classmethod
    def _coerce_code(cls, v):
        if isinstance(v, str):
            return {"code": v}
        return v

    def model_dump(self, **kwargs) -> dict:
        """Unwrap OfficeJSCode → plain string so downstream JSON/validator sees seg['code'] as str."""
        d = super().model_dump(**kwargs)
        if isinstance(d.get("code"), dict):
            d["code"] = d["code"].get("code", "")
        return d

    @property
    def code_str(self) -> str:
        """Convenience accessor returning the raw JS string."""
        return self.code.code if isinstance(self.code, OfficeJSCode) else str(self.code)


class SegmentList(BaseModel):
    segments: list[Segment] = Field(description="Ordered list of code segments to execute")


# -- Ask ----------------------------------------------------------------------

class StepSummary(BaseModel):
    """Partial segment shape sent by stepNavigator._onAskSend()."""
    model_config = ConfigDict(extra="allow")
    description: Optional[str] = None
    explanation: Optional[str] = None


class AskAnswer(BaseModel):
    answer:              str       = Field(description="Clear, concise answer in 1-3 sentences — describe what the user sees in the sheet (cell addresses, formula bar, visible values), never mention JavaScript or code")
    follow_up_questions: list[str] = Field(description="2 short suggested follow-up questions about what the user sees or how the result works in the sheet — no code references")


# -- Aspects (replaces Rubric) ------------------------------------------------
#
# A flat list of "aspects" — important dimensions the user should check.
# No hard/soft distinction. Used by the standalone Verify panel.

class Aspect(BaseModel):
    id:    str = Field(description="Unique aspect ID, e.g. 'a1', 'a2'")
    label: str = Field(description="Human-readable aspect description (1-2 sentences)")


class AspectList(BaseModel):
    aspects: list[Aspect] = Field(description="3-6 important aspects to review")


class VerifyResult(BaseModel):
    id:         str       = Field(description="Aspect ID")
    met:        bool      = Field(description="Whether the worksheet satisfies this aspect")
    reasoning:  str       = Field(description="One sentence explanation")
    references: list[str] = Field(description="Supporting cell ranges, e.g. ['A1:E1']")


class VerifyResultList(BaseModel):
    results: list[VerifyResult] = Field(description="One entry per aspect")


# Kept for backward compat with /code rubric_hint (no longer shown as gate,
# but still passed through so the LLM can be aware of aspects if present)
class RubricItem(BaseModel):
    id:      str  = Field(description="Aspect ID")
    label:   str  = Field(description="Human-readable aspect text")
    checked: bool = Field(default=False)


class Rubric(BaseModel):
    hard_requirements: list[RubricItem] = Field(default_factory=list)
    soft_requirements: list[RubricItem] = Field(default_factory=list)


class RubricHint(BaseModel):
    hard_must_satisfy: list[str] = Field(default_factory=list)
    soft_nice_to_have: list[str] = Field(default_factory=list)


# -- Signatures ----------------------------------------------------------------

class GenerateSegments(dspy.Signature):
    """
    Generate a sequence of Excel Office JS code segments that fully accomplish the task.

    Decompose into as many fine-grained segments as needed — one concern per segment
    (write data, apply formulas, format headers, format rows, totals, colour-coding,
    autofit, etc.). Prefer more segments over fewer; 5-10 is typical.

    Each segment: self-contained, single concern, clear explanation, 2-3 Q&A pairs,
    all tweakable constants as parameters[]. For column sizing always use
    range.getEntireColumn().format.autofitColumns() — never .autofit().
    Always call sheet.getRange() not range.getRange() — getRange() is a Worksheet method only.
    Never call range.getRow() — it does not exist. Use sheet.getRange('A1:Z1') with an explicit address.

    Follow all rules in js_hint exactly — they list known runtime errors to avoid.
    """
    user_message:  str              = dspy.InputField(desc="What the user wants to do in the spreadsheet")
    ws_context:    WorksheetContext = dspy.InputField(desc="Current worksheet state")
    rubric_hint:   RubricHint       = dspy.InputField(desc="Optional rubric requirements to satisfy (may be empty)")
    js_hint:       str              = dspy.InputField(desc="IMPORTANT: additional JS mistakes seen in recent runs that must be avoided — read carefully before writing any code (may be empty)")
    chat_history:  list[str]        = dspy.InputField(desc="Recent user messages for context (oldest first, may be empty)")

    result: SegmentList = dspy.OutputField()


class EditSegments(dspy.Signature):
    """
    Modify the given segment based on user feedback, then regenerate all downstream
    segments so they remain consistent. Output the edited segment first, then the regenerated remainder in order.

    Preserve or increase granularity — do not collapse steps.
    Follow all rules in js_hint exactly — they list known runtime errors to avoid.
    """
    user_message:       str              = dspy.InputField(desc="User's feedback describing the desired change")
    ws_context:         WorksheetContext = dspy.InputField(desc="Current worksheet state")
    original_segment:   Segment          = dspy.InputField(desc="The segment to edit")
    remaining_segments: list[Segment]    = dspy.InputField(desc="Segments that follow the edited one (may be empty)")
    js_hint:            str              = dspy.InputField(desc="IMPORTANT: additional JS mistakes seen in recent runs that must be avoided — read carefully before writing any code (may be empty)")
    chat_history:       list[str]        = dspy.InputField(desc="Recent user messages for context (oldest first, may be empty)")

    result: SegmentList = dspy.OutputField()


class AnswerQuestion(dspy.Signature):
    """
    Answer a follow-up question about a specific step in a spreadsheet automation plan.

    CRITICAL: The user is looking at their Excel sheet — they cannot see any JavaScript code.
    Frame every answer in terms of what the user observes in the sheet:
      - Cell addresses and range addresses (e.g. "column D", "row 3", "B2:B50")
      - Formula strings as they appear in the formula bar (e.g. =VLOOKUP(D2,Data!$A$2:$B$6,2,FALSE))
      - Visible values, text, formatting, dropdown options
      - Sheet tab names
    NEVER reference variable names, function names, or any JavaScript/code concepts.

    Be concise (1-3 sentences per answer). Suggest 2 natural follow-up questions that
    stay within the same "what does the user see / how does it work in the sheet" framing.
    """
    user_message:  str              = dspy.InputField(desc="The user's question about what they see in their sheet")
    ws_context:    WorksheetContext = dspy.InputField(desc="Current worksheet state including all sheet data")
    current_step:  StepSummary      = dspy.InputField(desc="Description and explanation of the step being asked about")
    history:       list[dict]       = dspy.InputField(desc="Prior Q&A turns [{q, a}] in this ask session (may be empty)")
    chat_history:  list[str]        = dspy.InputField(desc="Recent user messages for context (oldest first, may be empty)")

    result: AskAnswer = dspy.OutputField()


class ScaffoldAspects(dspy.Signature):
    """
    Given the user's chat history and current workbook state, identify
    *important aspects* the user should verify about the task.

    Aspects are thought-provoking dimensions that help the user overcome blind spots and hidden assumptions.
    Write each as a concise, specific, actionable verification
    (e.g. "Column headers should be consistent with the existing sheet naming convention").
    Focus on things the user might not have explicitly mentioned but that matter
    for the task (unknown unknowns, common spreadsheet pitfalls, data integrity,
    cross-sheet formula correctness, dropdown list sources).

    ws_context.sheetData.allSheets contains data from every worksheet tab —
    use it to notice cross-sheet dependencies the user may have overlooked.
    """
    user_message:  str              = dspy.InputField(desc="The user's original task description")
    ws_context:    WorksheetContext = dspy.InputField(desc="Current full worksheet state")
    chat_history:  list[str]        = dspy.InputField(desc="Recent user messages for context (oldest first, may be empty)")

    result: AspectList = dspy.OutputField()


class VerifyAspects(dspy.Signature):
    """
    Evaluate whether the current workbook satisfies each aspect.

    Use ALL available context:
    - ws_context.sheetData.allSheets contains data from every worksheet tab —
      check cross-sheet references and data that may live outside the active sheet.
    - chat_history contains the full conversation: what the user asked for, what
      was applied, and any corrections — use it to understand the intended outcome.
    - Be precise about cell references (include sheet name for non-active sheets,
      e.g. 'Data!A2:B6') and give one-sentence reasoning per item.
    - Cover every aspect in the list — do not skip any.
    """
    aspects:      AspectList       = dspy.InputField(desc="The aspects to evaluate against")
    ws_context:   WorksheetContext = dspy.InputField(desc="Full workbook state — all worksheets via sheetData.allSheets")
    chat_history: list[str]        = dspy.InputField(desc="Full conversation history (oldest first) — what the user asked for and what was done")

    result: VerifyResultList = dspy.OutputField()


class ChatResponse(dspy.Signature):
    """Answer a general Excel / spreadsheet question helpfully and concisely."""
    user_message:  str              = dspy.InputField(desc="User's question or request")
    ws_context:    WorksheetContext = dspy.InputField(desc="Current worksheet state")
    chat_history:  list[str]        = dspy.InputField(desc="Recent user messages for context (oldest first, may be empty)")

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


class AspectScaffoldProgram(dspy.Module):
    def __init__(self):
        self.predict = dspy.ChainOfThought(ScaffoldAspects)

    def forward(self, **kwargs) -> AspectList:
        return self.predict(**kwargs).result


class AspectVerifyProgram(dspy.Module):
    def __init__(self):
        self.predict = dspy.ChainOfThought(VerifyAspects)

    def forward(self, **kwargs) -> VerifyResultList:
        return self.predict(**kwargs).result


class ChatProgram(dspy.Module):
    def __init__(self):
        self.predict = dspy.ChainOfThought(ChatResponse)

    def forward(self, **kwargs) -> str:
        return self.predict(**kwargs).response


# -- Registry -----------------------------------------------------------------

_PROGRAMS: dict[str, type[dspy.Module]] = {
    "code":             SegmentProgram,
    "edit":             EditProgram,
    "ask":              AskProgram,
    "rubric_scaffold":  AspectScaffoldProgram,
    "rubric_verify":    AspectVerifyProgram,
    "chat":             ChatProgram,
}


@lru_cache(maxsize=None)
def get_program(endpoint: str) -> dspy.Module:
    cls = _PROGRAMS.get(endpoint)
    if cls is None:
        raise ValueError(f"No DSPy program for endpoint '{endpoint}'")
    return cls()
