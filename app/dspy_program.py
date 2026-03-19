"""
dspy_programs.py — DSPy signatures and compiled programs for SheetCheck.

Each endpoint gets a typed Signature (inputs → outputs) and a Program
(a dspy.Module that wraps the signature with chain-of-thought or other
predictors).

Usage
─────
    from dspy_programs import get_program
    prog   = get_program("code")
    result = prog(user_message=..., ws_context=..., hint=...)

All programs return plain Python dicts / lists matching the shapes that
server.py already expects, so the rest of the stack is unchanged.

DSPy config
───────────
Call configure_dspy() once at startup (called from utils.py).
"""

from __future__ import annotations

import json
import logging
from functools import lru_cache
from typing import Any

import dspy

logger = logging.getLogger(__name__)


# ── DSPy LM configuration ─────────────────────────────────────────────────────

def configure_dspy(provider: str, model: str, api_key: str) -> None:
    """
    Wire up the DSPy global LM.
    Called once per endpoint invocation (cheap — DSPy caches internally).

    DSPy uses LiteLLM routing, so provider prefixes work:
        anthropic/claude-...
        openai/gpt-...
        mistral/mistral-...
        google/gemini-...
    """
    lm_id = f"{provider}/{model}"
    try:
        lm = dspy.LM(lm_id, api_key=api_key, max_tokens=8096)
        dspy.configure(lm=lm)
        logger.debug("DSPy configured: %s", lm_id)
    except Exception as exc:
        logger.error("DSPy configure failed for %s: %s", lm_id, exc)
        raise


# ── Shared field descriptions ─────────────────────────────────────────────────

_SEG_SCHEMA = """JSON array of code segments. Each segment:
{
  "id":               "seg-N",
  "description":      "Short imperative label",
  "sheet_context":    ["<range>", ...],
  "explanation":      "1–2 sentences: inputs → outputs",
  "predecessors":     ["seg-id", ...],
  "qa_pairs":         [{"q":"Why ...?","a":"Because ..."}],
  "edit_suggestions": ["Suggestion 1", "Suggestion 2"],
  "parameters":       [{"label":"Label","key":"varName","value":42,"type":"number"}],
  "code":             "await Excel.run(async (ctx) => { ... await ctx.sync(); });"
}
Types for parameters.type: number | color | select (+ options:[]) | text
Respond with ONLY the raw JSON array — no markdown fences, no prose."""


# ── Signatures ────────────────────────────────────────────────────────────────

class GenerateSegments(dspy.Signature):
    """Generate Excel Office JS code segments for a spreadsheet task."""

    # Inputs
    user_message:  str = dspy.InputField(desc="What the user wants to do in the spreadsheet")
    ws_context:    str = dspy.InputField(desc="JSON-serialised worksheet context (cells, values, formats)")
    rubric_hint:   str = dspy.InputField(desc="Optional rubric requirements as JSON (may be empty string)")
    js_hint:       str = dspy.InputField(desc="Known JS mistakes / fixes to avoid (may be empty string)")
    schema:        str = dspy.InputField(desc="Expected output schema")

    # Output
    segments_json: str = dspy.OutputField(desc=_SEG_SCHEMA)


class EditSegments(dspy.Signature):
    """Edit an existing segment and regenerate all downstream segments."""

    user_message:       str = dspy.InputField(desc="User's feedback / edit request")
    ws_context:         str = dspy.InputField(desc="JSON-serialised worksheet context")
    original_segment:   str = dspy.InputField(desc="JSON of the segment being edited")
    remaining_segments: str = dspy.InputField(desc="JSON array of segments that follow the edited one")
    js_hint:            str = dspy.InputField(desc="Known JS mistakes / fixes to avoid (may be empty string)")
    schema:             str = dspy.InputField(desc="Expected output schema")

    # Output — same array shape: [edited_seg, ...regenerated_remainder]
    segments_json: str = dspy.OutputField(
        desc=_SEG_SCHEMA + "\nFirst element is the edited segment; the rest are the regenerated remainder."
    )


class AnswerQuestion(dspy.Signature):
    """Answer a follow-up question about a specific Excel automation step."""

    user_message: str = dspy.InputField(desc="The user's question")
    ws_context:   str = dspy.InputField(desc="JSON-serialised worksheet context")
    current_step: str = dspy.InputField(desc="JSON of the step the user is asking about")
    history:      str = dspy.InputField(desc="JSON array of prior conversation turns (may be empty)")

    answer_json: str = dspy.OutputField(
        desc='JSON object: {"answer":"...","follow_up_questions":["...","..."]}'
    )


class ScaffoldRubric(dspy.Signature):
    """Generate an initial grading rubric for a spreadsheet task."""

    user_message: str = dspy.InputField(desc="Description of the spreadsheet task")
    ws_context:   str = dspy.InputField(desc="JSON-serialised worksheet context")

    rubric_json: str = dspy.OutputField(
        desc=(
            'JSON object with hard_requirements and soft_requirements arrays. '
            'Each item: {"id":"h1","label":"...","checked":false}. '
            'Respond with ONLY the raw JSON object.'
        )
    )


class VerifyRubric(dspy.Signature):
    """Evaluate a worksheet against each rubric requirement."""

    rubric:     str = dspy.InputField(desc="JSON of the rubric (hard + soft requirements)")
    ws_context: str = dspy.InputField(desc="JSON-serialised worksheet state")

    results_json: str = dspy.OutputField(
        desc=(
            'JSON array, one entry per rubric item: '
            '{"id":"h1","met":true,"reasoning":"...","references":["A1:E1"]}. '
            'Respond with ONLY the raw JSON array.'
        )
    )


class ChatResponse(dspy.Signature):
    """Answer a general spreadsheet / Excel question helpfully and concisely."""

    user_message: str = dspy.InputField(desc="User's question or request")
    ws_context:   str = dspy.InputField(desc="JSON-serialised worksheet context")

    response: str = dspy.OutputField(desc="Helpful, concise answer — markdown OK")


# ── Programs (dspy.Module wrappers) ───────────────────────────────────────────

class SegmentProgram(dspy.Module):
    """Chain-of-thought program for /code."""
    def __init__(self):
        self.predict = dspy.ChainOfThought(GenerateSegments)

    def forward(self, **kwargs) -> str:
        return self.predict(**kwargs).segments_json


class EditProgram(dspy.Module):
    """Chain-of-thought program for /edit."""
    def __init__(self):
        self.predict = dspy.ChainOfThought(EditSegments)

    def forward(self, **kwargs) -> str:
        return self.predict(**kwargs).segments_json


class AskProgram(dspy.Module):
    def __init__(self):
        self.predict = dspy.ChainOfThought(AnswerQuestion)

    def forward(self, **kwargs) -> str:
        return self.predict(**kwargs).answer_json


class RubricScaffoldProgram(dspy.Module):
    def __init__(self):
        self.predict = dspy.ChainOfThought(ScaffoldRubric)

    def forward(self, **kwargs) -> str:
        return self.predict(**kwargs).rubric_json


class RubricVerifyProgram(dspy.Module):
    def __init__(self):
        self.predict = dspy.ChainOfThought(VerifyRubric)

    def forward(self, **kwargs) -> str:
        return self.predict(**kwargs).results_json


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
    """Return (and cache) the program instance for `endpoint`."""
    cls = _PROGRAMS.get(endpoint)
    if cls is None:
        raise ValueError(f"No DSPy program for endpoint '{endpoint}'")
    return cls()
