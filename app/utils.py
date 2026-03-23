"""
utils.py -- LLM call dispatch via DSPy programs.

All wire-format decisions are driven by the actual add-in JS:
  - WorksheetContext matches worksheetContext.js gather() output exactly
  - ask() receives step as {description, explanation} (StepSummary), not a full Segment
  - scaffold_rubric() returns the Rubric dict directly (not nested)
  - verify_rubric() returns a list of VerifyResult dicts (server wraps in {"results": [...]})

Threading: call_program() uses dspy.context(lm=...) -- a thread-local context manager --
instead of dspy.configure(), which is locked to the thread that first calls it and would
raise RuntimeError on every subsequent Flask request thread.
"""

from __future__ import annotations

import json
import logging
from typing import Any

from params import Provider, ENDPOINT_MODELS
from params import ANTHROPIC_API_KEY, OPENAI_API_KEY, MISTRAL_API_KEY, GEMINI_API_KEY

logger = logging.getLogger(__name__)

_PROVIDER_PREFIX: dict[Provider, str] = {
    Provider.ANTHROPIC: "anthropic",
    Provider.OPENAI:    "openai",
    Provider.MISTRAL:   "mistral",
    Provider.GOOGLE:    "gemini",
}

_PROVIDER_KEY: dict[Provider, str] = {
    Provider.ANTHROPIC: ANTHROPIC_API_KEY,
    Provider.OPENAI:    OPENAI_API_KEY,
    Provider.MISTRAL:   MISTRAL_API_KEY,
    Provider.GOOGLE:    GEMINI_API_KEY,
}


def call_program(endpoint: str, **kwargs) -> Any:
    """
    Run the DSPy program for `endpoint` using a thread-local LM context.

    Uses dspy.context(lm=...) instead of dspy.configure() so each Flask
    request thread can independently set its LM without racing on global state.
    Every call is logged via memor (llm_logger.py) for reproducibility.
    """
    import dspy
    from dspy_programs import get_lm, get_program
    from llm_logger import log_call

    cfg    = ENDPOINT_MODELS[endpoint]
    prefix = _PROVIDER_PREFIX[cfg.provider]
    key    = _PROVIDER_KEY[cfg.provider]
    lm     = get_lm(prefix, cfg.model.value, key, endpoint)
    model  = cfg.model.value

    with log_call(endpoint, model=model) as call_log:
        # Log a compact summary of the inputs as the prompt record
        prompt_summary = {k: (str(v)[:500] if not isinstance(v, str) else v[:500])
                          for k, v in kwargs.items()}
        call_log.set_prompt(json.dumps(prompt_summary, indent=2, default=str))

        with dspy.context(lm=lm):
            result = get_program(endpoint)(**kwargs)

        # Log a compact summary of the output as the response record
        call_log.set_response(str(result))

    return result


def _make_ws_context(ws_context: dict):
    """
    Construct a WorksheetContext from the raw dict the add-in sends.
    gather() in worksheetContext.js returns:
      { selection, sheetData: { usedRange: { address, values } }, namedRanges, sheetNames }
    extra='allow' on the model means unknown keys are preserved safely.
    """
    from dspy_programs import WorksheetContext
    return WorksheetContext(**ws_context) if ws_context else WorksheetContext()


def _validate_and_dump_segments(segment_list) -> list[dict]:
    """Run JS validation on each Segment, return plain dicts for JSON serialisation.
    Segment.model_dump() already unwraps OfficeJSCode → str in the 'code' field.
    """
    from js_validator import validate_js

    failures: list[str] = []
    dicts: list[dict] = []

    for seg in segment_list.segments:
        code_str = seg.code_str
        vr = validate_js(code_str, segment_id=seg.id)
        if not vr.valid:
            failures.append(f"  [{seg.id}] {'; '.join(vr.errors)}")
        elif vr.warnings:
            logger.warning("[%s] JS warnings: %s", seg.id, "; ".join(vr.warnings))
        dicts.append(seg.model_dump())

    if failures:
        raise ValueError("JS validation failed for segments:\n" + "\n".join(failures))

    return dicts


# -- Endpoint wrappers --------------------------------------------------------

def generate_segments(
    user_message: str,
    ws_context: dict,
    rubric: dict | None = None,
    chat_history: list[str] | None = None,
) -> list[dict]:
    from js_validator import mistakes_prompt_hint
    from dspy_programs import RubricHint

    result = call_program(
        "code",
        user_message=user_message,
        ws_context=_make_ws_context(ws_context),
        rubric_hint=RubricHint(),
        js_hint=mistakes_prompt_hint(),
        chat_history=chat_history or [],
    )
    return _validate_and_dump_segments(result)

def edit_segments(
    user_message: str,
    ws_context: dict,
    original_segment: dict,
    remaining_segments: list[dict],
    chat_history: list[str] | None = None,
) -> list[dict]:
    from js_validator import mistakes_prompt_hint
    from dspy_programs import Segment

    result = call_program(
        "edit",
        user_message=user_message or "Apply user feedback to this segment and update the remainder.",
        ws_context=_make_ws_context(ws_context),
        original_segment=Segment(**original_segment),
        remaining_segments=[Segment(**s) for s in remaining_segments],
        js_hint=mistakes_prompt_hint(),
        chat_history=chat_history or [],
    )
    return _validate_and_dump_segments(result)


def ask_question(
    user_message: str,
    ws_context: dict,
    step: dict,
    history: list,
    chat_history: list[str] | None = None,
) -> dict:
    """
    step arrives as { description, explanation } from stepNavigator._onAskSend().
    We use StepSummary (not full Segment) so the model isn't asked to reconstruct
    fields it was never given.
    """
    from dspy_programs import StepSummary

    result = call_program(
        "ask",
        user_message=user_message,
        ws_context=_make_ws_context(ws_context),
        current_step=StepSummary(**step) if step else StepSummary(),
        history=history,
        chat_history=chat_history or [],
    )
    return result.model_dump()


def scaffold_rubric(
    user_message: str,
    ws_context: dict,
    chat_history: list[str] | None = None,
) -> dict:
    """Returns the rubric dict directly -- matches what rubricManager.setRubric() expects."""
    result = call_program(
        "rubric_scaffold",
        user_message=user_message,
        ws_context=_make_ws_context(ws_context),
        chat_history=chat_history or [],
    )
    return result.model_dump()



def verify_rubric(
    rubric: dict,
    ws_context: dict,
    chat_history: list[str] | None = None,
) -> list[dict]:
    """
    `rubric` is { aspects: [{id, label}, ...] } sent by the front-end.
    Builds an AspectList and calls the rubric_verify program.
    server.py wraps the return value in {"results": [...]}.
    """
    from dspy_programs import Aspect, AspectList

    aspects_raw = rubric.get("aspects", [])
    aspect_list = AspectList(aspects=[Aspect(**a) for a in aspects_raw])
    result = call_program(
        "rubric_verify",
        aspects=aspect_list,
        ws_context=_make_ws_context(ws_context),
        chat_history=chat_history or [],
    )
    return [r.model_dump() for r in result.results]


def chat_response(
    user_message: str,
    ws_context: dict,
    chat_history: list[str] | None = None,
) -> str:
    return call_program(
        "chat",
        user_message=user_message,
        ws_context=_make_ws_context(ws_context),
        chat_history=chat_history or [],
    )
