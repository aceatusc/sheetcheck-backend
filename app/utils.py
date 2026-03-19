"""
utils.py — LLM call dispatch via DSPy programs.

With Pydantic-typed DSPy signatures, programs now return typed objects
directly. This means:
  - No more parse_json() / parse_segments() for structured endpoints
  - No more manual json.dumps() for structured inputs — pass the model
  - JS validation still runs, now against Segment objects
  - build_user_prompt() is gone — /chat passes fields directly too
"""

from __future__ import annotations

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
    """Configure DSPy for the right provider/model and run the program."""
    from dspy_programs import configure_dspy, get_program

    cfg    = ENDPOINT_MODELS[endpoint]
    prefix = _PROVIDER_PREFIX[cfg.provider]
    key    = _PROVIDER_KEY[cfg.provider]

    configure_dspy(prefix, cfg.model.value, key)
    return get_program(endpoint)(**kwargs)


def _validate_and_dump_segments(segment_list) -> list[dict]:
    """
    Run JS validation on each Segment in a SegmentList, then return
    plain dicts suitable for JSON serialisation.
    """
    from js_validator import validate_js
    from dspy_programs import Segment

    failures: list[str] = []
    dicts: list[dict] = []

    for seg in segment_list.segments:
        vr = validate_js(seg.code, segment_id=seg.id)
        if not vr.valid:
            failures.append(f"  [{seg.id}] {'; '.join(vr.errors)}")
        elif vr.warnings:
            logger.warning("[%s] JS warnings: %s", seg.id, "; ".join(vr.warnings))
        dicts.append(seg.model_dump())

    if failures:
        raise ValueError("JS validation failed for segments:\n" + "\n".join(failures))

    return dicts


# ── Convenience wrappers used by server.py ────────────────────────────────────

def generate_segments(
    user_message: str,
    ws_context: dict,
    rubric: dict | None = None,
) -> list[dict]:
    from js_validator import mistakes_prompt_hint
    from dspy_programs import WorksheetContext, RubricHint

    rubric_hint = RubricHint()
    if rubric:
        rubric_hint = RubricHint(
            hard_must_satisfy=[r["label"] for r in rubric.get("hard_requirements", [])],
            soft_nice_to_have=[r["label"] for r in rubric.get("soft_requirements", [])],
        )

    result = call_program(
        "code",
        user_message=user_message,
        ws_context=WorksheetContext(**ws_context) if ws_context else WorksheetContext(),
        rubric_hint=rubric_hint,
        js_hint=mistakes_prompt_hint(),
    )
    return _validate_and_dump_segments(result)


def edit_segments(
    user_message: str,
    ws_context: dict,
    original_segment: dict,
    remaining_segments: list[dict],
) -> list[dict]:
    from js_validator import mistakes_prompt_hint
    from dspy_programs import WorksheetContext, Segment

    result = call_program(
        "edit",
        user_message=user_message or "Apply user feedback to this segment and update the remainder.",
        ws_context=WorksheetContext(**ws_context) if ws_context else WorksheetContext(),
        original_segment=Segment(**original_segment),
        remaining_segments=[Segment(**s) for s in remaining_segments],
        js_hint=mistakes_prompt_hint(),
    )
    return _validate_and_dump_segments(result)


def ask_question(
    user_message: str,
    ws_context: dict,
    step: dict,
    history: list,
) -> dict:
    from dspy_programs import WorksheetContext, Segment

    result = call_program(
        "ask",
        user_message=user_message,
        ws_context=WorksheetContext(**ws_context) if ws_context else WorksheetContext(),
        current_step=Segment(**step) if step else Segment(id="", description="", sheet_context=[], explanation="", code=""),
        history=history,
    )
    return result.model_dump()


def scaffold_rubric(user_message: str, ws_context: dict) -> dict:
    from dspy_programs import WorksheetContext

    result = call_program(
        "rubric_scaffold",
        user_message=user_message,
        ws_context=WorksheetContext(**ws_context) if ws_context else WorksheetContext(),
    )
    return result.model_dump()


def verify_rubric(rubric: dict, ws_context: dict) -> list[dict]:
    from dspy_programs import WorksheetContext, Rubric, RubricItem

    rubric_model = Rubric(
        hard_requirements=[RubricItem(**r) for r in rubric.get("hard_requirements", [])],
        soft_requirements=[RubricItem(**r) for r in rubric.get("soft_requirements", [])],
    )
    result = call_program(
        "rubric_verify",
        rubric=rubric_model,
        ws_context=WorksheetContext(**ws_context) if ws_context else WorksheetContext(),
    )
    return [r.model_dump() for r in result.results]


def chat_response(user_message: str, ws_context: dict) -> str:
    from dspy_programs import WorksheetContext

    return call_program(
        "chat",
        user_message=user_message,
        ws_context=WorksheetContext(**ws_context) if ws_context else WorksheetContext(),
    )
