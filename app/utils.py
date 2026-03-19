"""
utils.py — LLM call dispatch via DSPy programs.

Changes from original
─────────────────────
- call_llm() → call_program()  (DSPy-backed, typed I/O)
- parse_segments() now also runs JS validation via js_validator
- build_user_prompt() kept for backwards-compat but is lighter —
  DSPy programs receive structured fields rather than one big prompt blob
- JS validation errors cause a ValueError that server.py surfaces as 502
  (same as before), with the validation detail included in the message
"""

from __future__ import annotations

import json
import logging
import re
from typing import Any

from params import Provider, ENDPOINT_MODELS
from params import ANTHROPIC_API_KEY, OPENAI_API_KEY, MISTRAL_API_KEY, GEMINI_API_KEY

logger = logging.getLogger(__name__)

# Provider → LiteLLM prefix used by DSPy
_PROVIDER_PREFIX: dict[Provider, str] = {
    Provider.ANTHROPIC: "anthropic",
    Provider.OPENAI:    "openai",
    Provider.MISTRAL:   "mistral",
    Provider.GOOGLE:    "gemini",
}

# Provider → which API key to pass
_PROVIDER_KEY: dict[Provider, str] = {
    Provider.ANTHROPIC: ANTHROPIC_API_KEY,
    Provider.OPENAI:    OPENAI_API_KEY,
    Provider.MISTRAL:   MISTRAL_API_KEY,
    Provider.GOOGLE:    GEMINI_API_KEY,
}

# Required segment fields
_REQUIRED_SEG_FIELDS = {"id", "description", "sheet_context", "explanation", "code"}

# Optional defaults to inject
_SEG_DEFAULTS: dict[str, Any] = {
    "predecessors":     [],
    "affordances":      [],
    "alternatives":     [],
    "qa_pairs":         [],
    "edit_suggestions": [],
    "parameters":       [],
    "undo_code":        "",
}


# ── DSPy segment schema example (injected into GenerateSegments / EditSegments)

def _segment_schema_example() -> str:
    """Compact schema hint passed to DSPy programs that produce segments."""
    from params import STUB_SEGMENTS
    return json.dumps(STUB_SEGMENTS[:2])  # 2 examples is enough context


# ── Public helpers ────────────────────────────────────────────────────────────

def build_user_prompt(user_message: str, ws_context: dict, extra: dict | None = None) -> str:
    """
    Legacy helper retained for /chat and any callers that pass a raw prompt.
    For segment-producing endpoints, prefer call_program() directly.
    """
    ctx_block = json.dumps(ws_context, indent=2) if ws_context else "{}"
    parts = [f"Worksheet context:\n```json\n{ctx_block}\n```\n\nUser request: {user_message}"]
    if extra:
        parts.append(f"\nExtra context:\n```json\n{json.dumps(extra, indent=2)}\n```")
    return "\n".join(parts)


def call_program(endpoint: str, **kwargs) -> str:
    """
    Configure DSPy for the right provider/model, then run the program.
    Returns the raw string output from the program.
    """
    from dspy_programs import configure_dspy, get_program

    cfg = ENDPOINT_MODELS[endpoint]
    prefix = _PROVIDER_PREFIX[cfg.provider]
    api_key = _PROVIDER_KEY[cfg.provider]

    configure_dspy(prefix, cfg.model.value, api_key)

    prog = get_program(endpoint)
    return prog(**kwargs)


def parse_json(raw_text: str) -> Any:
    cleaned = re.sub(r"^```(?:json)?\s*", "", raw_text.strip(), flags=re.IGNORECASE)
    cleaned = re.sub(r"\s*```$", "", cleaned.strip())
    return json.loads(cleaned)


def parse_segments(raw_text: str, validate: bool = True) -> list[dict]:
    """
    Parse a JSON array of segments, fill defaults, and optionally
    run JS validation on each segment's `code` field.
    """
    segments = parse_json(raw_text)
    if not isinstance(segments, list):
        raise ValueError("LLM response is not a JSON array")

    for i, seg in enumerate(segments):
        missing = _REQUIRED_SEG_FIELDS - seg.keys()
        if missing:
            raise ValueError(f"Segment {i} missing fields: {missing}")
        for k, v in _SEG_DEFAULTS.items():
            seg.setdefault(k, type(v)())  # empty list / empty string

    if validate:
        from js_validator import validate_segments, mistakes_prompt_hint  # noqa: F401
        validate_segments(segments)

    # Strip internal _validation key before returning to client
    for seg in segments:
        seg.pop("_validation", None)

    return segments


# ── Convenience wrappers used by server.py ────────────────────────────────────

def generate_segments(
    user_message: str,
    ws_context: dict,
    rubric: dict | None = None,
    js_hint: str = "",
) -> list[dict]:
    """
    Shared logic for /code and /edit (code generation path).
    Returns validated, parsed segment list.
    """
    from js_validator import mistakes_prompt_hint

    hint = js_hint or mistakes_prompt_hint()

    rubric_hint = ""
    if rubric:
        hard = [r["label"] for r in rubric.get("hard_requirements", [])]
        soft = [r["label"] for r in rubric.get("soft_requirements", [])]
        rubric_hint = json.dumps({"hard_must_satisfy": hard, "soft_nice_to_have": soft})

    raw = call_program(
        "code",
        user_message=user_message,
        ws_context=json.dumps(ws_context, indent=2) if ws_context else "{}",
        rubric_hint=rubric_hint,
        js_hint=hint,
        schema=_segment_schema_example(),
    )
    return parse_segments(raw)


def edit_segments(
    user_message: str,
    ws_context: dict,
    original_segment: dict,
    remaining_segments: list[dict],
    js_hint: str = "",
) -> list[dict]:
    """
    /edit path: modify a segment and regenerate its downstream chain.
    """
    from js_validator import mistakes_prompt_hint

    hint = js_hint or mistakes_prompt_hint()
    raw = call_program(
        "edit",
        user_message=user_message or "Apply user feedback to this segment and update the remainder.",
        ws_context=json.dumps(ws_context, indent=2) if ws_context else "{}",
        original_segment=json.dumps(original_segment),
        remaining_segments=json.dumps(remaining_segments),
        js_hint=hint,
        schema=_segment_schema_example(),
    )
    return parse_segments(raw)
