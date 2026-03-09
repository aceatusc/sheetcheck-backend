""" Server Utilities."""

import json
import re

from params import LLM_PROVIDER, SYSTEM_PROMPT
from params import (
    ANTHROPIC_API_KEY,
    OPENAI_API_KEY,
    MISTRAL_API_KEY,
)
from params import (
    ANTHROPIC_MODEL,
    OPENAI_MODEL,
    MISTRAL_MODEL,
)

# ---------------------------------------------------------------------------
# Prompt builder
# ---------------------------------------------------------------------------

def build_user_prompt(user_message: str, ws_context: dict) -> str:
    ctx_block = json.dumps(ws_context, indent=2) if ws_context else "{}"
    return (
        f"Worksheet context:\n```json\n{ctx_block}\n```\n\n"
        f"User request: {user_message}"
    )

# ---------------------------------------------------------------------------
# LLM call
# ---------------------------------------------------------------------------

def call_llm(user_prompt: str) -> str:
    if LLM_PROVIDER == "anthropic":
        return _call_anthropic(user_prompt)
    elif LLM_PROVIDER == "openai":
        return _call_openai(user_prompt)
    elif LLM_PROVIDER == "mistralai":
        return _call_mistralai(user_prompt)
    else:
        raise ValueError(f"Unknown LLM_PROVIDER: {LLM_PROVIDER}")


def _call_anthropic(user_prompt: str) -> str:
    import anthropic
    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    message = client.messages.create(
        model=ANTHROPIC_MODEL,
        # max_tokens=4096,
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": user_prompt}],
    )
    return message.content[0].text


def _call_openai(user_prompt: str) -> str:
    from openai import OpenAI
    client = OpenAI(api_key=OPENAI_API_KEY)
    response = client.chat.completions.create(
        model=OPENAI_MODEL,
        # max_tokens=4096,
        messages=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user",   "content": user_prompt},
        ],
    )
    return response.choices[0].message.content


def _call_mistralai(user_prompt: str) -> str:
    from mistralai import Mistral
    client = Mistral(api_key=MISTRAL_API_KEY)
    response = client.chat.complete(
        model=MISTRAL_MODEL,
        # max_tokens=4096,
        messages=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user",   "content": user_prompt},
        ],
    )
    return response.choices[0].message.content

# ---------------------------------------------------------------------------
# Response parser
# ---------------------------------------------------------------------------

def parse_segments(raw_text: str) -> list:
    cleaned = re.sub(r"^```(?:json)?\s*", "", raw_text.strip(), flags=re.IGNORECASE)
    cleaned = re.sub(r"\s*```$", "", cleaned.strip())

    segments = json.loads(cleaned)
    if not isinstance(segments, list):
        raise ValueError("LLM response is not a JSON array")

    required_fields = {"id", "description", "sheet_context", "explanation", "code"}
    for i, seg in enumerate(segments):
        missing = required_fields - seg.keys()
        if missing:
            raise ValueError(f"Segment {i} missing fields: {missing}")

    return segments
