"""Server utilities — LLM call dispatch."""

import json
import re

from params import Provider, ENDPOINT_MODELS, SYSTEM_PROMPTS
from params import ANTHROPIC_API_KEY, OPENAI_API_KEY, MISTRAL_API_KEY, GEMINI_API_KEY


def build_user_prompt(user_message: str, ws_context: dict, extra: dict = None) -> str:
    ctx_block = json.dumps(ws_context, indent=2) if ws_context else "{}"
    parts = [f"Worksheet context:\n```json\n{ctx_block}\n```\n\nUser request: {user_message}"]
    if extra:
        parts.append(f"\nExtra context:\n```json\n{json.dumps(extra, indent=2)}\n```")
    return "\n".join(parts)


def call_llm(endpoint: str, user_prompt: str) -> str:
    cfg = ENDPOINT_MODELS[endpoint]
    system = SYSTEM_PROMPTS[endpoint]
    if cfg.provider == Provider.ANTHROPIC:
        return _call_anthropic(system, user_prompt, cfg.model)
    elif cfg.provider == Provider.OPENAI:
        return _call_openai(system, user_prompt, cfg.model)
    elif cfg.provider == Provider.MISTRAL:
        return _call_mistralai(system, user_prompt, cfg.model)
    elif cfg.provider == Provider.GOOGLE:
        return _call_google(system, user_prompt, cfg.model)
    else:
        raise ValueError(f"Unknown provider: {cfg.provider}")


def _call_anthropic(system: str, user_prompt: str, model: str) -> str:
    import anthropic
    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    msg = client.messages.create(model=model, max_tokens=8096, system=system,
                                  messages=[{"role":"user","content":user_prompt}])
    return msg.content[0].text


def _call_openai(system: str, user_prompt: str, model: str) -> str:
    from openai import OpenAI
    client = OpenAI(api_key=OPENAI_API_KEY)
    r = client.chat.completions.create(model=model, max_tokens=8096,
        messages=[{"role":"system","content":system},{"role":"user","content":user_prompt}])
    return r.choices[0].message.content


def _call_mistralai(system: str, user_prompt: str, model: str) -> str:
    from mistralai import Mistral
    client = Mistral(api_key=MISTRAL_API_KEY)
    r = client.chat.complete(model=model,
        messages=[{"role":"system","content":system},{"role":"user","content":user_prompt}])
    return r.choices[0].message.content


def _call_google(system: str, user_prompt: str, model: str) -> str:
    from google import genai
    from google.genai import types
    client = genai.Client(api_key=GEMINI_API_KEY)
    chat = client.chats.create(model=model,
        config=types.GenerateContentConfig(system_instruction=system))
    return chat.send_message(user_prompt).text


def parse_json(raw_text: str):
    cleaned = re.sub(r"^```(?:json)?\s*", "", raw_text.strip(), flags=re.IGNORECASE)
    cleaned = re.sub(r"\s*```$", "", cleaned.strip())
    return json.loads(cleaned)


def parse_segments(raw_text: str) -> list:
    segments = parse_json(raw_text)
    if not isinstance(segments, list):
        raise ValueError("LLM response is not a JSON array")
    required = {"id","description","sheet_context","explanation","code"}
    for i, seg in enumerate(segments):
        missing = required - seg.keys()
        if missing:
            raise ValueError(f"Segment {i} missing fields: {missing}")
        seg.setdefault("predecessors", [])
        seg.setdefault("affordances", [])
        seg.setdefault("alternatives", [])
        seg.setdefault("qa_pairs", [])
        seg.setdefault("undo_code", "")
    return segments
