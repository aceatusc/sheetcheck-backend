"""
llm_logger.py -- Structured LLM call logging via memor.

Each call_program() invocation produces one memor Session saved as JSON to
the logs/ directory. Sessions capture:
  - The endpoint name and model used (as the session title)
  - The rendered prompt DSPy sent (Prompt message)
  - The raw LLM response (Response message) with inference time and model tag
  - Timestamps, token estimates, and any error that occurred

Directory layout
----------------
  logs/
    code/
      2025-01-15T14-32-01_abc123.json
    edit/
      ...
    ask/
      ...

Usage
-----
    from llm_logger import log_call

    with log_call("code", model="gemini/gemini-3-flash-preview") as logger:
        result = program(**kwargs)
        logger.set_response(raw_text, tokens=...)

The context manager handles timing, saves on exit, and swallows logging
errors so they never affect the main request path.
"""

from __future__ import annotations

import json
import logging
import time
from contextlib import contextmanager
from datetime import datetime, timezone
from pathlib import Path
from typing import Generator, Optional

logger = logging.getLogger(__name__)

LOGS_DIR = Path(__file__).parent / "logs"


# -- memor import with graceful degradation -----------------------------------
#
# If memor is not installed, all logging calls become no-ops so the server
# continues to work. Install with: uv add memor

try:
    from memor import Session, Prompt, Response, LLMModel  # type: ignore
    _MEMOR_AVAILABLE = True
except ImportError:
    _MEMOR_AVAILABLE = False
    logger.warning("[LLMLogger] memor not installed -- LLM call logging disabled. Run: uv add memor")


# -- LLMModel mapping ---------------------------------------------------------
#
# memor.LLMModel is an enum of known models. For models it doesn't know about
# (Gemini 3, Mistral variants) we fall back to LLMModel.UNKNOWN if available,
# otherwise pass None.

def _resolve_llm_model(model_str: str) -> Optional[object]:
    if not _MEMOR_AVAILABLE:
        return None
    _MODEL_MAP = {
        "claude-opus-4-6":               LLMModel.CLAUDE_3_OPUS    if hasattr(LLMModel, "CLAUDE_3_OPUS")    else None,
        "claude-sonnet-4-6":             LLMModel.CLAUDE_3_5_SONNET if hasattr(LLMModel, "CLAUDE_3_5_SONNET") else None,
        "claude-haiku-4-5":              LLMModel.CLAUDE_3_HAIKU    if hasattr(LLMModel, "CLAUDE_3_HAIKU")    else None,
        "gemini-2.5-flash":              LLMModel.GEMINI_2_5_FLASH if hasattr(LLMModel, "GEMINI_2_5_FLASH") else None,
        "gemini-3-flash-preview":       LLMModel.GEMINI_3_FLASH   if hasattr(LLMModel, "GEMINI_3_FLASH")   else None,
        "gemini-3.1-flash-lite-preview":LLMModel.GEMINI_3_1_FLASH_LITE if hasattr(LLMModel, "GEMINI_3_1_FLASH_LITE") else None,
        "gemini-3.1-pro-preview":     LLMModel.GEMINI_3_1_PRO     if hasattr(LLMModel, "GEMINI_3_1_PRO")     else None,
        "ministral-3b-2512":            LLMModel.MISTRAL_3B_2512  if hasattr(LLMModel, "MISTRAL_3B_2512")  else None,
        "mistral-small-2506":           LLMModel.MISTRAL_SMALL_INSTRUCT if hasattr(LLMModel, "MISTRAL_SMALL_2506") else None,
        "mistral-large-2512":           LLMModel.MISTRAL_LARGE_INSTRUCT if hasattr(LLMModel, "MISTRAL_LARGE_2512") else None,
        "gpt-4o":                        LLMModel.GPT_4O            if hasattr(LLMModel, "GPT_4O")            else None,
    }
    return _MODEL_MAP.get(model_str)   # None for unknown models -- memor accepts None


# -- CallLogger ---------------------------------------------------------------

class CallLogger:
    """Collects data for one LLM call. Used as a context manager via log_call()."""

    def __init__(self, endpoint: str, model: str):
        self.endpoint  = endpoint
        self.model     = model
        self._prompt   = ""
        self._response = ""
        self._error: Optional[str] = None
        self._start    = time.monotonic()
        self._elapsed  = 0.0

    def set_prompt(self, text: str) -> None:
        """Call with the rendered prompt string before the LLM call."""
        self._prompt = text

    def set_response(self, text: str) -> None:
        """Call with the raw response string after the LLM call."""
        self._response = text

    def set_error(self, exc: Exception) -> None:
        self._error = f"{type(exc).__name__}: {exc}"

    def _stop_timer(self) -> None:
        self._elapsed = time.monotonic() - self._start

    def save(self) -> None:
        """Build a memor Session and save it to logs/<endpoint>/."""
        if not _MEMOR_AVAILABLE:
            return
        try:
            self._stop_timer()
            ts      = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H-%M-%S")
            uid     = hex(int(time.monotonic() * 1e6))[-6:]
            title   = f"{self.endpoint} | {self.model}"
            out_dir = LOGS_DIR / self.endpoint
            out_dir.mkdir(parents=True, exist_ok=True)
            out_path = out_dir / f"{ts}_{uid}.json"

            llm_model = _resolve_llm_model(self.model)

            prompt_msg = Prompt(message=self._prompt or "(no prompt captured)")
            resp_text  = self._response if not self._error else f"ERROR: {self._error}"
            resp_msg   = Response(
                message=resp_text or "(empty response)",
                model=llm_model,
                inference_time=round(self._elapsed, 3),
            )

            session = Session(title=title, messages=[prompt_msg, resp_msg])
            session.save(str(out_path))
            logger.debug("[LLMLogger] Saved: %s", out_path)
        except Exception as exc:
            # Logging must never affect the request path
            logger.warning("[LLMLogger] Failed to save session: %s", exc)


@contextmanager
def log_call(endpoint: str, model: str) -> Generator[CallLogger, None, None]:
    """
    Context manager that creates a CallLogger, yields it to the caller,
    then saves the session on exit regardless of success or failure.

    Usage:
        with log_call("code", model="gemini-3-flash") as cl:
            cl.set_prompt(rendered_prompt)
            result = program(**kwargs)
            cl.set_response(str(result))
    """
    cl = CallLogger(endpoint, model)
    try:
        yield cl
    except Exception as exc:
        cl.set_error(exc)
        raise
    finally:
        cl._stop_timer()
        cl.save()
