"""
logger.py — Structured request/response logging for SheetCheck.

Writes to:
  logs/sheetcheck.log   — rotating file, one JSON line per event
  stdout                — human-readable coloured summary

Usage:
    from logger import log_request, log_response, log_error, log_info
"""

import json
import logging
import logging.handlers
import os
import time
import traceback
from datetime import datetime, timezone


# ── Setup ─────────────────────────────────────────────────────────────────────

LOG_DIR  = os.path.join(os.path.dirname(__file__), "..", "logs")
LOG_FILE = os.path.join(LOG_DIR, "sheetcheck.log")
os.makedirs(LOG_DIR, exist_ok=True)

# Root logger for this app
_logger = logging.getLogger("sheetcheck")
_logger.setLevel(logging.DEBUG)
_logger.propagate = False


# ── File handler — rotating, JSON lines ───────────────────────────────────────

class _JsonFormatter(logging.Formatter):
    def format(self, record):
        payload = {
            "ts":      datetime.now(timezone.utc).isoformat(),
            "level":   record.levelname,
            "event":   record.getMessage(),
        }
        if hasattr(record, "extra"):
            payload.update(record.extra)
        if record.exc_info:
            payload["traceback"] = self.formatException(record.exc_info)
        return json.dumps(payload, ensure_ascii=False)

_file_handler = logging.handlers.RotatingFileHandler(
    LOG_FILE, maxBytes=10 * 1024 * 1024, backupCount=5, encoding="utf-8"
)
_file_handler.setLevel(logging.DEBUG)
_file_handler.setFormatter(_JsonFormatter())
_logger.addHandler(_file_handler)


# ── Console handler — human-readable ─────────────────────────────────────────

_RESET  = "\033[0m"
_COLORS = {
    "DEBUG":    "\033[36m",   # cyan
    "INFO":     "\033[32m",   # green
    "WARNING":  "\033[33m",   # yellow
    "ERROR":    "\033[31m",   # red
    "CRITICAL": "\033[35m",   # magenta
}

class _ConsoleFormatter(logging.Formatter):
    def format(self, record):
        color = _COLORS.get(record.levelname, "")
        ts    = datetime.now().strftime("%H:%M:%S")
        msg   = record.getMessage()
        extra = getattr(record, "extra", {})
        parts = [f"{color}[{ts}] {record.levelname:<8}{_RESET} {msg}"]
        if extra:
            for k, v in extra.items():
                if k in ("endpoint", "status", "duration_ms", "ip", "error"):
                    parts.append(f"  {k}={v}")
        return "\n".join(parts)

_console_handler = logging.StreamHandler()
_console_handler.setLevel(logging.INFO)
_console_handler.setFormatter(_ConsoleFormatter())
_logger.addHandler(_console_handler)


# ── Public helpers ────────────────────────────────────────────────────────────

def _emit(level: str, event: str, **kwargs):
    record = _logger.makeRecord(
        name="sheetcheck", level=getattr(logging, level),
        fn="", lno=0, msg=event, args=(), exc_info=None,
    )
    record.extra = kwargs
    _logger.handle(record)


def log_info(event: str, **kwargs):
    _emit("INFO", event, **kwargs)


def log_request(endpoint: str, ip: str, body: dict):
    """Log an incoming API request (sanitise sensitive fields)."""
    safe_body = {k: v for k, v in body.items() if k not in ("code", "undo_code")}
    # Truncate large fields
    for k in ("context", "segment", "remaining_segments", "rubric"):
        if k in safe_body and isinstance(safe_body[k], (dict, list)):
            raw = json.dumps(safe_body[k])
            safe_body[k] = raw[:200] + "…" if len(raw) > 200 else safe_body[k]
    _emit("INFO", f"→ {endpoint}", endpoint=endpoint, ip=ip, body=safe_body)


def log_response(endpoint: str, status: int, duration_ms: float, summary: str = ""):
    """Log a completed API response."""
    level = "INFO" if status < 400 else ("WARNING" if status < 500 else "ERROR")
    _emit(level, f"← {endpoint} {status}", endpoint=endpoint, status=status,
          duration_ms=round(duration_ms, 1), summary=summary)


def log_llm(endpoint: str, raw_response: str):
    """Log the raw LLM response (first 500 chars to avoid huge files)."""
    preview = raw_response[:500] + ("…" if len(raw_response) > 500 else "")
    _emit("DEBUG", f"LLM response [{endpoint}]", endpoint=endpoint, raw_preview=preview,
          raw_length=len(raw_response))


def log_error(endpoint: str, exc: Exception, raw: str = None):
    """Log an error with full traceback."""
    record = _logger.makeRecord(
        name="sheetcheck", level=logging.ERROR,
        fn="", lno=0, msg=f"✗ {endpoint} error: {exc}",
        args=(), exc_info=None,
    )
    record.extra = {
        "endpoint": endpoint,
        "error":    str(exc),
        "traceback": traceback.format_exc(),
    }
    if raw:
        record.extra["raw_preview"] = raw[:300]
    _logger.handle(record)
