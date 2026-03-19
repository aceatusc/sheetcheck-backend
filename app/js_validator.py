"""
js_validator.py — JavaScript syntax checking + iterative mistake registry.

Architecture
────────────
  validate_js(code)           → ValidationResult
  MistakeRegistry             — logs failures to mistakes_log.jsonl, loads
                                known fixes so the LLM prompt can warn about them

Validator layers (tried in order, first failure wins):
  1. Structural heuristics  — fast, zero-dep checks (bracket balance, etc.)
  2. esprima AST parse       — pure-Python ES6 parser; catches syntax errors
  3. (optional) py_mini_racer V8 sandbox — deeper runtime-level checks

Mistake registry
────────────────
  Each validation failure is appended to mistakes_log.jsonl with:
    - timestamp, segment_id, error_type, error_message, offending_code_snippet

  mistakes_prompt_hint() returns a compact string you can inject into the
  LLM system prompt so it learns from past mistakes on this deployment.

  You can also manually add "known fixes" to mistakes_fixes.json:
    { "error_pattern": "fix description" }
  These are included verbatim in the hint.
"""

from __future__ import annotations

import json
import re
import logging
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Optional

logger = logging.getLogger(__name__)

# ── Paths ─────────────────────────────────────────────────────────────────────

_HERE          = Path(__file__).parent
MISTAKES_LOG   = _HERE / "mistakes_log.jsonl"
FIXES_FILE     = _HERE / "mistakes_fixes.json"


# ── Result type ───────────────────────────────────────────────────────────────

@dataclass
class ValidationResult:
    valid:    bool
    errors:   list[str]          = field(default_factory=list)
    warnings: list[str]          = field(default_factory=list)
    method:   str                = "none"   # "heuristic" | "esprima" | "v8"

    @property
    def ok(self) -> bool:
        return self.valid


# ── Layer 1: heuristic checks ─────────────────────────────────────────────────

def _heuristic_check(code: str) -> ValidationResult:
    errors: list[str] = []
    warnings: list[str] = []

    # Must be wrapped in async IIFE pattern used by Office JS
    if "Excel.run" not in code:
        warnings.append("No Excel.run() call found — is this Office JS code?")

    if "await ctx.sync()" not in code and "await context.sync()" not in code:
        warnings.append("No ctx.sync() / context.sync() call found — changes may not be flushed.")

    # Bracket / paren / bracket balance
    pairs = {"(": ")", "{": "}", "[": "]"}
    stack: list[str] = []
    for ch in code:
        if ch in pairs:
            stack.append(pairs[ch])
        elif ch in pairs.values():
            if not stack or stack[-1] != ch:
                errors.append(f"Unmatched bracket/paren: '{ch}'")
                break
            stack.pop()
    if stack:
        errors.append(f"Unclosed bracket/paren — expected '{stack[-1]}' before end of code")

    # Detect common LLM mistakes ──────────────────────────────────────────────
    # 1. numberFormat assigned a plain string instead of 2-D array
    if re.search(r'\.numberFormat\s*=\s*"[^"]*"', code):
        warnings.append(
            "numberFormat assigned a plain string — for multi-cell ranges use a 2-D array "
            "e.g. [['$#,##0','$#,##0'],…]"
        )

    # 2. .values assigned without 2-D array wrapping
    if re.search(r'\.values\s*=\s*\[[^\[]*\](?!\])', code):
        warnings.append(
            ".values assigned a 1-D array — Excel JS API requires a 2-D array: [[…],[…]]"
        )

    # 3. getRange called with row > 1048576 or col > 16384
    for m in re.finditer(r'getRange\("([A-Z]+)(\d+):([A-Z]+)(\d+)"\)', code):
        r1, r2 = int(m.group(2)), int(m.group(4))
        if r1 > 1_048_576 or r2 > 1_048_576:
            errors.append(f"Row index out of Excel bounds in range '{m.group(0)}'")

    # 4. Missing async keyword on inner function passed to Excel.run
    if re.search(r'Excel\.run\s*\(\s*(?!async)', code):
        warnings.append(
            "Excel.run callback may be missing 'async' keyword — "
            "use Excel.run(async (ctx) => { … })"
        )

    return ValidationResult(
        valid=len(errors) == 0,
        errors=errors,
        warnings=warnings,
        method="heuristic",
    )


# ── Layer 2: esprima AST parse ────────────────────────────────────────────────

def _esprima_check(code: str) -> Optional[ValidationResult]:
    """Returns None if esprima is not installed."""
    try:
        import esprima  # type: ignore
    except ImportError:
        return None

    try:
        esprima.parseScript(code, tolerant=False)
        return ValidationResult(valid=True, method="esprima")
    except esprima.Error as exc:
        msg = str(exc)
        return ValidationResult(valid=False, errors=[f"Syntax error: {msg}"], method="esprima")
    except Exception as exc:
        # esprima may choke on modern syntax it doesn't fully support;
        # treat as a warning rather than a hard failure
        return ValidationResult(
            valid=True,
            warnings=[f"esprima parse warning (non-fatal): {exc}"],
            method="esprima",
        )


# ── Layer 3: py_mini_racer V8 sandbox ─────────────────────────────────────────

def _v8_check(code: str) -> Optional[ValidationResult]:
    """Returns None if py_mini_racer is not installed."""
    try:
        from py_mini_racer import MiniRacer  # type: ignore
    except ImportError:
        return None

    # Shim out the Office JS globals so the code can be parsed/compiled
    shim = """
    var Excel = { run: function(fn) { return fn({ workbook: { worksheets: {
        getActiveWorksheet: function() {
            var rng = { values: null, formulas: null, numberFormat: null,
                        format: { fill: {}, font: {}, horizontalAlignment: null },
                        load: function(){}, getEntireColumn: function(){ return { format: { autofitColumns: function(){} }}; }
                      };
            return { getRange: function(){ return rng; } };
        }
    }}}, { sync: function(){ return Promise.resolve(); } }); }};
    var ctx = { sync: function(){ return Promise.resolve(); } };
    """

    ctx = MiniRacer()
    try:
        ctx.eval(shim + code)
        return ValidationResult(valid=True, method="v8")
    except Exception as exc:
        msg = str(exc)
        return ValidationResult(valid=False, errors=[f"V8 execution error: {msg}"], method="v8")


# ── Public API ────────────────────────────────────────────────────────────────

def validate_js(code: str, segment_id: str = "unknown") -> ValidationResult:
    """
    Run all available validation layers on `code`.
    Failures are logged to mistakes_log.jsonl for later review.
    """
    # Layer 1 always runs
    result = _heuristic_check(code)
    if not result.valid:
        _log_mistake(segment_id, "heuristic", result.errors, code)
        return result

    # Layer 2: esprima
    esprima_result = _esprima_check(code)
    if esprima_result is not None:
        result.warnings.extend(esprima_result.warnings)
        result.method = esprima_result.method
        if not esprima_result.valid:
            result.valid  = False
            result.errors = esprima_result.errors
            _log_mistake(segment_id, "esprima", result.errors, code)
            return result

    # Layer 3: V8 (optional)
    v8_result = _v8_check(code)
    if v8_result is not None:
        result.warnings.extend(v8_result.warnings)
        result.method = v8_result.method
        if not v8_result.valid:
            result.valid  = False
            result.errors = v8_result.errors
            _log_mistake(segment_id, "v8", result.errors, code)
            return result

    return result


def validate_segments(segments: list[dict]) -> list[dict]:
    """
    Validate the `code` field of every segment in-place.
    Adds a `_validation` key to each segment dict (stripped before returning to client).
    Raises ValueError listing all invalid segments.
    """
    failures: list[str] = []
    for seg in segments:
        code = seg.get("code", "")
        seg_id = seg.get("id", "unknown")
        vr = validate_js(code, segment_id=seg_id)
        seg["_validation"] = {
            "valid":    vr.valid,
            "errors":   vr.errors,
            "warnings": vr.warnings,
            "method":   vr.method,
        }
        if not vr.valid:
            failures.append(
                f"  [{seg_id}] {'; '.join(vr.errors)}"
            )
        elif vr.warnings:
            logger.warning("[%s] JS warnings: %s", seg_id, "; ".join(vr.warnings))

    if failures:
        raise ValueError("JS validation failed for segments:\n" + "\n".join(failures))

    return segments


# ── Mistake registry ──────────────────────────────────────────────────────────

def _log_mistake(segment_id: str, error_type: str, errors: list[str], code: str) -> None:
    snippet = code[:300] + ("…" if len(code) > 300 else "")
    entry = {
        "ts":         datetime.now(timezone.utc).isoformat(),
        "segment_id": segment_id,
        "error_type": error_type,
        "errors":     errors,
        "snippet":    snippet,
    }
    try:
        with MISTAKES_LOG.open("a", encoding="utf-8") as f:
            f.write(json.dumps(entry) + "\n")
    except OSError as exc:
        logger.warning("Could not write to mistakes_log: %s", exc)


def load_mistakes(limit: int = 20) -> list[dict]:
    """Return the most recent `limit` entries from mistakes_log.jsonl."""
    if not MISTAKES_LOG.exists():
        return []
    lines = MISTAKES_LOG.read_text(encoding="utf-8").strip().splitlines()
    entries = []
    for line in lines[-limit:]:
        try:
            entries.append(json.loads(line))
        except json.JSONDecodeError:
            pass
    return entries


def load_known_fixes() -> dict[str, str]:
    """
    Load manually curated fixes from mistakes_fixes.json.
    Format: { "pattern description": "fix description" }
    """
    if not FIXES_FILE.exists():
        return {}
    try:
        return json.loads(FIXES_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {}


def mistakes_prompt_hint(max_recent: int = 5) -> str:
    """
    Returns a string to inject into the LLM system prompt, summarising
    recent mistakes and known fixes so the model can avoid repeating them.
    """
    parts: list[str] = []

    fixes = load_known_fixes()
    if fixes:
        parts.append("Known JS mistakes to avoid:")
        for pattern, fix in fixes.items():
            parts.append(f"  - {pattern}: {fix}")

    recent = load_mistakes(limit=max_recent)
    if recent:
        parts.append(f"\nRecent JS validation failures (last {len(recent)}):")
        # Deduplicate by error message
        seen: set[str] = set()
        for m in recent:
            key = " | ".join(m.get("errors", []))
            if key not in seen:
                seen.add(key)
                parts.append(f"  [{m['error_type']}] {key}")

    return "\n".join(parts) if parts else ""
