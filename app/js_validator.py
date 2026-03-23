"""
js_validator.py — JavaScript syntax checking + iterative mistake registry.
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

_HERE        = Path(__file__).parent
MISTAKES_LOG = _HERE / "mistakes_log.jsonl"
FIXES_FILE   = _HERE / "mistakes_fixes.json"


@dataclass
class ValidationResult:
    valid:    bool
    errors:   list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    method:   str       = "none"

    @property
    def ok(self) -> bool:
        return self.valid


def _heuristic_check(code: str) -> ValidationResult:
    errors: list[str] = []
    warnings: list[str] = []

    if "Excel.run" not in code:
        warnings.append("No Excel.run() call found — is this Office JS code?")

    if "await ctx.sync()" not in code and "await context.sync()" not in code:
        warnings.append("No ctx.sync() / context.sync() call found — changes may not be flushed.")

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

    if re.search(r'\.numberFormat\s*=\s*"[^"]*"', code):
        warnings.append(
            "numberFormat assigned a plain string — for multi-cell ranges use a 2-D array "
            "e.g. [['$#,##0','$#,##0'],…]"
        )

    if re.search(r'\.values\s*=\s*\[[^\[]*\](?!\])', code):
        warnings.append(
            ".values assigned a 1-D array — Excel JS API requires a 2-D array: [[…],[…]]"
        )

    if re.search(r'\.conditionalFormatting\b', code):
        errors.append(
            "conditionalFormatting API used — not supported in Office JS add-ins. "
            "Apply formatting explicitly in a loop instead."
        )

    if re.search(r'\.autofit\s*\(\s*\)', code):
        errors.append(
            "autofit() is not a function in Office JS — use .format.autofitColumns() instead. "
            "BAD: range.getEntireColumn().autofit(). "
            "GOOD: range.getEntireColumn().format.autofitColumns()"
        )

    # Detect range.getRange(...) — getRange is a Worksheet method, not a Range method.
    # Look for a variable that is likely a range (assigned via getRange/getCell/getUsedRange)
    # and then has .getRange() called on it.
    range_vars = set(re.findall(
        r'(?:const|let|var)\s+(\w+)\s*=\s*\w+\.(?:getRange|getCell|getUsedRange|getUsedRangeOrNullObject|getEntireColumn|getEntireRow)\b',
        code
    ))
    for var in range_vars:
        if re.search(rf'\b{re.escape(var)}\.getRange\s*\(', code):
            errors.append(
                f"range.getRange() is not a function — getRange() belongs to Worksheet, not Range. "
                f"BAD: {var}.getRange('A1'). GOOD: sheet.getRange('A1')"
            )

    # Dimension mismatch: detect getRange("A1:Cx") paired with a literal array
    # whose row count visibly doesn't match the range row span.
    for m in re.finditer(
        r'getRange\("[A-Z]+(\d+):[A-Z]+(\d+)"\)[^;]*?\.(?:values|formulas|numberFormat)\s*=\s*(\[)',
        code, re.DOTALL
    ):
        start_row, end_row = int(m.group(1)), int(m.group(2))
        range_rows = end_row - start_row + 1
        # Count top-level inner arrays in the literal that follows
        literal_start = m.start(3)
        depth, inner_count, i = 0, 0, literal_start
        while i < len(code):
            if code[i] == '[':
                if depth == 1:
                    inner_count += 1
                depth += 1
            elif code[i] == ']':
                depth -= 1
                if depth == 0:
                    break
            i += 1
        if inner_count > 0 and inner_count != range_rows:
            warnings.append(
                f"Array dimensions may not match range size: range spans {range_rows} row(s) "
                f"but array has {inner_count} inner array(s). "
                "Ensure the 2-D array is exactly [rows × cols]."
            )

    for m in re.finditer(r'getRange\("([A-Z]+)(\d+):([A-Z]+)(\d+)"\)', code):
        r1, r2 = int(m.group(2)), int(m.group(4))
        if r1 > 1_048_576 or r2 > 1_048_576:
            errors.append(f"Row index out of Excel bounds in range '{m.group(0)}'")

    if re.search(r'Excel\.run\s*\(\s*(?!async)', code):
        warnings.append(
            "Excel.run callback may be missing 'async' keyword — "
            "use Excel.run(async (ctx) => { … })"
        )

    return ValidationResult(valid=len(errors) == 0, errors=errors, warnings=warnings, method="heuristic")


def _esprima_check(code: str) -> Optional[ValidationResult]:
    """Returns None if esprima is not installed."""
    try:
        import esprima  # type: ignore
    except ImportError:
        return None

    # Office JS uses top-level `await Excel.run(...)` which is not valid in
    # ES2017 script context (what esprima supports). Wrapping in an async IIFE
    # puts the await inside an async function, which is legal, while still
    # catching real syntax errors in the body.
    wrapped = f"(async function(){{{code}}})();"
    try:
        esprima.parseScript(wrapped, tolerant=False)
        return ValidationResult(valid=True, method="esprima")
    except esprima.Error as exc:
        return ValidationResult(valid=False, errors=[f"Syntax error: {exc}"], method="esprima")
    except Exception as exc:
        return ValidationResult(valid=True, warnings=[f"esprima parse warning (non-fatal): {exc}"], method="esprima")


def _v8_check(code: str) -> Optional[ValidationResult]:
    """Returns None if py_mini_racer is not installed."""
    try:
        from py_mini_racer import MiniRacer  # type: ignore
    except ImportError:
        return None

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
        return ValidationResult(valid=False, errors=[f"V8 execution error: {exc}"], method="v8")


def validate_js(code: str, segment_id: str = "unknown") -> ValidationResult:
    result = _heuristic_check(code)
    if not result.valid:
        _log_mistake(segment_id, "heuristic", result.errors, code)
        return result

    esprima_result = _esprima_check(code)
    if esprima_result is not None:
        result.warnings.extend(esprima_result.warnings)
        result.method = esprima_result.method
        if not esprima_result.valid:
            result.valid  = False
            result.errors = esprima_result.errors
            _log_mistake(segment_id, "esprima", result.errors, code)
            return result

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
    failures: list[str] = []
    for seg in segments:
        code   = seg.get("code", "")
        seg_id = seg.get("id", "unknown")
        vr = validate_js(code, segment_id=seg_id)
        seg["_validation"] = {"valid": vr.valid, "errors": vr.errors, "warnings": vr.warnings, "method": vr.method}
        if not vr.valid:
            failures.append(f"  [{seg_id}] {'; '.join(vr.errors)}")
        elif vr.warnings:
            logger.warning("[%s] JS warnings: %s", seg_id, "; ".join(vr.warnings))

    if failures:
        logger.warning("JS validation failed for segments:\n" + "\n".join(failures))
    return segments


def _log_mistake(segment_id: str, error_type: str, errors: list[str], code: str) -> None:
    entry = {
        "ts": datetime.now(timezone.utc).isoformat(),
        "segment_id": segment_id,
        "error_type": error_type,
        "errors": errors,
        "snippet": code[:300] + ("…" if len(code) > 300 else ""),
    }
    try:
        with MISTAKES_LOG.open("a", encoding="utf-8") as f:
            f.write(json.dumps(entry) + "\n")
    except OSError as exc:
        logger.warning("Could not write to mistakes_log: %s", exc)


def load_mistakes(limit: int = 20) -> list[dict]:
    if not MISTAKES_LOG.exists():
        return []
    entries = []
    for line in MISTAKES_LOG.read_text(encoding="utf-8").strip().splitlines()[-limit:]:
        try:
            entries.append(json.loads(line))
        except json.JSONDecodeError:
            pass
    return entries


def load_known_fixes() -> dict[str, str]:
    if not FIXES_FILE.exists():
        return {}
    try:
        return json.loads(FIXES_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {}


def mistakes_prompt_hint(max_recent: int = 5) -> str:
    parts: list[str] = []

    fixes = load_known_fixes()
    if fixes:
        parts.append("Known JS mistakes to avoid:")
        for pattern, fix in fixes.items():
            parts.append(f"  - {pattern}: {fix}")

    recent = load_mistakes(limit=max_recent)
    if recent:
        parts.append(f"\nRecent JS validation failures (last {len(recent)}):")
        seen: set[str] = set()
        for m in recent:
            key = " | ".join(m.get("errors", []))
            if key not in seen:
                seen.add(key)
                parts.append(f"  [{m['error_type']}] {key}")

    return "\n".join(parts) if parts else ""
