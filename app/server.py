"""
server.py -- SheetCheck API proxy
"""

import logging
import threading
from functools import wraps

from flask import Flask, Blueprint, jsonify, request
from flask_cors import CORS

from params import SHARED_SECRET, STUB_RUBRIC, STUB_ASK, STUB_EDIT, STUB_VERIFY, ENDPOINT_MODELS, Provider
from params import ANTHROPIC_API_KEY, OPENAI_API_KEY, MISTRAL_API_KEY, GEMINI_API_KEY
from stubs import STUBS, STUB_RUBRICS, STUB_VERIFIES
from utils import (
    generate_segments,
    edit_segments,
    ask_question,
    scaffold_rubric,
    verify_rubric,
    chat_response,
    _PROVIDER_PREFIX,
    _PROVIDER_KEY,
)

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app   = Flask(__name__)
CORS(app)
addin = Blueprint("addin", __name__, url_prefix="/addin")


# -- Startup warm-up ----------------------------------------------------------
#
# On first request, several things are slow:
#   1. Heavy imports (dspy, pydantic, anthropic SDK, etc.) loaded lazily
#   2. dspy.LM() opening a connection and validating against the provider API
#   3. dspy.Module / ChainOfThought objects being instantiated
#
# We pre-warm all of these in a background thread at startup so the first real
# user request hits only the LLM round-trip latency, not initialization overhead.
# The background thread runs after the Flask app is fully set up, so it does not
# block startup or affect gunicorn worker forking.

def _warmup():
    try:
        logger.info("[Warmup] Pre-warming DSPy programs and LM clients...")
        import dspy
        from dspy_programs import get_lm, get_program

        endpoints = list(ENDPOINT_MODELS.keys())
        for endpoint in endpoints:
            cfg    = ENDPOINT_MODELS[endpoint]
            prefix = _PROVIDER_PREFIX[cfg.provider]
            key    = _PROVIDER_KEY[cfg.provider]
            # Cache the LM client (opens provider connection)
            get_lm(prefix, cfg.model.value, key, endpoint)
            # Cache the program instance (instantiates ChainOfThought)
            get_program(endpoint)
            logger.info("[Warmup]   %s -> %s/%s", endpoint, prefix, cfg.model.value)

        # Also force-import the validator so esprima is ready
        from js_validator import load_known_fixes, load_mistakes
        load_known_fixes()
        load_mistakes(limit=1)

        logger.info("[Warmup] Done -- all programs and LM clients ready.")
    except Exception as exc:
        # Warmup failure is non-fatal -- requests will still work, just slower
        # on the first call.
        logger.warning("[Warmup] Failed (non-fatal): %s", exc)

threading.Thread(target=_warmup, daemon=True, name="dspy-warmup").start()


# -- Stub message mapping -----------------------------------------------------
#
# Maps natural-language demo prompts (shown on demo buttons in the UI) to
# their stub key. The legacy "stub:KEY" prefix still works as a fallback.

STUB_MESSAGES: dict[str, str] = {
    "build a profit and loss dashboard":        "pnl",
    "build a tax plan":                         "sales",
    "build an inventory summary":               "inventory",
}


def _resolve_stub(message: str) -> str | None:
    """Return stub key for a message, or None if not a stub trigger."""
    msg = message.strip()
    if msg.lower().startswith("stub:"):
        return msg[5:].strip().lower()
    return STUB_MESSAGES.get(msg.lower())


# ── Decorators ────────────────────────────────────────────────────────────────

def require_auth(f):
    @wraps(f)
    def wrapper(*args, **kwargs):
        if request.headers.get("X-Addin-Secret", "") != SHARED_SECRET:
            return jsonify({"error": "Unauthorized"}), 401
        return f(*args, **kwargs)
    return wrapper


def require_json(f):
    @wraps(f)
    def wrapper(*args, **kwargs):
        body = request.get_json(silent=True)
        if not body:
            return jsonify({"error": "Invalid JSON body"}), 400
        return f(body, *args, **kwargs)
    return wrapper


def handle_llm_errors(f):
    @wraps(f)
    def wrapper(*args, **kwargs):
        try:
            return f(*args, **kwargs)
        except ValueError as exc:
            logger.error("Validation / parse error: %s", exc)
            return jsonify({"error": str(exc)}), 502
        except Exception as exc:
            logger.exception("Unexpected error in %s", f.__name__)
            return jsonify({"error": str(exc)}), 502
    return wrapper


# ── Routes ────────────────────────────────────────────────────────────────────

@addin.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok"})


@addin.route("/code", methods=["POST"])
@require_auth
@require_json
@handle_llm_errors
def code(body: dict):
    message = body.get("message", "").strip()
    context = body.get("context", {})
    rubric  = body.get("rubric", None)

    if not message:
        return jsonify({"error": "message is required"}), 400

    stub_key = _resolve_stub(message)
    if stub_key:
        segs = STUBS.get(stub_key)
        if segs:
            return jsonify({"segments": segs})

    chat_history = body.get("chat_history", [])
    segments = generate_segments(message, context, rubric=rubric, chat_history=chat_history)
    return jsonify({"segments": segments})


@addin.route("/ask", methods=["POST"])
@require_auth
@require_json
@handle_llm_errors
def ask(body: dict):
    message = body.get("message", "").strip()
    context = body.get("context", {})
    step    = body.get("step", {})
    history = body.get("history", [])

    if not message:
        return jsonify({"error": "message is required"}), 400
    if message.lower() == "test":
        return jsonify(STUB_ASK)

    chat_history = body.get("chat_history", [])
    return jsonify(ask_question(message, context, step, history, chat_history=chat_history))


@addin.route("/edit", methods=["POST"])
@require_auth
@require_json
@handle_llm_errors
def edit(body: dict):
    message            = body.get("message", "").strip()
    context            = body.get("context", {})
    original_segment   = body.get("segment", {})
    remaining_segments = body.get("remaining_segments", [])

    if message.lower() == "test":
        return jsonify({"segments": STUB_EDIT})

    chat_history = body.get("chat_history", [])
    segments = edit_segments(message, context, original_segment, remaining_segments, chat_history=chat_history)
    return jsonify({"segments": segments})


@addin.route("/rubric/scaffold", methods=["POST"])
@require_auth
@require_json
@handle_llm_errors
def rubric_scaffold(body: dict):
    message = body.get("message", "").strip()
    context = body.get("context", {})

    if not message:
        return jsonify(STUB_RUBRIC)
    stub_key = _resolve_stub(message)
    if stub_key:
        return jsonify(STUB_RUBRICS.get(stub_key, STUB_RUBRIC))

    chat_history = body.get("chat_history", [])
    return jsonify(scaffold_rubric(message, context, chat_history=chat_history))


@addin.route("/rubric/verify", methods=["POST"])
@require_auth
@require_json
@handle_llm_errors
def rubric_verify(body: dict):
    rubric  = body.get("rubric", {})
    context = body.get("context", {})

    stub_key = rubric.get("stub_key", "")
    if stub_key and stub_key in STUB_VERIFIES:
        return jsonify({"results": STUB_VERIFIES[stub_key]})
    if not context.get("sheetData"):
        return jsonify({"results": STUB_VERIFY})

    chat_history = body.get("chat_history", [])
    return jsonify({"results": verify_rubric(rubric, context, chat_history=chat_history)})


@addin.route("/interactions", methods=["POST"])
@require_json
def interactions(body: dict):
    """
    Receive a session interaction log from the front-end.
    Accepts unauthenticated requests because sendBeacon cannot set custom
    headers — the payload is non-sensitive telemetry only.
    Written as NDJSON (one event per line) to logs/interactions/.
    """
    import json
    from pathlib import Path

    session_id    = body.get("session_id", "unknown")
    session_start = body.get("session_start", "")
    events        = body.get("events", [])

    if not events:
        return jsonify({"ok": True, "events": 0})

    out_dir = Path(__file__).parent / "logs" / "interactions"
    out_dir.mkdir(parents=True, exist_ok=True)

    # One file per session: <start_date>_<session_id>.ndjson
    date_slug = session_start[:10] if session_start else "unknown"
    out_path  = out_dir / f"{date_slug}_{session_id}.ndjson"

    with out_path.open("w", encoding="utf-8") as f:
        # Header line with session metadata
        f.write(json.dumps({
            "session_id":    session_id,
            "session_start": session_start,
            "event_count":   len(events),
        }) + "\n")
        # One event per line
        for ev in events:
            f.write(json.dumps(ev) + "\n")

    logger.info("[Interactions] session=%s  events=%d  -> %s", session_id, len(events), out_path.name)
    return jsonify({"ok": True, "events": len(events)})


@addin.route("/chat", methods=["POST"])
@require_auth
@require_json
@handle_llm_errors
def chat(body: dict):
    message = body.get("message", "").strip()
    context = body.get("context", {})

    if not message:
        return jsonify({"error": "message is required"}), 400
    if message.lower() == "test":
        return jsonify({"response": "[stub] The assistant is ready to help with your spreadsheet questions!"})

    chat_history = body.get("chat_history", [])
    return jsonify({"response": chat_response(message, context, chat_history=chat_history)})


app.register_blueprint(addin)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8883, debug=True)
