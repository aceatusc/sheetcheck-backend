"""
server.py — SheetCheck API proxy
"""

import logging
from functools import wraps

from flask import Flask, Blueprint, jsonify, request
from flask_cors import CORS

from params import SHARED_SECRET, STUB_RUBRIC, STUB_ASK, STUB_EDIT, STUB_VERIFY
from stubs import STUBS, STUB_RUBRICS, STUB_VERIFIES
from utils import (
    generate_segments,
    edit_segments,
    ask_question,
    scaffold_rubric,
    verify_rubric,
    chat_response,
)

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app   = Flask(__name__)
CORS(app)
addin = Blueprint("addin", __name__, url_prefix="/addin")


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

    if message.lower().startswith("stub:"):
        key  = message[5:].strip().lower()
        segs = STUBS.get(key)
        if segs:
            return jsonify({"segments": segs})
        return jsonify({"error": f"Unknown stub '{key}'. Valid: {list(STUBS.keys())}"}), 400

    segments = generate_segments(message, context, rubric=rubric)
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

    return jsonify(ask_question(message, context, step, history))


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

    segments = edit_segments(message, context, original_segment, remaining_segments)
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
    if message.lower().startswith("stub:"):
        key = message[5:].strip().lower()
        return jsonify(STUB_RUBRICS.get(key, STUB_RUBRIC))

    return jsonify(scaffold_rubric(message, context))


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

    return jsonify({"results": verify_rubric(rubric, context)})


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

    return jsonify({"response": chat_response(message, context)})


app.register_blueprint(addin)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8883, debug=True)
