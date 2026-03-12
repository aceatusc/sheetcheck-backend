"""
server.py — SheetCheck API proxy
Routes:
  GET  /addin/health
  POST /addin/code            — generate code segments from user task
  POST /addin/ask             — follow-up Q&A about a specific step
  POST /addin/edit            — modify a specific segment based on feedback
  POST /addin/rubric/scaffold — generate initial rubric for a task
  POST /addin/rubric/verify   — evaluate worksheet against rubric
  POST /addin/chat            — general LLM chat proxy
"""

import time
from flask import Flask, Blueprint, jsonify, request
from flask_cors import CORS

from params import SHARED_SECRET, STUB_SEGMENTS, STUB_RUBRIC, STUB_ASK, STUB_EDIT, STUB_VERIFY
from utils import build_user_prompt, call_llm, parse_segments, parse_json
from logger import log_info, log_request, log_response, log_llm, log_error

app   = Flask(__name__)
CORS(app)
addin = Blueprint("addin", __name__, url_prefix="/addin")

log_info("SheetCheck server starting")


# ── Auth helper ───────────────────────────────────────────────────────────────

def _auth():
    if request.headers.get("X-Addin-Secret", "") != SHARED_SECRET:
        log_response(request.path, 401, 0, "Unauthorized")
        return jsonify({"error": "Unauthorized"}), 401
    return None

def _body():
    b = request.get_json(silent=True)
    if not b:
        log_response(request.path, 400, 0, "Invalid JSON body")
        return None, (jsonify({"error": "Invalid JSON body"}), 400)
    return b, None

def _ip():
    return request.headers.get("X-Forwarded-For", request.remote_addr)


# ── Routes ────────────────────────────────────────────────────────────────────

@addin.route("/health", methods=["GET"])
def health():
    log_info("health check", ip=_ip())
    return jsonify({"status": "ok"})


@addin.route("/code", methods=["POST"])
def code():
    if (e := _auth()): return e
    body, err = _body()
    if err: return err

    t0      = time.monotonic()
    message = body.get("message", "").strip()
    context = body.get("context", {})
    rubric  = body.get("rubric", None)

    log_request("/addin/code", _ip(), body)

    if not message:
        log_response("/addin/code", 400, 0, "missing message")
        return jsonify({"error": "message is required"}), 400

    if message.lower() == "test":
        ms = (time.monotonic() - t0) * 1000
        log_response("/addin/code", 200, ms, f"stub — {len(STUB_SEGMENTS)} segments")
        return jsonify({"segments": STUB_SEGMENTS})

    extra  = {"rubric": rubric} if rubric else None
    prompt = build_user_prompt(message, context, extra)
    raw    = None
    try:
        raw      = call_llm("code", prompt)
        log_llm("code", raw)
        segments = parse_segments(raw)
        ms       = (time.monotonic() - t0) * 1000
        log_response("/addin/code", 200, ms, f"{len(segments)} segments")
        return jsonify({"segments": segments})
    except Exception as exc:
        log_error("/addin/code", exc, raw)
        return jsonify({"error": str(exc), "raw": raw}), 502


@addin.route("/ask", methods=["POST"])
def ask():
    if (e := _auth()): return e
    body, err = _body()
    if err: return err

    t0      = time.monotonic()
    message = body.get("message", "").strip()
    context = body.get("context", {})
    step    = body.get("step", {})
    history = body.get("history", [])

    log_request("/addin/ask", _ip(), body)

    if not message:
        log_response("/addin/ask", 400, 0, "missing message")
        return jsonify({"error": "message is required"}), 400

    if message.lower() == "test":
        ms = (time.monotonic() - t0) * 1000
        log_response("/addin/ask", 200, ms, "stub")
        return jsonify(STUB_ASK)

    extra  = {"current_step": step, "conversation_history": history}
    prompt = build_user_prompt(message, context, extra)
    raw    = None
    try:
        raw    = call_llm("ask", prompt)
        log_llm("ask", raw)
        result = parse_json(raw)
        ms     = (time.monotonic() - t0) * 1000
        log_response("/addin/ask", 200, ms)
        return jsonify(result)
    except Exception as exc:
        log_error("/addin/ask", exc, raw)
        return jsonify({"error": str(exc), "raw": raw}), 502


@addin.route("/edit", methods=["POST"])
def edit():
    if (e := _auth()): return e
    body, err = _body()
    if err: return err

    t0                 = time.monotonic()
    message            = body.get("message", "").strip()
    context            = body.get("context", {})
    original_segment   = body.get("segment", {})
    preferred_alt_id   = body.get("preferred_alt_id", None)
    remaining_segments = body.get("remaining_segments", [])

    log_request("/addin/edit", _ip(), body)

    if message.lower() == "test":
        ms = (time.monotonic() - t0) * 1000
        log_response("/addin/edit", 200, ms, "stub")
        return jsonify({"segments": [STUB_EDIT]})

    extra = {
        "original_segment":      original_segment,
        "remaining_segments":    remaining_segments,
        "preferred_alternative": preferred_alt_id,
        "user_feedback":         message,
    }
    prompt = build_user_prompt(message or "Apply user feedback to this segment.", context, extra)
    raw    = None
    try:
        raw      = call_llm("edit", prompt)
        log_llm("edit", raw)
        segments = parse_segments(raw)
        ms       = (time.monotonic() - t0) * 1000
        log_response("/addin/edit", 200, ms, f"{len(segments)} segments in chain")
        return jsonify({"segments": segments})
    except Exception as exc:
        log_error("/addin/edit", exc, raw)
        return jsonify({"error": str(exc), "raw": raw}), 502


@addin.route("/rubric/scaffold", methods=["POST"])
def rubric_scaffold():
    if (e := _auth()): return e
    body, err = _body()
    if err: return err

    t0      = time.monotonic()
    message = body.get("message", "").strip()
    context = body.get("context", {})

    log_request("/addin/rubric/scaffold", _ip(), body)

    if not message or message.lower() == "test":
        ms = (time.monotonic() - t0) * 1000
        log_response("/addin/rubric/scaffold", 200, ms, "stub")
        return jsonify(STUB_RUBRIC)

    prompt = build_user_prompt(message, context)
    raw    = None
    try:
        raw    = call_llm("rubric_scaffold", prompt)
        log_llm("rubric_scaffold", raw)
        rubric = parse_json(raw)
        ms     = (time.monotonic() - t0) * 1000
        hard   = len(rubric.get("hard_requirements", []))
        soft   = len(rubric.get("soft_requirements", []))
        log_response("/addin/rubric/scaffold", 200, ms, f"{hard} hard, {soft} soft")
        return jsonify(rubric)
    except Exception as exc:
        log_error("/addin/rubric/scaffold", exc, raw)
        return jsonify({"error": str(exc), "raw": raw}), 502


@addin.route("/rubric/verify", methods=["POST"])
def rubric_verify():
    if (e := _auth()): return e
    body, err = _body()
    if err: return err

    t0      = time.monotonic()
    rubric  = body.get("rubric", {})
    context = body.get("context", {})

    log_request("/addin/rubric/verify", _ip(), body)

    if not context.get("sheetData"):
        ms = (time.monotonic() - t0) * 1000
        log_response("/addin/rubric/verify", 200, ms, "stub — no sheet data")
        return jsonify({"results": STUB_VERIFY})

    extra  = {"rubric": rubric}
    prompt = build_user_prompt("Verify the worksheet satisfies each rubric item.", context, extra)
    raw    = None
    try:
        raw    = call_llm("rubric_verify", prompt)
        log_llm("rubric_verify", raw)
        result = parse_json(raw)
        ms     = (time.monotonic() - t0) * 1000
        log_response("/addin/rubric/verify", 200, ms, f"{len(result)} results")
        return jsonify({"results": result})
    except Exception as exc:
        log_error("/addin/rubric/verify", exc, raw)
        return jsonify({"error": str(exc), "raw": raw}), 502


@addin.route("/chat", methods=["POST"])
def chat():
    if (e := _auth()): return e
    body, err = _body()
    if err: return err

    t0      = time.monotonic()
    message = body.get("message", "").strip()
    context = body.get("context", {})

    log_request("/addin/chat", _ip(), body)

    if not message:
        log_response("/addin/chat", 400, 0, "missing message")
        return jsonify({"error": "message is required"}), 400

    if message.lower() == "test":
        ms = (time.monotonic() - t0) * 1000
        log_response("/addin/chat", 200, ms, "stub")
        return jsonify({"response": "[stub] The assistant is ready to help with your spreadsheet questions!"})

    prompt = build_user_prompt(message, context)
    raw    = None
    try:
        raw = call_llm("chat", prompt)
        log_llm("chat", raw)
        ms  = (time.monotonic() - t0) * 1000
        log_response("/addin/chat", 200, ms)
        return jsonify({"response": raw})
    except Exception as exc:
        log_error("/addin/chat", exc, raw)
        return jsonify({"error": str(exc)}), 502


app.register_blueprint(addin)
if __name__ == "__main__":
    log_info("Starting Flask dev server.")
    app.run(host="0.0.0.0", port=8883, debug=True)
