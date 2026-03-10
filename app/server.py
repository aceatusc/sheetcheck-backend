"""
server.py
---------
Lightweight Flask proxy that sits between the SheetCheck front
and the LLM API.  Responsibilities:
  - Authenticate requests via a shared secret header
  - Return stub segments when failed or user text is 'test' (no LLM called)
  - Build the system prompt with worksheet context
  - Call the LLM API
  - Parse the response into a CodeSegment[] JSON array
  - Return it to the caller
"""

from flask import Flask, Blueprint, jsonify, request
from flask_cors import CORS

from params import SHARED_SECRET, STUB_SEGMENTS
from utils import build_user_prompt, call_llm, parse_segments

# ---------------------------------------------------------------------------
# App + Blueprint setup
# ---------------------------------------------------------------------------

app   = Flask(__name__)
CORS(app)  # allow requests from the Office add-in iframe origin

addin = Blueprint("addin", __name__, url_prefix="/addin")

# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------

@addin.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok"})


@addin.route("/chat", methods=["POST"])
def chat():
    # ── Auth ──────────────────────────────────────────────────────────────────
    secret = request.headers.get("X-Addin-Secret", "")
    if secret != SHARED_SECRET:
        return jsonify({"error": "Unauthorized"}), 401

    # ── Parse body ────────────────────────────────────────────────────────────
    body = request.get_json(silent=True)
    if not body:
        return jsonify({"error": "Invalid JSON body"}), 400

    user_message = body.get("message", "").strip()
    ws_context   = body.get("context", {})

    if not user_message:
        return jsonify({"error": "message is required"}), 400

    # ── Test mode: return stub without calling the LLM ────────────────────────
    if user_message.lower() == "test":
        return jsonify({"segments": STUB_SEGMENTS})

    # ── Build prompt and call LLM ─────────────────────────────────────────────
    user_prompt = build_user_prompt(user_message, ws_context)

    try:
        raw_text = call_llm(user_prompt)
    except Exception as exc:
        return jsonify({"error": f"LLM call failed: {exc}"}), 502

    try:
        segments = parse_segments(raw_text)
    except Exception as exc:
        return jsonify({"error": f"Could not parse LLM response: {exc}",
                        "raw": raw_text}), 422

    return jsonify({"segments": segments})


@addin.route("/code", methods=["POST"])
def code():
    # ── Auth ──────────────────────────────────────────────────────────────────
    secret = request.headers.get("X-Addin-Secret", "")
    if secret != SHARED_SECRET:
        return jsonify({"error": "Unauthorized"}), 401

    # ── Parse body ────────────────────────────────────────────────────────────
    body = request.get_json(silent=True)
    if not body:
        return jsonify({"error": "Invalid JSON body"}), 400

    user_message = body.get("message", "").strip()
    ws_context   = body.get("context", {})

    if not user_message:
        return jsonify({"error": "message is required"}), 400

    # ── Test mode: return stub without calling the LLM ────────────────────────
    if user_message.lower() == "test":
        return jsonify({"segments": STUB_SEGMENTS})

    # ── Build prompt and call LLM ─────────────────────────────────────────────
    user_prompt = build_user_prompt(user_message, ws_context)

    try:
        raw_text = call_llm(user_prompt)
    except Exception as exc:
        return jsonify({"error": f"LLM call failed: {exc}"}), 502

    try:
        segments = parse_segments(raw_text)
    except Exception as exc:
        return jsonify({"error": f"Could not parse LLM response: {exc}",
                        "raw": raw_text}), 422

    return jsonify({"segments": segments})


app.register_blueprint(addin)
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8883, debug=True)
