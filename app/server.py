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

from flask import Flask, Blueprint, jsonify, request
from flask_cors import CORS

from params import SHARED_SECRET, STUB_SEGMENTS, STUB_RUBRIC, STUB_ASK, STUB_EDIT, STUB_VERIFY
from utils import build_user_prompt, call_llm, parse_segments, parse_json

app   = Flask(__name__)
CORS(app)
addin = Blueprint("addin", __name__, url_prefix="/addin")


# ── Auth helper ───────────────────────────────────────────────────────────────

def _auth():
    if request.headers.get("X-Addin-Secret","") != SHARED_SECRET:
        return jsonify({"error":"Unauthorized"}), 401
    return None

def _body():
    b = request.get_json(silent=True)
    if not b:
        return None, (jsonify({"error":"Invalid JSON body"}), 400)
    return b, None


# ── Routes ────────────────────────────────────────────────────────────────────

@addin.route("/health", methods=["GET"])
def health():
    return jsonify({"status":"ok"})


@addin.route("/code", methods=["POST"])
def code():
    if (e := _auth()): return e
    body, err = _body();
    if err: return err

    message  = body.get("message","").strip()
    context  = body.get("context", {})
    rubric   = body.get("rubric", None)   # optional rubric passed for context
    if not message:
        return jsonify({"error":"message is required"}), 400

    if message.lower() == "test":
        return jsonify({"segments": STUB_SEGMENTS})

    extra = {"rubric": rubric} if rubric else None
    prompt = build_user_prompt(message, context, extra)
    try:
        raw = call_llm("code", prompt)
        segments = parse_segments(raw)
    except Exception as exc:
        print(raw)
        print(str(exc))
        return jsonify({"error": str(exc)}), 502

    return jsonify({"segments": segments})


@addin.route("/ask", methods=["POST"])
def ask():
    if (e := _auth()): return e
    body, err = _body();
    if err: return err

    message  = body.get("message","").strip()
    context  = body.get("context", {})
    step     = body.get("step", {})
    history  = body.get("history", [])
    if not message:
        return jsonify({"error":"message is required"}), 400

    if message.lower() == "test":
        return jsonify(STUB_ASK)

    extra = {"current_step": step, "conversation_history": history}
    prompt = build_user_prompt(message, context, extra)
    try:
        raw    = call_llm("ask", prompt)
        result = parse_json(raw)
    except Exception as exc:
        print(raw)
        print(str(exc))
        return jsonify({"error": str(exc)}), 502

    return jsonify(result)


@addin.route("/edit", methods=["POST"])
def edit():
    if (e := _auth()): return e
    body, err = _body();
    if err: return err

    message           = body.get("message","").strip()
    context           = body.get("context", {})
    original_segment  = body.get("segment", {})
    preferred_alt_id  = body.get("preferred_alt_id", None)

    if message.lower() == "test":
        return jsonify({"segment": STUB_EDIT})

    extra = {"original_segment": original_segment, "preferred_alternative": preferred_alt_id, "user_feedback": message}
    prompt = build_user_prompt(message or "Apply user feedback to this segment.", context, extra)
    try:
        raw     = call_llm("edit", prompt)
        segment = parse_json(raw)
    except Exception as exc:
        print(raw)
        print(str(exc))
        return jsonify({"error": str(exc)}), 502

    return jsonify({"segment": segment})


@addin.route("/rubric/scaffold", methods=["POST"])
def rubric_scaffold():
    if (e := _auth()): return e
    body, err = _body();
    if err: return err

    message = body.get("message","").strip()
    context = body.get("context", {})

    if not message or message.lower() == "test":
        return jsonify(STUB_RUBRIC)

    prompt = build_user_prompt(message, context)
    try:
        raw    = call_llm("rubric_scaffold", prompt)
        rubric = parse_json(raw)
    except Exception as exc:
        print(raw)
        print(str(exc))
        return jsonify({"error": str(exc)}), 502

    return jsonify(rubric)


@addin.route("/rubric/verify", methods=["POST"])
def rubric_verify():
    if (e := _auth()): return e
    body, err = _body();
    if err: return err

    rubric  = body.get("rubric", {})
    context = body.get("context", {})

    # Stub: if the rubric matches the default stub structure, return stub results
    if not context.get("sheetData"):
        return jsonify({"results": STUB_VERIFY})

    extra = {"rubric": rubric}
    prompt = build_user_prompt("Verify the worksheet satisfies each rubric item.", context, extra)
    try:
        raw    = call_llm("rubric_verify", prompt)
        result = parse_json(raw)
    except Exception as exc:
        print(raw)
        print(str(exc))
        return jsonify({"error": str(exc)}), 502

    return jsonify({"results": result})


@addin.route("/chat", methods=["POST"])
def chat():
    if (e := _auth()): return e
    body, err = _body();
    if err: return err

    message = body.get("message","").strip()
    context = body.get("context", {})
    if not message:
        return jsonify({"error":"message is required"}), 400

    if message.lower() == "test":
        return jsonify({"response": "[stub] The assistant is ready to help with your spreadsheet questions!"})

    prompt = build_user_prompt(message, context)
    try:
        raw = call_llm("chat", prompt)
    except Exception as exc:
        print(raw)
        print(str(exc))
        return jsonify({"error": str(exc)}), 502

    return jsonify({"response": raw})


app.register_blueprint(addin)
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8883, debug=True)
