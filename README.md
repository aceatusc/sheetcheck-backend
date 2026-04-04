# Pista Backend

Flask-based API backend for Pista, providing LLM-powered spreadsheet analysis and validation.

## Setup

```bash
# Create and activate virtual environment with uv
uv venv
source .venv/bin/activate

# Install dependencies
uv sync

# Set environment variables (.env or export)
export ANTHROPIC_API_KEY=your_key
export OPENAI_API_KEY=your_key
export MISTRAL_API_KEY=your_key
export GEMINI_API_KEY=your_key
```

## Running

```bash
# Development
python app/server.py

# Production (gunicorn)
gunicorn -c gunicorn.conf.py app.server:app
```

The server runs on `http://localhost:5000` by default.

## Project Structure

- **server.py** — Flask API with endpoints for segment generation, editing, Q&A, rubric scaffolding/verification, and chat
- **dspy_programs.py** — DSPy signatures and programs for LLM operations with support for Anthropic, OpenAI, Mistral, and Google
- **utils.py** — LLM dispatch layer, thread-safe context management, and wire format handling
- **js_validator.py** — JavaScript syntax validation and iterative mistake tracking
- **params.py** — Configuration, API keys, model endpoints, and constants
- **stubs.py** — Mock data for testing/development
- **llm_logger.py** — Logging utilities for LLM interactions
- **gunicorn.conf.py** — Production server configuration

## API Endpoints

- `POST /addin/generate` — Generate segments from data
- `POST /addin/edit` — Edit existing segments
- `POST /addin/ask` — Ask questions about content
- `POST /addin/scaffold-rubric` — Create rubric structure
- `POST /addin/verify-rubric` — Validate rubric accuracy
- `POST /addin/chat` — Chat interactions
