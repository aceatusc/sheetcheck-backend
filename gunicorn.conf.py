"""
gunicorn.conf.py — production server config for SheetCheck backend.

Run with:
    uv run gunicorn -c gunicorn.conf.py server:app
"""

bind         = "0.0.0.0:8883"
workers      = 1          # rule of thumb: 2 * CPU cores + 1
timeout      = 600        # 10 min — must be >= nginx proxy_read_timeout
keepalive    = 10
worker_class = "sync"     # sync is fine; LLM calls are I/O-bound per worker
accesslog    = "-"        # stdout — picked up by your existing logging
errorlog     = "-"
loglevel     = "info"
