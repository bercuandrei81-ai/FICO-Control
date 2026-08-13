"""Repository-level Python startup hook for FICO Control.

Render can start the service from the repository root. Load the main backend
application and register the Verificare Scor routes on that exact FastAPI app.
"""

try:
    from backend import app_mobile_api as _app_mobile_api
    from backend.score_check import register_score_check

    register_score_check(_app_mobile_api.app)
except Exception as exc:
    print(
        "SCORE_CHECK_ROOT_STARTUP_ERROR:",
        type(exc).__name__,
        str(exc)[:500],
        flush=True,
    )
