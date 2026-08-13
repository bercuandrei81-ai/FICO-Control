"""Backend package initialization for FICO Control."""

# Keep the existing Mentor/Cortex name matcher active.
from . import sitecustomize  # noqa: F401

# Load the main FastAPI application once, then register the separate
# Verificare Scor module on the same app instance.
from . import app_mobile_api as _app_mobile_api  # noqa: E402
from .score_check import register_score_check  # noqa: E402

register_score_check(_app_mobile_api.app)
