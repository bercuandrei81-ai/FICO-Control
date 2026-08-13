"""Backend package initialization for FICO Control."""

# Ensure the Mentor Check V2 name matcher is loaded when Render imports
# backend.app_mobile_api from the repository root.
from . import sitecustomize  # noqa: F401
