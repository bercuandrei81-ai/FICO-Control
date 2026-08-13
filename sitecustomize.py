"""Repository-level Python startup hook.

FICO Control customizations are loaded from backend.__init__ when the backend
package is imported. Keeping this root hook lightweight avoids initializing the
web application during unrelated Python startup commands.
"""
