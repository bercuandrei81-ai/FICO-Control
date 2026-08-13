"""Repository-level startup hook for FICO Control."""

try:
    from starlette.responses import HTMLResponse as _HTMLResponse

    _ORIGINAL_HTML_RESPONSE_INIT = _HTMLResponse.__init__

    def _fico_html_response_init(self, content=None, *args, **kwargs):
        if isinstance(content, str):
            marker = '<a class="side-link" href="/admin/mentor?d='
            if (
                'side-nav-links' in content
                and '/admin/score-check' not in content
                and marker in content
            ):
                start = content.find(marker)
                end = content.find('</a>', start)
                if end != -1:
                    end += 4
                    button = '\n        <a class="side-link" href="/admin/score-check"><i></i>Verificare Scor</a>'
                    content = content[:end] + button + content[end:]
        _ORIGINAL_HTML_RESPONSE_INIT(self, content, *args, **kwargs)

    if not getattr(_HTMLResponse, '_fico_score_button_patched', False):
        _HTMLResponse.__init__ = _fico_html_response_init
        _HTMLResponse._fico_score_button_patched = True

    from backend import app_mobile_api as _app_mobile_api
    from backend.score_check import register_score_check
    register_score_check(_app_mobile_api.app)

except Exception as exc:
    print('SCORE_CHECK_ROOT_STARTUP_ERROR:', type(exc).__name__, str(exc)[:500], flush=True)
