"""Safer Cortex <-> Mentor name matching for FICO Control."""

import re
import unicodedata
import difflib as _difflib

_ORIGINAL_SEQUENCE_MATCHER = _difflib.SequenceMatcher


def _normalize_token(value: str) -> str:
    value = unicodedata.normalize("NFKD", value or "")
    value = "".join(ch for ch in value if not unicodedata.combining(ch))
    value = re.sub(r"[^a-zA-Z0-9]", "", value).casefold()
    return value


def _tokenize(value: str) -> list[str]:
    return [token for token in (_normalize_token(part) for part in str(value or "").split()) if token]


def _token_match(a: str, b: str) -> bool:
    if not a or not b:
        return False
    if a == b:
        return True
    if len(a) == 1:
        return b.startswith(a)
    if len(b) == 1:
        return a.startswith(b)
    if len(a) >= 2 and len(b) >= 2 and (a.startswith(b) or b.startswith(a)):
        return True
    if min(len(a), len(b)) >= 4:
        return _ORIGINAL_SEQUENCE_MATCHER(None, a, b).ratio() >= 0.86
    return False


def _name_token_score(a: str, b: str) -> float:
    left = _tokenize(a)
    right = _tokenize(b)
    if not left or not right:
        return 0.0
    used = set()
    matched = 0
    exact = 0
    for token in left:
        best_index = None
        best_exact = False
        for index, candidate in enumerate(right):
            if index in used:
                continue
            if _token_match(token, candidate):
                best_index = index
                best_exact = token == candidate
                if best_exact:
                    break
        if best_index is not None:
            used.add(best_index)
            matched += 1
            if best_exact:
                exact += 1
    smaller = min(len(left), len(right))
    larger = max(len(left), len(right))
    if smaller >= 2 and matched == smaller:
        coverage_penalty = min(0.08, 0.025 * (larger - smaller))
        exact_bonus = 0.02 * (exact / smaller)
        return min(0.98, 0.94 - coverage_penalty + exact_bonus)
    if matched == 1:
        return 0.66 if smaller == 1 else 0.58
    return 0.0


class SmartSequenceMatcher(_ORIGINAL_SEQUENCE_MATCHER):
    def ratio(self):
        base = super().ratio()
        if isinstance(self.a, str) and isinstance(self.b, str):
            smart = _name_token_score(self.a, self.b)
            if smart:
                return max(base, smart)
        return base


_difflib.SequenceMatcher = SmartSequenceMatcher

# Render often starts from /backend, so this startup hook must also patch the
# HTMLResponse used by app_mobile_api. The score button is inserted immediately
# after Mentor Check on the Admin Dashboard.
try:
    from starlette.responses import HTMLResponse as _HTMLResponse
    _ORIGINAL_HTML_RESPONSE_INIT = _HTMLResponse.__init__

    def _fico_html_response_init(self, content=None, *args, **kwargs):
        if isinstance(content, str) and 'side-nav-links' in content and '/admin/score-check' not in content:
            marker = '<a class="side-link" href="/admin/mentor?d='
            start = content.find(marker)
            if start != -1:
                end = content.find('</a>', start)
                if end != -1:
                    end += 4
                    button = '\n        <a class="side-link" href="/admin/score-check"><i></i>Verificare Scor</a>'
                    content = content[:end] + button + content[end:]
        _ORIGINAL_HTML_RESPONSE_INIT(self, content, *args, **kwargs)

    if not getattr(_HTMLResponse, '_fico_score_button_patched', False):
        _HTMLResponse.__init__ = _fico_html_response_init
        _HTMLResponse._fico_score_button_patched = True
except Exception as exc:
    print('SCORE_CHECK_HTML_PATCH_ERROR:', type(exc).__name__, str(exc)[:500], flush=True)

# Register score routes on every FastAPI app created after this hook loads.
try:
    from fastapi import FastAPI as _FastAPI
    _ORIGINAL_FASTAPI_INIT = _FastAPI.__init__

    def _fico_fastapi_init(self, *args, **kwargs):
        _ORIGINAL_FASTAPI_INIT(self, *args, **kwargs)
        try:
            try:
                from score_check import register_score_check
            except ImportError:
                from backend.score_check import register_score_check
            register_score_check(self)
        except Exception as exc:
            print('SCORE_CHECK_STARTUP_ERROR:', type(exc).__name__, str(exc)[:500], flush=True)

    if not getattr(_FastAPI, '_fico_score_check_patched', False):
        _FastAPI.__init__ = _fico_fastapi_init
        _FastAPI._fico_score_check_patched = True
except Exception as exc:
    print('SCORE_CHECK_PATCH_ERROR:', type(exc).__name__, str(exc)[:500], flush=True)
