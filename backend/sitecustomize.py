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

# Render often starts from /backend. Patch every HTMLResponse so the Score
# Verification entry appears in the dashboard side menu and in module top bars.
try:
    from starlette.responses import HTMLResponse as _HTMLResponse
    _ORIGINAL_HTML_RESPONSE_INIT = _HTMLResponse.__init__

    def _inject_score_link(content: str) -> str:
        if '/admin/score-check' in content:
            return content

        # Main Admin Dashboard left menu: insert after Mentor Check.
        if 'side-nav-links' in content:
            marker = '<a class="side-link" href="/admin/mentor?d='
            start = content.find(marker)
            if start != -1:
                end = content.find('</a>', start)
                if end != -1:
                    end += 4
                    button = '\n        <a class="side-link" href="/admin/score-check"><i></i>Verificare Scor</a>'
                    content = content[:end] + button + content[end:]
                    return content

        # Mentor Check top navigation.
        mentor_marker = '<a class="btn btn-light" href="/admin/hours">Control ore</a>'
        if mentor_marker in content:
            score_button = '<a class="btn btn-light" href="/admin/score-check">Verificare Scor</a>'
            return content.replace(mentor_marker, score_button + mentor_marker, 1)

        # Hours page top navigation.
        hours_marker = '<a class="btn btn-light" href="/admin/pod-ccc">POD & CCC</a>'
        if 'Control ore șoferi' in content and hours_marker in content:
            score_button = '<a class="btn btn-light" href="/admin/score-check">Verificare Scor</a>'
            return content.replace(hours_marker, score_button + hours_marker, 1)

        # POD & CCC navigation.
        pod_marker = '<a class="pc-btn" href="/admin/hours">Control ore</a>'
        if 'POD & CCC' in content and pod_marker in content:
            score_button = '<a class="pc-btn" href="/admin/score-check">Verificare Scor</a>'
            return content.replace(pod_marker, score_button + pod_marker, 1)

        # Concessions navigation.
        concessions_marker = '<a class="cn-btn" href="/admin/pod-ccc">POD & CCC</a>'
        if '<h1>Concesii</h1>' in content and concessions_marker in content:
            score_button = '<a class="cn-btn" href="/admin/score-check">Verificare Scor</a>'
            return content.replace(concessions_marker, score_button + concessions_marker, 1)

        # Atlas navigation.
        atlas_marker = '<a\n        class="atlas-btn"\n        href="/admin/hours"\n      >Control ore</a>'
        if '<h1>Atlas Paket</h1>' in content and atlas_marker in content:
            score_button = '<a class="atlas-btn" href="/admin/score-check">Verificare Scor</a>'
            return content.replace(atlas_marker, score_button + atlas_marker, 1)

        return content

    def _fico_html_response_init(self, content=None, *args, **kwargs):
        if isinstance(content, str):
            content = _inject_score_link(content)
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
