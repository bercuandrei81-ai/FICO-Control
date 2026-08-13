"""Runtime patch for safer Cortex <-> Mentor name matching.

Loaded automatically by Python through the standard sitecustomize hook.
It only changes difflib.SequenceMatcher.ratio() for name-like strings and
otherwise delegates to the standard library implementation.
"""

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
    return [
        token
        for token in (_normalize_token(part) for part in str(value or "").split())
        if token
    ]


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

    # Match tokens one-to-one, allowing initials, prefixes and small typos.
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

    # Two or more matched name parts are a strong signal. This handles
    # missing middle names such as "Andrei Dumitrascu" vs
    # "Andrei Cristinel Dumitrascu" and reversed first/last-name order.
    if smaller >= 2 and matched == smaller:
        coverage_penalty = min(0.08, 0.025 * (larger - smaller))
        exact_bonus = 0.02 * (exact / smaller)
        return min(0.98, 0.94 - coverage_penalty + exact_bonus)

    # One-token matches should never become automatic "connected" results.
    if matched == 1:
        return 0.66 if smaller == 1 else 0.58

    return 0.0


class SmartSequenceMatcher(_ORIGINAL_SEQUENCE_MATCHER):
    def ratio(self):
        base = super().ratio()
        a = self.a
        b = self.b

        if isinstance(a, str) and isinstance(b, str):
            smart = _name_token_score(a, b)
            if smart:
                return max(base, smart)

        return base


_difflib.SequenceMatcher = SmartSequenceMatcher
