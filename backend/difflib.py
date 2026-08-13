"""Compatibility wrapper around Python's stdlib difflib.

FICO Control imports ``SequenceMatcher`` directly in app_mobile_api.py.  This
module keeps the normal difflib API but makes SequenceMatcher name-aware so
Cortex names can safely match shorter Mentor names such as:

    Raed Sabri <-> Raed Darwish Sabri
    Octavian Florin Stan <-> Florin Stan

It also supports reversed order, accents, initials and small spelling changes.
"""

import importlib.util
import os
import re
import sysconfig
import unicodedata

_stdlib_path = os.path.join(sysconfig.get_path("stdlib"), "difflib.py")
_spec = importlib.util.spec_from_file_location("_fico_stdlib_difflib", _stdlib_path)
_stdlib = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_stdlib)

# Re-export the normal stdlib difflib API first.
for _name in dir(_stdlib):
    if not _name.startswith("__"):
        globals()[_name] = getattr(_stdlib, _name)

_OriginalSequenceMatcher = _stdlib.SequenceMatcher


def _normalize(value):
    value = unicodedata.normalize("NFKD", str(value or ""))
    value = "".join(ch for ch in value if not unicodedata.combining(ch))
    value = value.casefold()
    value = re.sub(r"[^a-z0-9]+", " ", value)
    return " ".join(value.split())


def _tokens(value):
    return [token for token in _normalize(value).split() if token]


def _token_similarity(left, right):
    if left == right:
        return 1.0
    if not left or not right:
        return 0.0
    if len(left) == 1 or len(right) == 1:
        return 0.94 if left.startswith(right) or right.startswith(left) else 0.0
    if len(left) >= 2 and len(right) >= 2 and (left.startswith(right) or right.startswith(left)):
        return 0.93
    if min(len(left), len(right)) >= 4:
        return _OriginalSequenceMatcher(None, left, right).ratio()
    return 0.0


def _name_score(left_value, right_value):
    left = _tokens(left_value)
    right = _tokens(right_value)
    if not left or not right:
        return 0.0

    # Exact normalized string.
    if left == right:
        return 1.0

    # Match tokens one-to-one, independent of order.  Prefer exact matches.
    used = set()
    matched_scores = []

    for token in left:
        best_index = None
        best_score = 0.0
        for index, candidate in enumerate(right):
            if index in used:
                continue
            score = _token_similarity(token, candidate)
            if score > best_score:
                best_index = index
                best_score = score
                if score == 1.0:
                    break
        if best_index is not None and best_score >= 0.84:
            used.add(best_index)
            matched_scores.append(best_score)

    smaller_count = min(len(left), len(right))
    larger_count = max(len(left), len(right))
    matched_count = len(matched_scores)

    # Strong rule: every token from the shorter name is present in the longer
    # name. This is the important Mentor case where a middle name is missing.
    if smaller_count >= 2 and matched_count >= smaller_count:
        average_quality = sum(matched_scores) / matched_count
        extra_tokens = larger_count - smaller_count
        penalty = min(0.06, 0.02 * extra_tokens)
        return min(0.99, 0.96 * average_quality - penalty + 0.02)

    # If both names contain at least three tokens, two very strong shared tokens
    # can still be a review-quality match, but never force a green result.
    if matched_count >= 2:
        average_quality = sum(matched_scores) / matched_count
        return min(0.84, 0.72 + 0.10 * average_quality)

    # One shared token is insufficient to identify a driver.
    return 0.0


class SequenceMatcher(_OriginalSequenceMatcher):
    def ratio(self):
        base = super().ratio()
        if isinstance(self.a, str) and isinstance(self.b, str):
            smart = _name_score(self.a, self.b)
            return max(base, smart)
        return base


# Keep star-import behaviour compatible with stdlib difflib.
__all__ = list(getattr(_stdlib, "__all__", []))
if "SequenceMatcher" not in __all__:
    __all__.append("SequenceMatcher")
