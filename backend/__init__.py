"""Backend package initialization for FICO Control."""

import re

# Keep the existing Mentor/Cortex name matcher active.
from . import sitecustomize  # noqa: F401

# Load the main FastAPI application once, then register the separate
# Verificare Scor module on the same app instance.
from . import app_mobile_api as _app_mobile_api  # noqa: E402
from .score_check import register_score_check  # noqa: E402

register_score_check(_app_mobile_api.app)


# ---------------------------------------------------------------------------
# Atlas Paket V4 safety/recovery patch
# ---------------------------------------------------------------------------

_atlas_original_build_assignments = _app_mobile_api.atlas_build_assignments
_atlas_original_page_html = _app_mobile_api.atlas_page_html


def _atlas_expected_total(text):
    upper = str(text or "").upper()
    for pattern in (
        r"TOTAL\s*PACKAGES\s*[:\-]?\s*(\d{1,4})",
        r"TOTAL\s*[:\-]?\s*(\d{1,4})\s*PACKAGES",
    ):
        match = re.search(pattern, upper)
        if match:
            try:
                return int(match.group(1))
            except Exception:
                pass
    return None


def _atlas_page_info(text):
    match = re.search(
        r"\bPAGE\s*(\d+)\s*(?:OF|/)\s*(\d+)\b",
        str(text or ""),
        flags=re.IGNORECASE,
    )
    if not match:
        return None, None
    return int(match.group(1)), int(match.group(2))


def _atlas_robust_build_assignments(cortex_routes, pages):
    assignments, review = _atlas_original_build_assignments(
        cortex_routes,
        pages,
    )

    seen_tracking = {
        item.get("tracking")
        for item in assignments
        if item.get("tracking")
    }
    seen_tracking.update(
        item.get("tracking")
        for item in review
        if item.get("tracking") and item.get("tracking") != "—"
    )

    pages_sorted = sorted(
        pages,
        key=lambda page: (
            page.get("page_number") is None,
            page.get("page_number")
            if page.get("page_number") is not None
            else page.get("upload_index", 0),
            page.get("upload_index", 0),
        ),
    )

    current_route = None

    for page in pages_sorted:
        filename = page.get("filename") or "Atlas"
        text = str(page.get("text") or "")

        for raw_line in text.splitlines():
            line = raw_line.strip()
            if not line:
                continue

            routes = _app_mobile_api.atlas_route_codes_from_text(line)
            if routes:
                cortex_candidates = [
                    route for route in routes if route in cortex_routes
                ]
                current_route = (
                    cortex_candidates[0]
                    if len(cortex_candidates) == 1
                    else routes[0]
                )

            tracking_ids = _app_mobile_api.atlas_tracking_ids_from_text(line)

            for tracking_id in tracking_ids:
                if tracking_id in seen_tracking:
                    continue

                seen_tracking.add(tracking_id)

                if not current_route:
                    review.append({
                        "file": filename,
                        "route": "—",
                        "tracking": tracking_id,
                        "reason": "Tracking ID fără o rută identificată",
                    })
                    continue

                if current_route not in cortex_routes:
                    review.append({
                        "file": filename,
                        "route": current_route,
                        "tracking": tracking_id,
                        "reason": "Ruta nu există în Excelul Cortex",
                    })
                    continue

                assignments.append({
                    "driver": cortex_routes[current_route],
                    "route": current_route,
                    "tracking": tracking_id,
                    "file": filename,
                })

    assignments.sort(
        key=lambda item: (
            item["driver"].casefold(),
            item["route"],
            item["tracking"],
        )
    )

    page_info = []
    for page in pages_sorted:
        number = page.get("page_number")
        total = page.get("page_total")
        if number is None or total is None:
            number, total = _atlas_page_info(page.get("text"))
        if number is not None and total is not None:
            page_info.append((number, total))

    if page_info:
        expected_pages = max(total for _, total in page_info)
        received_pages = {number for number, _ in page_info}
        missing_pages = [
            number
            for number in range(1, expected_pages + 1)
            if number not in received_pages
        ]
        if missing_pages:
            review.append({
                "file": "Atlas",
                "route": "—",
                "tracking": "—",
                "reason": (
                    "Lipsesc paginile Atlas: "
                    + ", ".join(str(number) for number in missing_pages)
                    + f" din {expected_pages}"
                ),
            })

    expected_totals = [
        _atlas_expected_total(page.get("text"))
        for page in pages_sorted
    ]
    expected_totals = [value for value in expected_totals if value]
    expected_total = max(expected_totals) if expected_totals else None

    detected_ids = {
        item.get("tracking")
        for item in assignments
        if item.get("tracking") and item.get("tracking") != "—"
    }
    detected_ids.update(
        item.get("tracking")
        for item in review
        if item.get("tracking") and item.get("tracking") != "—"
    )
    detected_total = len(detected_ids)

    if expected_total and detected_total != expected_total:
        difference = expected_total - detected_total

        if difference > 0:
            reason = (
                f"VERIFICARE TOTAL: Amazon arată {expected_total} pachete, "
                f"dar sistemul a identificat {detected_total}. "
                f"Lipsesc {difference} pachete."
            )
        else:
            reason = (
                f"VERIFICARE TOTAL: Amazon arată {expected_total} pachete, "
                f"dar sistemul a identificat {detected_total}. "
                "Verifică duplicatele OCR."
            )

        review.append({
            "file": "Atlas",
            "route": "—",
            "tracking": "—",
            "reason": reason,
        })

    return assignments, review


def _atlas_robust_page_html(assignments=None, review=None, error=""):
    review = review or []
    page = _atlas_original_page_html(assignments, review, error)

    important_warnings = [
        str(item.get("reason") or "")
        for item in review
        if str(item.get("reason") or "").startswith("VERIFICARE TOTAL")
        or str(item.get("reason") or "").startswith("Lipsesc paginile")
    ]

    if important_warnings:
        banner = (
            '<div style="margin:18px 0;padding:16px 18px;border-radius:13px;'
            'background:#fff1f0;border:2px solid #d92d20;color:#b42318;'
            'font-weight:900;font-size:15px">⚠ '
            + "<br>⚠ ".join(
                _app_mobile_api.html.escape(text)
                for text in important_warnings
            )
            + "</div>"
        )

        page = page.replace(
            '<section class="atlas-stats">',
            banner + '<section class="atlas-stats">',
            1,
        )

    page = page.replace(
        'name="atlas_images"\n          accept=',
        'id="atlasImagesInput"\n          name="atlas_images"\n          accept=',
        1,
    )

    selection_script = r"""
<script>
(() => {
  const input = document.getElementById('atlasImagesInput');
  if (!input) return;

  const info = document.createElement('div');
  info.style.cssText =
    'margin-top:8px;font-size:12px;font-weight:900;color:#147a42;line-height:1.4';

  input.insertAdjacentElement('afterend', info);

  const refresh = () => {
    const files = Array.from(input.files || []);
    info.textContent = files.length
      ? `${files.length} poze selectate: ${files.map(file => file.name).join(' · ')}`
      : '0 poze selectate';
  };

  input.addEventListener('change', refresh);
  refresh();
})();
</script>
"""

    page = page.replace("</body>", selection_script + "</body>", 1)
    return page


_app_mobile_api.atlas_build_assignments = _atlas_robust_build_assignments
_app_mobile_api.atlas_page_html = _atlas_robust_page_html

print("ATLAS_PAKET_V4_ROBUST_VALIDATION_LOADED", flush=True)
