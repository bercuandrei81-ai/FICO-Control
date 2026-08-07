from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import HTMLResponse, RedirectResponse, StreamingResponse, FileResponse
from pydantic import BaseModel
from datetime import date, datetime
import sqlite3
import io
import csv
import html
import openpyxl
import os
import uuid
import unicodedata
from difflib import SequenceMatcher
import json
import urllib.request
import urllib.parse
import re

DB = "fico.db"
UPLOAD_DIR = "uploads"
MAX_IMAGE_BYTES = 10 * 1024 * 1024
ALLOWED_IMAGE_TYPES = {"image/jpeg", "image/png", "image/webp"}

os.makedirs(UPLOAD_DIR, exist_ok=True)

app = FastAPI(title="FICO Control")


def db():
    conn = sqlite3.connect(DB)
    conn.row_factory = sqlite3.Row
    return conn


def ensure_column(conn, table, column, definition):
    columns = {
        row["name"]
        for row in conn.execute(f"PRAGMA table_info({table})").fetchall()
    }
    if column not in columns:
        conn.execute(f"ALTER TABLE {table} ADD COLUMN {column} {definition}")


def init_db():
    conn = db()
    conn.executescript("""
    CREATE TABLE IF NOT EXISTS drivers(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT UNIQUE NOT NULL,
        active INTEGER DEFAULT 1
    );

    CREATE TABLE IF NOT EXISTS daily_required(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        work_date TEXT NOT NULL,
        driver_id INTEGER NOT NULL,
        UNIQUE(work_date, driver_id),
        FOREIGN KEY(driver_id) REFERENCES drivers(id)
    );

    CREATE TABLE IF NOT EXISTS submissions(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        work_date TEXT NOT NULL,
        driver_id INTEGER NOT NULL,
        fico_score INTEGER NOT NULL,
        submitted_at TEXT NOT NULL,
        UNIQUE(work_date, driver_id),
        FOREIGN KEY(driver_id) REFERENCES drivers(id)
    );
    """)

    ensure_column(conn, "submissions", "entered_full_name", "TEXT")
    ensure_column(conn, "submissions", "proof_filename", "TEXT")
    ensure_column(conn, "submissions", "proof_original_name", "TEXT")
    ensure_column(conn, "submissions", "detected_fico_score", "INTEGER")
    ensure_column(conn, "submissions", "verification_status", "TEXT")
    ensure_column(conn, "submissions", "name_match_score", "REAL")

    conn.commit()
    conn.close()


init_db()



def normalize_name(value: str) -> str:
    value = unicodedata.normalize("NFKD", value or "")
    value = "".join(ch for ch in value if not unicodedata.combining(ch))
    value = value.replace("-", " ")
    value = " ".join(value.strip().split())
    return value.casefold()


def name_similarity(a: str, b: str) -> float:
    a_n = normalize_name(a)
    b_n = normalize_name(b)

    if not a_n or not b_n:
        return 0.0

    # Exact / strong prefix matching gets priority.
    if a_n == b_n:
        return 1.0

    if b_n.startswith(a_n) or a_n.startswith(b_n):
        shorter = min(len(a_n), len(b_n))
        longer = max(len(a_n), len(b_n))
        return 0.90 + (0.10 * shorter / max(longer, 1))

    # Compare token initials and general similarity.
    ratio = SequenceMatcher(None, a_n, b_n).ratio()

    a_tokens = a_n.split()
    b_tokens = b_n.split()

    token_score = 0.0
    if a_tokens and b_tokens:
        matches = 0
        for i, token in enumerate(a_tokens):
            if i >= len(b_tokens):
                break
            target = b_tokens[i]
            if target.startswith(token) or token.startswith(target):
                matches += 1
        token_score = matches / max(len(a_tokens), len(b_tokens))

    return max(ratio, token_score * 0.96)


def find_required_driver(conn, work_date: str, entered_name: str):
    entered = " ".join((entered_name or "").strip().split())
    target = normalize_name(entered)

    rows = conn.execute("""
        SELECT d.id, d.name
        FROM daily_required r
        JOIN drivers d ON d.id = r.driver_id
        WHERE r.work_date = ?
        ORDER BY d.name
    """, (work_date,)).fetchall()

    # Exact match first.
    for row in rows:
        if normalize_name(row["name"]) == target:
            return row, 1.0, False

    scored = []
    for row in rows:
        score = name_similarity(entered, row["name"])
        scored.append((score, row))

    scored.sort(key=lambda item: item[0], reverse=True)

    if not scored:
        return None, 0.0, False

    best_score, best_row = scored[0]
    second_score = scored[1][0] if len(scored) > 1 else 0.0

    # Accept short but clear prefixes such as "Elvis V".
    # Reject ambiguous matches when the top two are too close.
    if best_score >= 0.82:
        ambiguous = second_score >= 0.80 and (best_score - second_score) < 0.08
        if ambiguous:
            return None, best_score, True
        return best_row, best_score, False

    return None, best_score, False


def extract_fico_candidates_from_text(text_value: str) -> list[int]:
    if not text_value:
        return []

    candidates = []
    for match in re.findall(r"(?<!\d)(\d{3})(?!\d)", text_value):
        try:
            value = int(match)
        except ValueError:
            continue

        # FICO in this workflow is capped at 850.
        if 300 <= value <= 850:
            candidates.append(value)

    return candidates


def detect_fico_from_image_bytes(raw: bytes, content_type: str) -> tuple[int | None, str]:
    """
    Optional OCR integration.

    If environment variable OCRSPACE_API_KEY is configured on Render,
    the image is sent to OCR.Space and the returned text is searched for
    a plausible FICO score. Without an OCR key, the function returns
    'ocr_not_configured' and Admin will request manual verification.
    """
    api_key = os.getenv("OCRSPACE_API_KEY", "").strip()

    if not api_key:
        return None, "ocr_not_configured"

    boundary = "----FICOControlBoundary7MA4YWxkTrZu0gW"
    ext = "jpg"
    if content_type == "image/png":
        ext = "png"
    elif content_type == "image/webp":
        ext = "webp"

    parts = []

    def add_field(name, value):
        parts.append(
            f"--{boundary}\r\n"
            f'Content-Disposition: form-data; name="{name}"\r\n\r\n'
            f"{value}\r\n".encode("utf-8")
        )

    add_field("apikey", api_key)
    add_field("language", "eng")
    add_field("isOverlayRequired", "false")
    add_field("OCREngine", "2")

    file_header = (
        f"--{boundary}\r\n"
        f'Content-Disposition: form-data; name="file"; filename="fico.{ext}"\r\n'
        f"Content-Type: {content_type}\r\n\r\n"
    ).encode("utf-8")

    parts.append(file_header + raw + b"\r\n")
    parts.append(f"--{boundary}--\r\n".encode("utf-8"))
    body = b"".join(parts)

    req = urllib.request.Request(
        "https://api.ocr.space/parse/image",
        data=body,
        headers={"Content-Type": f"multipart/form-data; boundary={boundary}"},
        method="POST"
    )

    try:
        with urllib.request.urlopen(req, timeout=20) as response:
            payload = json.loads(response.read().decode("utf-8", errors="replace"))
    except Exception:
        return None, "ocr_error"

    parsed = payload.get("ParsedResults") or []
    combined = "\n".join(
        str(item.get("ParsedText") or "")
        for item in parsed
    )

    candidates = extract_fico_candidates_from_text(combined)

    if not candidates:
        return None, "ocr_unreadable"

    # Usually the FICO score is the highest relevant 3-digit number in
    # the screenshot. Prefer 850 when present, otherwise the highest.
    detected = max(candidates)
    return detected, "ocr_ok"


async def save_proof_image(proof: UploadFile) -> tuple[str, str, bytes]:
    if proof.content_type not in ALLOWED_IMAGE_TYPES:
        raise HTTPException(status_code=400, detail="invalid_image_type")

    raw = await proof.read()

    if not raw:
        raise HTTPException(status_code=400, detail="empty_image")

    if len(raw) > MAX_IMAGE_BYTES:
        raise HTTPException(status_code=400, detail="image_too_large")

    extension = {
        "image/jpeg": ".jpg",
        "image/png": ".png",
        "image/webp": ".webp"
    }[proof.content_type]

    filename = f"{uuid.uuid4().hex}{extension}"
    path = os.path.join(UPLOAD_DIR, filename)

    with open(path, "wb") as f:
        f.write(raw)

    return filename, (proof.filename or "proof"), raw


class SubmissionIn(BaseModel):
    driver_id: int
    fico_score: int


@app.get("/api/health")
def health():
    return {"ok": True, "service": "FICO Control"}


@app.get("/api/drivers/today")
def drivers_today():
    today = date.today().isoformat()
    conn = db()

    rows = conn.execute("""
        SELECT d.id, d.name
        FROM daily_required r
        JOIN drivers d ON d.id = r.driver_id
        WHERE r.work_date = ?
        ORDER BY d.name
    """, (today,)).fetchall()

    conn.close()

    return {
        "date": today,
        "drivers": [{"id": r["id"], "name": r["name"]} for r in rows]
    }


# Kept temporarily for compatibility with the existing mobile prototype.
@app.post("/api/submissions")
def submit_mobile_legacy(payload: SubmissionIn):
    if payload.fico_score < 0 or payload.fico_score > 850:
        raise HTTPException(status_code=400, detail="invalid_score")

    today = date.today().isoformat()
    conn = db()

    required = conn.execute(
        "SELECT 1 FROM daily_required WHERE work_date=? AND driver_id=?",
        (today, payload.driver_id)
    ).fetchone()

    if not required:
        conn.close()
        raise HTTPException(status_code=400, detail="not_required")

    existing = conn.execute(
        "SELECT 1 FROM submissions WHERE work_date=? AND driver_id=?",
        (today, payload.driver_id)
    ).fetchone()

    if existing:
        conn.close()
        raise HTTPException(status_code=409, detail="already_sent")

    conn.execute("""
        INSERT INTO submissions(
            work_date, driver_id, fico_score, submitted_at
        ) VALUES(?,?,?,?)
    """, (
        today,
        payload.driver_id,
        payload.fico_score,
        datetime.now().isoformat(timespec="seconds")
    ))

    conn.commit()
    conn.close()
    return {"ok": True}



@app.post("/api/submissions/photo")
async def submit_mobile_photo(
    full_name: str = Form(...),
    fico_score: int = Form(...),
    proof: UploadFile = File(...)
):
    if fico_score < 300 or fico_score > 850:
        raise HTTPException(status_code=400, detail="invalid_score")

    today = date.today().isoformat()
    conn = db()

    driver, match_score, ambiguous = find_required_driver(conn, today, full_name)

    if ambiguous:
        conn.close()
        raise HTTPException(status_code=400, detail="ambiguous_name")

    if not driver:
        conn.close()
        raise HTTPException(status_code=400, detail="driver_not_found_today")

    existing = conn.execute(
        "SELECT 1 FROM submissions WHERE work_date=? AND driver_id=?",
        (today, driver["id"])
    ).fetchone()

    if existing:
        conn.close()
        raise HTTPException(status_code=409, detail="already_sent")

    try:
        filename, original_name, raw = await save_proof_image(proof)
    except Exception:
        conn.close()
        raise

    detected_score, ocr_state = detect_fico_from_image_bytes(
        raw,
        proof.content_type or "image/jpeg"
    )

    if detected_score is None:
        verification_status = "manual_review"
    elif detected_score == fico_score:
        verification_status = "verified"
    else:
        verification_status = "mismatch"

    conn.execute("""
        INSERT INTO submissions(
            work_date,
            driver_id,
            fico_score,
            submitted_at,
            entered_full_name,
            proof_filename,
            proof_original_name,
            detected_fico_score,
            verification_status,
            name_match_score
        ) VALUES(?,?,?,?,?,?,?,?,?,?)
    """, (
        today,
        driver["id"],
        fico_score,
        datetime.now().isoformat(timespec="seconds"),
        " ".join(full_name.strip().split()),
        filename,
        original_name,
        detected_score,
        verification_status,
        match_score
    ))

    conn.commit()
    conn.close()

    return {
        "ok": True,
        "driver": driver["name"],
        "name_match_score": round(match_score, 3),
        "detected_fico_score": detected_score,
        "verification_status": verification_status,
        "ocr_state": ocr_state
    }


def extract_driver_names_from_xlsx(raw: bytes) -> list[str]:
    workbook = openpyxl.load_workbook(
        io.BytesIO(raw),
        read_only=True,
        data_only=True
    )

    ws = workbook["Strecken"] if "Strecken" in workbook.sheetnames else workbook[workbook.sheetnames[0]]
    rows = ws.iter_rows(values_only=True)
    headers = next(rows, None)

    if not headers:
        return []

    normalized = [
        str(h).strip().lower() if h is not None else ""
        for h in headers
    ]

    candidates = [
        "name des fahrers",
        "fahrername",
        "fahrer",
        "driver name",
        "driver",
        "name",
        "nume șofer",
        "nume sofer",
        "nume"
    ]

    name_index = None
    for candidate in candidates:
        if candidate in normalized:
            name_index = normalized.index(candidate)
            break

    if name_index is None:
        raise ValueError("driver_column_not_found")

    names, seen = [], set()

    for row in rows:
        if name_index >= len(row) or row[name_index] is None:
            continue

        raw_name = str(row[name_index]).strip()
        if not raw_name:
            continue

        # Cortex can show the assigned driver followed by rescue/helper drivers
        # in the same cell, usually separated by "|" or line breaks.
        # For FICO Control, ONLY the first driver belongs to that route.
        parts = re.split(r"[|\n\r]+", raw_name)
        name = next((part.strip() for part in parts if part.strip()), "")

        if not name:
            continue

        key = normalize_name(name)
        if key not in seen:
            seen.add(key)
            names.append(name)

    return names


@app.post("/admin/upload")
async def upload_daily_list(
    work_date: str = Form(...),
    file: UploadFile = File(...)
):
    filename = (file.filename or "").lower()

    if not filename.endswith(".xlsx"):
        return RedirectResponse(
            f"/admin?d={work_date}&error=unsupported",
            status_code=303
        )

    raw = await file.read()

    try:
        names = extract_driver_names_from_xlsx(raw)
    except Exception:
        return RedirectResponse(
            f"/admin?d={work_date}&error=import",
            status_code=303
        )

    if not names:
        return RedirectResponse(
            f"/admin?d={work_date}&error=empty",
            status_code=303
        )

    conn = db()
    conn.execute("DELETE FROM daily_required WHERE work_date=?", (work_date,))

    for name in names:
        conn.execute("INSERT OR IGNORE INTO drivers(name) VALUES(?)", (name,))
        driver = conn.execute("SELECT id FROM drivers WHERE name=?", (name,)).fetchone()
        conn.execute(
            "INSERT OR IGNORE INTO daily_required(work_date, driver_id) VALUES(?,?)",
            (work_date, driver["id"])
        )

    conn.commit()
    conn.close()

    return RedirectResponse(
        f"/admin?d={work_date}&imported={len(names)}",
        status_code=303
    )



@app.post("/submit")
async def submit_web(
    full_name: str = Form(...),
    fico_score: int = Form(...),
    proof: UploadFile = File(...)
):
    if fico_score < 300 or fico_score > 850:
        return RedirectResponse("/?error=invalid_score", status_code=303)

    today = date.today().isoformat()
    conn = db()

    driver, match_score, ambiguous = find_required_driver(conn, today, full_name)

    if ambiguous:
        conn.close()
        return RedirectResponse("/?error=ambiguous_name", status_code=303)

    if not driver:
        conn.close()
        return RedirectResponse("/?error=name_not_found", status_code=303)

    existing = conn.execute(
        "SELECT 1 FROM submissions WHERE work_date=? AND driver_id=?",
        (today, driver["id"])
    ).fetchone()

    if existing:
        conn.close()
        return RedirectResponse("/?error=already_sent", status_code=303)

    try:
        filename, original_name, raw = await save_proof_image(proof)
    except HTTPException as exc:
        conn.close()
        return RedirectResponse(f"/?error={exc.detail}", status_code=303)

    detected_score, ocr_state = detect_fico_from_image_bytes(
        raw,
        proof.content_type or "image/jpeg"
    )

    if detected_score is None:
        verification_status = "manual_review"
    elif detected_score == fico_score:
        verification_status = "verified"
    else:
        verification_status = "mismatch"

    conn.execute("""
        INSERT INTO submissions(
            work_date,
            driver_id,
            fico_score,
            submitted_at,
            entered_full_name,
            proof_filename,
            proof_original_name,
            detected_fico_score,
            verification_status,
            name_match_score
        ) VALUES(?,?,?,?,?,?,?,?,?,?)
    """, (
        today,
        driver["id"],
        fico_score,
        datetime.now().isoformat(timespec="seconds"),
        " ".join(full_name.strip().split()),
        filename,
        original_name,
        detected_score,
        verification_status,
        match_score
    ))

    conn.commit()
    conn.close()

    return RedirectResponse("/?success=1", status_code=303)


@app.get("/proof/{filename}")
def proof_image(filename: str):
    # Prevent path traversal.
    safe_name = os.path.basename(filename)
    path = os.path.join(UPLOAD_DIR, safe_name)

    if not os.path.isfile(path):
        raise HTTPException(status_code=404, detail="proof_not_found")

    return FileResponse(path)


@app.get("/", response_class=HTMLResponse)
def driver_page():
    today = date.today().isoformat()

    page = f"""
<!doctype html>
<html lang="ro">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>FICO Control</title>
<style>
*{{box-sizing:border-box}}
body{{margin:0;font-family:Arial,sans-serif;background:#f4f6f8;color:#17212b}}
.card{{width:min(92%,500px);margin:45px auto;background:#fff;padding:30px;border-radius:20px;box-shadow:0 12px 35px rgba(0,0,0,.08)}}
.topbar{{display:flex;align-items:center;justify-content:space-between;gap:16px}}
.brand{{font-size:13px;font-weight:800;letter-spacing:2px;white-space:nowrap}}
.languages{{display:flex;gap:5px;background:#f4f6f8;padding:4px;border-radius:10px}}
.lang-btn{{width:auto;margin:0;padding:7px 9px;border:0;border-radius:7px;background:transparent;color:#667085;font-size:12px;font-weight:800;cursor:pointer}}
.lang-btn.active{{background:#17212b;color:#fff}}
h1{{font-size:31px;margin:22px 0 8px}}
.muted{{color:#667085;line-height:1.45}}
label{{display:block;margin:20px 0 8px;font-weight:700}}
input,button{{width:100%;padding:14px;border-radius:11px;border:1px solid #d8dde3;font-size:16px}}
input[type=file]{{background:#fff}}
.proof-input{{position:absolute;width:1px;height:1px;opacity:0;pointer-events:none}}
.proof-actions{{display:grid;grid-template-columns:1fr 1fr;gap:10px}}
.proof-action{{width:100%;padding:14px 12px;border-radius:11px;border:1px solid #d8dde3;background:#fff;color:#17212b;font-size:15px;font-weight:800;cursor:pointer}}
.proof-action.primary{{background:#17212b;color:#fff;border-color:#17212b}}
.proof-selected{{display:none;margin-top:10px;padding:10px 12px;border-radius:9px;background:#eaf7ef;color:#146c43;font-size:13px;font-weight:700;word-break:break-word}}
.proof-selected.show{{display:block}}
.submit{{margin-top:22px;background:#17212b;color:#fff;border:0;font-weight:800;cursor:pointer}}
.hint{{font-size:13px;color:#667085;margin-top:7px}}
.notice{{display:none;margin:18px 0 0;padding:12px 14px;border-radius:10px;font-size:14px;font-weight:700;line-height:1.4}}
.notice.success{{display:block;background:#eaf7ef;color:#146c43}}
.notice.error{{display:block;background:#fdeeee;color:#b42318}}
@media(max-width:420px){{
  .card{{padding:23px}}
  .topbar{{align-items:flex-start}}
  .brand{{font-size:12px}}
  .lang-btn{{padding:7px}}
  .proof-actions{{grid-template-columns:1fr}}
}}
</style>
</head>
<body>
<main class="card">
<div class="topbar">
  <div class="brand">FICO CONTROL</div>
  <div class="languages" aria-label="Language">
    <button type="button" class="lang-btn" data-lang="ro">RO</button>
    <button type="button" class="lang-btn" data-lang="de">DE</button>
    <button type="button" class="lang-btn" data-lang="en">EN</button>
  </div>
</div>

<h1 data-i18n="title">Trimite scorul FICO</h1>
<p class="muted">
  {today}<br>
  <span data-i18n="intro">Încarcă o poză sau un screenshot clar în care scorul FICO este vizibil.</span>
</p>

<div id="notice" class="notice"></div>

<form action="/submit" method="post" enctype="multipart/form-data">

<label data-i18n="photoLabel">1. Poză / Screenshot FICO</label>

<input
  id="proof"
  class="proof-input"
  type="file"
  name="proof"
  accept="image/*"
  required
>

<div class="proof-actions">
  <button
    id="cameraButton"
    class="proof-action primary"
    type="button"
    data-i18n="takePhoto"
    onclick="chooseProofSource('camera')"
  >Fă o poză acum</button>

  <button
    id="galleryButton"
    class="proof-action"
    type="button"
    data-i18n="chooseGallery"
    onclick="chooseProofSource('gallery')"
  >Alege din galerie</button>
</div>

<div id="proofSelected" class="proof-selected"></div>
<div class="hint" data-i18n="photoHint">Fă o poză acum sau alege un screenshot din galerie · maximum 10 MB</div>

<label data-i18n="nameLabel">2. Numele complet</label>
<input id="fullName" type="text" name="full_name" autocomplete="name" placeholder="Prenume și nume complet" required>
<div class="hint" data-i18n="nameHint">Scrie numele exact așa cum apare în lista de lucru.</div>

<label data-i18n="scoreLabel">3. Scor FICO</label>
<input id="ficoScore" type="number" name="fico_score" min="300" max="850" inputmode="numeric" placeholder="ex. 850" required>

<button class="submit" type="submit" data-i18n="submit">Trimite scorul și dovada</button>
</form>
</main>

<script>
const translations = {{
  ro: {{
    title: "Trimite scorul FICO",
    intro: "Încarcă o poză sau un screenshot clar în care scorul FICO este vizibil.",
    photoLabel: "1. Poză / Screenshot FICO",
    takePhoto: "Fă o poză acum",
    chooseGallery: "Alege din galerie",
    photoSelected: "Poză selectată:",
    photoHint: "Fă o poză acum sau alege un screenshot din galerie · maximum 10 MB",
    nameLabel: "2. Numele complet",
    nameHint: "Scrie numele exact așa cum apare în lista de lucru.",
    namePlaceholder: "Prenume și nume complet",
    scoreLabel: "3. Scor FICO",
    scorePlaceholder: "ex. 850",
    submit: "Trimite scorul și dovada",
    success: "Scorul FICO și dovada au fost trimise cu succes.",
    already_sent: "Ai trimis deja scorul FICO pentru astăzi.",
    name_not_found: "Numele introdus nu apare în lista șoferilor programați astăzi.",
    ambiguous_name: "Numele este prea scurt sau seamănă cu mai mulți șoferi. Te rog să scrii mai mult din numele complet.",
    invalid_score: "Scorul FICO introdus nu este valid.",
    invalid_image_type: "Te rog să încarci o imagine JPG, PNG sau WEBP.",
    empty_image: "Imaginea selectată este goală. Te rog să alegi alt fișier.",
    image_too_large: "Imaginea este prea mare. Dimensiunea maximă este 10 MB.",
    generic_error: "Trimiterea nu a putut fi finalizată. Te rog să încerci din nou."
  }},
  de: {{
    title: "FICO-Score senden",
    intro: "Lade ein klares Foto oder einen Screenshot hoch, auf dem der FICO-Score sichtbar ist.",
    photoLabel: "1. FICO Foto / Screenshot",
    takePhoto: "Jetzt Foto aufnehmen",
    chooseGallery: "Aus Galerie wählen",
    photoSelected: "Ausgewähltes Bild:",
    photoHint: "Jetzt ein Foto aufnehmen oder einen Screenshot aus der Galerie wählen · maximal 10 MB",
    nameLabel: "2. Vollständiger Name",
    nameHint: "Schreibe deinen Namen genau so, wie er in der Fahrer-/Arbeitsliste steht.",
    namePlaceholder: "Vor- und Nachname",
    scoreLabel: "3. FICO-Score",
    scorePlaceholder: "z. B. 850",
    submit: "Score und Nachweis senden",
    success: "FICO-Score und Nachweis wurden erfolgreich gesendet.",
    already_sent: "Du hast deinen FICO-Score für heute bereits gesendet.",
    name_not_found: "Der eingegebene Name steht heute nicht auf der Liste der eingeplanten Fahrer.",
    ambiguous_name: "Der Name ist zu kurz oder passt zu mehreren Fahrern. Bitte gib mehr vom vollständigen Namen ein.",
    invalid_score: "Der eingegebene FICO-Score ist ungültig.",
    invalid_image_type: "Bitte lade ein JPG-, PNG- oder WEBP-Bild hoch.",
    empty_image: "Die ausgewählte Bilddatei ist leer. Bitte wähle eine andere Datei.",
    image_too_large: "Das Bild ist zu groß. Die maximale Dateigröße beträgt 10 MB.",
    generic_error: "Die Übermittlung konnte nicht abgeschlossen werden. Bitte versuche es erneut."
  }},
  en: {{
    title: "Submit FICO Score",
    intro: "Upload a clear photo or screenshot where the FICO score is visible.",
    photoLabel: "1. FICO Photo / Screenshot",
    takePhoto: "Take a photo now",
    chooseGallery: "Choose from gallery",
    photoSelected: "Selected image:",
    photoHint: "Take a photo now or choose a screenshot from the gallery · maximum 10 MB",
    nameLabel: "2. Full Name",
    nameHint: "Enter your name exactly as it appears on the driver/work list.",
    namePlaceholder: "First and last name",
    scoreLabel: "3. FICO Score",
    scorePlaceholder: "e.g. 850",
    submit: "Submit score and proof",
    success: "Your FICO score and proof were submitted successfully.",
    already_sent: "You have already submitted your FICO score for today.",
    name_not_found: "The entered name is not on today's scheduled driver list.",
    ambiguous_name: "The name is too short or matches multiple drivers. Please enter more of the full name.",
    invalid_score: "The FICO score you entered is invalid.",
    invalid_image_type: "Please upload a JPG, PNG, or WEBP image.",
    empty_image: "The selected image is empty. Please choose another file.",
    image_too_large: "The image is too large. The maximum file size is 10 MB.",
    generic_error: "The submission could not be completed. Please try again."
  }}
}};


function chooseProofSource(source) {{
  const input = document.getElementById("proof");

  // Reset the value so selecting/taking the same image again still fires "change".
  input.value = "";

  if (source === "camera") {{
    // Android/iPhone browsers and WebViews interpret this as rear camera capture.
    input.setAttribute("capture", "environment");
  }} else {{
    // Without capture, the normal photo/file picker is opened.
    input.removeAttribute("capture");
  }}

  input.click();
}}

function updateSelectedProof() {{
  const input = document.getElementById("proof");
  const box = document.getElementById("proofSelected");
  const lang = localStorage.getItem("ficoLanguage") || detectInitialLanguage();
  const t = translations[lang] || translations.ro;

  if (input.files && input.files.length > 0) {{
    const name = input.files[0].name || "FICO";
    box.textContent = `${{t.photoSelected}} ${{name}}`;
    box.classList.add("show");
  }} else {{
    box.textContent = "";
    box.classList.remove("show");
  }}
}}

document.getElementById("proof").addEventListener("change", updateSelectedProof);

function detectInitialLanguage() {{
  const saved = localStorage.getItem("ficoLanguage");
  if (saved && translations[saved]) return saved;

  const browser = (navigator.language || "").toLowerCase();
  if (browser.startsWith("de")) return "de";
  if (browser.startsWith("en")) return "en";
  return "ro";
}}

function applyLanguage(lang) {{
  if (!translations[lang]) lang = "ro";
  localStorage.setItem("ficoLanguage", lang);
  document.documentElement.lang = lang;

  const t = translations[lang];
  document.querySelectorAll("[data-i18n]").forEach(el => {{
    const key = el.dataset.i18n;
    if (t[key]) el.textContent = t[key];
  }});

  document.getElementById("fullName").placeholder = t.namePlaceholder;
  document.getElementById("ficoScore").placeholder = t.scorePlaceholder;

  document.querySelectorAll(".lang-btn").forEach(btn => {{
    btn.classList.toggle("active", btn.dataset.lang === lang);
  }});

  if (document.getElementById("proof").files.length > 0) {{
    updateSelectedProof();
  }}

  renderNotice(lang);
}}

function renderNotice(lang) {{
  const params = new URLSearchParams(window.location.search);
  const notice = document.getElementById("notice");
  const t = translations[lang];

  notice.className = "notice";
  notice.textContent = "";

  if (params.get("success") === "1") {{
    notice.textContent = t.success;
    notice.classList.add("success");
    return;
  }}

  const error = params.get("error");
  if (error) {{
    notice.textContent = t[error] || t.generic_error;
    notice.classList.add("error");
  }}
}}

document.querySelectorAll(".lang-btn").forEach(btn => {{
  btn.addEventListener("click", () => applyLanguage(btn.dataset.lang));
}});

applyLanguage(detectInitialLanguage());
</script>
</body>
</html>
"""
    return HTMLResponse(page)



@app.get("/admin", response_class=HTMLResponse)
def admin_page(d: str | None = None, q: str | None = None):
    selected = d or date.today().isoformat()
    search = (q or "").strip()

    conn = db()

    rows = conn.execute("""
        SELECT d.name,
               s.fico_score,
               s.submitted_at,
               s.entered_full_name,
               s.proof_filename,
               s.proof_original_name,
               s.detected_fico_score,
               s.verification_status,
               s.name_match_score,
               CASE WHEN s.id IS NULL THEN 0 ELSE 1 END AS sent
        FROM daily_required r
        JOIN drivers d ON d.id = r.driver_id
        LEFT JOIN submissions s
          ON s.driver_id = d.id
         AND s.work_date = r.work_date
        WHERE r.work_date = ?
        ORDER BY sent ASC, d.name ASC
    """, (selected,)).fetchall()

    conn.close()

    total = len(rows)
    sent = sum(int(r["sent"]) for r in rows)
    missing = total - sent
    low_fico = sum(
        1 for r in rows
        if r["fico_score"] is not None and int(r["fico_score"]) < 800
    )

    needs_review = sum(
        1 for r in rows
        if r["verification_status"] in ("mismatch", "manual_review")
    )

    missing_names = [r["name"] for r in rows if not r["sent"]]

    filtered_rows = rows
    if search:
        key = normalize_name(search)
        filtered_rows = [
            r for r in rows
            if key in normalize_name(r["name"])
            or key in normalize_name(r["entered_full_name"] or "")
        ]

    table_rows = ""

    for r in filtered_rows:
        sent_bool = bool(r["sent"])
        status = "Trimis" if sent_bool else "Nu a trimis"
        status_class = "sent" if sent_bool else "missing"

        fico = r["fico_score"] if r["fico_score"] is not None else "—"
        hour = r["submitted_at"][11:16] if r["submitted_at"] else "—"

        score_class = ""
        score_badge = ""

        if r["fico_score"] is not None:
            score = int(r["fico_score"])
            if score < 800:
                score_class = "score-low"
                score_badge = '<span class="score-note danger">Scăzut</span>'
            elif score < 850:
                score_class = "score-mid"
                score_badge = '<span class="score-note warning">Sub 850</span>'
            else:
                score_class = "score-good"
                score_badge = '<span class="score-note good">850</span>'

        if r["proof_filename"]:
            proof_url = f'/proof/{html.escape(r["proof_filename"])}'
            proof = f'<a class="proof-btn" href="{proof_url}" target="_blank">Vezi poza</a>'
        else:
            proof = '<span class="dash">—</span>'

        entered_name = html.escape(r["entered_full_name"]) if r["entered_full_name"] else "—"

        detected_fico = (
            r["detected_fico_score"]
            if r["detected_fico_score"] is not None
            else "—"
        )

        verification_status = r["verification_status"] or (
            "manual_review" if sent_bool else ""
        )

        if verification_status == "verified":
            verification = '<span class="verify-pill verified">✓ Verificat</span>'
        elif verification_status == "mismatch":
            verification = '<span class="verify-pill mismatch">⚠ Verifică</span>'
        elif verification_status == "manual_review":
            verification = '<span class="verify-pill manual">? Manual</span>'
        else:
            verification = '<span class="dash">—</span>'

        table_rows += f"""
        <tr class="{'row-missing' if not sent_bool else ''}">
            <td>
                <div class="driver-name">{html.escape(r["name"])}</div>
                <div class="entered-name">Introdus: {entered_name}</div>
            </td>
            <td><span class="status-pill {status_class}">{status}</span></td>
            <td class="{score_class}">
                <div class="score-wrap">
                    <strong>{fico}</strong>
                    {score_badge}
                </div>
            </td>
            <td><strong>{detected_fico}</strong></td>
            <td>{verification}</td>
            <td>{hour}</td>
            <td>{proof}</td>
        </tr>
        """

    missing_js = "\n".join(missing_names).replace("\\", "\\\\").replace("`", "\\`")

    if filtered_rows:
        table_content = (
            '<table><thead><tr>'
            '<th>Șofer</th><th>Status</th><th>FICO introdus</th><th>FICO detectat</th><th>Verificare</th><th>Ora</th><th>Dovadă</th>'
            '</tr></thead><tbody>'
            + table_rows +
            '</tbody></table>'
        )
    else:
        table_content = '<div class="empty">Nu există șoferi pentru această selecție.</div>'

    page = f"""
<!doctype html>
<html lang="ro">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>FICO Control Admin</title>
<style>
*{{box-sizing:border-box}}
body{{margin:0;font-family:Arial,Helvetica,sans-serif;background:#f4f6f8;color:#17212b}}
.admin{{width:min(96%,1280px);margin:30px auto 60px}}
.topbar{{display:flex;align-items:flex-start;justify-content:space-between;gap:20px;margin-bottom:28px}}
.brand{{font-size:13px;font-weight:800;letter-spacing:2px}}
h1{{font-size:38px;margin:14px 0 0}}
.top-actions{{display:flex;gap:10px;flex-wrap:wrap;align-items:center}}
.btn,.btn-light{{display:inline-flex;align-items:center;justify-content:center;text-decoration:none;padding:12px 15px;border-radius:10px;font-size:14px;font-weight:800;cursor:pointer}}
.btn{{border:0;background:#17212b;color:#fff}}
.btn-light{{border:1px solid #d8dde3;background:#fff;color:#17212b}}
.stats{{display:grid;grid-template-columns:repeat(5,1fr);gap:15px;margin-bottom:20px}}
.card,.panel{{background:#fff;border-radius:16px;box-shadow:0 5px 18px rgba(0,0,0,.05)}}
.card{{padding:20px}}
.card strong{{display:block;font-size:32px;line-height:1;margin-bottom:8px}}
.card span{{color:#667085}}
.card.alert strong{{color:#d13b2e}}
.panel{{padding:20px;margin-bottom:15px}}
.controls{{display:flex;justify-content:space-between;gap:18px;flex-wrap:wrap}}
.control-group{{display:flex;gap:10px;flex-wrap:wrap;align-items:center}}
input,button{{padding:12px 14px;border-radius:10px;border:1px solid #d8dde3;font-size:15px}}
button{{cursor:pointer}}
button.primary{{border:0;background:#17212b;color:#fff;font-weight:800}}
.search{{min-width:260px}}
.import-title{{font-size:13px;color:#667085;font-weight:700;margin-bottom:10px}}
.table-wrap{{overflow-x:auto}}
table{{width:100%;border-collapse:collapse;min-width:850px}}
th,td{{text-align:left;padding:14px 13px;border-bottom:1px solid #eceff2;vertical-align:middle}}
th{{font-size:13px;color:#667085;text-transform:uppercase;letter-spacing:.4px}}
.driver-name{{font-weight:800}}
.entered-name{{color:#667085;font-size:12px;margin-top:4px}}
.status-pill{{display:inline-block;border-radius:999px;padding:6px 9px;font-size:12px;font-weight:800}}
.status-pill.sent{{background:#e9f8ef;color:#14804a}}
.status-pill.missing{{background:#fdeeee;color:#c9362b}}
.row-missing{{background:#fffafa}}
.score-wrap{{display:flex;align-items:center;gap:8px}}
.score-wrap strong{{font-size:17px}}
.score-note{{font-size:11px;font-weight:800;border-radius:999px;padding:4px 7px}}
.score-note.danger{{background:#fde2e1;color:#b42318}}
.score-note.warning{{background:#fff4d6;color:#8a5a00}}
.score-note.good{{background:#e9f8ef;color:#14804a}}
.proof-btn{{display:inline-block;text-decoration:none;color:#17212b;border:1px solid #d8dde3;background:#fff;padding:8px 10px;border-radius:8px;font-size:13px;font-weight:800}}
.proof-btn:hover{{background:#f4f6f8}}
.verify-pill{{display:inline-block;border-radius:999px;padding:6px 9px;font-size:12px;font-weight:800}}
.verify-pill.verified{{background:#e9f8ef;color:#14804a}}
.verify-pill.mismatch{{background:#fde2e1;color:#b42318}}
.verify-pill.manual{{background:#fff4d6;color:#8a5a00}}
.dash{{color:#98a2b3}}
.helper{{font-size:12px;color:#667085;margin-top:10px}}
.copy-ok{{display:none;margin-left:4px;color:#14804a;font-size:13px;font-weight:800}}
.copy-ok.show{{display:inline}}
.empty{{padding:35px;text-align:center;color:#667085}}
@media(max-width:850px){{.stats{{grid-template-columns:repeat(2,1fr)}}.topbar{{display:block}}.top-actions{{margin-top:18px}}}}
@media(max-width:520px){{.stats{{grid-template-columns:1fr}}h1{{font-size:30px}}.search{{min-width:100%;width:100%}}.control-group{{width:100%}}}}
</style>
</head>
<body>

<main class="admin">

<div class="topbar">
    <div>
        <div class="brand">FICO CONTROL</div>
        <h1>Admin Dashboard</h1>
    </div>

    <div class="top-actions">
        <a class="btn-light" href="/admin/export.xlsx?d={selected}">Export Excel</a>
        <button class="btn" type="button" onclick="copyMissing()">Copiază șoferii lipsă</button>
        <span id="copyOk" class="copy-ok">Copiat</span>
    </div>
</div>

<section class="stats">
    <div class="card"><strong>{total}</strong><span>Programați</span></div>
    <div class="card"><strong>{sent}</strong><span>Au trimis</span></div>
    <div class="card alert"><strong>{missing}</strong><span>Lipsesc</span></div>
    <div class="card alert"><strong>{low_fico}</strong><span>FICO sub 800</span></div>
    <div class="card alert"><strong>{needs_review}</strong><span>Necesită verificare</span></div>
</section>

<section class="panel">
    <div class="controls">
        <form class="control-group" action="/admin" method="get">
            <input type="date" name="d" value="{selected}">
            <input class="search" type="text" name="q" value="{html.escape(search)}" placeholder="Caută șofer...">
            <button class="primary" type="submit">Afișează</button>
            <a class="btn-light" href="/admin?d={selected}">Reset</a>
        </form>
    </div>
</section>

<section class="panel">
    <div class="import-title">Lista zilnică din Cortex</div>
    <form class="control-group" action="/admin/upload" method="post" enctype="multipart/form-data">
        <input type="date" name="work_date" value="{selected}" required>
        <input type="file" name="file" accept=".xlsx" required>
        <button class="primary" type="submit">Încarcă Excel Cortex</button>
    </form>
    <div class="helper">Aplicația extrage automat șoferii din coloana „Name des Fahrers”.</div>
</section>

<section class="panel table-wrap">
    {table_content}
</section>

</main>

<script>
const missingNames = `{missing_js}`;

async function copyMissing() {{
    const message = missingNames.trim();

    if (!message) {{
        alert("Toți șoferii au trimis scorul FICO.");
        return;
    }}

    try {{
        await navigator.clipboard.writeText(message);
        const el = document.getElementById("copyOk");
        el.classList.add("show");
        setTimeout(() => el.classList.remove("show"), 1800);
    }} catch (e) {{
        window.prompt("Copiază lista:", message);
    }}
}}
</script>

</body>
</html>
"""

    return HTMLResponse(page)


@app.get("/admin/export.xlsx")
def export_day_excel(d: str | None = None):
    selected = d or date.today().isoformat()
    conn = db()

    rows = conn.execute("""
        SELECT d.name,
               CASE WHEN s.id IS NULL THEN 'Nu a trimis' ELSE 'Trimis' END AS status,
               s.fico_score,
               s.submitted_at,
               COALESCE(s.entered_full_name, '') AS entered_full_name,
               COALESCE(s.proof_filename, '') AS proof_filename,
               s.detected_fico_score,
               COALESCE(s.verification_status, '') AS verification_status,
               s.name_match_score
        FROM daily_required r
        JOIN drivers d ON d.id = r.driver_id
        LEFT JOIN submissions s
          ON s.driver_id = d.id
         AND s.work_date = r.work_date
        WHERE r.work_date = ?
        ORDER BY
            CASE WHEN s.id IS NULL THEN 0 ELSE 1 END ASC,
            d.name ASC
    """, (selected,)).fetchall()

    conn.close()

    workbook = openpyxl.Workbook()
    ws = workbook.active
    ws.title = "FICO"

    headers = [
        "Șofer",
        "Status",
        "FICO introdus",
        "FICO detectat",
        "Verificare",
        "Ora trimiterii",
        "Nume introdus",
        "Potrivire nume",
        "Dovadă"
    ]
    ws.append(headers)

    for row in rows:
        proof = f"/proof/{row['proof_filename']}" if row["proof_filename"] else ""
        ws.append([
            row["name"],
            row["status"],
            row["fico_score"] if row["fico_score"] is not None else "",
            row["detected_fico_score"] if row["detected_fico_score"] is not None else "",
            row["verification_status"],
            row["submitted_at"] or "",
            row["entered_full_name"],
            round(float(row["name_match_score"]), 3) if row["name_match_score"] is not None else "",
            proof
        ])

    from openpyxl.styles import Font, PatternFill, Alignment

    header_fill = PatternFill("solid", fgColor="17212B")
    header_font = Font(color="FFFFFF", bold=True)

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(vertical="center")

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    widths = {
        "A": 30,
        "B": 16,
        "C": 15,
        "D": 15,
        "E": 20,
        "F": 23,
        "G": 30,
        "H": 15,
        "I": 30
    }

    for col, width in widths.items():
        ws.column_dimensions[col].width = width

    missing_fill = PatternFill("solid", fgColor="FDEEEE")
    low_fill = PatternFill("solid", fgColor="F4CCCC")

    for row_idx in range(2, ws.max_row + 1):
        status = ws[f"B{row_idx}"].value
        score = ws[f"C{row_idx}"].value

        if status == "Nu a trimis":
            for col_idx in range(1, 10):
                ws.cell(row=row_idx, column=col_idx).fill = missing_fill

        if isinstance(score, int) and score < 800:
            ws[f"C{row_idx}"].fill = low_fill
            ws[f"C{row_idx}"].font = Font(bold=True)

    output = io.BytesIO()
    workbook.save(output)
    output.seek(0)

    filename = f"FICO_{selected}.xlsx"

    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition":
            f'attachment; filename="{filename}"'
        }
    )
