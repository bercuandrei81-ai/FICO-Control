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

    conn.commit()
    conn.close()


init_db()


def normalize_name(value: str) -> str:
    value = unicodedata.normalize("NFKC", value or "")
    value = " ".join(value.strip().split())
    return value.casefold()


def find_required_driver(conn, work_date: str, entered_name: str):
    target = normalize_name(entered_name)

    rows = conn.execute("""
        SELECT d.id, d.name
        FROM daily_required r
        JOIN drivers d ON d.id = r.driver_id
        WHERE r.work_date = ?
    """, (work_date,)).fetchall()

    for row in rows:
        if normalize_name(row["name"]) == target:
            return row

    return None


async def save_proof_image(proof: UploadFile) -> tuple[str, str]:
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

    return filename, (proof.filename or "proof")


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
    if payload.fico_score < 0 or payload.fico_score > 1000:
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
    if fico_score < 0 or fico_score > 1000:
        raise HTTPException(status_code=400, detail="invalid_score")

    today = date.today().isoformat()
    conn = db()

    driver = find_required_driver(conn, today, full_name)

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
        filename, original_name = await save_proof_image(proof)
    except Exception:
        conn.close()
        raise

    conn.execute("""
        INSERT INTO submissions(
            work_date,
            driver_id,
            fico_score,
            submitted_at,
            entered_full_name,
            proof_filename,
            proof_original_name
        ) VALUES(?,?,?,?,?,?,?)
    """, (
        today,
        driver["id"],
        fico_score,
        datetime.now().isoformat(timespec="seconds"),
        " ".join(full_name.strip().split()),
        filename,
        original_name
    ))

    conn.commit()
    conn.close()

    return {"ok": True, "driver": driver["name"]}


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

        name = str(row[name_index]).strip()
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
    if fico_score < 0 or fico_score > 1000:
        return RedirectResponse("/?error=invalid_score", status_code=303)

    today = date.today().isoformat()
    conn = db()
    driver = find_required_driver(conn, today, full_name)

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
        filename, original_name = await save_proof_image(proof)
    except HTTPException as exc:
        conn.close()
        return RedirectResponse(f"/?error={exc.detail}", status_code=303)

    conn.execute("""
        INSERT INTO submissions(
            work_date,
            driver_id,
            fico_score,
            submitted_at,
            entered_full_name,
            proof_filename,
            proof_original_name
        ) VALUES(?,?,?,?,?,?,?)
    """, (
        today,
        driver["id"],
        fico_score,
        datetime.now().isoformat(timespec="seconds"),
        " ".join(full_name.strip().split()),
        filename,
        original_name
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
<input id="proof" type="file" name="proof" accept="image/jpeg,image/png,image/webp" required>
<div class="hint" data-i18n="photoHint">JPG, PNG sau WEBP · maximum 10 MB</div>

<label data-i18n="nameLabel">2. Numele complet</label>
<input id="fullName" type="text" name="full_name" autocomplete="name" placeholder="Prenume și nume complet" required>
<div class="hint" data-i18n="nameHint">Scrie numele exact așa cum apare în lista de lucru.</div>

<label data-i18n="scoreLabel">3. Scor FICO</label>
<input id="ficoScore" type="number" name="fico_score" min="0" max="1000" inputmode="numeric" placeholder="ex. 850" required>

<button class="submit" type="submit" data-i18n="submit">Trimite scorul și dovada</button>
</form>
</main>

<script>
const translations = {{
  ro: {{
    title: "Trimite scorul FICO",
    intro: "Încarcă o poză sau un screenshot clar în care scorul FICO este vizibil.",
    photoLabel: "1. Poză / Screenshot FICO",
    photoHint: "JPG, PNG sau WEBP · maximum 10 MB",
    nameLabel: "2. Numele complet",
    nameHint: "Scrie numele exact așa cum apare în lista de lucru.",
    namePlaceholder: "Prenume și nume complet",
    scoreLabel: "3. Scor FICO",
    scorePlaceholder: "ex. 850",
    submit: "Trimite scorul și dovada",
    success: "Scorul FICO și dovada au fost trimise cu succes.",
    already_sent: "Ai trimis deja scorul FICO pentru astăzi.",
    name_not_found: "Numele introdus nu apare în lista șoferilor programați astăzi.",
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
    photoHint: "JPG, PNG oder WEBP · maximal 10 MB",
    nameLabel: "2. Vollständiger Name",
    nameHint: "Schreibe deinen Namen genau so, wie er in der Fahrer-/Arbeitsliste steht.",
    namePlaceholder: "Vor- und Nachname",
    scoreLabel: "3. FICO-Score",
    scorePlaceholder: "z. B. 850",
    submit: "Score und Nachweis senden",
    success: "FICO-Score und Nachweis wurden erfolgreich gesendet.",
    already_sent: "Du hast deinen FICO-Score für heute bereits gesendet.",
    name_not_found: "Der eingegebene Name steht heute nicht auf der Liste der eingeplanten Fahrer.",
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
    photoHint: "JPG, PNG or WEBP · maximum 10 MB",
    nameLabel: "2. Full Name",
    nameHint: "Enter your name exactly as it appears on the driver/work list.",
    namePlaceholder: "First and last name",
    scoreLabel: "3. FICO Score",
    scorePlaceholder: "e.g. 850",
    submit: "Submit score and proof",
    success: "Your FICO score and proof were submitted successfully.",
    already_sent: "You have already submitted your FICO score for today.",
    name_not_found: "The entered name is not on today's scheduled driver list.",
    invalid_score: "The FICO score you entered is invalid.",
    invalid_image_type: "Please upload a JPG, PNG, or WEBP image.",
    empty_image: "The selected image is empty. Please choose another file.",
    image_too_large: "The image is too large. The maximum file size is 10 MB.",
    generic_error: "The submission could not be completed. Please try again."
  }}
}};

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
def admin_page(d: str | None = None):
    selected = d or date.today().isoformat()
    conn = db()

    rows = conn.execute("""
        SELECT d.name,
               s.fico_score,
               s.submitted_at,
               s.entered_full_name,
               s.proof_filename,
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

    table_rows = ""

    for r in rows:
        status = "Trimis" if r["sent"] else "Nu a trimis"
        status_class = "sent" if r["sent"] else "missing"
        fico = r["fico_score"] if r["fico_score"] is not None else "—"
        hour = r["submitted_at"][11:16] if r["submitted_at"] else "—"

        if r["proof_filename"]:
            proof = f'<a class="proof" href="/proof/{html.escape(r["proof_filename"])}" target="_blank">Vezi poza</a>'
        else:
            proof = "—"

        table_rows += f"""
        <tr>
            <td>{html.escape(r["name"])}</td>
            <td class="{status_class}">{status}</td>
            <td>{fico}</td>
            <td>{hour}</td>
            <td>{proof}</td>
        </tr>
        """

    page = f"""
<!doctype html>
<html lang="ro">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>FICO Control Admin</title>
<style>
*{{box-sizing:border-box}}
body{{margin:0;font-family:Arial,sans-serif;background:#f4f6f8;color:#17212b}}
.admin{{width:min(95%,1200px);margin:35px auto}}
.brand{{font-size:13px;font-weight:800;letter-spacing:2px}}
h1{{font-size:36px;margin:20px 0 35px}}
.stats{{display:grid;grid-template-columns:repeat(3,1fr);gap:15px;margin-bottom:25px}}
.card,.panel{{background:#fff;border-radius:16px;padding:20px;box-shadow:0 5px 18px rgba(0,0,0,.05)}}
.card strong{{display:block;font-size:32px}}
.card span{{color:#667085}}
.panel{{margin-bottom:15px}}
form{{display:flex;gap:10px;flex-wrap:wrap;align-items:center}}
input,button{{padding:12px 14px;border-radius:10px;border:1px solid #d8dde3;font-size:15px}}
button{{background:#17212b;color:#fff;border:0;font-weight:800;cursor:pointer}}
table{{width:100%;border-collapse:collapse}}
th,td{{text-align:left;padding:13px;border-bottom:1px solid #eceff2}}
.sent{{color:#14804a;font-weight:800}}
.missing{{color:#d13b2e;font-weight:800}}
.proof{{font-weight:700;color:#17212b}}
@media(max-width:700px){{.stats{{grid-template-columns:1fr}}.panel{{overflow-x:auto}}}}
</style>
</head>
<body>
<main class="admin">
<div class="brand">FICO CONTROL</div>
<h1>Admin Dashboard</h1>

<section class="stats">
<div class="card"><strong>{total}</strong><span>Programați</span></div>
<div class="card"><strong>{sent}</strong><span>Au trimis</span></div>
<div class="card"><strong>{missing}</strong><span>Lipsesc</span></div>
</section>

<section class="panel">
<form action="/admin" method="get">
<input type="date" name="d" value="{selected}">
<button type="submit">Vezi ziua</button>
</form>
<br>
<form action="/admin/upload" method="post" enctype="multipart/form-data">
<input type="date" name="work_date" value="{selected}" required>
<input type="file" name="file" accept=".xlsx" required>
<button type="submit">Încarcă Excel Cortex</button>
</form>
</section>

<section class="panel">
<table>
<thead>
<tr>
<th>Șofer</th>
<th>Status</th>
<th>FICO</th>
<th>Ora</th>
<th>Dovadă</th>
</tr>
</thead>
<tbody>
{table_rows}
</tbody>
</table>
</section>
</main>
</body>
</html>
"""
    return HTMLResponse(page)


@app.get("/admin/export")
def export_day(d: str | None = None):
    selected = d or date.today().isoformat()
    conn = db()

    rows = conn.execute("""
        SELECT d.name,
               CASE WHEN s.id IS NULL THEN 'Nu a trimis' ELSE 'Trimis' END AS status,
               COALESCE(s.fico_score, '') AS fico_score,
               COALESCE(s.submitted_at, '') AS submitted_at,
               COALESCE(s.proof_filename, '') AS proof_filename
        FROM daily_required r
        JOIN drivers d ON d.id = r.driver_id
        LEFT JOIN submissions s
          ON s.driver_id = d.id
         AND s.work_date = r.work_date
        WHERE r.work_date = ?
        ORDER BY status, d.name
    """, (selected,)).fetchall()

    conn.close()

    out = io.StringIO()
    writer = csv.writer(out)
    writer.writerow(["Driver", "Status", "FICO", "Submitted at", "Proof"])

    for r in rows:
        writer.writerow([
            r["name"],
            r["status"],
            r["fico_score"],
            r["submitted_at"],
            r["proof_filename"]
        ])

    out.seek(0)

    return StreamingResponse(
        iter([out.getvalue()]),
        media_type="text/csv",
        headers={
            "Content-Disposition":
            f'attachment; filename="fico_{selected}.csv"'
        }
    )
