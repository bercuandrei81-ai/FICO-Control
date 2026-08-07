from fastapi import FastAPI, Request, Form, UploadFile, File, HTTPException
from fastapi.responses import HTMLResponse, RedirectResponse, StreamingResponse
from fastapi.templating import Jinja2Templates
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel
from datetime import date, datetime
import sqlite3, csv, io
import openpyxl

DB = "fico.db"
app = FastAPI(title="FICO Control API")

templates = Jinja2Templates(directory="templates")
app.mount("/static", StaticFiles(directory="static"), name="static")


def db():
    conn = sqlite3.connect(DB)
    conn.row_factory = sqlite3.Row
    return conn


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
    conn.commit()
    conn.close()


init_db()


class SubmissionIn(BaseModel):
    driver_id: int
    fico_score: int


@app.get("/api/health")
def api_health():
    return {"ok": True}


@app.get("/api/drivers/today")
def api_today_drivers():
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


@app.post("/api/submissions")
def api_submit_fico(payload: SubmissionIn):
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
        INSERT INTO submissions(work_date, driver_id, fico_score, submitted_at)
        VALUES(?,?,?,?)
    """, (
        today,
        payload.driver_id,
        payload.fico_score,
        datetime.now().isoformat(timespec="seconds")
    ))
    conn.commit()
    conn.close()

    return {"ok": True}


@app.get("/", response_class=HTMLResponse)
def driver_form(request: Request):
    conn = db()
    today = date.today().isoformat()
    drivers = conn.execute("""
        SELECT d.id, d.name
        FROM daily_required r
        JOIN drivers d ON d.id=r.driver_id
        WHERE r.work_date=?
        ORDER BY d.name
    """, (today,)).fetchall()
    conn.close()

    return templates.TemplateResponse("driver.html", {
        "request": request,
        "drivers": drivers,
        "today": today
    })


@app.post("/submit")
def submit_fico(driver_id: int = Form(...), fico_score: int = Form(...)):
    today = date.today().isoformat()

    if fico_score < 0 or fico_score > 1000:
        return RedirectResponse("/?error=invalid_score", status_code=303)

    conn = db()

    required = conn.execute(
        "SELECT 1 FROM daily_required WHERE work_date=? AND driver_id=?",
        (today, driver_id)
    ).fetchone()

    if not required:
        conn.close()
        return RedirectResponse("/?error=not_required", status_code=303)

    existing = conn.execute(
        "SELECT 1 FROM submissions WHERE work_date=? AND driver_id=?",
        (today, driver_id)
    ).fetchone()

    if existing:
        conn.close()
        return RedirectResponse("/?error=already_sent", status_code=303)

    conn.execute("""
        INSERT INTO submissions(work_date, driver_id, fico_score, submitted_at)
        VALUES(?,?,?,?)
    """, (
        today,
        driver_id,
        fico_score,
        datetime.now().isoformat(timespec="seconds")
    ))
    conn.commit()
    conn.close()

    return RedirectResponse("/?success=1", status_code=303)


@app.get("/admin", response_class=HTMLResponse)
def admin_dashboard(request: Request, d: str | None = None):
    selected = d or date.today().isoformat()
    conn = db()

    rows = conn.execute("""
        SELECT d.name,
               s.fico_score,
               s.submitted_at,
               CASE WHEN s.id IS NULL THEN 0 ELSE 1 END AS sent
        FROM daily_required r
        JOIN drivers d ON d.id=r.driver_id
        LEFT JOIN submissions s
          ON s.driver_id=d.id AND s.work_date=r.work_date
        WHERE r.work_date=?
        ORDER BY sent ASC, d.name ASC
    """, (selected,)).fetchall()

    total = len(rows)
    sent = sum(r["sent"] for r in rows)
    missing = total - sent
    conn.close()

    return templates.TemplateResponse("admin.html", {
        "request": request,
        "rows": rows,
        "selected": selected,
        "total": total,
        "sent": sent,
        "missing": missing
    })


def extract_driver_names_from_xlsx(raw: bytes) -> list[str]:
    workbook = openpyxl.load_workbook(io.BytesIO(raw), read_only=True, data_only=True)
    ws = workbook["Strecken"] if "Strecken" in workbook.sheetnames else workbook[workbook.sheetnames[0]]

    rows = ws.iter_rows(values_only=True)
    headers = next(rows, None)
    if not headers:
        return []

    normalized = [str(h).strip().lower() if h is not None else "" for h in headers]
    candidates = [
        "name des fahrers", "fahrername", "fahrer",
        "driver name", "driver", "name",
        "nume șofer", "nume sofer", "nume"
    ]

    idx = None
    for candidate in candidates:
        if candidate in normalized:
            idx = normalized.index(candidate)
            break

    if idx is None:
        raise ValueError("driver_column_not_found")

    result, seen = [], set()
    for row in rows:
        if idx >= len(row) or row[idx] is None:
            continue
        name = str(row[idx]).strip()
        if not name:
            continue
        key = name.casefold()
        if key not in seen:
            seen.add(key)
            result.append(name)

    return result


@app.post("/admin/upload")
async def upload_daily_list(work_date: str = Form(...), file: UploadFile = File(...)):
    raw = await file.read()
    filename = (file.filename or "").lower()

    try:
        if filename.endswith(".xlsx"):
            names = extract_driver_names_from_xlsx(raw)
        else:
            return RedirectResponse(f"/admin?d={work_date}&error=unsupported", status_code=303)
    except Exception:
        return RedirectResponse(f"/admin?d={work_date}&error=import", status_code=303)

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

    return RedirectResponse(f"/admin?d={work_date}&imported={len(names)}", status_code=303)
