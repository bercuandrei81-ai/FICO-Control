from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import HTMLResponse, RedirectResponse, StreamingResponse
from pydantic import BaseModel
from datetime import date, datetime
import sqlite3
import io
import csv
import html
import openpyxl

DB = "fico.db"

app = FastAPI(title="FICO Control")


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


@app.post("/api/submissions")
def submit_mobile(payload: SubmissionIn):
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


def extract_driver_names_from_xlsx(raw: bytes) -> list[str]:
    workbook = openpyxl.load_workbook(
        io.BytesIO(raw),
        read_only=True,
        data_only=True
    )

    if "Strecken" in workbook.sheetnames:
        ws = workbook["Strecken"]
    else:
        ws = workbook[workbook.sheetnames[0]]

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

    names = []
    seen = set()

    for row in rows:
        if name_index >= len(row):
            continue

        value = row[name_index]

        if value is None:
            continue

        name = str(value).strip()

        if not name:
            continue

        key = name.casefold()

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

    conn.execute(
        "DELETE FROM daily_required WHERE work_date=?",
        (work_date,)
    )

    for name in names:
        conn.execute(
            "INSERT OR IGNORE INTO drivers(name) VALUES(?)",
            (name,)
        )

        driver = conn.execute(
            "SELECT id FROM drivers WHERE name=?",
            (name,)
        ).fetchone()

        conn.execute(
            """
            INSERT OR IGNORE INTO daily_required(work_date, driver_id)
            VALUES(?,?)
            """,
            (work_date, driver["id"])
        )

    conn.commit()
    conn.close()

    return RedirectResponse(
        f"/admin?d={work_date}&imported={len(names)}",
        status_code=303
    )


@app.post("/submit")
def submit_web(
    driver_id: int = Form(...),
    fico_score: int = Form(...)
):
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


@app.get("/", response_class=HTMLResponse)
def driver_page():
    today = date.today().isoformat()
    conn = db()

    drivers = conn.execute("""
        SELECT d.id, d.name
        FROM daily_required r
        JOIN drivers d ON d.id = r.driver_id
        WHERE r.work_date = ?
        ORDER BY d.name
    """, (today,)).fetchall()

    conn.close()

    options = '<option value="">Selectează numele</option>'

    for d in drivers:
        options += (
            f'<option value="{d["id"]}">'
            f'{html.escape(d["name"])}</option>'
        )

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
.card{{width:min(92%,470px);margin:55px auto;background:#fff;padding:30px;border-radius:20px;box-shadow:0 12px 35px rgba(0,0,0,.08)}}
.brand{{font-size:13px;font-weight:800;letter-spacing:2px}}
h1{{font-size:32px;margin:22px 0 8px}}
.muted{{color:#667085}}
label{{display:block;margin:20px 0 8px;font-weight:700}}
input,select,button{{width:100%;padding:14px;border-radius:11px;border:1px solid #d8dde3;font-size:16px}}
button{{margin-top:20px;background:#17212b;color:#fff;border:0;font-weight:800;cursor:pointer}}
</style>
</head>
<body>
<main class="card">
<div class="brand">FICO CONTROL</div>
<h1>Trimite scorul FICO</h1>
<p class="muted">{today}</p>
<form action="/submit" method="post">
<label>Numele șoferului</label>
<select name="driver_id" required>{options}</select>
<label>Scor FICO</label>
<input type="number" name="fico_score" min="0" max="1000" placeholder="ex. 850" required>
<button type="submit">Trimite scorul</button>
</form>
</main>
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

        table_rows += f"""
        <tr>
            <td>{html.escape(r["name"])}</td>
            <td class="{status_class}">{status}</td>
            <td>{fico}</td>
            <td>{hour}</td>
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
.admin{{width:min(94%,1150px);margin:35px auto}}
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
               COALESCE(s.submitted_at, '') AS submitted_at
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
    writer.writerow(["Driver", "Status", "FICO", "Submitted at"])

    for r in rows:
        writer.writerow([
            r["name"],
            r["status"],
            r["fico_score"],
            r["submitted_at"]
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
