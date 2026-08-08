from fastapi import FastAPI, UploadFile, File, Form, HTTPException, Request
from fastapi.responses import HTMLResponse, RedirectResponse, StreamingResponse, FileResponse
from pydantic import BaseModel
from datetime import date, datetime, timedelta, timezone
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
import base64
import secrets
import hmac
import hashlib

DB = "fico.db"
UPLOAD_DIR = "uploads"
MAX_IMAGE_BYTES = 10 * 1024 * 1024
ALLOWED_IMAGE_TYPES = {"image/jpeg", "image/png", "image/webp"}

ADMIN_COOKIE_NAME = "fico_admin_session"
ADMIN_SESSION_DAYS = 7
OWNER_VERIFY_MINUTES = 30
OWNER_CODE_MINUTES = 10

ADMIN_INITIAL_PASSWORD = os.getenv("ADMIN_INITIAL_PASSWORD", "").strip()
OWNER_EMAIL = os.getenv("OWNER_EMAIL", "").strip()
RESEND_API_KEY = os.getenv("RESEND_API_KEY", "").strip()
RESEND_FROM_EMAIL = os.getenv(
    "RESEND_FROM_EMAIL",
    "FICO Control <onboarding@resend.dev>"
).strip()

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

    CREATE TABLE IF NOT EXISTS unresolved_submissions(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        work_date TEXT NOT NULL,
        entered_full_name TEXT NOT NULL,
        fico_score INTEGER NOT NULL,
        submitted_at TEXT NOT NULL,
        proof_filename TEXT,
        proof_original_name TEXT,
        detected_fico_score INTEGER,
        verification_status TEXT,
        best_match_name TEXT,
        best_match_score REAL,
        match_reason TEXT
    );

    CREATE TABLE IF NOT EXISTS admin_settings(
        setting_key TEXT PRIMARY KEY,
        setting_value TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS admin_sessions(
        session_token TEXT PRIMARY KEY,
        display_name TEXT NOT NULL,
        normalized_name TEXT NOT NULL,
        created_at TEXT NOT NULL,
        last_seen TEXT NOT NULL,
        expires_at TEXT NOT NULL,
        revoked INTEGER DEFAULT 0,
        owner_verified_until TEXT
    );

    CREATE TABLE IF NOT EXISTS blocked_admin_names(
        normalized_name TEXT PRIMARY KEY,
        display_name TEXT NOT NULL,
        blocked_at TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS owner_email_codes(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        code_hash TEXT NOT NULL,
        created_at TEXT NOT NULL,
        expires_at TEXT NOT NULL,
        used INTEGER DEFAULT 0
    );
    """)

    ensure_column(conn, "submissions", "entered_full_name", "TEXT")
    ensure_column(conn, "submissions", "proof_filename", "TEXT")
    ensure_column(conn, "submissions", "proof_original_name", "TEXT")
    ensure_column(conn, "submissions", "detected_fico_score", "INTEGER")
    ensure_column(conn, "submissions", "verification_status", "TEXT")
    ensure_column(conn, "submissions", "name_match_score", "REAL")

    existing_password = conn.execute(
        "SELECT setting_value FROM admin_settings WHERE setting_key='shared_password_hash'"
    ).fetchone()

    if not existing_password and ADMIN_INITIAL_PASSWORD:
        salt = secrets.token_bytes(16)
        digest = hashlib.pbkdf2_hmac(
            "sha256",
            ADMIN_INITIAL_PASSWORD.encode("utf-8"),
            salt,
            250000
        )
        encoded = "pbkdf2_sha256$250000$" + base64.b64encode(salt).decode("ascii") + "$" + base64.b64encode(digest).decode("ascii")
        conn.execute(
            "INSERT INTO admin_settings(setting_key, setting_value) VALUES(?,?)",
            ("shared_password_hash", encoded)
        )

    conn.commit()
    conn.close()


init_db()


def utc_now():
    return datetime.now(timezone.utc)


def iso_utc(dt):
    return dt.astimezone(timezone.utc).isoformat(timespec="seconds")


def parse_iso_utc(value):
    if not value:
        return None
    try:
        parsed = datetime.fromisoformat(value)
        if parsed.tzinfo is None:
            parsed = parsed.replace(tzinfo=timezone.utc)
        return parsed.astimezone(timezone.utc)
    except Exception:
        return None


def normalize_admin_name(value: str) -> str:
    value = unicodedata.normalize("NFKD", value or "")
    value = "".join(ch for ch in value if not unicodedata.combining(ch))
    value = " ".join(value.strip().split())
    return value.casefold()


def hash_password(password: str) -> str:
    salt = secrets.token_bytes(16)
    iterations = 250000
    digest = hashlib.pbkdf2_hmac(
        "sha256",
        password.encode("utf-8"),
        salt,
        iterations
    )
    return (
        f"pbkdf2_sha256${iterations}$"
        f"{base64.b64encode(salt).decode('ascii')}$"
        f"{base64.b64encode(digest).decode('ascii')}"
    )


def verify_password(password: str, encoded: str) -> bool:
    try:
        algorithm, iterations_s, salt_b64, digest_b64 = encoded.split("$", 3)
        if algorithm != "pbkdf2_sha256":
            return False

        iterations = int(iterations_s)
        salt = base64.b64decode(salt_b64)
        expected = base64.b64decode(digest_b64)

        actual = hashlib.pbkdf2_hmac(
            "sha256",
            password.encode("utf-8"),
            salt,
            iterations
        )
        return hmac.compare_digest(actual, expected)
    except Exception:
        return False


def get_shared_password_hash(conn):
    row = conn.execute(
        "SELECT setting_value FROM admin_settings WHERE setting_key='shared_password_hash'"
    ).fetchone()
    return row["setting_value"] if row else None


def create_admin_session(conn, display_name: str) -> str:
    token = secrets.token_urlsafe(36)
    now = utc_now()
    expires = now + timedelta(days=ADMIN_SESSION_DAYS)

    conn.execute("""
        INSERT INTO admin_sessions(
            session_token,
            display_name,
            normalized_name,
            created_at,
            last_seen,
            expires_at,
            revoked
        ) VALUES(?,?,?,?,?,?,0)
    """, (
        token,
        display_name,
        normalize_admin_name(display_name),
        iso_utc(now),
        iso_utc(now),
        iso_utc(expires)
    ))
    conn.commit()
    return token


def get_valid_admin_session(token: str | None):
    if not token:
        return None

    conn = db()
    row = conn.execute("""
        SELECT session_token,
               display_name,
               normalized_name,
               created_at,
               last_seen,
               expires_at,
               revoked,
               owner_verified_until
        FROM admin_sessions
        WHERE session_token=?
    """, (token,)).fetchone()

    if not row:
        conn.close()
        return None

    blocked = conn.execute(
        "SELECT 1 FROM blocked_admin_names WHERE normalized_name=?",
        (row["normalized_name"],)
    ).fetchone()

    expires = parse_iso_utc(row["expires_at"])
    invalid = (
        bool(row["revoked"])
        or bool(blocked)
        or not expires
        or expires <= utc_now()
    )

    if invalid:
        conn.execute(
            "UPDATE admin_sessions SET revoked=1 WHERE session_token=?",
            (token,)
        )
        conn.commit()
        conn.close()
        return None

    conn.execute(
        "UPDATE admin_sessions SET last_seen=? WHERE session_token=?",
        (iso_utc(utc_now()), token)
    )
    conn.commit()
    conn.close()
    return dict(row)


def is_owner_verified(session):
    if not session:
        return False
    until = parse_iso_utc(session.get("owner_verified_until"))
    return bool(until and until > utc_now())


def mask_email(email: str) -> str:
    if not email or "@" not in email:
        return "neconfigurat"
    local, domain = email.split("@", 1)
    if len(local) <= 2:
        masked_local = local[:1] + "*"
    else:
        masked_local = local[:2] + "*" * max(2, len(local) - 2)
    return masked_local + "@" + domain


def hash_owner_code(code: str) -> str:
    return hashlib.sha256(code.encode("utf-8")).hexdigest()


def send_owner_verification_code(code: str):
    if not OWNER_EMAIL:
        raise RuntimeError("OWNER_EMAIL is not configured")
    if not RESEND_API_KEY:
        raise RuntimeError("RESEND_API_KEY is not configured")

    payload = json.dumps({
        "from": RESEND_FROM_EMAIL,
        "to": [OWNER_EMAIL],
        "subject": "FICO Control - Owner verification code",
        "html": (
            "<div style='font-family:Arial,sans-serif'>"
            "<h2>FICO Control</h2>"
            "<p>Codul pentru accesul Owner este:</p>"
            f"<div style='font-size:32px;font-weight:800;letter-spacing:5px'>{html.escape(code)}</div>"
            f"<p>Codul expiră în {OWNER_CODE_MINUTES} minute.</p>"
            "<p>Dacă nu ai cerut acest cod, îl poți ignora.</p>"
            "</div>"
        )
    }).encode("utf-8")

    req = urllib.request.Request(
        "https://api.resend.com/emails",
        data=payload,
        headers={
            "Authorization": f"Bearer {RESEND_API_KEY}",
            "Content-Type": "application/json"
        },
        method="POST"
    )

    with urllib.request.urlopen(req, timeout=20) as response:
        response.read()


def owner_gate_html(message: str = "", error: bool = False):
    message_html = ""
    if message:
        cls = "error" if error else "ok"
        message_html = f'<div class="{cls}">{html.escape(message)}</div>'

    return f"""
<!doctype html>
<html lang="ro">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>FICO Control Owner</title>
<style>
*{{box-sizing:border-box}}
body{{margin:0;background:#f4f6f8;font-family:Arial,sans-serif;color:#17212b}}
.wrap{{width:min(92%,520px);margin:55px auto}}
.card{{background:#fff;border-radius:20px;padding:28px;box-shadow:0 10px 30px rgba(0,0,0,.08)}}
.brand{{font-size:13px;font-weight:900;letter-spacing:2px}}
h1{{margin:18px 0 8px;font-size:30px}}
p{{color:#667085;line-height:1.5}}
input,button{{width:100%;padding:14px;border-radius:11px;font-size:16px;margin-top:10px}}
input{{border:1px solid #d8dde3}}
button{{border:0;background:#17212b;color:#fff;font-weight:800;cursor:pointer}}
.secondary{{background:#fff;color:#17212b;border:1px solid #d8dde3}}
.ok,.error{{padding:12px;border-radius:10px;margin:15px 0;font-weight:700}}
.ok{{background:#e9f8ef;color:#14804a}}
.error{{background:#fdeeee;color:#b42318}}
</style>
</head>
<body>
<div class="wrap">
<div class="card">
<div class="brand">FICO CONTROL · OWNER</div>
<h1>Confirmare prin email</h1>
<p>Controlul Owner este protejat separat. Codul este trimis numai la <strong>{html.escape(mask_email(OWNER_EMAIL))}</strong>.</p>
{message_html}
<form method="post" action="/admin/owner/send-code">
<button type="submit">Trimite codul pe email</button>
</form>
<form method="post" action="/admin/owner/verify-code">
<input name="code" inputmode="numeric" maxlength="6" placeholder="Cod din 6 cifre" required>
<button type="submit">Confirmă codul</button>
</form>
<a href="/admin" style="display:block;text-align:center;margin-top:16px;color:#17212b;font-weight:700">Înapoi la Admin</a>
</div>
</div>
</body>
</html>
"""


@app.middleware("http")
async def protect_admin_routes(request: Request, call_next):
    path = request.url.path

    public_admin_paths = {
        "/admin/login",
        "/admin/login/",
        "/admin/forgot-password",
        "/admin/forgot-password/",
        "/admin/forgot-password/send-code",
        "/admin/reset-password",
        "/admin/reset-password/",
        "/admin/setup-name",
        "/admin/setup-name/"
    }

    if path in public_admin_paths:
        return await call_next(request)

    protected = path.startswith("/admin") or path.startswith("/proof/")

    if not protected:
        return await call_next(request)

    token = request.cookies.get(ADMIN_COOKIE_NAME)
    session = get_valid_admin_session(token)

    if not session:
        next_path = path
        if request.url.query:
            next_path += "?" + request.url.query
        return RedirectResponse(
            "/admin/login?next=" + urllib.parse.quote(next_path, safe=""),
            status_code=303
        )

    request.state.admin_session = session
    return await call_next(request)


@app.get("/admin/login", response_class=HTMLResponse)
def admin_login_page(request: Request, next: str | None = None, error: str | None = None):
    current = get_valid_admin_session(request.cookies.get(ADMIN_COOKIE_NAME))
    if current:
        return RedirectResponse(next or "/admin", status_code=303)

    conn = db()
    configured = bool(get_shared_password_hash(conn))
    conn.close()

    error_text = ""
    if error == "invalid":
        error_text = "Numele sau parola nu sunt corecte."
    elif error == "blocked":
        error_text = "Accesul pentru acest nume este blocat."
    elif error == "name":
        error_text = "Introdu numele tău."
    elif not configured:
        error_text = "Parola Admin nu este încă configurată în Render."

    next_value = html.escape(next or "/admin", quote=True)
    error_html = f'<div class="error">{html.escape(error_text)}</div>' if error_text else ""

    page = f"""
<!doctype html>
<html lang="ro">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>FICO Control Admin Login</title>
<style>
*{{box-sizing:border-box}}
body{{margin:0;background:#f4f6f8;font-family:Arial,sans-serif;color:#17212b}}
.wrap{{width:min(92%,470px);margin:65px auto}}
.card{{background:#fff;border-radius:22px;padding:30px;box-shadow:0 12px 35px rgba(0,0,0,.08)}}
.brand{{font-size:13px;font-weight:900;letter-spacing:2px}}
h1{{font-size:31px;margin:22px 0 8px}}
.subtitle{{color:#667085;line-height:1.5;margin-bottom:22px}}
label{{display:block;font-weight:800;margin:16px 0 7px}}
input,button{{width:100%;padding:14px;border-radius:11px;font-size:16px}}
input{{border:1px solid #d8dde3}}
.password-wrap{{position:relative}}
.password-wrap input{{padding-right:54px}}
.eye-btn{{position:absolute;right:8px;top:50%;transform:translateY(-50%);width:42px;height:42px;margin:0;padding:0;border:0;background:transparent;color:#667085;font-size:20px;cursor:pointer;display:flex;align-items:center;justify-content:center}}
.eye-btn:hover{{background:#f2f4f7}}
button{{margin-top:22px;border:0;background:#17212b;color:#fff;font-weight:800;cursor:pointer}}
.error{{margin:14px 0;padding:12px;border-radius:10px;background:#fdeeee;color:#b42318;font-weight:700}}
.langs{{display:flex;gap:6px;justify-content:flex-end}}
.langs button{{width:auto;margin:0;padding:7px 9px;background:#f2f4f7;color:#667085}}
.langs button.active{{background:#17212b;color:#fff}}
.forgot{{display:block;text-align:right;margin-top:11px;color:#475467;text-decoration:none;font-weight:700;font-size:14px}}
.forgot:hover{{text-decoration:underline}}
</style>
</head>
<body>
<div class="wrap">
<div class="card">
<div class="langs">
<button type="button" data-lang="ro" class="active">RO</button>
<button type="button" data-lang="de">DE</button>
<button type="button" data-lang="en">EN</button>
</div>

<div class="brand">FICO CONTROL</div>
<h1 id="title">Admin Login</h1>
<p class="subtitle" id="subtitle">Introdu numele tău și parola comună pentru a intra în Admin Dashboard.</p>

{error_html}

<form method="post" action="/admin/login" autocomplete="off">
<input type="hidden" name="next_path" value="{next_value}">

<label id="nameLabel">Numele tău</label>
<input name="display_name" autocomplete="off" required>

<label id="passwordLabel">Parola</label>
<div class="password-wrap">
<input id="adminPassword" name="password" type="password" autocomplete="new-password" required>
<button class="eye-btn" id="togglePassword" type="button" aria-label="Arată parola" title="Arată parola">◉</button>
</div>

<a class="forgot" id="forgotLink" href="/admin/forgot-password">Ai uitat parola?</a>
<button type="submit" id="loginButton">Intră în Admin</button>
</form>
</div>
</div>

<script>
const T = {{
 ro: {{
   title:"Admin Login",
   subtitle:"Introdu numele tău și parola comună pentru a intra în Admin Dashboard.",
   name:"Numele tău",
   password:"Parola",
   forgot:"Ai uitat parola?",
   button:"Intră în Admin"
 }},
 de: {{
   title:"Admin Login",
   subtitle:"Gib deinen Namen und das gemeinsame Passwort ein, um das Admin Dashboard zu öffnen.",
   name:"Dein Name",
   password:"Passwort",
   forgot:"Passwort vergessen?",
   button:"Admin öffnen"
 }},
 en: {{
   title:"Admin Login",
   subtitle:"Enter your name and the shared password to open the Admin Dashboard.",
   name:"Your name",
   password:"Password",
   forgot:"Forgot password?",
   button:"Open Admin"
 }}
}};

document.querySelectorAll("[data-lang]").forEach(btn => {{
 btn.addEventListener("click", () => {{
   document.querySelectorAll("[data-lang]").forEach(x => x.classList.remove("active"));
   btn.classList.add("active");
   const t=T[btn.dataset.lang];
   document.getElementById("title").textContent=t.title;
   document.getElementById("subtitle").textContent=t.subtitle;
   document.getElementById("nameLabel").textContent=t.name;
   document.getElementById("passwordLabel").textContent=t.password;
   document.getElementById("forgotLink").textContent=t.forgot;
   document.getElementById("loginButton").textContent=t.button;
 }});
}});

const passwordInput = document.getElementById("adminPassword");
const togglePassword = document.getElementById("togglePassword");

togglePassword.addEventListener("click", () => {{
  const showing = passwordInput.type === "text";
  passwordInput.type = showing ? "password" : "text";
  togglePassword.textContent = showing ? "◉" : "◎";
  togglePassword.setAttribute(
    "aria-label",
    showing ? "Arată parola" : "Ascunde parola"
  );
  togglePassword.setAttribute(
    "title",
    showing ? "Arată parola" : "Ascunde parola"
  );
}});
</script>
</body>
</html>
"""
    response = HTMLResponse(page)
    response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    response.headers["Pragma"] = "no-cache"
    return response


@app.post("/admin/login")
def admin_login(
    display_name: str = Form(...),
    password: str = Form(...),
    next_path: str = Form("/admin")
):
    clean_name = " ".join((display_name or "").strip().split())
    if not clean_name:
        return RedirectResponse("/admin/login?error=name", status_code=303)

    normalized = normalize_admin_name(clean_name)
    conn = db()

    blocked = conn.execute(
        "SELECT 1 FROM blocked_admin_names WHERE normalized_name=?",
        (normalized,)
    ).fetchone()

    if blocked:
        conn.close()
        return RedirectResponse("/admin/login?error=blocked", status_code=303)

    password_hash = get_shared_password_hash(conn)

    db_password_ok = bool(
        password_hash and verify_password(password, password_hash)
    )

    env_password_ok = bool(
        ADMIN_INITIAL_PASSWORD
        and hmac.compare_digest(password, ADMIN_INITIAL_PASSWORD)
    )

    if not db_password_ok and not env_password_ok:
        conn.close()
        return RedirectResponse("/admin/login?error=invalid", status_code=303)

    if env_password_ok and not db_password_ok:
        conn.execute("""
            INSERT INTO admin_settings(setting_key, setting_value)
            VALUES('shared_password_hash', ?)
            ON CONFLICT(setting_key) DO UPDATE SET
                setting_value=excluded.setting_value
        """, (hash_password(ADMIN_INITIAL_PASSWORD),))
        conn.commit()

    token = create_admin_session(conn, clean_name)
    conn.close()

    safe_next = next_path if next_path.startswith("/admin") else "/admin"
    response = RedirectResponse(safe_next, status_code=303)

    # Session-only cookie: no Max-Age / Expires.
    # It is not designed to persist as a remembered login.
    response.set_cookie(
        ADMIN_COOKIE_NAME,
        token,
        httponly=True,
        secure=True,
        samesite="lax"
    )

    # Remove any old browser-memory cookie from the previous version.
    response.delete_cookie("fico_admin_display_name")
    response.delete_cookie("fico_admin_pending")

    return response


@app.post("/admin/logout")
def admin_logout(request: Request):
    token = request.cookies.get(ADMIN_COOKIE_NAME)
    if token:
        conn = db()
        conn.execute(
            "UPDATE admin_sessions SET revoked=1 WHERE session_token=?",
            (token,)
        )
        conn.commit()
        conn.close()

    response = RedirectResponse("/admin/login", status_code=303)
    response.delete_cookie(ADMIN_COOKIE_NAME)
    response.delete_cookie("fico_admin_display_name")
    response.delete_cookie("fico_admin_pending")
    response.headers["Cache-Control"] = "no-store"
    return response


@app.get("/admin/forgot-password", response_class=HTMLResponse)
def forgot_password_page(message: str | None = None, error: str | None = None):
    message_html = ""
    if message:
        message_html = f'<div class="ok">{html.escape(message)}</div>'
    if error:
        message_html = f'<div class="error">{html.escape(error)}</div>'

    return HTMLResponse(f"""
<!doctype html>
<html lang="ro">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>FICO Control - Ai uitat parola?</title>
<style>
*{{box-sizing:border-box}}
body{{margin:0;background:#f4f6f8;font-family:Arial,sans-serif;color:#17212b}}
.wrap{{width:min(92%,500px);margin:65px auto}}
.card{{background:#fff;border-radius:22px;padding:30px;box-shadow:0 12px 35px rgba(0,0,0,.08)}}
.brand{{font-size:13px;font-weight:900;letter-spacing:2px}}
h1{{font-size:30px;margin:22px 0 8px}}
p{{color:#667085;line-height:1.5}}
input,button{{width:100%;padding:14px;border-radius:11px;font-size:16px}}
input{{border:1px solid #d8dde3;margin-top:10px}}
button{{margin-top:16px;border:0;background:#17212b;color:#fff;font-weight:800;cursor:pointer}}
.ok,.error{{margin:14px 0;padding:12px;border-radius:10px;font-weight:700}}
.ok{{background:#e9f8ef;color:#14804a}}
.error{{background:#fdeeee;color:#b42318}}
.back{{display:block;text-align:center;margin-top:16px;color:#17212b;font-weight:700;text-decoration:none}}
</style>
</head>
<body>
<div class="wrap"><div class="card">
<div class="brand">FICO CONTROL</div>
<h1>Ai uitat parola?</h1>
<p>Parola poate fi resetată numai cu un cod trimis la emailul Owner: <strong>{html.escape(mask_email(OWNER_EMAIL))}</strong>.</p>
{message_html}
<form method="post" action="/admin/forgot-password/send-code">
<button type="submit">Trimite codul pe email</button>
</form>
<form method="post" action="/admin/reset-password">
<input name="code" inputmode="numeric" maxlength="6" placeholder="Cod din 6 cifre" required>
<input name="new_password" type="password" minlength="8" placeholder="Parola nouă (minimum 8 caractere)" required>
<button type="submit">Setează parola nouă</button>
</form>
<a class="back" href="/admin/login">Înapoi la login</a>
</div></div>
</body>
</html>
""")


@app.post("/admin/forgot-password/send-code")
def forgot_password_send_code():
    if not OWNER_EMAIL or not RESEND_API_KEY:
        return forgot_password_page(
            error="Emailul Owner sau cheia de email nu este configurată."
        )

    code = f"{secrets.randbelow(1000000):06d}"
    now = utc_now()
    expires = now + timedelta(minutes=OWNER_CODE_MINUTES)

    conn = db()
    conn.execute("UPDATE owner_email_codes SET used=1 WHERE used=0")
    conn.execute(
        """
        INSERT INTO owner_email_codes(code_hash, created_at, expires_at, used)
        VALUES(?,?,?,0)
        """,
        (hash_owner_code(code), iso_utc(now), iso_utc(expires))
    )
    conn.commit()
    conn.close()

    try:
        send_owner_verification_code(code)
    except Exception:
        return forgot_password_page(
            error="Codul nu a putut fi trimis. Verifică setările de email."
        )

    return forgot_password_page(
        message=f"Codul a fost trimis la {mask_email(OWNER_EMAIL)}."
    )


@app.post("/admin/reset-password")
def reset_password_with_email_code(
    code: str = Form(...),
    new_password: str = Form(...)
):
    clean_code = re.sub(r"\D", "", code or "")
    if len(clean_code) != 6:
        return forgot_password_page(error="Cod invalid.")

    if len(new_password) < 8:
        return forgot_password_page(
            error="Parola trebuie să aibă minimum 8 caractere."
        )

    conn = db()
    row = conn.execute(
        """
        SELECT id, code_hash, expires_at
        FROM owner_email_codes
        WHERE used=0
        ORDER BY id DESC
        LIMIT 1
        """
    ).fetchone()

    if not row:
        conn.close()
        return forgot_password_page(error="Nu există un cod activ.")

    expires = parse_iso_utc(row["expires_at"])
    valid = (
        expires
        and expires > utc_now()
        and hmac.compare_digest(
            row["code_hash"],
            hash_owner_code(clean_code)
        )
    )

    if not valid:
        conn.close()
        return forgot_password_page(
            error="Codul este greșit sau a expirat."
        )

    conn.execute(
        "UPDATE owner_email_codes SET used=1 WHERE id=?",
        (row["id"],)
    )
    conn.execute(
        """
        INSERT INTO admin_settings(setting_key, setting_value)
        VALUES('shared_password_hash', ?)
        ON CONFLICT(setting_key) DO UPDATE SET
            setting_value=excluded.setting_value
        """,
        (hash_password(new_password),)
    )
    # Force everyone to use the new password.
    conn.execute("UPDATE admin_sessions SET revoked=1")
    conn.commit()
    conn.close()

    return RedirectResponse("/admin/login", status_code=303)


@app.post("/admin/owner/send-code")
def owner_send_code(request: Request):
    if not OWNER_EMAIL or not RESEND_API_KEY:
        return HTMLResponse(
            owner_gate_html(
                "OWNER_EMAIL sau RESEND_API_KEY nu este configurat în Render.",
                error=True
            ),
            status_code=503
        )

    code = f"{secrets.randbelow(1000000):06d}"
    now = utc_now()
    expires = now + timedelta(minutes=OWNER_CODE_MINUTES)

    conn = db()
    conn.execute("UPDATE owner_email_codes SET used=1 WHERE used=0")
    conn.execute("""
        INSERT INTO owner_email_codes(
            code_hash, created_at, expires_at, used
        ) VALUES(?,?,?,0)
    """, (
        hash_owner_code(code),
        iso_utc(now),
        iso_utc(expires)
    ))
    conn.commit()
    conn.close()

    try:
        send_owner_verification_code(code)
    except Exception as exc:
        return HTMLResponse(
            owner_gate_html(
                "Codul nu a putut fi trimis. Verifică setările de email din Render.",
                error=True
            ),
            status_code=502
        )

    return HTMLResponse(
        owner_gate_html(
            f"Codul a fost trimis la {mask_email(OWNER_EMAIL)}."
        )
    )


@app.post("/admin/owner/verify-code")
def owner_verify_code(request: Request, code: str = Form(...)):
    clean_code = re.sub(r"\\D", "", code or "")
    if len(clean_code) != 6:
        return HTMLResponse(
            owner_gate_html("Cod invalid.", error=True),
            status_code=400
        )

    conn = db()
    row = conn.execute("""
        SELECT id, code_hash, expires_at
        FROM owner_email_codes
        WHERE used=0
        ORDER BY id DESC
        LIMIT 1
    """).fetchone()

    if not row:
        conn.close()
        return HTMLResponse(
            owner_gate_html("Nu există un cod activ.", error=True),
            status_code=400
        )

    expires = parse_iso_utc(row["expires_at"])
    valid = (
        expires
        and expires > utc_now()
        and hmac.compare_digest(
            row["code_hash"],
            hash_owner_code(clean_code)
        )
    )

    if not valid:
        conn.close()
        return HTMLResponse(
            owner_gate_html("Codul este greșit sau a expirat.", error=True),
            status_code=400
        )

    token = request.cookies.get(ADMIN_COOKIE_NAME)
    owner_until = utc_now() + timedelta(minutes=OWNER_VERIFY_MINUTES)

    conn.execute(
        "UPDATE owner_email_codes SET used=1 WHERE id=?",
        (row["id"],)
    )
    conn.execute(
        "UPDATE admin_sessions SET owner_verified_until=? WHERE session_token=?",
        (iso_utc(owner_until), token)
    )
    conn.commit()
    conn.close()

    return RedirectResponse("/admin/owner", status_code=303)


@app.get("/admin/owner", response_class=HTMLResponse)
def admin_owner_page(request: Request):
    session = getattr(request.state, "admin_session", None)
    if not is_owner_verified(session):
        return HTMLResponse(owner_gate_html())

    conn = db()
    active = conn.execute("""
        SELECT display_name,
               normalized_name,
               created_at,
               last_seen,
               expires_at,
               COUNT(*) AS session_count
        FROM admin_sessions
        WHERE revoked=0
          AND expires_at > ?
        GROUP BY normalized_name, display_name
        ORDER BY last_seen DESC
    """, (iso_utc(utc_now()),)).fetchall()

    blocked = conn.execute("""
        SELECT display_name, normalized_name, blocked_at
        FROM blocked_admin_names
        ORDER BY blocked_at DESC
    """).fetchall()
    conn.close()

    rows = ""
    for r in active:
        rows += f"""
        <tr>
          <td><strong>{html.escape(r["display_name"])}</strong></td>
          <td>{html.escape(r["last_seen"][0:16].replace("T"," "))}</td>
          <td>{r["session_count"]}</td>
          <td>
            <form method="post" action="/admin/owner/revoke" style="display:inline">
              <input type="hidden" name="display_name" value="{html.escape(r["display_name"], quote=True)}">
              <button class="small" type="submit">Deconectează</button>
            </form>
            <form method="post" action="/admin/owner/block" style="display:inline">
              <input type="hidden" name="display_name" value="{html.escape(r["display_name"], quote=True)}">
              <button class="small danger" type="submit">Blochează</button>
            </form>
          </td>
        </tr>
        """

    blocked_rows = ""
    for r in blocked:
        blocked_rows += f"""
        <tr>
          <td><strong>{html.escape(r["display_name"])}</strong></td>
          <td>{html.escape(r["blocked_at"][0:16].replace("T"," "))}</td>
          <td>
            <form method="post" action="/admin/owner/unblock">
              <input type="hidden" name="display_name" value="{html.escape(r["display_name"], quote=True)}">
              <button class="small" type="submit">Deblochează</button>
            </form>
          </td>
        </tr>
        """

    page = f"""
<!doctype html>
<html lang="ro">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>FICO Control Owner</title>
<style>
*{{box-sizing:border-box}}
body{{margin:0;background:#f4f6f8;font-family:Arial,sans-serif;color:#17212b}}
.wrap{{width:min(95%,1100px);margin:35px auto 60px}}
.top{{display:flex;justify-content:space-between;align-items:center;gap:15px;flex-wrap:wrap}}
.brand{{font-size:13px;font-weight:900;letter-spacing:2px}}
h1{{font-size:36px;margin:14px 0 25px}}
.card{{background:#fff;border-radius:16px;padding:20px;margin-bottom:16px;box-shadow:0 5px 18px rgba(0,0,0,.05)}}
table{{width:100%;border-collapse:collapse}}
th,td{{text-align:left;padding:13px;border-bottom:1px solid #eceff2}}
input,button{{padding:11px 13px;border-radius:9px;border:1px solid #d8dde3}}
button{{cursor:pointer;font-weight:800}}
.small{{background:#fff;color:#17212b}}
.danger{{background:#fdeeee;color:#b42318;border-color:#f4b8b3}}
.primary{{background:#17212b;color:#fff;border:0}}
a{{color:#17212b;font-weight:800;text-decoration:none}}
.password-row{{display:flex;gap:10px;flex-wrap:wrap}}
.password-row input{{min-width:260px;flex:1}}
</style>
</head>
<body>
<main class="wrap">
<div class="top">
<div>
<div class="brand">FICO CONTROL · OWNER</div>
<h1>Control acces Admin</h1>
</div>
<a href="/admin">Înapoi la Dashboard</a>
</div>

<section class="card">
<h2>Schimbă parola comună</h2>
<p>Confirmarea Owner este valabilă temporar după codul primit pe email. Schimbarea parolei va deconecta toate celelalte sesiuni.</p>
<form class="password-row" method="post" action="/admin/owner/change-password">
<input type="password" name="new_password" minlength="8" placeholder="Parola nouă (minimum 8 caractere)" required>
<button class="primary" type="submit">Schimbă parola</button>
</form>
</section>

<section class="card">
<h2>Persoane conectate</h2>
<table>
<thead><tr><th>Nume</th><th>Ultima activitate</th><th>Sesiuni</th><th>Acțiuni</th></tr></thead>
<tbody>{rows or '<tr><td colspan="4">Nicio sesiune activă.</td></tr>'}</tbody>
</table>
</section>

<section class="card">
<h2>Nume blocate</h2>
<table>
<thead><tr><th>Nume</th><th>Blocat la</th><th>Acțiune</th></tr></thead>
<tbody>{blocked_rows or '<tr><td colspan="3">Niciun nume blocat.</td></tr>'}</tbody>
</table>
</section>
</main>
</body>
</html>
"""
    return HTMLResponse(page)


def require_owner_verified(request: Request):
    session = getattr(request.state, "admin_session", None)
    if not is_owner_verified(session):
        raise HTTPException(status_code=403, detail="owner_verification_required")
    return session


@app.post("/admin/owner/revoke")
def owner_revoke_user(request: Request, display_name: str = Form(...)):
    require_owner_verified(request)
    normalized = normalize_admin_name(display_name)

    conn = db()
    conn.execute(
        "UPDATE admin_sessions SET revoked=1 WHERE normalized_name=?",
        (normalized,)
    )
    conn.commit()
    conn.close()
    return RedirectResponse("/admin/owner", status_code=303)


@app.post("/admin/owner/block")
def owner_block_user(request: Request, display_name: str = Form(...)):
    require_owner_verified(request)
    clean = " ".join((display_name or "").strip().split())
    normalized = normalize_admin_name(clean)

    conn = db()
    conn.execute("""
        INSERT INTO blocked_admin_names(
            normalized_name, display_name, blocked_at
        ) VALUES(?,?,?)
        ON CONFLICT(normalized_name) DO UPDATE SET
            display_name=excluded.display_name,
            blocked_at=excluded.blocked_at
    """, (normalized, clean, iso_utc(utc_now())))
    conn.execute(
        "UPDATE admin_sessions SET revoked=1 WHERE normalized_name=?",
        (normalized,)
    )
    conn.commit()
    conn.close()
    return RedirectResponse("/admin/owner", status_code=303)


@app.post("/admin/owner/unblock")
def owner_unblock_user(request: Request, display_name: str = Form(...)):
    require_owner_verified(request)
    normalized = normalize_admin_name(display_name)

    conn = db()
    conn.execute(
        "DELETE FROM blocked_admin_names WHERE normalized_name=?",
        (normalized,)
    )
    conn.commit()
    conn.close()
    return RedirectResponse("/admin/owner", status_code=303)


@app.post("/admin/owner/change-password")
def owner_change_password(
    request: Request,
    new_password: str = Form(...)
):
    session = require_owner_verified(request)

    if len(new_password) < 8:
        raise HTTPException(
            status_code=400,
            detail="password_must_have_at_least_8_characters"
        )

    token = request.cookies.get(ADMIN_COOKIE_NAME)
    conn = db()
    conn.execute("""
        INSERT INTO admin_settings(setting_key, setting_value)
        VALUES('shared_password_hash', ?)
        ON CONFLICT(setting_key) DO UPDATE SET
            setting_value=excluded.setting_value
    """, (hash_password(new_password),))

    conn.execute(
        "UPDATE admin_sessions SET revoked=1 WHERE session_token<>?",
        (token,)
    )
    conn.commit()
    conn.close()

    return RedirectResponse("/admin/owner", status_code=303)




def normalize_name(value: str) -> str:
    value = unicodedata.normalize("NFKD", value or "")
    value = "".join(ch for ch in value if not unicodedata.combining(ch))
    value = value.replace("-", " ")
    value = " ".join(value.strip().split())
    return value.casefold()


def token_matches(entered_token: str, full_token: str) -> bool:
    if not entered_token or not full_token:
        return False

    if entered_token == full_token:
        return True

    if len(entered_token) == 1:
        return full_token.startswith(entered_token)

    if len(entered_token) >= 2 and full_token.startswith(entered_token):
        return True

    if len(entered_token) >= 4:
        return SequenceMatcher(None, entered_token, full_token).ratio() >= 0.84

    return False


def token_subsequence_score(entered_tokens: list[str], full_tokens: list[str]) -> float:
    if not entered_tokens or not full_tokens:
        return 0.0

    matched_positions = []
    search_from = 0

    for entered_token in entered_tokens:
        found = None
        for idx in range(search_from, len(full_tokens)):
            if token_matches(entered_token, full_tokens[idx]):
                found = idx
                break

        if found is None:
            return 0.0

        matched_positions.append(found)
        search_from = found + 1

    exact = sum(
        1
        for entered_token, idx in zip(entered_tokens, matched_positions)
        if entered_token == full_tokens[idx]
    )
    exact_ratio = exact / len(entered_tokens)

    first_last_bonus = 0.0
    if len(entered_tokens) >= 2:
        if (
            token_matches(entered_tokens[0], full_tokens[0])
            and token_matches(entered_tokens[-1], full_tokens[-1])
        ):
            first_last_bonus = 0.06

    return min(1.0, 0.88 + 0.06 * exact_ratio + first_last_bonus)


def unordered_token_score(entered_tokens: list[str], full_tokens: list[str]) -> float:
    if not entered_tokens or not full_tokens:
        return 0.0

    used = set()
    exact = 0

    for entered_token in entered_tokens:
        best_idx = None
        for idx, full_token in enumerate(full_tokens):
            if idx in used:
                continue
            if token_matches(entered_token, full_token):
                best_idx = idx
                if entered_token == full_token:
                    exact += 1
                break

        if best_idx is None:
            return 0.0

        used.add(best_idx)

    exact_ratio = exact / len(entered_tokens)
    return min(0.94, 0.86 + 0.08 * exact_ratio)


def name_similarity(a: str, b: str) -> float:
    a_n = normalize_name(a)
    b_n = normalize_name(b)

    if not a_n or not b_n:
        return 0.0

    if a_n == b_n:
        return 1.0

    if b_n.startswith(a_n) or a_n.startswith(b_n):
        shorter = min(len(a_n), len(b_n))
        longer = max(len(a_n), len(b_n))
        return 0.92 + (0.08 * shorter / max(longer, 1))

    entered_tokens = a_n.split()
    full_tokens = b_n.split()

    sequence_score = token_subsequence_score(entered_tokens, full_tokens)
    unordered_score = unordered_token_score(entered_tokens, full_tokens)
    ratio = SequenceMatcher(None, a_n, b_n).ratio()

    return max(sequence_score, unordered_score, ratio * 0.90)


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

    # Accept clear partial names such as:
    # "Alain Sery" -> "Alain Gnebehi Kamin Sery"
    # "Elvis V"    -> "Elvis Velcu"
    #
    # If two drivers match almost equally well, do not guess.
    if best_score >= 0.84:
        ambiguous = second_score >= 0.84 and (best_score - second_score) < 0.06
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

    clean_entered_name = " ".join((full_name or "").strip().split())
    if not clean_entered_name:
        raise HTTPException(status_code=400, detail="missing_name")

    today = date.today().isoformat()
    conn = db()

    driver, match_score, ambiguous = find_required_driver(
        conn,
        today,
        clean_entered_name
    )

    # Always receive and save the proof first. Even if the name cannot be
    # identified safely, the submission must not be lost.
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

    # If we cannot safely identify the driver, keep the submission in a
    # separate review queue instead of rejecting it.
    if not driver:
        best_match_name = None

        # Find the highest candidate only as a hint for the admin.
        candidates = conn.execute("""
            SELECT d.name
            FROM daily_required r
            JOIN drivers d ON d.id = r.driver_id
            WHERE r.work_date = ?
            ORDER BY d.name
        """, (today,)).fetchall()

        best_hint_score = 0.0
        for candidate in candidates:
            candidate_score = name_similarity(
                clean_entered_name,
                candidate["name"]
            )
            if candidate_score > best_hint_score:
                best_hint_score = candidate_score
                best_match_name = candidate["name"]

        reason = "ambiguous_name" if ambiguous else "unrecognized_name"

        conn.execute("""
            INSERT INTO unresolved_submissions(
                work_date,
                entered_full_name,
                fico_score,
                submitted_at,
                proof_filename,
                proof_original_name,
                detected_fico_score,
                verification_status,
                best_match_name,
                best_match_score,
                match_reason
            ) VALUES(?,?,?,?,?,?,?,?,?,?,?)
        """, (
            today,
            clean_entered_name,
            fico_score,
            datetime.now().isoformat(timespec="seconds"),
            filename,
            original_name,
            detected_score,
            verification_status,
            best_match_name,
            best_hint_score,
            reason
        ))

        conn.commit()
        conn.close()

        return {
            "ok": True,
            "needs_name_review": True,
            "entered_name": clean_entered_name,
            "best_match_name": best_match_name,
            "best_match_score": round(best_hint_score, 3),
            "detected_fico_score": detected_score,
            "verification_status": verification_status,
            "ocr_state": ocr_state
        }

    existing = conn.execute(
        "SELECT 1 FROM submissions WHERE work_date=? AND driver_id=?",
        (today, driver["id"])
    ).fetchone()

    if existing:
        conn.close()
        raise HTTPException(status_code=409, detail="already_sent")

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
        clean_entered_name,
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
.proof-input-native{{position:absolute;left:-9999px;width:1px;height:1px;opacity:0}}
.proof-action{{display:block;width:100%;box-sizing:border-box;padding:14px 12px;border-radius:11px;border:1px solid #d8dde3;background:#fff;color:#17212b;font-size:15px;font-weight:800;cursor:pointer;text-align:center}}
.proof-action.primary{{background:#17212b;color:#fff;border-color:#17212b}}
.camera-label{{margin-top:2px}}
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

<label class="proof-action primary camera-label" for="proof" data-i18n="takePhoto">Fă o poză acum</label>
<input
  id="proof"
  class="proof-input-native"
  type="file"
  name="proof"
  accept="image/*"
  capture="environment"
  required
>
<div id="proofSelected" class="proof-selected"></div>
<div class="hint" data-i18n="photoHint">Fă o poză clară în care scorul FICO este vizibil · maximum 10 MB</div>

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
    photoSelected: "Poză selectată:",
    photoHint: "Fă o poză clară în care scorul FICO este vizibil · maximum 10 MB",
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
    photoSelected: "Ausgewähltes Bild:",
    photoHint: "Nimm ein klares Foto auf, auf dem der FICO-Score sichtbar ist · maximal 10 MB",
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
    photoSelected: "Selected image:",
    photoHint: "Take a clear photo where the FICO score is visible · maximum 10 MB",
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
def admin_page(request: Request, d: str | None = None, q: str | None = None):
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

    unresolved_rows = conn.execute("""
        SELECT id,
               entered_full_name,
               fico_score,
               submitted_at,
               proof_filename,
               proof_original_name,
               detected_fico_score,
               verification_status,
               best_match_name,
               best_match_score,
               match_reason
        FROM unresolved_submissions
        WHERE work_date = ?
        ORDER BY submitted_at DESC
    """, (selected,)).fetchall()

    conn.close()

    total = len(rows)
    sent = sum(int(r["sent"]) for r in rows)
    missing = total - sent
    low_fico = sum(
        1 for r in rows
        if r["fico_score"] is not None and int(r["fico_score"]) < 800
    )

    needs_review = (
        sum(
            1 for r in rows
            if r["verification_status"] in ("mismatch", "manual_review")
        )
        + len(unresolved_rows)
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

    unresolved_table = ""
    if unresolved_rows:
        unresolved_items = ""

        for u in unresolved_rows:
            entered = html.escape(u["entered_full_name"] or "—")
            best_match = html.escape(u["best_match_name"] or "Nicio potrivire sigură")
            best_score = (
                f'{round(float(u["best_match_score"]) * 100)}%'
                if u["best_match_score"] is not None
                else "—"
            )
            fico_unresolved = (
                u["fico_score"]
                if u["fico_score"] is not None
                else "—"
            )
            detected_unresolved = (
                u["detected_fico_score"]
                if u["detected_fico_score"] is not None
                else "—"
            )
            unresolved_hour = (
                u["submitted_at"][11:16]
                if u["submitted_at"]
                else "—"
            )

            if u["proof_filename"]:
                unresolved_proof_url = (
                    f'/proof/{html.escape(u["proof_filename"])}'
                )
                unresolved_proof = (
                    f'<a class="proof-btn" href="{unresolved_proof_url}" '
                    f'target="_blank">Vezi poza</a>'
                )
            else:
                unresolved_proof = '<span class="dash">—</span>'

            if u["verification_status"] == "mismatch":
                fico_check = '<span class="verify-pill mismatch">⚠ FICO diferit</span>'
            elif u["verification_status"] == "verified":
                fico_check = '<span class="verify-pill verified">✓ FICO verificat</span>'
            else:
                fico_check = '<span class="verify-pill manual">? FICO manual</span>'

            unresolved_items += f"""
            <tr class="row-name-review">
                <td>
                    <div class="driver-name">⚠ {entered}</div>
                    <div class="entered-name">
                        Posibil: {best_match} · potrivire {best_score}
                    </div>
                </td>
                <td><span class="verify-pill mismatch">⚠ Nume de verificat</span></td>
                <td><strong>{fico_unresolved}</strong></td>
                <td><strong>{detected_unresolved}</strong></td>
                <td>{fico_check}</td>
                <td>{unresolved_hour}</td>
                <td>{unresolved_proof}</td>
            </tr>
            """

        unresolved_table = f"""
        <section class="name-review-section">
            <div class="review-title">
                ⚠ Trimiteri cu nume neidentificat ({len(unresolved_rows)})
            </div>
            <div class="review-subtitle">
                Poza și scorul au fost salvate. Verifică manual cui aparține trimiterea.
            </div>
            <table>
                <thead>
                    <tr>
                        <th>Nume introdus</th>
                        <th>Atenție</th>
                        <th>FICO introdus</th>
                        <th>FICO detectat</th>
                        <th>Verificare</th>
                        <th>Ora</th>
                        <th>Dovadă</th>
                    </tr>
                </thead>
                <tbody>{unresolved_items}</tbody>
            </table>
        </section>
        """

    missing_js = "\\n".join(missing_names).replace("\\\\", "\\\\\\\\").replace("`", "\\\\`")

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
.name-review-section{{margin:0 0 20px;border:2px solid #f59e0b;border-radius:16px;overflow:hidden;background:#fffaf0}}
.review-title{{font-size:18px;font-weight:900;color:#9a5a00;padding:16px 18px 4px}}
.review-subtitle{{font-size:13px;color:#8a6500;padding:0 18px 14px}}
.row-name-review{{background:#fff8e6}}
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
        <div style="margin-top:8px;color:#667085;font-size:13px">
            Conectat ca: <strong>{html.escape(getattr(request.state, "admin_session", {}).get("display_name", "Admin"))}</strong>
        </div>
    </div>

    <div class="top-actions">
        <a class="btn-light" href="/admin/owner">Owner</a>
        <form method="post" action="/admin/logout" style="margin:0">
            <button class="btn-light" type="submit">Ieșire</button>
        </form>
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
    {unresolved_table}
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
