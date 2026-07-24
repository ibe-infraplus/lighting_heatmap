from __future__ import annotations

import html
import os
import secrets
import time
from datetime import datetime, timedelta, timezone
from functools import wraps
from pathlib import Path
from urllib.parse import urlsplit

import psycopg
from flask import (
    Flask,
    Response,
    abort,
    redirect,
    render_template,
    request,
    session,
    url_for,
)
from psycopg.rows import dict_row
from werkzeug.middleware.proxy_fix import ProxyFix
from werkzeug.security import check_password_hash, generate_password_hash


APP_ROOT = Path(__file__).resolve().parent
DASHBOARD_FILE = APP_ROOT / "lighting_dashboard.html"


def env_bool(name: str, default: bool = False) -> bool:
    value = os.getenv(name)
    if value is None:
        return default
    return value.strip().lower() in {"1", "true", "yes", "on"}


def normalized_base_path() -> str:
    value = os.getenv("APP_BASE_PATH", "/lighting_heatmap").strip()
    if not value.startswith("/"):
        value = f"/{value}"
    return value.rstrip("/") or "/lighting_heatmap"


BASE_PATH = normalized_base_path()


def database_connection() -> psycopg.Connection:
    return psycopg.connect(
        host=os.environ["DB_HOST"],
        port=int(os.getenv("DB_PORT", "5432")),
        dbname=os.environ["DB_NAME"],
        user=os.environ["DB_USER"],
        password=os.environ["DB_PASSWORD"],
        sslmode=os.getenv("DB_SSLMODE", "prefer"),
        connect_timeout=6,
        application_name="lighting_heatmap",
        row_factory=dict_row,
    )


def init_database() -> None:
    last_error: Exception | None = None
    for attempt in range(1, 11):
        try:
            with database_connection() as conn:
                with conn.cursor() as cur:
                    cur.execute(
                        """
                        CREATE TABLE IF NOT EXISTS app_users (
                            id BIGSERIAL PRIMARY KEY,
                            username VARCHAR(80) NOT NULL,
                            password_hash TEXT NOT NULL,
                            display_name VARCHAR(120) NOT NULL,
                            is_active BOOLEAN NOT NULL DEFAULT TRUE,
                            failed_attempts INTEGER NOT NULL DEFAULT 0,
                            locked_until TIMESTAMPTZ,
                            last_login_at TIMESTAMPTZ,
                            created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
                            updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
                        )
                        """
                    )
                    cur.execute(
                        """
                        CREATE UNIQUE INDEX IF NOT EXISTS app_users_username_lower_idx
                        ON app_users (LOWER(username))
                        """
                    )
                    cur.execute(
                        """
                        CREATE TABLE IF NOT EXISTS app_login_audit (
                            id BIGSERIAL PRIMARY KEY,
                            username VARCHAR(80) NOT NULL,
                            user_id BIGINT REFERENCES app_users(id) ON DELETE SET NULL,
                            success BOOLEAN NOT NULL,
                            ip_address INET,
                            user_agent TEXT,
                            created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
                        )
                        """
                    )

                    admin_username = os.getenv("APP_ADMIN_USERNAME", "").strip()
                    admin_password = os.getenv("APP_ADMIN_PASSWORD", "")
                    admin_display_name = os.getenv("APP_ADMIN_DISPLAY_NAME", "ผู้ดูแลระบบ").strip()
                    if admin_username and admin_password:
                        if len(admin_password) < 10:
                            raise RuntimeError("APP_ADMIN_PASSWORD ต้องมีอย่างน้อย 10 ตัวอักษร")
                        cur.execute(
                            """
                            INSERT INTO app_users (username, password_hash, display_name)
                            SELECT %s, %s, %s
                            WHERE NOT EXISTS (
                                SELECT 1 FROM app_users WHERE LOWER(username) = LOWER(%s)
                            )
                            ON CONFLICT DO NOTHING
                            """,
                            (
                                admin_username,
                                generate_password_hash(admin_password),
                                admin_display_name or admin_username,
                                admin_username,
                            ),
                        )
                conn.commit()
            return
        except (psycopg.OperationalError, psycopg.InterfaceError) as exc:
            last_error = exc
            if attempt == 10:
                break
            time.sleep(2)
    raise RuntimeError("ไม่สามารถเชื่อมต่อ PostgreSQL เพื่อเตรียมระบบ Login") from last_error


def safe_next_url(value: str | None) -> str:
    if not value:
        return BASE_PATH
    parsed = urlsplit(value)
    if parsed.scheme or parsed.netloc or not parsed.path.startswith(BASE_PATH):
        return BASE_PATH
    return value


def client_ip() -> str | None:
    value = request.remote_addr
    return value if value and value not in {"-", "unknown"} else None


def create_csrf_token() -> str:
    token = session.get("_csrf_token")
    if not token:
        token = secrets.token_urlsafe(32)
        session["_csrf_token"] = token
    return token


def valid_csrf_token() -> bool:
    expected = session.get("_csrf_token", "")
    supplied = request.form.get("csrf_token", "")
    return bool(expected and supplied and secrets.compare_digest(expected, supplied))


def login_required(view):
    @wraps(view)
    def wrapped(*args, **kwargs):
        if not session.get("user_id"):
            return redirect(url_for("login", next=request.full_path.rstrip("?")))
        return view(*args, **kwargs)

    return wrapped


def inject_logout_control(document: str) -> str:
    username = html.escape(str(session.get("display_name") or session.get("username") or "ผู้ใช้งาน"))
    csrf_token = html.escape(create_csrf_token(), quote=True)
    logout_url = html.escape(url_for("logout"), quote=True)
    control = f"""
    <style>
      .auth-session-control {{
        position: fixed; z-index: 40; left: 72px; bottom: 16px;
        display: flex; align-items: center; gap: 9px; padding: 7px 8px 7px 12px;
        border: 1px solid rgba(209,217,230,.92); border-radius: 16px;
        background: rgba(255,255,255,.96); box-shadow: 0 6px 22px rgba(15,35,62,.16);
        color: #344054; font-family: "Noto Sans Thai","Leelawadee UI",Tahoma,sans-serif;
        backdrop-filter: blur(12px);
      }}
      .auth-session-user {{ max-width: 150px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; font-size: 11px; font-weight: 700; }}
      .auth-session-control button {{
        border: 0; border-radius: 11px; padding: 7px 10px; cursor: pointer;
        color: #b42318; background: #fff1f0; font: inherit; font-size: 11px; font-weight: 800;
      }}
      .auth-session-control button:hover {{ background: #ffe1de; }}
      @media (max-width: 700px) {{
        .auth-session-control {{ left: 64px; bottom: 10px; }}
        .auth-session-user {{ display: none; }}
      }}
    </style>
    <form class="auth-session-control" action="{logout_url}" method="post">
      <input type="hidden" name="csrf_token" value="{csrf_token}" />
      <span class="auth-session-user" title="{username}">{username}</span>
      <button type="submit">ออกจากระบบ</button>
    </form>
    """
    return document.replace("</body>", f"{control}</body>", 1)


def create_app() -> Flask:
    secret_key = os.getenv("SECRET_KEY", "")
    if len(secret_key) < 32:
        raise RuntimeError("SECRET_KEY ต้องมีอย่างน้อย 32 ตัวอักษร")

    init_database()

    app = Flask(__name__, template_folder=str(APP_ROOT / "templates"))
    app.secret_key = secret_key
    app.wsgi_app = ProxyFix(app.wsgi_app, x_for=1, x_proto=1, x_host=1)
    app.config.update(
        SESSION_COOKIE_NAME="lighting_session",
        SESSION_COOKIE_HTTPONLY=True,
        SESSION_COOKIE_SECURE=env_bool("SESSION_COOKIE_SECURE", True),
        SESSION_COOKIE_SAMESITE="Lax",
        SESSION_COOKIE_PATH=BASE_PATH,
        PERMANENT_SESSION_LIFETIME=timedelta(hours=8),
        MAX_CONTENT_LENGTH=32 * 1024,
    )

    @app.context_processor
    def template_context():
        return {"csrf_token": create_csrf_token, "base_path": BASE_PATH}

    @app.after_request
    def security_headers(response: Response):
        response.headers["X-Content-Type-Options"] = "nosniff"
        response.headers["X-Frame-Options"] = "SAMEORIGIN"
        response.headers["Referrer-Policy"] = "strict-origin-when-cross-origin"
        response.headers["Permissions-Policy"] = "camera=(), microphone=(), geolocation=()"
        response.headers["Content-Security-Policy"] = (
            "default-src 'self'; "
            "script-src 'self' 'unsafe-inline' https://unpkg.com; "
            "style-src 'self' 'unsafe-inline' https://unpkg.com; "
            "img-src 'self' data: blob: https:; "
            "connect-src 'self' https:; "
            "font-src 'self' data: https:; "
            "frame-ancestors 'self'; base-uri 'self'; form-action 'self'"
        )
        if request.endpoint in {"login", "dashboard"}:
            response.headers["Cache-Control"] = "no-store, private"
        return response

    @app.get("/health")
    def health():
        try:
            with database_connection() as conn:
                with conn.cursor() as cur:
                    cur.execute("SELECT 1")
                    cur.fetchone()
            return {"status": "ok", "database": "connected"}
        except psycopg.Error:
            return {"status": "unhealthy", "database": "unavailable"}, 503

    @app.get("/")
    def root():
        return redirect(BASE_PATH)

    @app.get("/login")
    def legacy_login():
        return redirect(url_for("login"))

    @app.route(f"{BASE_PATH}/login", methods=["GET", "POST"])
    def login():
        if session.get("user_id"):
            return redirect(BASE_PATH)

        error = ""
        username = ""
        next_url = safe_next_url(request.values.get("next"))

        if request.method == "POST":
            if not valid_csrf_token():
                abort(400)

            username = request.form.get("username", "").strip()[:80]
            password = request.form.get("password", "")
            now = datetime.now(timezone.utc)

            if not username or not password:
                error = "กรุณากรอกชื่อผู้ใช้และรหัสผ่าน"
            else:
                with database_connection() as conn:
                    with conn.cursor() as cur:
                        cur.execute(
                            """
                            SELECT id, username, display_name, password_hash, is_active,
                                   failed_attempts, locked_until
                            FROM app_users
                            WHERE LOWER(username) = LOWER(%s)
                            FOR UPDATE
                            """,
                            (username,),
                        )
                        user = cur.fetchone()
                        success = False
                        locked = False

                        if user and user["is_active"]:
                            locked_until = user["locked_until"]
                            locked = bool(locked_until and locked_until > now)
                            success = not locked and check_password_hash(user["password_hash"], password)

                            if success:
                                cur.execute(
                                    """
                                    UPDATE app_users
                                    SET failed_attempts = 0, locked_until = NULL,
                                        last_login_at = NOW(), updated_at = NOW()
                                    WHERE id = %s
                                    """,
                                    (user["id"],),
                                )
                            elif not locked:
                                previous_attempts = 0 if locked_until and locked_until <= now else int(user["failed_attempts"] or 0)
                                next_attempts = previous_attempts + 1
                                next_lock = now + timedelta(minutes=15) if next_attempts >= 5 else None
                                cur.execute(
                                    """
                                    UPDATE app_users
                                    SET failed_attempts = %s, locked_until = %s, updated_at = NOW()
                                    WHERE id = %s
                                    """,
                                    (next_attempts, next_lock, user["id"]),
                                )

                        cur.execute(
                            """
                            INSERT INTO app_login_audit
                                (username, user_id, success, ip_address, user_agent)
                            VALUES (%s, %s, %s, %s, %s)
                            """,
                            (
                                username,
                                user["id"] if user else None,
                                success,
                                client_ip(),
                                request.user_agent.string[:500],
                            ),
                        )
                    conn.commit()

                if success and user:
                    session.clear()
                    session.permanent = True
                    session["user_id"] = user["id"]
                    session["username"] = user["username"]
                    session["display_name"] = user["display_name"]
                    create_csrf_token()
                    return redirect(next_url)

                error = (
                    "บัญชีถูกพักชั่วคราว กรุณาลองใหม่ภายหลัง"
                    if locked
                    else "ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง"
                )

        return render_template(
            "login.html",
            error=error,
            username=username,
            next_url=next_url,
        )

    @app.post(f"{BASE_PATH}/logout")
    @login_required
    def logout():
        if not valid_csrf_token():
            abort(400)
        session.clear()
        return redirect(url_for("login"))

    @app.get(BASE_PATH)
    @login_required
    def dashboard():
        if not DASHBOARD_FILE.exists():
            return "ยังไม่ได้สร้างไฟล์ Dashboard", 503
        document = DASHBOARD_FILE.read_text(encoding="utf-8")
        return Response(inject_logout_control(document), content_type="text/html; charset=utf-8")

    @app.get(f"{BASE_PATH}/")
    def dashboard_slash():
        return redirect(BASE_PATH)

    return app


app = create_app()
