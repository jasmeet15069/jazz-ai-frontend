
"""
JAZZ Uncensored AI — v7.1 (Production Edition)
===============================================
✅ ALL v7.0 features preserved (Subscription Management, RAG, Vision, etc.)
🔧 BUG FIX: Schema migration system — no more "no such column" on upgrades
✨ NEW FEATURES:
   - Chat Sessions: Named, persistent conversation threads
   - API Keys: Users can generate API keys for programmatic access
   - System Announcements: Admin can broadcast messages to all/specific users
   - User Activity Log: Full audit trail of user actions (login, upload, etc.)
   - Password Change + Reset: Secure token-based password reset flow
   - Chat Search: Full-text search across chat history
   - GDPR: Export all user data + account self-deletion
   - Admin Impersonation: Admins can get a scoped token to debug as any user
   - Website Versioning: Keep last 5 versions of each site, restore any
   - Notification Inbox: Users receive in-app notifications (subscription changes, etc.)
   - Bulk Document Delete: Delete multiple docs at once
   - Global Rate Limit Config: Admin can adjust limits live
   - System Health History: Track and expose rolling health stats
"""

from __future__ import annotations

import asyncio
import base64
import csv
import json
import logging
import os
import re
import secrets
import shutil
import subprocess
import sys
import tempfile
import time
import traceback
import uuid
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor
from contextlib import asynccontextmanager
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, AsyncGenerator, Dict, List, Optional, Tuple

import aiosqlite
import fitz
import openpyxl
from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.triggers.cron import CronTrigger
from apscheduler.triggers.interval import IntervalTrigger
from cryptography.fernet import Fernet
from docx import Document as DocxDocument
from dotenv import load_dotenv
from fastapi import Depends, FastAPI, File, HTTPException, Query, Request, UploadFile, BackgroundTasks
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse, StreamingResponse
from fastapi.security import HTTPAuthorizationCredentials, HTTPBearer
from fastapi.staticfiles import StaticFiles
from jose import JWTError, jwt
from openai import OpenAI
from passlib.context import CryptContext
from pydantic import BaseModel, Field

import chromadb
from chromadb.utils import embedding_functions

load_dotenv()

# ─────────────────────────────────────────────────────────────────────────────
# Logging
# ─────────────────────────────────────────────────────────────────────────────

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s — %(message)s",
    handlers=[logging.FileHandler("jazz.log"), logging.StreamHandler()],
)
logger = logging.getLogger("jazz")

# ─────────────────────────────────────────────────────────────────────────────
# Config helpers
# ─────────────────────────────────────────────────────────────────────────────

def _env(key: str, default: str = "") -> str:
    val = os.getenv(key, default)
    if not val:
        logger.warning("[CONFIG] %s not set", key)
    return val

SECRET_KEY: str = _env("JWT_SECRET_KEY")
if not SECRET_KEY:
    SECRET_KEY = secrets.token_urlsafe(64)
    logger.warning("[CONFIG] JWT_SECRET_KEY not set — using ephemeral key")

ALGORITHM                   = "HS256"
ACCESS_TOKEN_EXPIRE_MINUTES = 60 * 24 * 7

_FERNET_KEY_RAW: str = os.getenv("FERNET_KEY", "")
if not _FERNET_KEY_RAW:
    _FERNET_KEY_RAW = Fernet.generate_key().decode()
    logger.warning("[CONFIG] FERNET_KEY not set — add to .env: FERNET_KEY=%s", _FERNET_KEY_RAW)

_fernet = Fernet(
    _FERNET_KEY_RAW.encode() if isinstance(_FERNET_KEY_RAW, str) else _FERNET_KEY_RAW
)

def encrypt_creds(data: dict) -> str:
    return _fernet.encrypt(json.dumps(data).encode()).decode()

def decrypt_creds(token: str) -> dict:
    return json.loads(_fernet.decrypt(token.encode()).decode())

ADMIN_EMAIL:    str = os.getenv("ADMIN_EMAIL", "jasmeet.15069@gmail.com")
ADMIN_PASSWORD: str = _env("ADMIN_PASSWORD")
GROQ_API_KEY:   str = _env("GROQ_API_KEY")
HF_TOKEN:       str = _env("HF_TOKEN")
DB_PATH:        str = str(Path(os.getenv("DB_PATH", "./jazz.db")).resolve())

MYSQL_ENABLED: bool = bool(os.getenv("MYSQL_HOST"))
MYSQL_CONFIG: Dict[str, Any] = {
    "host":               os.getenv("MYSQL_HOST", ""),
    "port":               int(os.getenv("MYSQL_PORT", "3306")),
    "user":               os.getenv("MYSQL_USER", ""),
    "password":           os.getenv("MYSQL_PASSWORD", ""),
    "database":           os.getenv("MYSQL_DATABASE", ""),
    "ssl_disabled":       False,
    "ssl_verify_cert":    False,
    "connection_timeout": 10,
    "autocommit":         True,
}

ALLOWED_ORIGINS: List[str] = [
    o.strip() for o in os.getenv("ALLOWED_ORIGINS", "*").split(",") if o.strip()
]

CHUNK_SIZE      = 500
CHUNK_OVERLAP   = 50
TOP_K_RETRIEVAL = 4

WORKSPACE_DIR = Path("workspace")
WORKSPACE_DIR.mkdir(exist_ok=True)
SITES_DIR = WORKSPACE_DIR / "sites"
SITES_DIR.mkdir(exist_ok=True)
SITE_VERSIONS_DIR = WORKSPACE_DIR / "site_versions"
SITE_VERSIONS_DIR.mkdir(exist_ok=True)
MAX_SITE_VERSIONS = 5

ALLOWED_EXTENSIONS: set = {
    ".html", ".css", ".js", ".py", ".txt", ".md", ".json", ".xml",
    ".yaml", ".yml", ".csv", ".sql", ".sh", ".bat", ".conf", ".ini",
    ".log", ".svg", ".ts", ".jsx", ".tsx",
}

MAX_FILE_SIZE        = 10 * 1024 * 1024
MAX_IMAGE_SIZE       = 20 * 1024 * 1024
ALLOWED_IMAGE_TYPES  = {"image/jpeg", "image/png", "image/gif", "image/webp"}
PASSWORD_RESET_EXPIRY_MINUTES = 30
API_KEY_PREFIX       = "jazz_"

# ── Subscription tier limits ───────────────────────────────────────────────
SUBSCRIPTION_LIMITS: Dict[str, Dict[str, Any]] = {
    "free": {
        "messages_per_day":  50,
        "documents":         5,
        "max_file_mb":       10,
        "mysql_access":      False,
        "image_analysis":    True,
        "agent_jobs":        1,
        "websites":          3,
        "code_runs_per_day": 20,
    },
    "pro": {
        "messages_per_day":  500,
        "documents":         50,
        "max_file_mb":       50,
        "mysql_access":      True,
        "image_analysis":    True,
        "agent_jobs":        10,
        "websites":          50,
        "code_runs_per_day": 200,
    },
    "enterprise": {
        "messages_per_day":  -1,
        "documents":         -1,
        "max_file_mb":       100,
        "mysql_access":      True,
        "image_analysis":    True,
        "agent_jobs":        -1,
        "websites":          -1,
        "code_runs_per_day": -1,
    },
    "paused": {
        "messages_per_day":  0,
        "documents":         0,
        "max_file_mb":       0,
        "mysql_access":      False,
        "image_analysis":    False,
        "agent_jobs":        0,
        "websites":          0,
        "code_runs_per_day": 0,
    },
}

MAX_MEMORIES      = 100
_SERVER_START     = datetime.utcnow()
_executor         = ThreadPoolExecutor(max_workers=8)
CODE_EXEC_TIMEOUT = 10
CODE_EXEC_MAX_OUT = 50_000

_PLAYWRIGHT_AVAILABLE: bool = False
try:
    from playwright.sync_api import sync_playwright as _pw_check  # noqa: F401
    _PLAYWRIGHT_AVAILABLE = True
    logger.info("[PLAYWRIGHT] Available ✅")
except ImportError:
    logger.warning("[PLAYWRIGHT] Not installed ❌")

# ─────────────────────────────────────────────────────────────────────────────
# Security
# ─────────────────────────────────────────────────────────────────────────────

pwd_context = CryptContext(schemes=["bcrypt"], deprecated="auto")
security    = HTTPBearer(auto_error=False)

# ─────────────────────────────────────────────────────────────────────────────
# SQLite helpers
# ─────────────────────────────────────────────────────────────────────────────

def db():
    return aiosqlite.connect(DB_PATH)

async def db_execute(sql: str, params: tuple = ()) -> None:
    async with db() as conn:
        conn.row_factory = aiosqlite.Row
        await conn.execute(sql, params)
        await conn.commit()

async def db_fetchone(sql: str, params: tuple = ()) -> Optional[Dict]:
    async with db() as conn:
        conn.row_factory = aiosqlite.Row
        async with conn.execute(sql, params) as cur:
            row = await cur.fetchone()
            return dict(row) if row else None

async def db_fetchall(sql: str, params: tuple = ()) -> List[Dict]:
    async with db() as conn:
        conn.row_factory = aiosqlite.Row
        async with conn.execute(sql, params) as cur:
            rows = await cur.fetchall()
            return [dict(r) for r in rows]

async def db_insert(sql: str, params: tuple = ()) -> int:
    async with db() as conn:
        conn.row_factory = aiosqlite.Row
        async with conn.execute(sql, params) as cur:
            last_id = cur.lastrowid
        await conn.commit()
        return last_id

async def db_count(sql: str, params: tuple = ()) -> int:
    row = await db_fetchone(sql, params)
    return list(row.values())[0] if row else 0

# ─────────────────────────────────────────────────────────────────────────────
# Schema — CREATE TABLE IF NOT EXISTS (idempotent; migrations handle columns)
# ─────────────────────────────────────────────────────────────────────────────

SCHEMA_SQL = """
PRAGMA journal_mode=WAL;
PRAGMA foreign_keys=ON;

CREATE TABLE IF NOT EXISTS users (
    id                      TEXT PRIMARY KEY,
    email                   TEXT UNIQUE NOT NULL,
    password_hash           TEXT NOT NULL,
    full_name               TEXT DEFAULT '',
    role                    TEXT DEFAULT 'client',
    subscription            TEXT DEFAULT 'free',
    account_status          TEXT DEFAULT 'active',
    memory_enabled          INTEGER DEFAULT 1,
    avatar_url              TEXT DEFAULT '',
    preferred_model         TEXT DEFAULT 'uncensored',
    subscription_expires_at TEXT DEFAULT NULL,
    pause_reason            TEXT DEFAULT NULL,
    subscription_changed_by TEXT DEFAULT NULL,
    subscription_changed_at TEXT DEFAULT NULL,
    upgrade_requested       TEXT DEFAULT NULL,
    upgrade_requested_at    TEXT DEFAULT NULL,
    upgrade_note            TEXT DEFAULT NULL,
    deleted_at              TEXT DEFAULT NULL,
    created_at              TEXT DEFAULT (datetime('now')),
    updated_at              TEXT DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS subscription_audit_log (
    id               TEXT PRIMARY KEY,
    user_id          TEXT NOT NULL,
    changed_by       TEXT NOT NULL,
    changed_by_email TEXT DEFAULT '',
    action           TEXT NOT NULL,
    old_value        TEXT DEFAULT '',
    new_value        TEXT DEFAULT '',
    note             TEXT DEFAULT '',
    created_at       TEXT DEFAULT (datetime('now')),
    FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS custom_plans (
    id                TEXT PRIMARY KEY,
    name              TEXT UNIQUE NOT NULL,
    display_name      TEXT NOT NULL,
    messages_per_day  INTEGER DEFAULT 100,
    documents         INTEGER DEFAULT 10,
    max_file_mb       INTEGER DEFAULT 25,
    mysql_access      INTEGER DEFAULT 0,
    image_analysis    INTEGER DEFAULT 1,
    agent_jobs        INTEGER DEFAULT 2,
    websites          INTEGER DEFAULT 10,
    code_runs_per_day INTEGER DEFAULT 50,
    price_note        TEXT DEFAULT '',
    created_by        TEXT DEFAULT '',
    created_at        TEXT DEFAULT (datetime('now')),
    updated_at        TEXT DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS chat_sessions (
    id          TEXT PRIMARY KEY,
    user_id     TEXT NOT NULL,
    title       TEXT NOT NULL DEFAULT 'New Chat',
    model_type  TEXT DEFAULT 'uncensored',
    is_pinned   INTEGER DEFAULT 0,
    message_count INTEGER DEFAULT 0,
    last_message_at TEXT DEFAULT NULL,
    created_at  TEXT DEFAULT (datetime('now')),
    updated_at  TEXT DEFAULT (datetime('now')),
    FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS chat_history (
    id               TEXT PRIMARY KEY,
    user_id          TEXT NOT NULL,
    session_id       TEXT DEFAULT NULL,
    message          TEXT NOT NULL,
    response         TEXT NOT NULL,
    query_type       TEXT DEFAULT 'chat',
    agent_used       TEXT DEFAULT 'general_agent',
    model_used       TEXT DEFAULT 'uncensored',
    response_time_ms INTEGER DEFAULT 0,
    metadata         TEXT DEFAULT '{}',
    created_at       TEXT DEFAULT (datetime('now')),
    FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE,
    FOREIGN KEY (session_id) REFERENCES chat_sessions(id) ON DELETE SET NULL
);

CREATE TABLE IF NOT EXISTS documents (
    id                TEXT PRIMARY KEY,
    user_id           TEXT NOT NULL,
    filename          TEXT NOT NULL,
    original_name     TEXT NOT NULL,
    file_type         TEXT DEFAULT '',
    file_size         INTEGER DEFAULT 0,
    chunk_count       INTEGER DEFAULT 0,
    chroma_collection TEXT DEFAULT 'user_documents',
    uploaded_at       TEXT DEFAULT (datetime('now')),
    FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS user_memories (
    id         TEXT PRIMARY KEY,
    user_id    TEXT NOT NULL,
    content    TEXT NOT NULL,
    category   TEXT DEFAULT 'general',
    source     TEXT DEFAULT 'manual',
    created_at TEXT DEFAULT (datetime('now')),
    updated_at TEXT DEFAULT (datetime('now')),
    FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS mysql_query_logs (
    id                TEXT PRIMARY KEY,
    user_id           TEXT NOT NULL,
    natural_language  TEXT NOT NULL,
    generated_sql     TEXT NOT NULL,
    execution_success INTEGER DEFAULT 1,
    row_count         INTEGER DEFAULT 0,
    execution_time_ms INTEGER DEFAULT 0,
    created_at        TEXT DEFAULT (datetime('now')),
    FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS agent_jobs (
    id                TEXT PRIMARY KEY,
    user_id           TEXT NOT NULL,
    name              TEXT NOT NULL,
    description       TEXT DEFAULT '',
    job_type          TEXT NOT NULL,
    cron_schedule     TEXT NOT NULL,
    time_window_start TEXT DEFAULT '11:00',
    time_window_end   TEXT DEFAULT '12:00',
    credentials_enc   TEXT DEFAULT '',
    parameters        TEXT DEFAULT '{}',
    status            TEXT DEFAULT 'idle',
    enabled           INTEGER DEFAULT 1,
    last_run_at       TEXT DEFAULT NULL,
    last_run_status   TEXT DEFAULT NULL,
    next_run_at       TEXT DEFAULT NULL,
    total_runs        INTEGER DEFAULT 0,
    total_applied     INTEGER DEFAULT 0,
    created_at        TEXT DEFAULT (datetime('now')),
    updated_at        TEXT DEFAULT (datetime('now')),
    FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS agent_job_logs (
    id           TEXT PRIMARY KEY,
    job_id       TEXT NOT NULL,
    user_id      TEXT NOT NULL,
    status       TEXT DEFAULT 'running',
    output       TEXT DEFAULT '',
    applied_jobs TEXT DEFAULT '[]',
    error        TEXT DEFAULT NULL,
    started_at   TEXT DEFAULT (datetime('now')),
    completed_at TEXT DEFAULT NULL,
    duration_sec INTEGER DEFAULT 0,
    FOREIGN KEY (job_id) REFERENCES agent_jobs(id) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS websites (
    id          TEXT PRIMARY KEY,
    user_id     TEXT NOT NULL,
    title       TEXT NOT NULL,
    description TEXT DEFAULT '',
    filename    TEXT NOT NULL,
    html_size   INTEGER DEFAULT 0,
    prompt_used TEXT DEFAULT '',
    version     INTEGER DEFAULT 1,
    created_at  TEXT DEFAULT (datetime('now')),
    updated_at  TEXT DEFAULT (datetime('now')),
    FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS website_versions (
    id         TEXT PRIMARY KEY,
    site_id    TEXT NOT NULL,
    user_id    TEXT NOT NULL,
    version    INTEGER NOT NULL,
    filename   TEXT NOT NULL,
    html_size  INTEGER DEFAULT 0,
    note       TEXT DEFAULT '',
    created_at TEXT DEFAULT (datetime('now')),
    FOREIGN KEY (site_id) REFERENCES websites(id) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS prompt_templates (
    id         TEXT PRIMARY KEY,
    user_id    TEXT NOT NULL,
    title      TEXT NOT NULL,
    content    TEXT NOT NULL,
    category   TEXT DEFAULT 'general',
    is_public  INTEGER DEFAULT 0,
    use_count  INTEGER DEFAULT 0,
    created_at TEXT DEFAULT (datetime('now')),
    FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS code_runs (
    id          TEXT PRIMARY KEY,
    user_id     TEXT NOT NULL,
    language    TEXT DEFAULT 'python',
    code        TEXT NOT NULL,
    stdout      TEXT DEFAULT '',
    stderr      TEXT DEFAULT '',
    exit_code   INTEGER DEFAULT 0,
    duration_ms INTEGER DEFAULT 0,
    created_at  TEXT DEFAULT (datetime('now')),
    FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS api_keys (
    id          TEXT PRIMARY KEY,
    user_id     TEXT NOT NULL,
    name        TEXT NOT NULL,
    key_hash    TEXT NOT NULL,
    key_prefix  TEXT NOT NULL,
    last_used_at TEXT DEFAULT NULL,
    expires_at  TEXT DEFAULT NULL,
    is_active   INTEGER DEFAULT 1,
    created_at  TEXT DEFAULT (datetime('now')),
    FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS system_announcements (
    id           TEXT PRIMARY KEY,
    title        TEXT NOT NULL,
    body         TEXT NOT NULL,
    type         TEXT DEFAULT 'info',
    target       TEXT DEFAULT 'all',
    is_active    INTEGER DEFAULT 1,
    created_by   TEXT DEFAULT '',
    expires_at   TEXT DEFAULT NULL,
    created_at   TEXT DEFAULT (datetime('now')),
    updated_at   TEXT DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS user_notifications (
    id         TEXT PRIMARY KEY,
    user_id    TEXT NOT NULL,
    title      TEXT NOT NULL,
    body       TEXT NOT NULL,
    type       TEXT DEFAULT 'info',
    is_read    INTEGER DEFAULT 0,
    action_url TEXT DEFAULT '',
    created_at TEXT DEFAULT (datetime('now')),
    FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS user_activity_log (
    id         TEXT PRIMARY KEY,
    user_id    TEXT NOT NULL,
    action     TEXT NOT NULL,
    resource   TEXT DEFAULT '',
    detail     TEXT DEFAULT '',
    ip_address TEXT DEFAULT '',
    user_agent TEXT DEFAULT '',
    created_at TEXT DEFAULT (datetime('now')),
    FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS password_reset_tokens (
    id         TEXT PRIMARY KEY,
    user_id    TEXT NOT NULL,
    token_hash TEXT NOT NULL,
    expires_at TEXT NOT NULL,
    used       INTEGER DEFAULT 0,
    created_at TEXT DEFAULT (datetime('now')),
    FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS health_history (
    id         TEXT PRIMARY KEY,
    ts         TEXT NOT NULL,
    sqlite_ok  INTEGER DEFAULT 1,
    chroma_ok  INTEGER DEFAULT 1,
    mysql_ok   INTEGER DEFAULT 0,
    uptime_sec INTEGER DEFAULT 0,
    created_at TEXT DEFAULT (datetime('now'))
);

-- Indexes
CREATE INDEX IF NOT EXISTS idx_chat_user       ON chat_history(user_id, created_at);
CREATE INDEX IF NOT EXISTS idx_chat_session    ON chat_history(session_id, created_at);
CREATE INDEX IF NOT EXISTS idx_docs_user       ON documents(user_id);
CREATE INDEX IF NOT EXISTS idx_mem_user        ON user_memories(user_id);
CREATE INDEX IF NOT EXISTS idx_ajob_user       ON agent_jobs(user_id);
CREATE INDEX IF NOT EXISTS idx_ajlog_job       ON agent_job_logs(job_id, started_at);
CREATE INDEX IF NOT EXISTS idx_sites_user      ON websites(user_id);
CREATE INDEX IF NOT EXISTS idx_site_ver        ON website_versions(site_id, version);
CREATE INDEX IF NOT EXISTS idx_prompts_user    ON prompt_templates(user_id);
CREATE INDEX IF NOT EXISTS idx_code_user       ON code_runs(user_id);
CREATE INDEX IF NOT EXISTS idx_sub_audit       ON subscription_audit_log(user_id, created_at);
CREATE INDEX IF NOT EXISTS idx_sub_expires     ON users(subscription_expires_at);
CREATE INDEX IF NOT EXISTS idx_acct_status     ON users(account_status);
CREATE INDEX IF NOT EXISTS idx_api_keys_user   ON api_keys(user_id);
CREATE INDEX IF NOT EXISTS idx_notif_user      ON user_notifications(user_id, is_read);
CREATE INDEX IF NOT EXISTS idx_activity_user   ON user_activity_log(user_id, created_at);
CREATE INDEX IF NOT EXISTS idx_sessions_user   ON chat_sessions(user_id, updated_at);
CREATE INDEX IF NOT EXISTS idx_chat_fts_msg    ON chat_history(user_id, message);
"""

# ─────────────────────────────────────────────────────────────────────────────
# Schema Migration System — THE CORE BUG FIX
# Adds missing columns to existing tables without losing any data.
# SQLite doesn't support ALTER TABLE ADD COLUMN IF NOT EXISTS, so we wrap
# each ALTER in a try/except — it will silently skip already-existing columns.
# ─────────────────────────────────────────────────────────────────────────────

# Format: (table_name, column_name, column_definition)
MIGRATIONS: List[Tuple[str, str, str]] = [
    # users — v7.0 new columns
    ("users", "subscription_expires_at", "TEXT DEFAULT NULL"),
    ("users", "pause_reason",            "TEXT DEFAULT NULL"),
    ("users", "subscription_changed_by", "TEXT DEFAULT NULL"),
    ("users", "subscription_changed_at", "TEXT DEFAULT NULL"),
    ("users", "upgrade_requested",       "TEXT DEFAULT NULL"),
    ("users", "upgrade_requested_at",    "TEXT DEFAULT NULL"),
    ("users", "upgrade_note",            "TEXT DEFAULT NULL"),
    ("users", "account_status",          "TEXT DEFAULT 'active'"),
    # users — v7.1 new columns
    ("users", "deleted_at",              "TEXT DEFAULT NULL"),
    # websites — v7.1 versioning
    ("websites", "version",              "INTEGER DEFAULT 1"),
    # chat_history — v7.1 sessions
    ("chat_history", "session_id",       "TEXT DEFAULT NULL"),
]

async def run_migrations() -> None:
    """
    Idempotent column migrations. Safe to run on every startup.
    Skips columns that already exist — no data loss ever.
    """
    import sqlite3 as _s3

    applied = 0
    for table, column, col_def in MIGRATIONS:
        try:
            with _s3.connect(DB_PATH) as conn:
                conn.execute(f"ALTER TABLE {table} ADD COLUMN {column} {col_def}")
                conn.commit()
            logger.info("[MIGRATION] + %s.%s", table, column)
            applied += 1
        except _s3.OperationalError as exc:
            if "duplicate column name" in str(exc).lower():
                pass  # Already exists — skip silently
            else:
                logger.warning("[MIGRATION] %s.%s: %s", table, column, exc)

    if applied:
        logger.info("[MIGRATION] Applied %d column(s) ✅", applied)
    else:
        logger.info("[MIGRATION] Schema up-to-date ✅")


async def init_db() -> None:
    # 1. Create all tables (idempotent)
    async with db() as conn:
        await conn.executescript(SCHEMA_SQL)
        await conn.commit()

    # 2. Run column migrations for existing tables
    await run_migrations()

    # 3. Reset stuck running jobs
    await db_execute(
        "UPDATE agent_jobs SET status='idle', updated_at=? WHERE status='running'",
        (datetime.utcnow().isoformat(),),
    )

    # 4. Create admin user if needed
    existing = await db_fetchone("SELECT id FROM users WHERE email=?", (ADMIN_EMAIL,))
    if not existing and ADMIN_PASSWORD:
        uid   = str(uuid.uuid4())
        phash = pwd_context.hash(ADMIN_PASSWORD)
        await db_execute(
            "INSERT INTO users (id,email,password_hash,full_name,role,subscription,account_status) VALUES (?,?,?,?,?,?,?)",
            (uid, ADMIN_EMAIL, phash, "Admin", "admin", "enterprise", "active"),
        )
        logger.info("[SQLITE] Admin user created: %s", ADMIN_EMAIL)

    logger.info("[SQLITE] Initialized ✅  path=%s", DB_PATH)

# ─────────────────────────────────────────────────────────────────────────────
# ChromaDB
# ─────────────────────────────────────────────────────────────────────────────

chroma_client:            Optional[chromadb.PersistentClient] = None
documents_collection:     Any = None
sqlite_schema_collection: Any = None
_embedding_fn:            Any = None

try:
    _embedding_fn = embedding_functions.SentenceTransformerEmbeddingFunction(
        model_name="all-MiniLM-L6-v2"
    )
    chroma_client = chromadb.PersistentClient(path="./chroma_db")
    documents_collection = chroma_client.get_or_create_collection(
        name="user_documents", embedding_function=_embedding_fn
    )
    sqlite_schema_collection = chroma_client.get_or_create_collection(
        name="sqlite_schema", embedding_function=_embedding_fn
    )
    logger.info("[CHROMADB] Initialized ✅")
except Exception as exc:
    logger.error("[CHROMADB] Init failed: %s ❌", exc)

# ─────────────────────────────────────────────────────────────────────────────
# MySQL pool
# ─────────────────────────────────────────────────────────────────────────────

_mysql_pool: Any = None

def _init_mysql_pool() -> None:
    global _mysql_pool
    if not MYSQL_ENABLED:
        return
    try:
        from mysql.connector import pooling as mp
        _mysql_pool = mp.MySQLConnectionPool(
            pool_name="jazz_pool", pool_size=5, pool_reset_session=True, **MYSQL_CONFIG
        )
        logger.info("[MYSQL] Pool created ✅")
    except Exception as exc:
        logger.error("[MYSQL] Pool failed: %s ❌", exc)

def get_mysql_connection() -> Any:
    global _mysql_pool
    if _mysql_pool is None:
        _init_mysql_pool()
    if _mysql_pool is None:
        return None
    try:
        return _mysql_pool.get_connection()
    except Exception:
        _init_mysql_pool()
        try:
            return _mysql_pool.get_connection() if _mysql_pool else None
        except Exception:
            return None

# ─────────────────────────────────────────────────────────────────────────────
# Rate limiter
# ─────────────────────────────────────────────────────────────────────────────

class RateLimiter:
    _WINDOW_SECONDS = 86_400

    def __init__(self) -> None:
        self._window: Dict[str, List[float]] = defaultdict(list)

    def _cleanup(self, key: str) -> None:
        now = time.time()
        self._window[key] = [t for t in self._window[key] if now - t < self._WINDOW_SECONDS]

    def check(self, uid: str, tier: str, resource: str = "messages_per_day") -> bool:
        limits = _get_effective_limits(tier)
        limit  = limits.get(resource, 50)
        if limit == -1:
            return True
        if limit == 0:
            return False
        key = f"{uid}:{resource}"
        self._cleanup(key)
        if len(self._window[key]) >= limit:
            return False
        self._window[key].append(time.time())
        return True

    def usage_today(self, uid: str, resource: str = "messages_per_day") -> int:
        key = f"{uid}:{resource}"
        self._cleanup(key)
        return len(self._window[key])

rate_limiter = RateLimiter()

# ─────────────────────────────────────────────────────────────────────────────
# Subscription helpers
# ─────────────────────────────────────────────────────────────────────────────

def _get_effective_limits(tier: str) -> Dict[str, Any]:
    if tier in SUBSCRIPTION_LIMITS:
        return SUBSCRIPTION_LIMITS[tier]
    import sqlite3 as _s3
    try:
        with _s3.connect(DB_PATH) as conn:
            conn.row_factory = _s3.Row
            row = conn.execute("SELECT * FROM custom_plans WHERE name=?", (tier,)).fetchone()
            if row:
                d = dict(row)
                return {
                    "messages_per_day":  d["messages_per_day"],
                    "documents":         d["documents"],
                    "max_file_mb":       d["max_file_mb"],
                    "mysql_access":      bool(d["mysql_access"]),
                    "image_analysis":    bool(d["image_analysis"]),
                    "agent_jobs":        d["agent_jobs"],
                    "websites":          d["websites"],
                    "code_runs_per_day": d["code_runs_per_day"],
                }
    except Exception as exc:
        logger.error("[LIMITS] Custom plan lookup failed: %s", exc)
    return SUBSCRIPTION_LIMITS["free"]

async def _get_effective_limits_async(tier: str) -> Dict[str, Any]:
    if tier in SUBSCRIPTION_LIMITS:
        return SUBSCRIPTION_LIMITS[tier]
    row = await db_fetchone("SELECT * FROM custom_plans WHERE name=?", (tier,))
    if row:
        return {
            "messages_per_day":  row["messages_per_day"],
            "documents":         row["documents"],
            "max_file_mb":       row["max_file_mb"],
            "mysql_access":      bool(row["mysql_access"]),
            "image_analysis":    bool(row["image_analysis"]),
            "agent_jobs":        row["agent_jobs"],
            "websites":          row["websites"],
            "code_runs_per_day": row["code_runs_per_day"],
        }
    return SUBSCRIPTION_LIMITS["free"]

async def get_user_subscription(user_id: str) -> str:
    row = await db_fetchone(
        "SELECT subscription, account_status, subscription_expires_at FROM users WHERE id=?",
        (user_id,),
    )
    if not row:
        return "free"
    if row["account_status"] in ("paused", "banned"):
        return "paused"
    expires = row.get("subscription_expires_at")
    if expires:
        try:
            if datetime.utcnow() > datetime.fromisoformat(expires):
                await _auto_expire_subscription(user_id, row["subscription"])
                return "free"
        except Exception:
            pass
    return row["subscription"] or "free"

async def _auto_expire_subscription(user_id: str, old_tier: str) -> None:
    now = datetime.utcnow().isoformat()
    await db_execute(
        "UPDATE users SET subscription='free', subscription_expires_at=NULL, "
        "subscription_changed_by='system', subscription_changed_at=?, updated_at=? WHERE id=?",
        (now, now, user_id),
    )
    await _write_audit_log(
        user_id=user_id, changed_by="system", changed_by_email="system",
        action="expire", old_value=old_tier, new_value="free",
        note="Subscription expired — auto-downgraded to free",
    )
    logger.info("[SUBSCRIPTION] User %s auto-expired from %s to free", user_id, old_tier)

async def _write_audit_log(
    user_id: str, changed_by: str, changed_by_email: str,
    action: str, old_value: str, new_value: str, note: str = "",
) -> None:
    await db_execute(
        "INSERT INTO subscription_audit_log (id,user_id,changed_by,changed_by_email,action,old_value,new_value,note) VALUES (?,?,?,?,?,?,?,?)",
        (str(uuid.uuid4()), user_id, changed_by, changed_by_email, action, old_value, new_value, note),
    )

async def is_memory_enabled(user_id: str) -> bool:
    row = await db_fetchone("SELECT memory_enabled FROM users WHERE id=?", (user_id,))
    return bool((row or {}).get("memory_enabled", 1))

async def check_account_active(user_id: str) -> None:
    row    = await db_fetchone("SELECT account_status FROM users WHERE id=?", (user_id,))
    status = (row or {}).get("account_status", "active")
    if status == "paused":
        raise HTTPException(403, "Your account is paused. Contact support.")
    if status == "banned":
        raise HTTPException(403, "Your account has been suspended.")

# ─────────────────────────────────────────────────────────────────────────────
# Notification helpers
# ─────────────────────────────────────────────────────────────────────────────

async def send_notification(user_id: str, title: str, body: str, notif_type: str = "info", action_url: str = "") -> None:
    await db_execute(
        "INSERT INTO user_notifications (id,user_id,title,body,type,action_url) VALUES (?,?,?,?,?,?)",
        (str(uuid.uuid4()), user_id, title, body, notif_type, action_url),
    )

async def log_activity(user_id: str, action: str, resource: str = "", detail: str = "", ip: str = "", ua: str = "") -> None:
    await db_execute(
        "INSERT INTO user_activity_log (id,user_id,action,resource,detail,ip_address,user_agent) VALUES (?,?,?,?,?,?,?)",
        (str(uuid.uuid4()), user_id, action, resource, detail[:500], ip[:100], ua[:200]),
    )

# ─────────────────────────────────────────────────────────────────────────────
# Pydantic models
# ─────────────────────────────────────────────────────────────────────────────

class ChatRequest(BaseModel):
    message:    str           = Field(..., min_length=1, max_length=8000)
    model_type: str           = Field("uncensored", pattern="^(censored|uncensored)$")
    use_rag:    bool          = True
    image_url:  Optional[str] = None
    session_id: Optional[str] = None

class ChatResponse(BaseModel):
    response:       str
    response_time:  float
    agent_used:     str
    tokens_used:    Optional[int] = None
    memory_updated: bool          = False
    new_memories:   List[str]     = []
    session_id:     Optional[str] = None

class UserSignup(BaseModel):
    email:     str           = Field(..., pattern=r"^[^@]+@[^@]+\.[^@]+$")
    password:  str           = Field(..., min_length=8)
    full_name: Optional[str] = None

class UserLogin(BaseModel):
    email:    str
    password: str

class TokenResponse(BaseModel):
    access_token: str
    token_type:   str = "bearer"
    user:         Dict[str, Any]

class ProfileUpdate(BaseModel):
    full_name:       Optional[str]  = None
    avatar_url:      Optional[str]  = None
    preferred_model: Optional[str]  = Field(None, pattern="^(censored|uncensored)$")
    memory_enabled:  Optional[bool] = None

class PasswordChangeRequest(BaseModel):
    current_password: str = Field(..., min_length=1)
    new_password:     str = Field(..., min_length=8)

class PasswordResetRequest(BaseModel):
    email: str = Field(..., pattern=r"^[^@]+@[^@]+\.[^@]+$")

class PasswordResetConfirm(BaseModel):
    token:        str = Field(..., min_length=10)
    new_password: str = Field(..., min_length=8)

class DocumentResponse(BaseModel):
    id:            str
    filename:      str
    original_name: str
    file_type:     str
    file_size:     int
    chunk_count:   int
    uploaded_at:   str

class MemoryCreate(BaseModel):
    content:  str           = Field(..., min_length=1, max_length=500)
    category: Optional[str] = "general"

class MemoryUpdate(BaseModel):
    content: str = Field(..., min_length=1, max_length=500)

class MemorySettingsUpdate(BaseModel):
    enabled: bool

class FileOperationRequest(BaseModel):
    filename:  str
    content:   Optional[str] = None
    operation: str           = Field(..., pattern="^(create|read|update|delete|list)$")

# Chat session models
class SessionCreate(BaseModel):
    title:      str           = Field("New Chat", min_length=1, max_length=200)
    model_type: Optional[str] = Field("uncensored", pattern="^(censored|uncensored)$")

class SessionUpdate(BaseModel):
    title:     Optional[str]  = Field(None, min_length=1, max_length=200)
    is_pinned: Optional[bool] = None

# API Key models
class APIKeyCreate(BaseModel):
    name:       str           = Field(..., min_length=1, max_length=100)
    expires_at: Optional[str] = None  # ISO date or None for non-expiring

# Announcement models
class AnnouncementCreate(BaseModel):
    title:      str            = Field(..., min_length=1, max_length=200)
    body:       str            = Field(..., min_length=1, max_length=2000)
    type:       str            = Field("info", pattern="^(info|warning|success|danger)$")
    target:     str            = Field("all", pattern="^(all|free|pro|enterprise)$")
    is_active:  bool           = True
    expires_at: Optional[str]  = None

# Notification models
class NotificationMarkRead(BaseModel):
    notification_ids: Optional[List[str]] = None  # None = mark all

# Subscription management models
class SubscriptionUpdate(BaseModel):
    user_id:    str
    tier:       str
    expires_at: Optional[str]  = None
    note:       Optional[str]  = ""

class BulkSubscriptionUpdate(BaseModel):
    user_ids:   List[str]
    tier:       str
    expires_at: Optional[str] = None
    note:       Optional[str] = ""

class AccountStatusUpdate(BaseModel):
    user_id: str
    status:  str = Field(..., pattern="^(active|paused|banned)$")
    reason:  Optional[str] = ""

class CustomPlanCreate(BaseModel):
    name:               str   = Field(..., min_length=2, max_length=40, pattern=r"^[a-z0-9_]+$")
    display_name:       str   = Field(..., min_length=2, max_length=80)
    messages_per_day:   int   = Field(100, ge=-1)
    documents:          int   = Field(10,  ge=-1)
    max_file_mb:        int   = Field(25,  ge=1)
    mysql_access:       bool  = False
    image_analysis:     bool  = True
    agent_jobs:         int   = Field(2,   ge=-1)
    websites:           int   = Field(10,  ge=-1)
    code_runs_per_day:  int   = Field(50,  ge=-1)
    price_note:         str   = ""

class CustomPlanUpdate(BaseModel):
    display_name:       Optional[str]  = None
    messages_per_day:   Optional[int]  = None
    documents:          Optional[int]  = None
    max_file_mb:        Optional[int]  = None
    mysql_access:       Optional[bool] = None
    image_analysis:     Optional[bool] = None
    agent_jobs:         Optional[int]  = None
    websites:           Optional[int]  = None
    code_runs_per_day:  Optional[int]  = None
    price_note:         Optional[str]  = None

class UpgradeRequest(BaseModel):
    requested_tier: str = Field(..., pattern="^(free|pro|enterprise)$")
    note:           Optional[str] = ""

class UpgradeDecision(BaseModel):
    user_id:  str
    approve:  bool
    note:     Optional[str] = ""

class AgentJobCreate(BaseModel):
    name:               str                      = Field(..., min_length=1, max_length=100)
    description:        Optional[str]            = ""
    job_type:           str                      = Field(..., pattern="^(naukri_apply|linkedin_apply|custom)$")
    cron_schedule:      str
    time_window_start:  Optional[str]            = "11:00"
    time_window_end:    Optional[str]            = "12:00"
    credentials:        Dict[str, Any]
    parameters:         Optional[Dict[str, Any]] = {}

class AgentJobUpdate(BaseModel):
    name:              Optional[str]            = None
    description:       Optional[str]            = None
    cron_schedule:     Optional[str]            = None
    time_window_start: Optional[str]            = None
    time_window_end:   Optional[str]            = None
    credentials:       Optional[Dict[str, Any]] = None
    parameters:        Optional[Dict[str, Any]] = None
    enabled:           Optional[bool]           = None

class WebsiteBuildRequest(BaseModel):
    description: str           = Field(..., min_length=10, max_length=2000)
    title:       Optional[str] = None
    style:       Optional[str] = "modern"
    model_type:  str           = Field("censored", pattern="^(censored|uncensored)$")

class WebsiteUpdateRequest(BaseModel):
    instructions: str      = Field(..., min_length=5, max_length=1000)
    model_type:   str      = Field("censored", pattern="^(censored|uncensored)$")

class WebsiteRestoreRequest(BaseModel):
    version: int = Field(..., ge=1)

class CodeRunRequest(BaseModel):
    code:     str = Field(..., min_length=1, max_length=20_000)
    language: str = Field("python", pattern="^(python|javascript|bash)$")

class SQLiteQueryRequest(BaseModel):
    question: str = Field(..., min_length=1, max_length=500)

class SQLiteRawRequest(BaseModel):
    sql:      str  = Field(..., min_length=1, max_length=2000)
    readonly: bool = True

class PromptCreate(BaseModel):
    title:     str           = Field(..., min_length=1, max_length=100)
    content:   str           = Field(..., min_length=1, max_length=5000)
    category:  Optional[str] = "general"
    is_public: bool          = False

class PromptUpdate(BaseModel):
    title:     Optional[str]  = None
    content:   Optional[str]  = None
    category:  Optional[str]  = None
    is_public: Optional[bool] = None

class MySQLQueryRequest(BaseModel):
    question: str = Field(..., min_length=1, max_length=1000)

class MySQLQueryResponse(BaseModel):
    success:           bool
    question:          Optional[str]        = None
    generated_sql:     Optional[str]        = None
    results:           List[Dict[str, Any]] = []
    row_count:         int                  = 0
    execution_time_ms: int                  = 0
    error:             Optional[str]        = None

class BulkDocumentDelete(BaseModel):
    document_ids: List[str] = Field(..., min_length=1)

# ─────────────────────────────────────────────────────────────────────────────
# Auth helpers
# ─────────────────────────────────────────────────────────────────────────────

def create_access_token(data: Dict, expires_delta: Optional[timedelta] = None) -> str:
    to_encode = data.copy()
    to_encode["exp"] = datetime.utcnow() + (
        expires_delta or timedelta(minutes=ACCESS_TOKEN_EXPIRE_MINUTES)
    )
    return jwt.encode(to_encode, SECRET_KEY, algorithm=ALGORITHM)

def verify_token(token: str) -> Optional[Dict]:
    try:
        return jwt.decode(token, SECRET_KEY, algorithms=[ALGORITHM])
    except JWTError:
        return None

async def _verify_api_key(raw_key: str) -> Optional[Dict]:
    """Verify a raw API key and return the user dict, or None if invalid."""
    if not raw_key.startswith(API_KEY_PREFIX):
        return None
    key_hash = pwd_context.hash(raw_key)
    # We can't reverse bcrypt, so we store a fast SHA-256 prefix hash
    import hashlib
    fast_hash = hashlib.sha256(raw_key.encode()).hexdigest()
    row = await db_fetchone(
        "SELECT ak.*, u.role, u.account_status, u.subscription FROM api_keys ak "
        "JOIN users u ON ak.user_id=u.id "
        "WHERE ak.key_hash=? AND ak.is_active=1",
        (fast_hash,),
    )
    if not row:
        return None
    if row.get("expires_at") and datetime.utcnow() > datetime.fromisoformat(row["expires_at"]):
        return None
    # Update last_used_at async (fire-and-forget)
    asyncio.create_task(db_execute(
        "UPDATE api_keys SET last_used_at=? WHERE id=?",
        (datetime.utcnow().isoformat(), row["id"]),
    ))
    return {"sub": row["user_id"], "email": "", "role": row["role"], "via_api_key": True}

async def get_current_user(
    credentials: Optional[HTTPAuthorizationCredentials] = Depends(security),
) -> Dict:
    if not credentials:
        raise HTTPException(401, "Authorization header missing", headers={"WWW-Authenticate": "Bearer"})
    token = credentials.credentials
    # Try API key first
    if token.startswith(API_KEY_PREFIX):
        payload = await _verify_api_key(token)
        if not payload:
            raise HTTPException(401, "Invalid or expired API key")
        return payload
    # Then JWT
    payload = verify_token(token)
    if not payload:
        raise HTTPException(401, "Invalid or expired token", headers={"WWW-Authenticate": "Bearer"})
    return payload

async def require_admin(
    credentials: Optional[HTTPAuthorizationCredentials] = Depends(security),
) -> Dict:
    if not credentials:
        raise HTTPException(401, "Authorization header missing")
    payload = verify_token(credentials.credentials)
    if not payload:
        raise HTTPException(401, "Invalid or expired token")
    if payload.get("role") != "admin":
        raise HTTPException(403, "Admin access required")
    return payload

# ─────────────────────────────────────────────────────────────────────────────
# Memory system
# ─────────────────────────────────────────────────────────────────────────────

async def get_user_memories(user_id: str) -> List[Dict]:
    return await db_fetchall(
        "SELECT * FROM user_memories WHERE user_id=? ORDER BY created_at ASC", (user_id,)
    )

def format_memories_for_prompt(memories: List[Dict]) -> str:
    if not memories:
        return ""
    lines = ["[User Memory:]"]
    for m in memories:
        lines.append(f"• ({m.get('category','general')}) {m['content']}")
    return "\n".join(lines)

def _sync_extract_memories(user_id: str, user_message: str, ai_response: str) -> List[str]:
    import sqlite3 as _s3

    def _count():
        with _s3.connect(DB_PATH) as c:
            r = c.execute("SELECT COUNT(*) FROM user_memories WHERE user_id=?", (user_id,)).fetchone()
            return r[0] if r else 0

    def _existing():
        with _s3.connect(DB_PATH) as c:
            return [r[0].lower() for r in c.execute("SELECT content FROM user_memories WHERE user_id=?", (user_id,)).fetchall()]

    def _insert(mid, content, category):
        with _s3.connect(DB_PATH) as c:
            c.execute("INSERT INTO user_memories (id,user_id,content,category,source) VALUES (?,?,?,?,?)",
                      (mid, user_id, content, category, "auto"))
            c.commit()

    def _mem_enabled():
        with _s3.connect(DB_PATH) as c:
            r = c.execute("SELECT memory_enabled FROM users WHERE id=?", (user_id,)).fetchone()
            return bool(r[0]) if r else True

    try:
        if not _mem_enabled() or _count() >= MAX_MEMORIES:
            return []
        client, model = get_model_client("censored")
        prompt = (
            f"Extract 0-3 memorable facts about the USER.\n"
            f"User: {user_message}\nAI: {ai_response[:500]}\n"
            'Return ONLY JSON: [{"content":"...","category":"preference|profession|goal|personal|technical|general"}]\n'
            "If nothing memorable: []"
        )
        comp      = client.chat.completions.create(model=model, messages=[{"role":"user","content":prompt}], max_tokens=300, temperature=0.1)
        raw       = re.sub(r"```json|```", "", comp.choices[0].message.content.strip()).strip()
        extracted = json.loads(raw)
        if not isinstance(extracted, list):
            return []
        existing      = _existing()
        new_memories: List[str] = []
        for item in extracted[:3]:
            if not isinstance(item, dict) or not item.get("content"):
                continue
            content  = item["content"].strip()[:500]
            category = item.get("category", "general")
            if any(content.lower() in e or e in content.lower() for e in existing):
                continue
            _insert(str(uuid.uuid4()), content, category)
            new_memories.append(content)
        return new_memories
    except Exception as exc:
        logger.error("[MEMORY] %s", exc)
        return []

async def extract_and_save_memories(user_id: str, user_message: str, ai_response: str) -> List[str]:
    loop = asyncio.get_event_loop()
    return await loop.run_in_executor(_executor, _sync_extract_memories, user_id, user_message, ai_response)

# ─────────────────────────────────────────────────────────────────────────────
# Document processing / RAG
# ─────────────────────────────────────────────────────────────────────────────

def chunk_text(text: str, size=CHUNK_SIZE, overlap=CHUNK_OVERLAP) -> List[str]:
    chunks, start = [], 0
    while start < len(text):
        chunk = text[start:start+size]
        if chunk.strip():
            chunks.append(chunk)
        start += size - overlap
    return chunks

def extract_text_from_pdf(path):
    try:
        with fitz.open(path) as pdf:
            return "\n".join(page.get_text() for page in pdf)
    except Exception as exc:
        logger.error("[PDF] %s", exc); return ""

def extract_text_from_docx(path):
    try:
        doc = DocxDocument(path)
        return "\n".join(p.text for p in doc.paragraphs)
    except Exception as exc:
        logger.error("[DOCX] %s", exc); return ""

def extract_text_from_excel(path):
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
        parts = []
        for sheet in wb.worksheets:
            parts.append(f"\nSheet: {sheet.title}")
            for row in sheet.iter_rows(values_only=True):
                row_str = " | ".join(str(c) for c in row if c is not None)
                if row_str.strip():
                    parts.append(row_str)
        return "\n".join(parts)
    except Exception as exc:
        logger.error("[EXCEL] %s", exc); return ""

def process_document(path, mime):
    if "pdf"  in mime: return extract_text_from_pdf(path)
    if "word" in mime or "docx" in mime: return extract_text_from_docx(path)
    if "excel" in mime or "sheet" in mime: return extract_text_from_excel(path)
    if "csv" in mime:
        try:
            with open(path, "r", encoding="utf-8") as f:
                return "\n".join(" | ".join(row) for row in csv.reader(f))
        except Exception: return ""
    try: return Path(path).read_text(encoding="utf-8")
    except Exception: return ""

def get_relevant_context(user_id, query, top_k=TOP_K_RETRIEVAL):
    if documents_collection is None: return ""
    try:
        results = documents_collection.query(query_texts=[query], n_results=top_k, where={"user_id": user_id})
        if not results or not results["documents"] or not results["documents"][0]: return ""
        parts = []
        for i, doc in enumerate(results["documents"][0]):
            meta = results["metadatas"][0][i] if results.get("metadatas") else {}
            parts.append(f"[Source: {meta.get('filename','Unknown')}]\n{doc}")
        return "\n\n---\n\n".join(parts)
    except Exception as exc:
        logger.error("[RAG] %s", exc); return ""

_last_reindex_at = None

def index_sqlite_schema():
    global _last_reindex_at
    if sqlite_schema_collection is None: return False
    import sqlite3 as _s3
    try:
        with _s3.connect(DB_PATH) as conn:
            tables = [r[0] for r in conn.execute("SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%'").fetchall()]
        existing = sqlite_schema_collection.get()
        if existing["ids"]:
            sqlite_schema_collection.delete(ids=existing["ids"])
        chunks, ids, metas = [], [], []
        with _s3.connect(DB_PATH) as conn:
            for tbl in tables:
                try:
                    cols    = conn.execute(f"PRAGMA table_info(`{tbl}`)").fetchall()
                    samples = conn.execute(f"SELECT * FROM `{tbl}` LIMIT 3").fetchall()
                    col_info = ", ".join(f"{c[1]}({c[2]})" + (",PK" if c[5] else "") for c in cols)
                    text = f"Table: {tbl} | Columns: {col_info}"
                    if samples:
                        text += f" | Samples: {json.dumps([list(r) for r in samples], default=str)}"
                    chunks.append(text); ids.append(f"tbl_{tbl}"); metas.append({"table_name": tbl, "type": "schema"})
                except Exception as exc:
                    logger.warning("[SQLITE-SCHEMA] Skipping %s: %s", tbl, exc)
        if chunks:
            sqlite_schema_collection.add(documents=chunks, ids=ids, metadatas=metas)
        _last_reindex_at = datetime.utcnow().isoformat()
        logger.info("[SQLITE] Schema indexed — %d tables ✅", len(chunks))
        return True
    except Exception as exc:
        logger.error("[SQLITE-SCHEMA] %s", exc); return False

def get_schema_context(query, top_k=3):
    if sqlite_schema_collection is None: return ""
    try:
        results = sqlite_schema_collection.query(query_texts=[query], n_results=top_k)
        if results and results["documents"] and results["documents"][0]:
            return "\n\n".join(results["documents"][0])
        return ""
    except Exception as exc:
        logger.error("[SQLITE-SCHEMA] %s", exc); return ""

_DB_KEYWORDS = {"table","column","row","select","count","show","describe","database","schema","query","fetch","list","find","get","records","how many","what is in","show me","tell me about","get me","find all","data from"}

def is_db_intent(message):
    low = message.lower()
    return any(kw in low for kw in _DB_KEYWORDS)

# ─────────────────────────────────────────────────────────────────────────────
# AI clients
# ─────────────────────────────────────────────────────────────────────────────

def get_model_client(model_type: str) -> Tuple[OpenAI, str]:
    if model_type == "censored":
        return OpenAI(base_url="https://api.groq.com/openai/v1", api_key=GROQ_API_KEY), "llama-3.3-70b-versatile"
    return OpenAI(base_url="https://router.huggingface.co/v1", api_key=HF_TOKEN), "dphn/Dolphin-Mistral-24B-Venice-Edition:featherless-ai"

def get_vision_client() -> Tuple[OpenAI, str]:
    return OpenAI(base_url="https://router.huggingface.co/v1", api_key=HF_TOKEN), "Qwen/Qwen2.5-VL-7B-Instruct:fastest"

async def analyze_image_pipeline(image_data, user_message, model_type, user_id, memories):
    t0 = time.time()
    vc, vm = get_vision_client()
    vision_prompt = (
        'Analyze this image and return JSON:\n'
        '{"description":"...","objects":[],"text_in_image":null,"scene_type":"photo/diagram/screenshot/art/chart/other",'
        f'"technical_details":"...","user_question_relevance":"re: {user_message}"}}'
    )
    vision_json = {}
    try:
        vcomp = vc.chat.completions.create(
            model=vm,
            messages=[{"role":"user","content":[{"type":"text","text":vision_prompt},{"type":"image_url","image_url":{"url":image_data}}]}],
            max_tokens=600, temperature=0.1,
        )
        raw         = re.sub(r"```json|```","",vcomp.choices[0].message.content.strip()).strip()
        vision_json = json.loads(raw)
    except Exception as exc:
        logger.error("[VISION] %s", exc)
        vision_json = {"description":"Image provided","scene_type":"unknown","user_question_relevance":user_message}
    ai_client, ai_model = get_model_client(model_type)
    memory_block        = format_memories_for_prompt(memories)
    combined_prompt = (
        f"{memory_block}\n## Image Analysis:\n{json.dumps(vision_json,indent=2)}\n## User:\n{user_message}\nProvide a thorough response."
    )
    comp  = ai_client.chat.completions.create(model=ai_model, messages=[{"role":"user","content":combined_prompt}], max_tokens=1024, temperature=0.7)
    final = comp.choices[0].message.content or ""
    logger.info("[VISION] Pipeline %.2fs", time.time()-t0)
    return final, json.dumps(vision_json)

# ─────────────────────────────────────────────────────────────────────────────
# Agent routing
# ─────────────────────────────────────────────────────────────────────────────

class Agent:
    def __init__(self, name, system_prompt, keywords):
        self.name = name; self.system_prompt = system_prompt; self.keywords = keywords
    def build_prompt(self, message, context="", memory_block=""):
        parts = [self.system_prompt]
        if memory_block: parts.append(f"\n{memory_block}\n")
        if context:      parts.append(f"\nRelevant context:\n{context}\n")
        parts.append(f"\nUser: {message}")
        if context:      parts.append("\nAnswer using the context where relevant.")
        return "\n".join(parts)

AGENTS: List[Agent] = [
    Agent("coding_agent","You are a senior software engineer. Write clean, efficient, well-documented code. Always include error handling and explain your approach briefly.",
          ["code","python","script","program","function","class","debug","algorithm","api","bug","error","javascript","typescript","sql","refactor","implement","build","develop","fix"]),
    Agent("research_agent","You are an expert researcher. Provide comprehensive, structured explanations with bullet points and clear headings.",
          ["research","explain","analyze","investigate","compare","review","report","study","summarize","difference","how does","why does","what is","overview","deep dive"]),
    Agent("vision_agent","You are an expert visual analyst. Analyze images with precision and detail.",
          ["image","photo","picture","screenshot","diagram","chart","see","look","show","visual","analyze image","what is this"]),
    Agent("file_agent","You are a file system expert. Handle file operations safely and explain what you're doing.",
          ["file","read","write","save","open","create","delete","folder","directory","path","document","upload","download","rename"]),
    Agent("data_agent","You are a data analyst. Interpret data and generate actionable insights with clear formatting.",
          ["data","chart","graph","plot","statistics","average","sum","trend","insight","analysis","metrics","kpi","dashboard"]),
    Agent("website_agent","You are an expert web developer and UI/UX designer.",
          ["website","webpage","landing page","portfolio","html","css","web app","frontend","design","layout","ui","ux","responsive"]),
    Agent("general_agent","You are JAZZ — a sharp, uncensored AI assistant with personality. You are direct, witty, and genuinely helpful. Be concise unless depth is needed. Use markdown formatting.",
          ["hello","hi","help","chat","talk","general"]),
]
DEFAULT_AGENT = AGENTS[-1]

def route_message(message, has_image=False):
    if has_image:
        return next(a for a in AGENTS if a.name == "vision_agent")
    low = message.lower()
    best, best_score = DEFAULT_AGENT, 0
    for agent in AGENTS[:-1]:
        score = sum(1 for kw in agent.keywords if kw in low)
        if score > best_score:
            best, best_score = agent, score
    return best

# ─────────────────────────────────────────────────────────────────────────────
# NL-to-SQL
# ─────────────────────────────────────────────────────────────────────────────

_FORBIDDEN_SQL = ["INSERT","UPDATE","DELETE","DROP","ALTER","CREATE","TRUNCATE"]

async def run_nl_to_sql(question, user_id):
    t0         = time.time()
    schema_ctx = get_schema_context(question)
    client, model = get_model_client("censored")
    prompt     = f"MySQL schema:\n{schema_ctx}\n\nConvert to safe SELECT only:\n{question}\nReturn ONLY SQL."
    comp       = client.chat.completions.create(model=model, messages=[{"role":"user","content":prompt}], max_tokens=250, temperature=0.05)
    sql        = re.sub(r"```sql|```","",comp.choices[0].message.content.strip()).strip()
    if any(kw in sql.upper() for kw in _FORBIDDEN_SQL) or not sql.upper().startswith("SELECT"):
        return MySQLQueryResponse(success=False, question=question, error="Only SELECT queries are allowed")
    if "LIMIT" not in sql.upper():
        sql += " LIMIT 200"
    conn = get_mysql_connection()
    if not conn:
        return MySQLQueryResponse(success=False, question=question, error="MySQL unavailable")
    try:
        cur = conn.cursor(dictionary=True)
        cur.execute(sql); rows = cur.fetchall()
        elapsed = int((time.time()-t0)*1000)
        cur.close(); conn.close()
        await db_insert(
            "INSERT INTO mysql_query_logs (id,user_id,natural_language,generated_sql,execution_success,row_count,execution_time_ms) VALUES (?,?,?,?,?,?,?)",
            (str(uuid.uuid4()),user_id,question,sql,1,len(rows),elapsed),
        )
        return MySQLQueryResponse(success=True,question=question,generated_sql=sql,results=rows,row_count=len(rows),execution_time_ms=elapsed)
    except Exception as exc:
        try: conn.close()
        except Exception: pass
        return MySQLQueryResponse(success=False,question=question,generated_sql=sql,error=str(exc))

# ─────────────────────────────────────────────────────────────────────────────
# File manager
# ─────────────────────────────────────────────────────────────────────────────

class FileManager:
    def __init__(self, base):
        self.base = Path(base).resolve(); self.base.mkdir(parents=True, exist_ok=True)
    def _safe(self, filename):
        name = os.path.basename(filename)
        if Path(name).suffix.lower() not in ALLOWED_EXTENSIONS: raise ValueError("Extension not allowed")
        resolved = (self.base/name).resolve()
        if not str(resolved).startswith(str(self.base)): raise ValueError("Path traversal denied")
        return resolved
    def create(self,f,c): p=self._safe(f); return {"success":False,"message":"exists"} if p.exists() else (p.write_text(c,"utf-8"),{"success":True,"filename":f})[1]
    def read(self,f):     p=self._safe(f); return {"success":False,"message":"not found"} if not p.exists() else {"success":True,"filename":f,"content":p.read_text("utf-8")}
    def update(self,f,c): p=self._safe(f); return {"success":False,"message":"not found"} if not p.exists() else (p.write_text(c,"utf-8"),{"success":True})[1]
    def delete(self,f):   p=self._safe(f); return {"success":False,"message":"not found"} if not p.exists() else (p.unlink(),{"success":True})[1]
    def list_files(self): return sorted([{"name":f.name,"size":f.stat().st_size,"modified":datetime.fromtimestamp(f.stat().st_mtime).isoformat()} for f in self.base.iterdir() if f.is_file()],key=lambda x:x["modified"],reverse=True)

file_mgr = FileManager(WORKSPACE_DIR)

# ─────────────────────────────────────────────────────────────────────────────
# Website builder
# ─────────────────────────────────────────────────────────────────────────────

_STYLE_HINTS = {
    "modern":        "Clean, bold typography, white space, subtle shadows, sans-serif, CSS Grid layouts",
    "retro":         "80s/90s aesthetic, pixel fonts, neon colors on dark bg, CRT scanline effects",
    "minimal":       "Extreme whitespace, single accent color, fine typography, zero decoration",
    "glassmorphism": "Frosted glass panels, backdrop-filter blur, translucency, light gradients",
    "brutalist":     "Raw HTML feel, bold borders, high contrast, asymmetric layout, exposed structure",
}
_WEBSITE_SYSTEM_PROMPT = """You are an elite frontend developer. Generate COMPLETE, SELF-CONTAINED single-file HTML websites. All CSS in <style>, all JS in <script>. May use CDN. Responsive, impressive. Return ONLY HTML."""

async def _build_website_html(description,title,style,model_type):
    client,model = get_model_client(model_type)
    comp = client.chat.completions.create(
        model=model,
        messages=[{"role":"system","content":_WEBSITE_SYSTEM_PROMPT},{"role":"user","content":f"Build: {description}\nTitle: {title}\nStyle: {style} — {_STYLE_HINTS.get(style,'')}\nReturn ONLY raw HTML."}],
        max_tokens=4096, temperature=0.8,
    )
    html = comp.choices[0].message.content or ""
    html = re.sub(r"^```html\s*","",html.strip(),flags=re.IGNORECASE)
    html = re.sub(r"```\s*$","",html.strip())
    return html.strip()

async def _update_website_html(current_html,instructions,model_type):
    client,model = get_model_client(model_type)
    comp = client.chat.completions.create(
        model=model,
        messages=[{"role":"system","content":_WEBSITE_SYSTEM_PROMPT},{"role":"user","content":f"Existing site:\n{current_html[:8000]}\n\nApply: {instructions}\n\nReturn COMPLETE updated HTML."}],
        max_tokens=4096, temperature=0.7,
    )
    html = comp.choices[0].message.content or ""
    html = re.sub(r"^```html\s*","",html.strip(),flags=re.IGNORECASE)
    html = re.sub(r"```\s*$","",html.strip())
    return html.strip()

async def _save_site_version(site_id: str, user_id: str, current_version: int, html: str, note: str = "") -> None:
    """Archive current site HTML as a version, pruning old ones beyond MAX_SITE_VERSIONS."""
    ver_filename = f"{site_id}_v{current_version}.html"
    ver_path     = SITE_VERSIONS_DIR / ver_filename
    ver_path.write_text(html, encoding="utf-8")
    await db_execute(
        "INSERT INTO website_versions (id,site_id,user_id,version,filename,html_size,note) VALUES (?,?,?,?,?,?,?)",
        (str(uuid.uuid4()), site_id, user_id, current_version, ver_filename, len(html), note),
    )
    # Prune old versions
    old_versions = await db_fetchall(
        "SELECT id,filename FROM website_versions WHERE site_id=? ORDER BY version DESC",
        (site_id,),
    )
    if len(old_versions) > MAX_SITE_VERSIONS:
        for old in old_versions[MAX_SITE_VERSIONS:]:
            old_path = SITE_VERSIONS_DIR / old["filename"]
            if old_path.exists():
                old_path.unlink(missing_ok=True)
            await db_execute("DELETE FROM website_versions WHERE id=?", (old["id"],))

# ─────────────────────────────────────────────────────────────────────────────
# Code runner
# ─────────────────────────────────────────────────────────────────────────────

_PYTHON_BLOCKED = [r"import\s+os",r"import\s+subprocess",r"import\s+sys",r"__import__",r"open\s*\(",r"exec\s*\(",r"eval\s*\(",r"compile\s*\(",r"importlib",r"shutil",r"socket",r"requests\.",r"urllib"]

def _is_code_safe(code, language):
    if language == "python":
        for p in _PYTHON_BLOCKED:
            if re.search(p, code, re.IGNORECASE):
                return False, f"Blocked: {p}"
    return True, ""

def _run_code_sync(code, language):
    t0  = time.time()
    cmd = [sys.executable,"-c",code] if language=="python" else ["node","-e",code] if language=="javascript" else ["bash","-c",code]
    try:
        result = subprocess.run(cmd,capture_output=True,text=True,timeout=CODE_EXEC_TIMEOUT,cwd=tempfile.gettempdir())
        return {"stdout":result.stdout[:CODE_EXEC_MAX_OUT],"stderr":result.stderr[:CODE_EXEC_MAX_OUT],"exit_code":result.returncode,"duration_ms":int((time.time()-t0)*1000)}
    except subprocess.TimeoutExpired:
        return {"stdout":"","stderr":f"⏱ Timed out ({CODE_EXEC_TIMEOUT}s)","exit_code":-1,"duration_ms":CODE_EXEC_TIMEOUT*1000}
    except Exception as e:
        return {"stdout":"","stderr":str(e),"exit_code":-1,"duration_ms":0}

async def run_sqlite_nl_query(question):
    schema_ctx    = get_schema_context(question,top_k=5)
    client, model = get_model_client("censored")
    prompt = f"SQLite schema:\n{schema_ctx}\n\nWrite safe READ-ONLY SELECT for: {question}\nRULES: SELECT only. Add LIMIT 100.\nReturn ONLY SQL."
    comp   = client.chat.completions.create(model=model,messages=[{"role":"user","content":prompt}],max_tokens=300,temperature=0.05)
    sql    = re.sub(r"```sql|```","",comp.choices[0].message.content.strip()).strip()
    if not sql.upper().startswith("SELECT"):
        return {"success":False,"error":"Must be SELECT","sql":sql}
    return await _run_sqlite_raw(sql)

async def _run_sqlite_raw(sql):
    t0 = time.time()
    try:
        rows    = await db_fetchall(sql)
        elapsed = int((time.time()-t0)*1000)
        return {"success":True,"sql":sql,"rows":rows,"row_count":len(rows),"duration_ms":elapsed}
    except Exception as exc:
        return {"success":False,"sql":sql,"error":str(exc),"rows":[],"row_count":0}

# ─────────────────────────────────────────────────────────────────────────────
# Agent job engine
# ─────────────────────────────────────────────────────────────────────────────

class JobLogger:
    def __init__(self): self.lines=[]; self.applied=[]
    def log(self,msg):
        entry=f"[{datetime.now().strftime('%H:%M:%S')}] {msg}"
        self.lines.append(entry); logger.info("[JOB] %s",msg)
    def add_applied(self,job): self.applied.append(job)
    def output(self): return "\n".join(self.lines)

def _is_within_time_window(start_str,end_str):
    try:
        now=datetime.now(); fmt="%H:%M"
        start=datetime.strptime(start_str,fmt).replace(year=now.year,month=now.month,day=now.day)
        end  =datetime.strptime(end_str,fmt).replace(year=now.year,month=now.month,day=now.day)
        return start<=now<=end
    except Exception: return True

def _execute_job_sync(job_id,bypass_time_window=False):
    import sqlite3 as _s3
    def _db(sql,p=()):
        with _s3.connect(DB_PATH) as c:
            c.row_factory=_s3.Row; return c.execute(sql,p).fetchone()
    def _dbw(sql,p=()):
        with _s3.connect(DB_PATH) as c:
            c.execute(sql,p); c.commit()
    row=_db("SELECT * FROM agent_jobs WHERE id=?",(job_id,))
    if not row: return
    job=dict(row)
    if not bypass_time_window:
        if not _is_within_time_window(job.get("time_window_start","00:00"),job.get("time_window_end","23:59")):
            return
    now_iso=datetime.utcnow().isoformat(); log_id=str(uuid.uuid4())
    _dbw("UPDATE agent_jobs SET status='running',last_run_at=?,updated_at=? WHERE id=?",(now_iso,now_iso,job_id))
    _dbw("INSERT INTO agent_job_logs (id,job_id,user_id,status,output,started_at) VALUES (?,?,?,?,?,?)",(log_id,job_id,job["user_id"],"running","Starting\n",now_iso))
    jlog=JobLogger(); jlog.log(f"⚡ Job '{job['name']}' starting")
    t0=time.time(); applied=[]; error=None; final_status="completed"
    try:
        creds=decrypt_creds(job["credentials_enc"]) if job.get("credentials_enc") else {}
        params=json.loads(job.get("parameters") or "{}")
        if job["job_type"]=="naukri_apply":    applied=_run_naukri_agent(creds,params,jlog)
        elif job["job_type"]=="linkedin_apply": applied=_run_linkedin_agent(creds,params,jlog)
        else: jlog.log("⚠ Unknown job type")
        jlog.log(f"✅ Done. Applied: {len(applied)}")
    except Exception as exc:
        error=f"{exc}\n\n{traceback.format_exc()}"; final_status="failed"; jlog.log(f"❌ {exc}")
    elapsed=int(time.time()-t0); done_iso=datetime.utcnow().isoformat()
    _dbw("UPDATE agent_job_logs SET status=?,output=?,applied_jobs=?,error=?,completed_at=?,duration_sec=? WHERE id=?",(final_status,jlog.output(),json.dumps(applied,default=str),error,done_iso,elapsed,log_id))
    _dbw("UPDATE agent_jobs SET status='idle',last_run_at=?,last_run_status=?,total_runs=total_runs+1,total_applied=total_applied+?,updated_at=? WHERE id=?",(done_iso,final_status,len(applied),done_iso,job_id))

def _run_naukri_agent(creds,params,jlog):
    if not _PLAYWRIGHT_AVAILABLE: jlog.log("❌ Playwright not installed."); return []
    from playwright.sync_api import sync_playwright
    email=creds.get("email",""); password=creds.get("password","")
    roles=params.get("roles",["Data Analyst"]); location=params.get("location","")
    exp_min=params.get("experience_min",0); exp_max=params.get("experience_max",10)
    max_apply=min(params.get("max_apply",10),50); kw_exclude=[k.lower() for k in params.get("keywords_exclude",[])]
    applied=[]
    jlog.log(f"🚀 Naukri | roles={roles}")
    with sync_playwright() as pw:
        browser=pw.chromium.launch(headless=True,args=["--no-sandbox","--disable-dev-shm-usage"])
        ctx=browser.new_context(user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120",viewport={"width":1280,"height":800})
        page=ctx.new_page(); page.set_default_timeout(30_000)
        try:
            page.goto("https://www.naukri.com/nlogin/login",wait_until="domcontentloaded"); page.wait_for_timeout(2000)
            page.wait_for_selector("input[type='email'],#usernameField",timeout=10_000)
            page.fill("input[type='email'],#usernameField",email); page.wait_for_timeout(400)
            page.fill("input[type='password'],#passwordField",password); page.wait_for_timeout(400)
            page.click("button[type='submit']"); page.wait_for_timeout(4000)
            if "login" in page.url.lower(): jlog.log("❌ Login failed"); browser.close(); return []
            jlog.log("✅ Logged in")
        except Exception as exc: jlog.log(f"❌ {exc}"); browser.close(); return []
        total=0
        for role in roles:
            if total>=max_apply: break
            kw_enc=role.replace(" ","%20"); loc_enc=location.replace(" ","%20") if location else ""
            url=f"https://www.naukri.com/jobs?k={kw_enc}&l={loc_enc}&experience={exp_min}&to={exp_max}" if loc_enc else f"https://www.naukri.com/jobs?k={kw_enc}&experience={exp_min}&to={exp_max}"
            try: page.goto(url,wait_until="domcontentloaded"); page.wait_for_timeout(3000)
            except: continue
            cards=page.query_selector_all("article.jobTuple,div.srp-jobtuple-wrapper")
            for card in cards:
                if total>=max_apply: break
                try:
                    te=card.query_selector("a.title,a[class*='title']"); ce=card.query_selector("a.subTitle,a[class*='company']")
                    title=te.inner_text().strip() if te else "Unknown"; company=ce.inner_text().strip() if ce else "Unknown"
                    job_url=te.get_attribute("href") if te else ""
                    if any(k in title.lower() for k in kw_exclude) or not job_url: continue
                    jp=ctx.new_page()
                    try:
                        jp.goto(job_url,wait_until="domcontentloaded"); jp.wait_for_timeout(2000)
                        ab=jp.query_selector("button:has-text('Apply'),button:has-text('Easy Apply')")
                        if not ab: jp.close(); continue
                        ab.click(); jp.wait_for_timeout(2500)
                        for _ in range(4):
                            sb=jp.query_selector("button:has-text('Apply'),button:has-text('Submit'),button:has-text('Continue')")
                            if not sb: break
                            try: sb.click(); jp.wait_for_timeout(1500)
                            except: break
                        applied.append({"title":title,"company":company,"url":job_url,"applied_at":datetime.utcnow().isoformat(),"status":"applied"})
                        total+=1; jlog.log(f"✅ Applied: {title} @ {company} ({total}/{max_apply})")
                    except Exception as exc: jlog.log(f"⚠ {exc}")
                    finally:
                        try: jp.close()
                        except: pass
                except Exception as exc: jlog.log(f"⚠ Card: {exc}")
        browser.close()
    jlog.log(f"🏁 Done. Applied: {total}")
    return applied

def _run_linkedin_agent(creds,params,jlog):
    if not _PLAYWRIGHT_AVAILABLE: jlog.log("❌ Playwright not installed."); return []
    from playwright.sync_api import sync_playwright
    email=creds.get("email",""); password=creds.get("password","")
    roles=params.get("roles",["Data Analyst"]); location=params.get("location","India"); max_apply=min(params.get("max_apply",5),25)
    applied=[]
    with sync_playwright() as pw:
        browser=pw.chromium.launch(headless=True,args=["--no-sandbox"])
        ctx=browser.new_context(user_agent="Mozilla/5.0"); page=ctx.new_page(); page.set_default_timeout(30_000)
        try:
            page.goto("https://www.linkedin.com/login",wait_until="domcontentloaded")
            page.fill("#username",email); page.fill("#password",password); page.click("button[type='submit']"); page.wait_for_timeout(4000)
            if "checkpoint" in page.url or "login" in page.url: jlog.log("❌ LinkedIn login failed"); browser.close(); return []
            jlog.log("✅ LinkedIn logged in")
        except Exception as exc: jlog.log(f"❌ {exc}"); browser.close(); return []
        total=0
        for role in roles:
            if total>=max_apply: break
            try:
                page.goto(f"https://www.linkedin.com/jobs/search/?keywords={role.replace(' ','%20')}&location={location.replace(' ','%20')}&f_LF=f_AL",wait_until="domcontentloaded"); page.wait_for_timeout(3000)
                cards=page.query_selector_all("li.jobs-search-results__list-item")
                for card in cards:
                    if total>=max_apply: break
                    try:
                        card.click(); page.wait_for_timeout(2000)
                        te=page.query_selector("h1.job-details-jobs-unified-top-card__job-title"); ce=page.query_selector("div.job-details-jobs-unified-top-card__company-name")
                        title=te.inner_text().strip() if te else "Unknown"; company=ce.inner_text().strip() if ce else "Unknown"
                        eb=page.query_selector("button:has-text('Easy Apply')")
                        if not eb: continue
                        eb.click(); page.wait_for_timeout(2000)
                        for _ in range(5):
                            nb=page.query_selector("button:has-text('Next'),button:has-text('Review'),button:has-text('Submit application')")
                            if not nb: break
                            nb.click(); page.wait_for_timeout(1500)
                        applied.append({"title":title,"company":company,"applied_at":datetime.utcnow().isoformat(),"status":"applied"})
                        total+=1; jlog.log(f"✅ Applied: {title} @ {company}")
                    except Exception as exc: jlog.log(f"⚠ {exc}")
            except Exception as exc: jlog.log(f"⚠ Search: {exc}")
        browser.close(); jlog.log(f"🏁 Done. Applied: {total}")
    return applied

async def execute_job(job_id,bypass_time_window=False):
    loop=asyncio.get_event_loop()
    await loop.run_in_executor(_executor,_execute_job_sync,job_id,bypass_time_window)

scheduler = BackgroundScheduler(daemon=True)

def _schedule_job(job):
    sid=f"agent_{job['id']}"
    try:
        parts=job["cron_schedule"].split()
        if len(parts)!=5: raise ValueError(f"Invalid cron: {job['cron_schedule']}")
        minute,hour,day,month,dow=parts
        trigger=CronTrigger(minute=minute,hour=hour,day=day,month=month,day_of_week=dow)
        def _wrapper(jid=job["id"]):
            import asyncio as _aio; loop=_aio.new_event_loop()
            try: loop.run_until_complete(execute_job(jid,bypass_time_window=False))
            finally: loop.close()
        if scheduler.get_job(sid): scheduler.remove_job(sid)
        scheduler.add_job(_wrapper,trigger=trigger,id=sid,replace_existing=True)
        logger.info("[SCHEDULER] Registered '%s'",job["name"])
    except Exception as exc: logger.error("[SCHEDULER] %s: %s",job["id"],exc)

def _unschedule_job(job_id):
    sid=f"agent_{job_id}"
    if scheduler.get_job(sid): scheduler.remove_job(sid)

def _load_all_jobs_into_scheduler():
    import sqlite3 as _s3
    try:
        with _s3.connect(DB_PATH) as conn:
            conn.row_factory=_s3.Row
            rows=conn.execute("SELECT * FROM agent_jobs WHERE enabled=1").fetchall()
        for row in rows: _schedule_job(dict(row))
        logger.info("[SCHEDULER] Loaded %d jobs ✅",len(rows))
    except Exception as exc: logger.error("[SCHEDULER] %s",exc)

def _run_expiry_check():
    import sqlite3 as _s3
    now_iso = datetime.utcnow().isoformat()
    try:
        with _s3.connect(DB_PATH) as conn:
            conn.row_factory = _s3.Row
            rows = conn.execute(
                "SELECT id, email, subscription FROM users WHERE subscription_expires_at IS NOT NULL AND subscription_expires_at <= ? AND subscription != 'free'",
                (now_iso,),
            ).fetchall()
        for row in rows:
            with _s3.connect(DB_PATH) as conn:
                conn.execute(
                    "UPDATE users SET subscription='free', subscription_expires_at=NULL, subscription_changed_by='system', subscription_changed_at=?, updated_at=? WHERE id=?",
                    (now_iso, now_iso, row["id"]),
                )
                conn.execute(
                    "INSERT INTO subscription_audit_log (id,user_id,changed_by,changed_by_email,action,old_value,new_value,note) VALUES (?,?,?,?,?,?,?,?)",
                    (str(uuid.uuid4()), row["id"], "system", "system", "expire", row["subscription"], "free", "Auto-expired"),
                )
                conn.commit()
            # Insert notification
            with _s3.connect(DB_PATH) as conn:
                conn.execute(
                    "INSERT INTO user_notifications (id,user_id,title,body,type) VALUES (?,?,?,?,?)",
                    (str(uuid.uuid4()), row["id"], "Subscription Expired",
                     f"Your {row['subscription']} subscription has expired. You're now on the Free plan.", "warning"),
                )
                conn.commit()
            logger.info("[EXPIRY] User %s expired from %s → free", row["id"], row["subscription"])
    except Exception as exc:
        logger.error("[EXPIRY] Check failed: %s", exc)

def _record_health():
    """Record a health snapshot every 5 minutes for the health history endpoint."""
    import sqlite3 as _s3
    try:
        db_ok    = True
        chroma_ok = chroma_client is not None
        mysql_ok  = _mysql_pool is not None
        uptime    = int((datetime.utcnow() - _SERVER_START).total_seconds())
        with _s3.connect(DB_PATH) as conn:
            conn.execute(
                "INSERT INTO health_history (id,ts,sqlite_ok,chroma_ok,mysql_ok,uptime_sec) VALUES (?,?,?,?,?,?)",
                (str(uuid.uuid4()), datetime.utcnow().isoformat(), 1 if db_ok else 0, 1 if chroma_ok else 0, 1 if mysql_ok else 0, uptime),
            )
            # Keep only last 288 records (24h at 5min intervals)
            conn.execute("DELETE FROM health_history WHERE id NOT IN (SELECT id FROM health_history ORDER BY created_at DESC LIMIT 288)")
            conn.commit()
    except Exception as exc:
        logger.warning("[HEALTH] Record failed: %s", exc)

# ─────────────────────────────────────────────────────────────────────────────
# FastAPI app
# ─────────────────────────────────────────────────────────────────────────────

@asynccontextmanager
async def lifespan(app: FastAPI) -> AsyncGenerator:
    logger.info("="*60)
    logger.info("[STARTUP] JAZZ AI v7.1 (Production Edition)")
    logger.info("  DB:    %s", DB_PATH)
    logger.info("  ADMIN: %s", ADMIN_EMAIL)
    logger.info("="*60)
    await init_db()
    index_sqlite_schema()
    if MYSQL_ENABLED: _init_mysql_pool()
    scheduler.add_job(index_sqlite_schema,   trigger=IntervalTrigger(hours=1),               id="index_sqlite",  replace_existing=True)
    scheduler.add_job(_run_expiry_check,     trigger=CronTrigger(hour=0, minute=5),           id="expiry_check",  replace_existing=True)
    scheduler.add_job(_record_health,        trigger=IntervalTrigger(minutes=5),              id="health_record", replace_existing=True)
    scheduler.start()
    _load_all_jobs_into_scheduler()
    logger.info("[STARTUP] All services ready ✅")
    yield
    logger.info("[SHUTDOWN] Stopping…")
    scheduler.shutdown(wait=False)
    _executor.shutdown(wait=False)

app = FastAPI(
    title="JAZZ AI",
    version="7.1.0",
    description="JAZZ Uncensored AI — Production Edition",
    lifespan=lifespan,
    docs_url="/api/docs",
    redoc_url=None,
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=ALLOWED_ORIGINS,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.middleware("http")
async def request_middleware(request: Request, call_next):
    rid      = str(uuid.uuid4())[:8]
    start    = time.time()
    response = await call_next(request)
    elapsed  = round((time.time() - start) * 1000)
    response.headers["X-Request-ID"]      = rid
    response.headers["X-Powered-By"]      = "JAZZ-AI-v7.1"
    response.headers["X-Response-Time-Ms"] = str(elapsed)
    return response

app.mount("/preview", StaticFiles(directory=str(SITES_DIR), html=True), name="preview")
_static_dir = Path("static")
if _static_dir.exists():
    app.mount("/static", StaticFiles(directory="static"), name="static")
    @app.get("/")
    async def root(): return FileResponse("static/index.html")
else:
    @app.get("/")
    async def root(): return {"status":"JAZZ AI v7.1","docs":"/api/docs"}

# ══════════════════════════════════════════════════════════════════════════════
#  AUTH
# ══════════════════════════════════════════════════════════════════════════════

@app.post("/auth/signup", response_model=TokenResponse)
async def signup(req: UserSignup, request: Request):
    if await db_fetchone("SELECT id FROM users WHERE email=?", (req.email,)):
        raise HTTPException(400, "Email already registered")
    uid   = str(uuid.uuid4())
    phash = pwd_context.hash(req.password)
    role  = "admin" if req.email == ADMIN_EMAIL else "client"
    sub   = "enterprise" if req.email == ADMIN_EMAIL else "free"
    await db_execute(
        "INSERT INTO users (id,email,password_hash,full_name,role,subscription,account_status) VALUES (?,?,?,?,?,?,?)",
        (uid, req.email, phash, req.full_name or "", role, sub, "active"),
    )
    await log_activity(uid, "signup", ip=request.client.host if request.client else "")
    token = create_access_token({"sub":uid,"email":req.email,"role":role})
    return TokenResponse(access_token=token, user={"id":uid,"email":req.email,"role":role,"subscription":sub,"full_name":req.full_name or ""})

@app.post("/auth/login", response_model=TokenResponse)
async def login(req: UserLogin, request: Request):
    ip = request.client.host if request.client else ""
    if req.email == ADMIN_EMAIL and ADMIN_PASSWORD and req.password == ADMIN_PASSWORD:
        user = await db_fetchone("SELECT * FROM users WHERE email=?", (ADMIN_EMAIL,))
        if not user:
            uid=str(uuid.uuid4()); phash=pwd_context.hash(ADMIN_PASSWORD)
            await db_execute("INSERT INTO users (id,email,password_hash,full_name,role,subscription,account_status) VALUES (?,?,?,?,?,?,?)",(uid,ADMIN_EMAIL,phash,"Admin","admin","enterprise","active"))
            user={"id":uid,"email":ADMIN_EMAIL,"role":"admin","subscription":"enterprise","full_name":"Admin","memory_enabled":1,"account_status":"active"}
        await log_activity(user["id"], "login", ip=ip)
        token=create_access_token({"sub":user["id"],"email":req.email,"role":"admin"})
        return TokenResponse(access_token=token,user={"id":user["id"],"email":req.email,"role":"admin","subscription":"enterprise","full_name":"Admin","memory_enabled":True,"account_status":"active"})
    user = await db_fetchone("SELECT * FROM users WHERE email=?", (req.email,))
    if not user: raise HTTPException(401,"Invalid credentials")
    if user.get("deleted_at"):                raise HTTPException(403,"Account deleted.")
    if user.get("account_status") == "banned": raise HTTPException(403,"Account suspended. Contact support.")
    if user.get("account_status") == "paused": raise HTTPException(403,"Account paused. Contact support.")
    try: ok = pwd_context.verify(req.password, user["password_hash"])
    except Exception: ok = False
    if not ok: raise HTTPException(401,"Invalid credentials")
    await log_activity(user["id"], "login", ip=ip)
    token=create_access_token({"sub":user["id"],"email":user["email"],"role":user["role"]})
    return TokenResponse(access_token=token,user={"id":user["id"],"email":user["email"],"role":user["role"],"subscription":user["subscription"],"full_name":user.get("full_name",""),"memory_enabled":bool(user.get("memory_enabled",1)),"account_status":user.get("account_status","active")})

@app.post("/auth/admin-login", response_model=TokenResponse)
async def admin_login(req: UserLogin, request: Request):
    if req.email != ADMIN_EMAIL: raise HTTPException(403,"Admin access only")
    return await login(req, request)

@app.get("/auth/me")
async def me(user: Dict = Depends(get_current_user)):
    row = await db_fetchone(
        "SELECT id,email,full_name,role,subscription,account_status,memory_enabled,avatar_url,preferred_model,subscription_expires_at,upgrade_requested,upgrade_requested_at,created_at FROM users WHERE id=?",
        (user["sub"],),
    )
    if not row: raise HTTPException(404,"User not found")
    row["memory_enabled"] = bool(row.get("memory_enabled",1))
    return row

@app.post("/auth/logout")
async def logout(user: Dict = Depends(get_current_user)):
    await log_activity(user["sub"], "logout")
    return {"message":"Logged out"}

@app.post("/auth/change-password")
async def change_password(req: PasswordChangeRequest, user: Dict = Depends(get_current_user)):
    """Authenticated user changes their own password."""
    row = await db_fetchone("SELECT password_hash FROM users WHERE id=?", (user["sub"],))
    if not row: raise HTTPException(404,"User not found")
    try: ok = pwd_context.verify(req.current_password, row["password_hash"])
    except Exception: ok = False
    if not ok: raise HTTPException(400,"Current password is incorrect")
    new_hash = pwd_context.hash(req.new_password)
    await db_execute(
        "UPDATE users SET password_hash=?, updated_at=? WHERE id=?",
        (new_hash, datetime.utcnow().isoformat(), user["sub"]),
    )
    await log_activity(user["sub"], "change_password")
    return {"success": True, "message": "Password changed successfully"}

@app.post("/auth/request-password-reset")
async def request_password_reset(req: PasswordResetRequest):
    """
    Initiates a password reset. In production, email the token.
    In this build, the token is returned directly (replace with email send).
    """
    user = await db_fetchone("SELECT id FROM users WHERE email=?", (req.email,))
    if not user:
        # Return success anyway to avoid email enumeration
        return {"success": True, "message": "If that email exists, a reset link has been sent."}
    raw_token   = secrets.token_urlsafe(32)
    import hashlib
    token_hash  = hashlib.sha256(raw_token.encode()).hexdigest()
    expires_at  = (datetime.utcnow() + timedelta(minutes=PASSWORD_RESET_EXPIRY_MINUTES)).isoformat()
    # Invalidate old tokens
    await db_execute("UPDATE password_reset_tokens SET used=1 WHERE user_id=? AND used=0", (user["id"],))
    await db_execute(
        "INSERT INTO password_reset_tokens (id,user_id,token_hash,expires_at) VALUES (?,?,?,?)",
        (str(uuid.uuid4()), user["id"], token_hash, expires_at),
    )
    # TODO: In production, send email with raw_token instead of returning it
    logger.info("[PASSWORD-RESET] Token for %s: %s", req.email, raw_token)
    return {
        "success": True,
        "message": "Reset token generated. (Dev mode: token returned directly — use email in production.)",
        "dev_token": raw_token,  # REMOVE IN PRODUCTION — send via email instead
        "expires_minutes": PASSWORD_RESET_EXPIRY_MINUTES,
    }

@app.post("/auth/confirm-password-reset")
async def confirm_password_reset(req: PasswordResetConfirm):
    """Complete password reset with the token from the email."""
    import hashlib
    token_hash = hashlib.sha256(req.token.encode()).hexdigest()
    now        = datetime.utcnow().isoformat()
    row = await db_fetchone(
        "SELECT * FROM password_reset_tokens WHERE token_hash=? AND used=0 AND expires_at > ?",
        (token_hash, now),
    )
    if not row: raise HTTPException(400,"Invalid or expired reset token")
    new_hash = pwd_context.hash(req.new_password)
    await db_execute(
        "UPDATE users SET password_hash=?, updated_at=? WHERE id=?",
        (new_hash, now, row["user_id"]),
    )
    await db_execute("UPDATE password_reset_tokens SET used=1 WHERE id=?", (row["id"],))
    await log_activity(row["user_id"], "password_reset_confirmed")
    return {"success": True, "message": "Password reset successfully. Please log in."}

@app.delete("/auth/account")
async def delete_own_account(user: Dict = Depends(get_current_user)):
    """GDPR: Soft-delete the authenticated user's account."""
    uid = user["sub"]
    row = await db_fetchone("SELECT role FROM users WHERE id=?", (uid,))
    if row and row.get("role") == "admin":
        raise HTTPException(400, "Admin accounts cannot be self-deleted.")
    now = datetime.utcnow().isoformat()
    # Soft delete — mark deleted, anonymize PII
    anon_email = f"deleted_{uid[:8]}@deleted.invalid"
    await db_execute(
        "UPDATE users SET deleted_at=?, email=?, full_name='[Deleted]', password_hash='', account_status='banned', updated_at=? WHERE id=?",
        (now, anon_email, now, uid),
    )
    await log_activity(uid, "account_self_deleted")
    return {"success": True, "message": "Your account has been deleted."}

@app.get("/user/export")
async def export_user_data(user: Dict = Depends(get_current_user)):
    """GDPR: Export everything associated with the authenticated user."""
    uid = user["sub"]
    profile     = await db_fetchone("SELECT id,email,full_name,role,subscription,created_at FROM users WHERE id=?", (uid,))
    chats       = await db_fetchall("SELECT message,response,agent_used,model_used,created_at FROM chat_history WHERE user_id=?", (uid,))
    memories    = await db_fetchall("SELECT content,category,source,created_at FROM user_memories WHERE user_id=?", (uid,))
    documents   = await db_fetchall("SELECT original_name,file_type,file_size,uploaded_at FROM documents WHERE user_id=?", (uid,))
    websites    = await db_fetchall("SELECT title,description,created_at FROM websites WHERE user_id=?", (uid,))
    prompts     = await db_fetchall("SELECT title,content,category,created_at FROM prompt_templates WHERE user_id=?", (uid,))
    activity    = await db_fetchall("SELECT action,resource,detail,created_at FROM user_activity_log WHERE user_id=? ORDER BY created_at DESC LIMIT 500", (uid,))
    export_data = {
        "export_date": datetime.utcnow().isoformat(),
        "profile":     profile,
        "chat_history": chats,
        "memories":     memories,
        "documents":    documents,
        "websites":     websites,
        "prompts":      prompts,
        "activity_log": activity,
    }
    content = json.dumps(export_data, indent=2, default=str)
    await log_activity(uid, "data_export")
    return StreamingResponse(
        iter([content.encode()]),
        media_type="application/json",
        headers={"Content-Disposition": f'attachment; filename="jazz_data_export_{uid[:8]}.json"'},
    )

# ══════════════════════════════════════════════════════════════════════════════
#  API KEYS
# ══════════════════════════════════════════════════════════════════════════════

@app.post("/api-keys")
async def create_api_key(req: APIKeyCreate, user: Dict = Depends(get_current_user)):
    """Generate a new API key for programmatic access."""
    uid = user["sub"]
    await check_account_active(uid)
    # Limit to 10 active keys per user
    count = await db_count("SELECT COUNT(*) FROM api_keys WHERE user_id=? AND is_active=1", (uid,))
    if count >= 10:
        raise HTTPException(400, "Maximum of 10 active API keys allowed")
    import hashlib
    raw_key  = API_KEY_PREFIX + secrets.token_urlsafe(40)
    key_hash = hashlib.sha256(raw_key.encode()).hexdigest()
    prefix   = raw_key[:12] + "..."
    kid      = str(uuid.uuid4())
    await db_execute(
        "INSERT INTO api_keys (id,user_id,name,key_hash,key_prefix,expires_at) VALUES (?,?,?,?,?,?)",
        (kid, uid, req.name, key_hash, prefix, req.expires_at),
    )
    await log_activity(uid, "api_key_created", detail=req.name)
    return {
        "success":    True,
        "id":         kid,
        "name":       req.name,
        "key":        raw_key,  # Shown ONCE — not stored in plaintext
        "prefix":     prefix,
        "expires_at": req.expires_at,
        "warning":    "Save this key now — it will NOT be shown again.",
    }

@app.get("/api-keys")
async def list_api_keys(user: Dict = Depends(get_current_user)):
    rows = await db_fetchall(
        "SELECT id,name,key_prefix,last_used_at,expires_at,is_active,created_at FROM api_keys WHERE user_id=? ORDER BY created_at DESC",
        (user["sub"],),
    )
    return {"api_keys": rows}

@app.delete("/api-keys/{key_id}")
async def revoke_api_key(key_id: str, user: Dict = Depends(get_current_user)):
    row = await db_fetchone("SELECT id FROM api_keys WHERE id=? AND user_id=?", (key_id, user["sub"]))
    if not row: raise HTTPException(404, "API key not found")
    await db_execute("UPDATE api_keys SET is_active=0 WHERE id=?", (key_id,))
    await log_activity(user["sub"], "api_key_revoked", detail=key_id)
    return {"success": True}

# ══════════════════════════════════════════════════════════════════════════════
#  PROFILE
# ══════════════════════════════════════════════════════════════════════════════

@app.get("/profile")
async def get_profile(user: Dict = Depends(get_current_user)):
    row = await db_fetchone(
        "SELECT id,email,full_name,role,subscription,account_status,memory_enabled,avatar_url,preferred_model,subscription_expires_at,upgrade_requested,created_at FROM users WHERE id=?",
        (user["sub"],),
    )
    if not row: raise HTTPException(404,"Profile not found")
    row["memory_enabled"] = bool(row.get("memory_enabled",1))
    return row

@app.put("/profile")
async def update_profile(req: ProfileUpdate, user: Dict = Depends(get_current_user)):
    data = req.model_dump(exclude_none=True)
    if not data: raise HTTPException(400,"Nothing to update")
    sets = ", ".join(f"{k}=?" for k in data)
    await db_execute(f"UPDATE users SET {sets}, updated_at=? WHERE id=?", tuple(data.values())+(datetime.utcnow().isoformat(),user["sub"]))
    return {"success":True}

@app.get("/user/stats")
async def user_stats(user: Dict = Depends(get_current_user)):
    uid    = user["sub"]
    tier   = await get_user_subscription(uid)
    limits = await _get_effective_limits_async(tier)
    row    = await db_fetchone("SELECT subscription_expires_at, account_status, upgrade_requested FROM users WHERE id=?", (uid,))
    lim    = lambda n: -1 if n == -1 else n
    return {
        "subscription":            tier,
        "account_status":          (row or {}).get("account_status","active"),
        "subscription_expires_at": (row or {}).get("subscription_expires_at"),
        "upgrade_requested":       (row or {}).get("upgrade_requested"),
        "messages_used_today":     rate_limiter.usage_today(uid),
        "messages_limit":          lim(limits["messages_per_day"]),
        "documents":               await db_count("SELECT COUNT(*) FROM documents WHERE user_id=?", (uid,)),
        "document_limit":          lim(limits["documents"]),
        "websites":                await db_count("SELECT COUNT(*) FROM websites WHERE user_id=?", (uid,)),
        "websites_limit":          lim(limits.get("websites",3)),
        "mysql_access":            limits["mysql_access"],
        "image_analysis":          limits["image_analysis"],
        "memory_count":            await db_count("SELECT COUNT(*) FROM user_memories WHERE user_id=?", (uid,)),
        "agent_jobs":              await db_count("SELECT COUNT(*) FROM agent_jobs WHERE user_id=?", (uid,)),
        "agent_jobs_limit":        lim(limits.get("agent_jobs",1)),
        "code_runs_today":         rate_limiter.usage_today(uid, "code_runs_per_day"),
        "code_runs_per_day":       lim(limits.get("code_runs_per_day",20)),
        "chat_sessions":           await db_count("SELECT COUNT(*) FROM chat_sessions WHERE user_id=?", (uid,)),
        "api_keys_active":         await db_count("SELECT COUNT(*) FROM api_keys WHERE user_id=? AND is_active=1", (uid,)),
        "unread_notifications":    await db_count("SELECT COUNT(*) FROM user_notifications WHERE user_id=? AND is_read=0", (uid,)),
        "playwright_available":    _PLAYWRIGHT_AVAILABLE,
        "version":                 "7.1.0",
    }

# ══════════════════════════════════════════════════════════════════════════════
#  CHAT SESSIONS (NEW)
# ══════════════════════════════════════════════════════════════════════════════

@app.post("/chat/sessions")
async def create_session(req: SessionCreate, user: Dict = Depends(get_current_user)):
    uid = user["sub"]
    await check_account_active(uid)
    sid = str(uuid.uuid4()); now = datetime.utcnow().isoformat()
    await db_execute(
        "INSERT INTO chat_sessions (id,user_id,title,model_type,created_at,updated_at) VALUES (?,?,?,?,?,?)",
        (sid, uid, req.title, req.model_type or "uncensored", now, now),
    )
    return await db_fetchone("SELECT * FROM chat_sessions WHERE id=?", (sid,))

@app.get("/chat/sessions")
async def list_sessions(user: Dict = Depends(get_current_user)):
    rows = await db_fetchall(
        "SELECT * FROM chat_sessions WHERE user_id=? ORDER BY is_pinned DESC, last_message_at DESC, created_at DESC",
        (user["sub"],),
    )
    return {"sessions": rows, "count": len(rows)}

@app.get("/chat/sessions/{session_id}")
async def get_session(session_id: str, page: int = 1, per_page: int = 50, user: Dict = Depends(get_current_user)):
    row = await db_fetchone("SELECT * FROM chat_sessions WHERE id=? AND user_id=?", (session_id, user["sub"]))
    if not row: raise HTTPException(404, "Session not found")
    offset = (page - 1) * per_page
    history = await db_fetchall(
        "SELECT * FROM chat_history WHERE session_id=? ORDER BY created_at ASC LIMIT ? OFFSET ?",
        (session_id, per_page, offset),
    )
    return {"session": row, "history": history, "page": page, "per_page": per_page}

@app.put("/chat/sessions/{session_id}")
async def update_session(session_id: str, req: SessionUpdate, user: Dict = Depends(get_current_user)):
    row = await db_fetchone("SELECT id FROM chat_sessions WHERE id=? AND user_id=?", (session_id, user["sub"]))
    if not row: raise HTTPException(404, "Session not found")
    updates = {}
    if req.title     is not None: updates["title"]     = req.title
    if req.is_pinned is not None: updates["is_pinned"] = 1 if req.is_pinned else 0
    if not updates: raise HTTPException(400, "Nothing to update")
    updates["updated_at"] = datetime.utcnow().isoformat()
    sets = ", ".join(f"{k}=?" for k in updates)
    await db_execute(f"UPDATE chat_sessions SET {sets} WHERE id=?", tuple(updates.values()) + (session_id,))
    return await db_fetchone("SELECT * FROM chat_sessions WHERE id=?", (session_id,))

@app.delete("/chat/sessions/{session_id}")
async def delete_session(session_id: str, user: Dict = Depends(get_current_user)):
    row = await db_fetchone("SELECT id FROM chat_sessions WHERE id=? AND user_id=?", (session_id, user["sub"]))
    if not row: raise HTTPException(404, "Session not found")
    await db_execute("DELETE FROM chat_history WHERE session_id=?", (session_id,))
    await db_execute("DELETE FROM chat_sessions WHERE id=?", (session_id,))
    return {"success": True}

@app.get("/chat/search")
async def search_chat(
    q:        str = Query(..., min_length=1, max_length=500),
    limit:    int = Query(20, le=100),
    user:     Dict = Depends(get_current_user),
):
    """Full-text search across the authenticated user's chat history."""
    uid  = user["sub"]
    term = f"%{q}%"
    rows = await db_fetchall(
        "SELECT id,message,response,agent_used,model_used,session_id,created_at FROM chat_history "
        "WHERE user_id=? AND (message LIKE ? OR response LIKE ?) ORDER BY created_at DESC LIMIT ?",
        (uid, term, term, limit),
    )
    return {"results": rows, "count": len(rows), "query": q}

# ══════════════════════════════════════════════════════════════════════════════
#  NOTIFICATIONS (NEW)
# ══════════════════════════════════════════════════════════════════════════════

@app.get("/notifications")
async def get_notifications(
    unread_only: bool = False,
    limit:       int  = Query(50, le=200),
    user:        Dict = Depends(get_current_user),
):
    uid  = user["sub"]
    sql  = "SELECT * FROM user_notifications WHERE user_id=?"
    params: tuple = (uid,)
    if unread_only:
        sql += " AND is_read=0"
    sql += " ORDER BY created_at DESC LIMIT ?"
    params += (limit,)
    rows  = await db_fetchall(sql, params)
    unread = await db_count("SELECT COUNT(*) FROM user_notifications WHERE user_id=? AND is_read=0", (uid,))
    return {"notifications": rows, "unread_count": unread}

@app.post("/notifications/mark-read")
async def mark_notifications_read(req: NotificationMarkRead, user: Dict = Depends(get_current_user)):
    uid = user["sub"]
    if req.notification_ids:
        for nid in req.notification_ids:
            await db_execute(
                "UPDATE user_notifications SET is_read=1 WHERE id=? AND user_id=?",
                (nid, uid),
            )
    else:
        await db_execute("UPDATE user_notifications SET is_read=1 WHERE user_id=?", (uid,))
    return {"success": True}

@app.delete("/notifications/{notif_id}")
async def delete_notification(notif_id: str, user: Dict = Depends(get_current_user)):
    await db_execute("DELETE FROM user_notifications WHERE id=? AND user_id=?", (notif_id, user["sub"]))
    return {"success": True}

# ══════════════════════════════════════════════════════════════════════════════
#  SYSTEM ANNOUNCEMENTS (NEW)
# ══════════════════════════════════════════════════════════════════════════════

@app.get("/announcements")
async def get_announcements(user: Dict = Depends(get_current_user)):
    """Active announcements visible to this user's subscription tier."""
    tier = await get_user_subscription(user["sub"])
    now  = datetime.utcnow().isoformat()
    rows = await db_fetchall(
        "SELECT * FROM system_announcements WHERE is_active=1 AND (expires_at IS NULL OR expires_at > ?) AND (target='all' OR target=?) ORDER BY created_at DESC LIMIT 20",
        (now, tier),
    )
    return {"announcements": rows}

@app.post("/admin/announcements")
async def create_announcement(req: AnnouncementCreate, admin: Dict = Depends(require_admin)):
    aid = str(uuid.uuid4()); now = datetime.utcnow().isoformat()
    await db_execute(
        "INSERT INTO system_announcements (id,title,body,type,target,is_active,created_by,expires_at,created_at,updated_at) VALUES (?,?,?,?,?,?,?,?,?,?)",
        (aid, req.title, req.body, req.type, req.target, 1 if req.is_active else 0, admin["sub"], req.expires_at, now, now),
    )
    # If targeting "all" or a specific tier, push notifications to relevant users
    if req.target == "all":
        users = await db_fetchall("SELECT id FROM users WHERE account_status='active' AND deleted_at IS NULL")
    else:
        users = await db_fetchall("SELECT id FROM users WHERE subscription=? AND account_status='active' AND deleted_at IS NULL", (req.target,))
    for u in users:
        await send_notification(u["id"], req.title, req.body[:300], req.type)
    return {"success": True, "id": aid, "notified_users": len(users)}

@app.get("/admin/announcements")
async def list_announcements(admin: Dict = Depends(require_admin)):
    rows = await db_fetchall("SELECT * FROM system_announcements ORDER BY created_at DESC")
    return {"announcements": rows}

@app.put("/admin/announcements/{aid}")
async def update_announcement(aid: str, req: AnnouncementCreate, admin: Dict = Depends(require_admin)):
    row = await db_fetchone("SELECT id FROM system_announcements WHERE id=?", (aid,))
    if not row: raise HTTPException(404, "Announcement not found")
    now = datetime.utcnow().isoformat()
    await db_execute(
        "UPDATE system_announcements SET title=?,body=?,type=?,target=?,is_active=?,expires_at=?,updated_at=? WHERE id=?",
        (req.title, req.body, req.type, req.target, 1 if req.is_active else 0, req.expires_at, now, aid),
    )
    return {"success": True}

@app.delete("/admin/announcements/{aid}")
async def delete_announcement(aid: str, admin: Dict = Depends(require_admin)):
    await db_execute("DELETE FROM system_announcements WHERE id=?", (aid,))
    return {"success": True}

# ══════════════════════════════════════════════════════════════════════════════
#  SUBSCRIPTION MANAGEMENT (v7.0 preserved + enhanced)
# ══════════════════════════════════════════════════════════════════════════════

@app.post("/subscription/request-upgrade")
async def request_upgrade(req: UpgradeRequest, user: Dict = Depends(get_current_user)):
    uid = user["sub"]
    await check_account_active(uid)
    now = datetime.utcnow().isoformat()
    await db_execute(
        "UPDATE users SET upgrade_requested=?, upgrade_requested_at=?, upgrade_note=?, updated_at=? WHERE id=?",
        (req.requested_tier, now, req.note or "", now, uid),
    )
    await _write_audit_log(
        user_id=uid, changed_by=uid, changed_by_email=user.get("email",""),
        action="upgrade_request", old_value="", new_value=req.requested_tier, note=req.note or "",
    )
    return {"success":True,"message":f"Upgrade request to '{req.requested_tier}' submitted."}

@app.get("/subscription/my-history")
async def my_subscription_history(user: Dict = Depends(get_current_user)):
    uid  = user["sub"]
    rows = await db_fetchall(
        "SELECT action,old_value,new_value,note,created_at FROM subscription_audit_log WHERE user_id=? ORDER BY created_at DESC LIMIT 50",
        (uid,),
    )
    return {"history": rows}

@app.put("/admin/subscription")
async def admin_set_subscription(req: SubscriptionUpdate, admin: Dict = Depends(require_admin)):
    user_row = await db_fetchone("SELECT subscription, email FROM users WHERE id=?", (req.user_id,))
    if not user_row: raise HTTPException(404,"User not found")
    all_tiers = list(SUBSCRIPTION_LIMITS.keys())
    custom_plans = await db_fetchall("SELECT name FROM custom_plans")
    all_tiers += [p["name"] for p in custom_plans]
    if req.tier not in all_tiers:
        raise HTTPException(400, f"Unknown tier '{req.tier}'. Available: {all_tiers}")
    now = datetime.utcnow().isoformat()
    await db_execute(
        "UPDATE users SET subscription=?, subscription_expires_at=?, subscription_changed_by=?, subscription_changed_at=?, upgrade_requested=NULL, updated_at=? WHERE id=?",
        (req.tier, req.expires_at, admin["sub"], now, now, req.user_id),
    )
    await _write_audit_log(
        user_id=req.user_id, changed_by=admin["sub"], changed_by_email=admin.get("email",""),
        action="grant", old_value=user_row["subscription"], new_value=req.tier, note=req.note or "",
    )
    await send_notification(
        req.user_id,
        "Subscription Updated",
        f"Your plan has been changed to {req.tier.upper()}" + (f" (expires {req.expires_at})" if req.expires_at else ""),
        "success",
    )
    return {"success":True,"user_id":req.user_id,"email":user_row["email"],"tier":req.tier,"expires_at":req.expires_at}

@app.put("/admin/subscription/bulk")
async def admin_bulk_subscription(req: BulkSubscriptionUpdate, admin: Dict = Depends(require_admin)):
    all_tiers = list(SUBSCRIPTION_LIMITS.keys()) + [p["name"] for p in await db_fetchall("SELECT name FROM custom_plans")]
    if req.tier not in all_tiers:
        raise HTTPException(400, f"Unknown tier '{req.tier}'")
    now = datetime.utcnow().isoformat()
    results = []
    for uid in req.user_ids:
        user_row = await db_fetchone("SELECT subscription, email FROM users WHERE id=?", (uid,))
        if not user_row:
            results.append({"user_id":uid,"success":False,"error":"Not found"}); continue
        await db_execute(
            "UPDATE users SET subscription=?, subscription_expires_at=?, subscription_changed_by=?, subscription_changed_at=?, updated_at=? WHERE id=?",
            (req.tier, req.expires_at, admin["sub"], now, now, uid),
        )
        await _write_audit_log(
            user_id=uid, changed_by=admin["sub"], changed_by_email=admin.get("email",""),
            action="grant", old_value=user_row["subscription"], new_value=req.tier, note=f"[Bulk] {req.note or ''}",
        )
        await send_notification(uid, "Subscription Updated", f"Your plan is now {req.tier.upper()}", "success")
        results.append({"user_id":uid,"email":user_row["email"],"success":True,"tier":req.tier})
    return {"success":True,"applied":len([r for r in results if r["success"]]),"results":results}

@app.put("/admin/account-status")
async def admin_set_account_status(req: AccountStatusUpdate, admin: Dict = Depends(require_admin)):
    user_row = await db_fetchone("SELECT account_status, email FROM users WHERE id=?", (req.user_id,))
    if not user_row: raise HTTPException(404,"User not found")
    now = datetime.utcnow().isoformat()
    if req.status in ("paused","banned"):
        await db_execute(
            "UPDATE users SET account_status=?, pause_reason=?, updated_at=? WHERE id=?",
            (req.status, req.reason or "", now, req.user_id),
        )
    else:
        await db_execute(
            "UPDATE users SET account_status='active', pause_reason=NULL, updated_at=? WHERE id=?",
            (now, req.user_id),
        )
    action = {"active":"unpause","paused":"pause","banned":"ban"}[req.status]
    await _write_audit_log(
        user_id=req.user_id, changed_by=admin["sub"], changed_by_email=admin.get("email",""),
        action=action, old_value=user_row["account_status"], new_value=req.status, note=req.reason or "",
    )
    notif_msgs = {"paused": "Your account has been temporarily paused.", "banned": "Your account has been suspended.", "active": "Your account has been reinstated."}
    await send_notification(req.user_id, "Account Status Changed", notif_msgs.get(req.status,""), "warning" if req.status != "active" else "success")
    return {"success":True,"user_id":req.user_id,"email":user_row["email"],"account_status":req.status}

@app.put("/admin/subscription/revoke")
async def admin_revoke_subscription(user_id: str, note: str = "", admin: Dict = Depends(require_admin)):
    user_row = await db_fetchone("SELECT subscription, email FROM users WHERE id=?", (user_id,))
    if not user_row: raise HTTPException(404,"User not found")
    now = datetime.utcnow().isoformat()
    await db_execute(
        "UPDATE users SET subscription='free', subscription_expires_at=NULL, subscription_changed_by=?, subscription_changed_at=?, updated_at=? WHERE id=?",
        (admin["sub"], now, now, user_id),
    )
    await _write_audit_log(
        user_id=user_id, changed_by=admin["sub"], changed_by_email=admin.get("email",""),
        action="revoke", old_value=user_row["subscription"], new_value="free", note=note,
    )
    await send_notification(user_id, "Subscription Revoked", "Your subscription has been revoked. You are now on the Free plan.", "warning")
    return {"success":True,"user_id":user_id,"email":user_row["email"],"subscription":"free"}

@app.get("/admin/subscription/audit-log")
async def admin_audit_log(
    user_id: Optional[str] = Query(None),
    limit:   int           = Query(100, le=500),
    admin:   Dict          = Depends(require_admin),
):
    if user_id:
        rows = await db_fetchall(
            "SELECT l.*,u.email FROM subscription_audit_log l LEFT JOIN users u ON l.user_id=u.id WHERE l.user_id=? ORDER BY l.created_at DESC LIMIT ?",
            (user_id, limit),
        )
    else:
        rows = await db_fetchall(
            "SELECT l.*,u.email FROM subscription_audit_log l LEFT JOIN users u ON l.user_id=u.id ORDER BY l.created_at DESC LIMIT ?",
            (limit,),
        )
    return {"audit_log": rows, "count": len(rows)}

@app.get("/admin/subscription/pending-upgrades")
async def admin_pending_upgrades(admin: Dict = Depends(require_admin)):
    rows = await db_fetchall(
        "SELECT id,email,full_name,subscription,upgrade_requested,upgrade_requested_at,upgrade_note FROM users WHERE upgrade_requested IS NOT NULL ORDER BY upgrade_requested_at ASC"
    )
    return {"pending": rows, "count": len(rows)}

@app.post("/admin/subscription/decide-upgrade")
async def admin_decide_upgrade(req: UpgradeDecision, admin: Dict = Depends(require_admin)):
    user_row = await db_fetchone("SELECT subscription, email, upgrade_requested FROM users WHERE id=?", (req.user_id,))
    if not user_row: raise HTTPException(404,"User not found")
    if not user_row.get("upgrade_requested"): raise HTTPException(400,"No upgrade request pending")
    now = datetime.utcnow().isoformat()
    if req.approve:
        new_tier = user_row["upgrade_requested"]
        await db_execute(
            "UPDATE users SET subscription=?, upgrade_requested=NULL, upgrade_requested_at=NULL, upgrade_note=NULL, subscription_changed_by=?, subscription_changed_at=?, updated_at=? WHERE id=?",
            (new_tier, admin["sub"], now, now, req.user_id),
        )
        await _write_audit_log(user_id=req.user_id,changed_by=admin["sub"],changed_by_email=admin.get("email",""),action="upgrade_approve",old_value=user_row["subscription"],new_value=new_tier,note=req.note or "")
        await send_notification(req.user_id, "Upgrade Approved! 🎉", f"Your plan has been upgraded to {new_tier.upper()}.", "success")
        return {"success":True,"approved":True,"user_id":req.user_id,"email":user_row["email"],"new_tier":new_tier}
    else:
        await db_execute("UPDATE users SET upgrade_requested=NULL, upgrade_requested_at=NULL, upgrade_note=NULL, updated_at=? WHERE id=?",(now,req.user_id))
        await _write_audit_log(user_id=req.user_id,changed_by=admin["sub"],changed_by_email=admin.get("email",""),action="upgrade_reject",old_value=user_row["subscription"],new_value=user_row["subscription"],note=req.note or "")
        await send_notification(req.user_id,"Upgrade Request Declined",req.note or "Your upgrade request was not approved at this time.","warning")
        return {"success":True,"approved":False,"user_id":req.user_id,"email":user_row["email"]}

@app.get("/admin/subscription/expiring-soon")
async def admin_expiring_soon(days: int = Query(7, ge=1, le=90), admin: Dict = Depends(require_admin)):
    cutoff = (datetime.utcnow() + timedelta(days=days)).isoformat()
    rows   = await db_fetchall(
        "SELECT id,email,full_name,subscription,subscription_expires_at FROM users WHERE subscription_expires_at IS NOT NULL AND subscription_expires_at <= ? AND subscription != 'free' ORDER BY subscription_expires_at ASC",
        (cutoff,),
    )
    return {"expiring_within_days": days, "users": rows, "count": len(rows)}

# ── Custom plans ──────────────────────────────────────────────────────────────

@app.post("/admin/plans")
async def create_custom_plan(req: CustomPlanCreate, admin: Dict = Depends(require_admin)):
    if await db_fetchone("SELECT id FROM custom_plans WHERE name=?", (req.name,)):
        raise HTTPException(400, f"Plan '{req.name}' already exists")
    if req.name in SUBSCRIPTION_LIMITS: raise HTTPException(400, f"'{req.name}' is a built-in tier")
    pid = str(uuid.uuid4()); now = datetime.utcnow().isoformat()
    await db_execute(
        "INSERT INTO custom_plans (id,name,display_name,messages_per_day,documents,max_file_mb,mysql_access,image_analysis,agent_jobs,websites,code_runs_per_day,price_note,created_by,created_at,updated_at) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
        (pid,req.name,req.display_name,req.messages_per_day,req.documents,req.max_file_mb,1 if req.mysql_access else 0,1 if req.image_analysis else 0,req.agent_jobs,req.websites,req.code_runs_per_day,req.price_note,admin["sub"],now,now),
    )
    return {"success":True,"plan_id":pid,"name":req.name,"display_name":req.display_name}

@app.get("/admin/plans")
async def list_custom_plans(admin: Dict = Depends(require_admin)):
    rows     = await db_fetchall("SELECT * FROM custom_plans ORDER BY created_at DESC")
    built_in = [{"name":k,"display_name":k.capitalize(),"built_in":True,**v} for k,v in SUBSCRIPTION_LIMITS.items()]
    return {"built_in_plans": built_in, "custom_plans": rows}

@app.get("/plans")
async def list_all_plans():
    custom   = await db_fetchall("SELECT name,display_name,messages_per_day,documents,max_file_mb,mysql_access,image_analysis,agent_jobs,websites,code_runs_per_day,price_note FROM custom_plans ORDER BY messages_per_day ASC")
    built_in = [{"name":k,"display_name":k.capitalize(),"built_in":True,**{kk:vv for kk,vv in v.items()}} for k,v in SUBSCRIPTION_LIMITS.items() if k != "paused"]
    return {"plans": built_in + [dict(r) for r in custom]}

@app.put("/admin/plans/{plan_name}")
async def update_custom_plan(plan_name: str, req: CustomPlanUpdate, admin: Dict = Depends(require_admin)):
    if not await db_fetchone("SELECT id FROM custom_plans WHERE name=?", (plan_name,)):
        raise HTTPException(404,"Custom plan not found")
    updates = {k:v for k,v in req.model_dump(exclude_none=True).items()}
    if "mysql_access"   in updates: updates["mysql_access"]   = 1 if updates["mysql_access"] else 0
    if "image_analysis" in updates: updates["image_analysis"] = 1 if updates["image_analysis"] else 0
    if not updates: raise HTTPException(400,"Nothing to update")
    updates["updated_at"] = datetime.utcnow().isoformat()
    sets = ", ".join(f"{k}=?" for k in updates)
    await db_execute(f"UPDATE custom_plans SET {sets} WHERE name=?", tuple(updates.values())+(plan_name,))
    return {"success":True,"plan":plan_name}

@app.delete("/admin/plans/{plan_name}")
async def delete_custom_plan(plan_name: str, admin: Dict = Depends(require_admin)):
    if not await db_fetchone("SELECT id FROM custom_plans WHERE name=?", (plan_name,)):
        raise HTTPException(404,"Custom plan not found")
    users_on_plan = await db_count("SELECT COUNT(*) FROM users WHERE subscription=?", (plan_name,))
    await db_execute("DELETE FROM custom_plans WHERE name=?", (plan_name,))
    return {"success":True,"plan_deleted":plan_name,"users_affected":users_on_plan}

# ── Admin user management ─────────────────────────────────────────────────────

@app.get("/admin/users")
async def admin_users(
    status:       Optional[str] = Query(None),
    subscription: Optional[str] = Query(None),
    search:       Optional[str] = Query(None),
    limit:        int           = Query(100, le=500),
    offset:       int           = Query(0, ge=0),
    admin:        Dict          = Depends(require_admin),
):
    where_clauses = ["deleted_at IS NULL"]
    params        = []
    if status:
        where_clauses.append("account_status=?"); params.append(status)
    if subscription:
        where_clauses.append("subscription=?");   params.append(subscription)
    if search:
        where_clauses.append("(email LIKE ? OR full_name LIKE ?)"); params += [f"%{search}%", f"%{search}%"]
    where_sql = "WHERE " + " AND ".join(where_clauses)
    rows = await db_fetchall(
        f"SELECT id,email,full_name,role,subscription,account_status,memory_enabled,subscription_expires_at,upgrade_requested,created_at FROM users {where_sql} ORDER BY created_at DESC LIMIT ? OFFSET ?",
        tuple(params) + (limit, offset),
    )
    total = await db_count(f"SELECT COUNT(*) FROM users {where_sql}", tuple(params))
    for r in rows:
        r["memory_enabled"] = bool(r.get("memory_enabled",1))
    return {"users": rows, "total": total, "limit": limit, "offset": offset}

@app.get("/admin/users/{user_id}")
async def admin_get_user(user_id: str, admin: Dict = Depends(require_admin)):
    row = await db_fetchone(
        "SELECT id,email,full_name,role,subscription,account_status,memory_enabled,subscription_expires_at,pause_reason,upgrade_requested,upgrade_requested_at,upgrade_note,subscription_changed_by,subscription_changed_at,created_at,updated_at FROM users WHERE id=?",
        (user_id,),
    )
    if not row: raise HTTPException(404,"User not found")
    row["memory_enabled"]        = bool(row.get("memory_enabled",1))
    row["subscription_history"]  = await db_fetchall("SELECT action,old_value,new_value,note,changed_by_email,created_at FROM subscription_audit_log WHERE user_id=? ORDER BY created_at DESC LIMIT 20",(user_id,))
    row["chat_count"]            = await db_count("SELECT COUNT(*) FROM chat_history WHERE user_id=?", (user_id,))
    row["memory_count"]          = await db_count("SELECT COUNT(*) FROM user_memories WHERE user_id=?", (user_id,))
    row["doc_count"]             = await db_count("SELECT COUNT(*) FROM documents WHERE user_id=?", (user_id,))
    row["website_count"]         = await db_count("SELECT COUNT(*) FROM websites WHERE user_id=?", (user_id,))
    row["session_count"]         = await db_count("SELECT COUNT(*) FROM chat_sessions WHERE user_id=?", (user_id,))
    return row

@app.post("/admin/impersonate/{user_id}")
async def admin_impersonate(user_id: str, admin: Dict = Depends(require_admin)):
    """
    Admin: generate a short-lived JWT scoped to another user for debugging.
    Token expires in 30 minutes and carries an 'impersonated_by' claim.
    """
    target = await db_fetchone("SELECT id,email,role FROM users WHERE id=?", (user_id,))
    if not target: raise HTTPException(404,"User not found")
    if target["role"] == "admin": raise HTTPException(400,"Cannot impersonate another admin")
    token = create_access_token(
        {"sub": target["id"], "email": target["email"], "role": target["role"], "impersonated_by": admin["sub"]},
        expires_delta=timedelta(minutes=30),
    )
    await log_activity(admin["sub"], "impersonate", resource=user_id, detail=f"Admin {admin.get('email','')} impersonated user {target['email']}")
    logger.warning("[IMPERSONATE] Admin %s impersonated user %s", admin.get("email",""), target["email"])
    return {
        "access_token":     token,
        "expires_minutes":  30,
        "target_user_id":   target["id"],
        "target_email":     target["email"],
        "warning":          "This token expires in 30 minutes and is logged.",
    }

@app.put("/admin/user-role")
async def admin_update_role(user_id: str, role: str, admin: Dict = Depends(require_admin)):
    if role not in ("admin","client"): raise HTTPException(400,"Role must be 'admin' or 'client'")
    await db_execute("UPDATE users SET role=?,updated_at=? WHERE id=?", (role, datetime.utcnow().isoformat(), user_id))
    return {"success":True}

@app.delete("/admin/user/{user_id}")
async def admin_delete_user(user_id: str, admin: Dict = Depends(require_admin)):
    if user_id == admin["sub"]: raise HTTPException(400,"Cannot delete your own account")
    await db_execute("DELETE FROM users WHERE id=?", (user_id,))
    return {"success":True}

@app.get("/admin/user-activity")
async def admin_user_activity(
    user_id: Optional[str] = Query(None),
    limit:   int           = Query(200, le=1000),
    admin:   Dict          = Depends(require_admin),
):
    if user_id:
        rows = await db_fetchall(
            "SELECT l.*,u.email FROM user_activity_log l LEFT JOIN users u ON l.user_id=u.id WHERE l.user_id=? ORDER BY l.created_at DESC LIMIT ?",
            (user_id, limit),
        )
    else:
        rows = await db_fetchall(
            "SELECT l.*,u.email FROM user_activity_log l LEFT JOIN users u ON l.user_id=u.id ORDER BY l.created_at DESC LIMIT ?",
            (limit,),
        )
    return {"activity": rows, "count": len(rows)}

@app.get("/admin/stats")
async def admin_stats(admin: Dict = Depends(require_admin)):
    today            = datetime.utcnow().strftime("%Y-%m-%d")
    sub_breakdown    = await db_fetchall(
        "SELECT subscription, COUNT(*) as count, account_status FROM users WHERE deleted_at IS NULL GROUP BY subscription, account_status ORDER BY count DESC"
    )
    return {
        "total_users":            await db_count("SELECT COUNT(*) FROM users WHERE deleted_at IS NULL"),
        "active_users":           await db_count("SELECT COUNT(*) FROM users WHERE account_status='active' AND deleted_at IS NULL"),
        "paused_users":           await db_count("SELECT COUNT(*) FROM users WHERE account_status='paused'"),
        "banned_users":           await db_count("SELECT COUNT(*) FROM users WHERE account_status='banned'"),
        "deleted_users":          await db_count("SELECT COUNT(*) FROM users WHERE deleted_at IS NOT NULL"),
        "pending_upgrades":       await db_count("SELECT COUNT(*) FROM users WHERE upgrade_requested IS NOT NULL"),
        "expiring_this_week":     await db_count("SELECT COUNT(*) FROM users WHERE subscription_expires_at IS NOT NULL AND subscription_expires_at <= ?", ((datetime.utcnow()+timedelta(days=7)).isoformat(),)),
        "subscription_breakdown": sub_breakdown,
        "total_documents":        await db_count("SELECT COUNT(*) FROM documents"),
        "chats_today":            await db_count("SELECT COUNT(*) FROM chat_history WHERE created_at >= ?", (today,)),
        "chat_sessions_total":    await db_count("SELECT COUNT(*) FROM chat_sessions"),
        "total_memories":         await db_count("SELECT COUNT(*) FROM user_memories"),
        "total_agent_jobs":       await db_count("SELECT COUNT(*) FROM agent_jobs"),
        "total_websites":         await db_count("SELECT COUNT(*) FROM websites"),
        "code_runs_today":        await db_count("SELECT COUNT(*) FROM code_runs WHERE created_at >= ?", (today,)),
        "sub_audit_entries":      await db_count("SELECT COUNT(*) FROM subscription_audit_log"),
        "custom_plans":           await db_count("SELECT COUNT(*) FROM custom_plans"),
        "active_api_keys":        await db_count("SELECT COUNT(*) FROM api_keys WHERE is_active=1"),
        "unread_notifications":   await db_count("SELECT COUNT(*) FROM user_notifications WHERE is_read=0"),
        "system_announcements":   await db_count("SELECT COUNT(*) FROM system_announcements WHERE is_active=1"),
        "mysql_enabled":          MYSQL_ENABLED,
        "playwright_available":   _PLAYWRIGHT_AVAILABLE,
        "system_status":          "healthy",
        "version":                "7.1.0",
        "uptime_since":           _SERVER_START.isoformat(),
    }

@app.get("/admin/chat-logs")
async def admin_chat_logs(limit: int = 100, admin: Dict = Depends(require_admin)):
    return {"logs": await db_fetchall("SELECT h.*,u.email FROM chat_history h LEFT JOIN users u ON h.user_id=u.id ORDER BY h.created_at DESC LIMIT ?", (limit,))}

@app.get("/admin/memories")
async def admin_memories(admin: Dict = Depends(require_admin)):
    return {"memories": await db_fetchall("SELECT m.*,u.email FROM user_memories m LEFT JOIN users u ON m.user_id=u.id ORDER BY m.created_at DESC LIMIT 200")}

@app.get("/admin/agent-jobs")
async def admin_all_jobs(admin: Dict = Depends(require_admin)):
    return {"jobs": await db_fetchall("SELECT j.*,u.email FROM agent_jobs j LEFT JOIN users u ON j.user_id=u.id ORDER BY j.created_at DESC")}

@app.get("/logs")
async def get_logs(user: Dict = Depends(require_admin)):
    try: return {"logs": Path("jazz.log").read_text().splitlines()[-200:]}
    except FileNotFoundError: return {"logs":[]}

# ══════════════════════════════════════════════════════════════════════════════
#  MEMORY
# ══════════════════════════════════════════════════════════════════════════════

@app.get("/memory")
async def list_memories(user: Dict = Depends(get_current_user)):
    uid=user["sub"]
    return {"memories":await get_user_memories(uid),"count":await db_count("SELECT COUNT(*) FROM user_memories WHERE user_id=?",(uid,)),"memory_enabled":await is_memory_enabled(uid),"max_memories":MAX_MEMORIES}

@app.post("/memory")
async def create_memory(req: MemoryCreate, user: Dict = Depends(get_current_user)):
    uid=user["sub"]
    if await db_count("SELECT COUNT(*) FROM user_memories WHERE user_id=?",(uid,)) >= MAX_MEMORIES:
        raise HTTPException(400,f"Memory limit ({MAX_MEMORIES}) reached")
    mid=str(uuid.uuid4())
    await db_execute("INSERT INTO user_memories (id,user_id,content,category,source) VALUES (?,?,?,?,?)",(mid,uid,req.content.strip(),req.category or "general","manual"))
    return {"success":True,"memory":await db_fetchone("SELECT * FROM user_memories WHERE id=?",(mid,))}

@app.put("/memory/{memory_id}")
async def update_memory(memory_id: str, req: MemoryUpdate, user: Dict = Depends(get_current_user)):
    row=await db_fetchone("SELECT id FROM user_memories WHERE id=? AND user_id=?",(memory_id,user["sub"]))
    if not row: raise HTTPException(404,"Memory not found")
    await db_execute("UPDATE user_memories SET content=?,updated_at=? WHERE id=?",(req.content.strip(),datetime.utcnow().isoformat(),memory_id))
    return {"success":True,"memory":await db_fetchone("SELECT * FROM user_memories WHERE id=?",(memory_id,))}

@app.delete("/memory/{memory_id}")
async def delete_memory(memory_id: str, user: Dict = Depends(get_current_user)):
    await db_execute("DELETE FROM user_memories WHERE id=? AND user_id=?",(memory_id,user["sub"])); return {"success":True}

@app.delete("/memory")
async def clear_all_memories(user: Dict = Depends(get_current_user)):
    await db_execute("DELETE FROM user_memories WHERE user_id=?",(user["sub"],)); return {"success":True}

@app.put("/memory/settings/toggle")
async def toggle_memory(req: MemorySettingsUpdate, user: Dict = Depends(get_current_user)):
    await db_execute("UPDATE users SET memory_enabled=?,updated_at=? WHERE id=?",(1 if req.enabled else 0,datetime.utcnow().isoformat(),user["sub"]))
    return {"success":True,"memory_enabled":req.enabled}

# ══════════════════════════════════════════════════════════════════════════════
#  IMAGES
# ══════════════════════════════════════════════════════════════════════════════

@app.post("/images/upload")
async def upload_image(file: UploadFile = File(...), user: Dict = Depends(get_current_user)):
    uid = user["sub"]
    await check_account_active(uid)
    if file.content_type not in ALLOWED_IMAGE_TYPES: raise HTTPException(400,f"Unsupported type: {file.content_type}")
    content=await file.read()
    if len(content)>MAX_IMAGE_SIZE: raise HTTPException(413,"Image too large")
    b64=base64.b64encode(content).decode("utf-8")
    return {"success":True,"data_url":f"data:{file.content_type};base64,{b64}","filename":file.filename,"size":len(content),"mime_type":file.content_type}

# ══════════════════════════════════════════════════════════════════════════════
#  CHAT
# ══════════════════════════════════════════════════════════════════════════════

async def _ensure_session(uid: str, session_id: Optional[str], model_type: str) -> Optional[str]:
    """Validate session_id belongs to user, or return None for sessionless chat."""
    if not session_id:
        return None
    row = await db_fetchone("SELECT id FROM chat_sessions WHERE id=? AND user_id=?", (session_id, uid))
    if not row:
        return None  # Invalid session — degrade gracefully
    return session_id

async def _update_session_meta(session_id: str) -> None:
    now = datetime.utcnow().isoformat()
    await db_execute(
        "UPDATE chat_sessions SET message_count=message_count+1, last_message_at=?, updated_at=? WHERE id=?",
        (now, now, session_id),
    )

@app.post("/chat", response_model=ChatResponse)
async def chat(req: ChatRequest, user: Dict = Depends(get_current_user)):
    t0=time.time(); uid=user["sub"]
    await check_account_active(uid)
    tier=await get_user_subscription(uid)
    if not rate_limiter.check(uid,tier): raise HTTPException(429,"Daily message limit reached. Upgrade your plan.")
    session_id = await _ensure_session(uid, req.session_id, req.model_type)
    mem_on=await is_memory_enabled(uid); memories=await get_user_memories(uid) if mem_on else []; memory_block=format_memories_for_prompt(memories); has_image=bool(req.image_url)
    if has_image:
        try:
            text,_=await analyze_image_pipeline(req.image_url,req.message,req.model_type,uid,memories)
            new_memories=await extract_and_save_memories(uid,req.message,text)
            chat_id = str(uuid.uuid4())
            await db_insert("INSERT INTO chat_history (id,user_id,session_id,message,response,query_type,agent_used,model_used,response_time_ms) VALUES (?,?,?,?,?,?,?,?,?)",(chat_id,uid,session_id,req.message,text,"image","vision_agent",req.model_type,int((time.time()-t0)*1000)))
            if session_id: await _update_session_meta(session_id)
            return ChatResponse(response=text,response_time=time.time()-t0,agent_used="vision_agent",memory_updated=bool(new_memories),new_memories=new_memories,session_id=session_id)
        except Exception as exc: raise HTTPException(500,f"Image analysis failed: {exc}")
    if is_db_intent(req.message) and MYSQL_ENABLED:
        limits=await _get_effective_limits_async(tier)
        if limits.get("mysql_access"):
            result=await run_nl_to_sql(req.message,uid)
            if result.success:
                text=f"**SQL:**\n```sql\n{result.generated_sql}\n```\n\n**Results** ({result.row_count} rows):\n"+json.dumps(result.results[:20],indent=2,default=str)
                return ChatResponse(response=text,response_time=time.time()-t0,agent_used="data_agent",session_id=session_id)
    agent=route_message(req.message,has_image); context=get_relevant_context(uid,req.message) if req.use_rag else ""
    prompt=agent.build_prompt(req.message,context,memory_block); ai,model=get_model_client(req.model_type)
    completion=ai.chat.completions.create(model=model,messages=[{"role":"user","content":prompt}],max_tokens=1024,temperature=0.7)
    text=completion.choices[0].message.content or ""
    new_memories=await extract_and_save_memories(uid,req.message,text)
    chat_id = str(uuid.uuid4())
    await db_insert("INSERT INTO chat_history (id,user_id,session_id,message,response,query_type,agent_used,model_used,response_time_ms) VALUES (?,?,?,?,?,?,?,?,?)",(chat_id,uid,session_id,req.message,text,"rag" if context else "chat",agent.name,req.model_type,int((time.time()-t0)*1000)))
    if session_id: await _update_session_meta(session_id)
    return ChatResponse(response=text,response_time=time.time()-t0,agent_used=agent.name,memory_updated=bool(new_memories),new_memories=new_memories,session_id=session_id)

@app.post("/chat-stream")
async def chat_stream(req: ChatRequest, user: Dict = Depends(get_current_user)):
    t0=time.time(); uid=user["sub"]
    await check_account_active(uid)
    tier=await get_user_subscription(uid)
    if not rate_limiter.check(uid,tier): raise HTTPException(429,"Daily limit reached.")
    session_id = await _ensure_session(uid, req.session_id, req.model_type)
    mem_on=await is_memory_enabled(uid); memories=await get_user_memories(uid) if mem_on else []; memory_block=format_memories_for_prompt(memories); has_image=bool(req.image_url)
    _sse_headers={"Cache-Control":"no-cache","X-Accel-Buffering":"no"}
    if has_image:
        async def _image_stream():
            yield f"data: {json.dumps({'type':'agent','agent':'vision_agent'})}\n\n"
            yield f"data: {json.dumps({'type':'status','message':'🔍 Analyzing image…'})}\n\n"
            try:
                text,_=await analyze_image_pipeline(req.image_url,req.message,req.model_type,uid,memories)
                words=text.split()
                for i in range(0,len(words),8):
                    chunk=" ".join(words[i:i+8])+(" " if i+8<len(words) else "")
                    yield f"data: {json.dumps({'type':'chunk','content':chunk})}\n\n"
                    await asyncio.sleep(0.01)
                new_memories=await extract_and_save_memories(uid,req.message,text)
                await db_insert("INSERT INTO chat_history (id,user_id,session_id,message,response,query_type,agent_used,model_used,response_time_ms) VALUES (?,?,?,?,?,?,?,?,?)",(str(uuid.uuid4()),uid,session_id,req.message,text,"image","vision_agent",req.model_type,int((time.time()-t0)*1000)))
                if session_id: await _update_session_meta(session_id)
                yield f"data: {json.dumps({'type':'done','response_time':time.time()-t0,'agent':'vision_agent','full_response':text,'memory_updated':bool(new_memories),'new_memories':new_memories,'session_id':session_id})}\n\n"
            except Exception as exc:
                yield f"data: {json.dumps({'type':'error','message':str(exc)})}\n\n"
        return StreamingResponse(_image_stream(),media_type="text/event-stream",headers=_sse_headers)
    agent=route_message(req.message,has_image); context=get_relevant_context(uid,req.message) if req.use_rag else ""
    prompt=agent.build_prompt(req.message,context,memory_block); ai,model=get_model_client(req.model_type)
    try: stream=ai.chat.completions.create(model=model,messages=[{"role":"user","content":prompt}],max_tokens=1024,temperature=0.7,stream=True)
    except Exception as exc: raise HTTPException(503,f"AI service error: {exc}")
    async def _generate():
        full=""
        try:
            yield f"data: {json.dumps({'type':'agent','agent':agent.name})}\n\n"
            for chunk in stream:
                if chunk.choices and chunk.choices[0].delta.content:
                    content=chunk.choices[0].delta.content; full+=content
                    yield f"data: {json.dumps({'type':'chunk','content':content})}\n\n"
            elapsed=time.time()-t0; new_memories=await extract_and_save_memories(uid,req.message,full)
            await db_insert("INSERT INTO chat_history (id,user_id,session_id,message,response,query_type,agent_used,model_used,response_time_ms) VALUES (?,?,?,?,?,?,?,?,?)",(str(uuid.uuid4()),uid,session_id,req.message,full,"rag" if context else "chat",agent.name,req.model_type,int(elapsed*1000)))
            if session_id: await _update_session_meta(session_id)
            yield f"data: {json.dumps({'type':'done','response_time':elapsed,'agent':agent.name,'full_response':full,'memory_updated':bool(new_memories),'new_memories':new_memories,'session_id':session_id})}\n\n"
        except Exception as exc:
            yield f"data: {json.dumps({'type':'error','message':str(exc)})}\n\n"
    return StreamingResponse(_generate(),media_type="text/event-stream",headers=_sse_headers)

@app.get("/chat/history")
async def chat_history(page: int = 1, per_page: int = 20, user: Dict = Depends(get_current_user)):
    uid=user["sub"]; offset=(page-1)*per_page
    rows=await db_fetchall("SELECT * FROM chat_history WHERE user_id=? AND session_id IS NULL ORDER BY created_at DESC LIMIT ? OFFSET ?",(uid,per_page,offset))
    total=await db_count("SELECT COUNT(*) FROM chat_history WHERE user_id=? AND session_id IS NULL",(uid,))
    return {"history":rows,"page":page,"per_page":per_page,"total":total}

@app.get("/chat/export")
async def export_chat(fmt: str = "markdown", user: Dict = Depends(get_current_user)):
    uid=user["sub"]; rows=await db_fetchall("SELECT * FROM chat_history WHERE user_id=? ORDER BY created_at ASC",(uid,))
    if fmt=="json":
        content=json.dumps(rows,indent=2,default=str); media_type="application/json"; filename=f"jazz_chat_{uid[:8]}.json"
    else:
        lines=[f"# JAZZ AI Chat Export\n_Exported: {datetime.utcnow().isoformat()}_\n"]
        for r in rows:
            lines.append(f"\n---\n**[{r.get('created_at','')}]** `{r.get('agent_used','')}`\n")
            lines.append(f"**You:** {r['message']}\n\n**JAZZ:** {r['response']}\n")
        content="\n".join(lines); media_type="text/markdown"; filename=f"jazz_chat_{uid[:8]}.md"
    return StreamingResponse(iter([content.encode()]),media_type=media_type,headers={"Content-Disposition":f'attachment; filename="{filename}"'})

# ══════════════════════════════════════════════════════════════════════════════
#  DOCUMENTS
# ══════════════════════════════════════════════════════════════════════════════

@app.post("/documents/upload")
async def upload_document(file: UploadFile = File(...), user: Dict = Depends(get_current_user)):
    if not documents_collection: raise HTTPException(503,"Vector DB unavailable")
    uid=user["sub"]
    await check_account_active(uid)
    tier=await get_user_subscription(uid); limits=await _get_effective_limits_async(tier)
    if limits["documents"] != -1:
        count=await db_count("SELECT COUNT(*) FROM documents WHERE user_id=?",(uid,))
        if count>=limits["documents"]: raise HTTPException(403,f"Document limit reached for {tier} plan")
    content=await file.read()
    if len(content)>limits["max_file_mb"]*1024*1024: raise HTTPException(413,"File too large")
    tmp=tempfile.mkdtemp()
    try:
        fpath=os.path.join(tmp,file.filename or "upload")
        with open(fpath,"wb") as f: f.write(content)
        mime=file.content_type or "application/octet-stream"
        text=process_document(fpath,mime)
        if not text.strip(): raise HTTPException(400,"Could not extract text")
        chunks=chunk_text(text); doc_id=str(uuid.uuid4())
        documents_collection.add(documents=chunks,ids=[f"{doc_id}_{i}" for i in range(len(chunks))],metadatas=[{"user_id":uid,"filename":file.filename,"doc_id":doc_id} for _ in chunks])
        now=datetime.utcnow().isoformat()
        await db_execute("INSERT INTO documents (id,user_id,filename,original_name,file_type,file_size,chunk_count,chroma_collection,uploaded_at) VALUES (?,?,?,?,?,?,?,?,?)",(doc_id,uid,file.filename or "",file.filename or "",mime,len(content),len(chunks),"user_documents",now))
        await log_activity(uid, "document_upload", resource=doc_id, detail=file.filename or "")
        return DocumentResponse(id=doc_id,filename=file.filename or "",original_name=file.filename or "",file_type=mime,file_size=len(content),chunk_count=len(chunks),uploaded_at=now)
    finally: shutil.rmtree(tmp,ignore_errors=True)

@app.get("/documents")
async def list_documents(user: Dict = Depends(get_current_user)):
    return {"documents":await db_fetchall("SELECT * FROM documents WHERE user_id=? ORDER BY uploaded_at DESC",(user["sub"],))}

@app.delete("/documents/{doc_id}")
async def delete_document(doc_id: str, user: Dict = Depends(get_current_user)):
    uid=user["sub"]
    if documents_collection:
        try:
            existing=documents_collection.get(where={"doc_id":doc_id,"user_id":uid})
            if existing["ids"]: documents_collection.delete(ids=existing["ids"])
        except Exception as exc: logger.warning("[CHROMA] %s",exc)
    await db_execute("DELETE FROM documents WHERE id=? AND user_id=?",(doc_id,uid))
    return {"success":True}

@app.post("/documents/bulk-delete")
async def bulk_delete_documents(req: BulkDocumentDelete, user: Dict = Depends(get_current_user)):
    uid  = user["sub"]
    deleted = []
    for doc_id in req.document_ids:
        row = await db_fetchone("SELECT id FROM documents WHERE id=? AND user_id=?", (doc_id, uid))
        if not row: continue
        if documents_collection:
            try:
                existing = documents_collection.get(where={"doc_id":doc_id,"user_id":uid})
                if existing["ids"]: documents_collection.delete(ids=existing["ids"])
            except Exception: pass
        await db_execute("DELETE FROM documents WHERE id=?", (doc_id,))
        deleted.append(doc_id)
    return {"success": True, "deleted": deleted, "count": len(deleted)}

# ══════════════════════════════════════════════════════════════════════════════
#  WEBSITE BUILDER (with versioning)
# ══════════════════════════════════════════════════════════════════════════════

@app.post("/websites/build")
async def build_website(req: WebsiteBuildRequest, user: Dict = Depends(get_current_user)):
    uid=user["sub"]
    await check_account_active(uid)
    tier=await get_user_subscription(uid); limits=await _get_effective_limits_async(tier)
    wlimit=limits.get("websites",3)
    if wlimit!=-1:
        count=await db_count("SELECT COUNT(*) FROM websites WHERE user_id=?",(uid,))
        if count>=wlimit: raise HTTPException(403,f"Website limit ({wlimit}) reached for {tier} plan")
    title=req.title or f"Site {datetime.utcnow().strftime('%Y%m%d_%H%M%S')}"; site_id=str(uuid.uuid4()); filename=f"{site_id}.html"
    try: html=await _build_website_html(req.description,title,req.style or "modern",req.model_type)
    except Exception as exc: raise HTTPException(500,f"Generation failed: {exc}")
    site_path=SITES_DIR/filename; site_path.write_text(html,encoding="utf-8")
    now=datetime.utcnow().isoformat()
    await db_execute("INSERT INTO websites (id,user_id,title,description,filename,html_size,prompt_used,version,created_at,updated_at) VALUES (?,?,?,?,?,?,?,?,?,?)",(site_id,uid,title,req.description[:500],filename,len(html),req.description,1,now,now))
    await log_activity(uid, "website_build", resource=site_id, detail=title)
    return {"success":True,"id":site_id,"title":title,"filename":filename,"preview_url":f"/preview/{filename}","html_size":len(html),"version":1,"created_at":now}

@app.put("/websites/{site_id}")
async def update_website(site_id: str, req: WebsiteUpdateRequest, user: Dict = Depends(get_current_user)):
    uid=user["sub"]
    row=await db_fetchone("SELECT * FROM websites WHERE id=? AND user_id=?",(site_id,uid))
    if not row: raise HTTPException(404,"Website not found")
    site_path=SITES_DIR/row["filename"]
    if not site_path.exists(): raise HTTPException(404,"Site file missing")
    current_html = site_path.read_text(encoding="utf-8")
    current_ver  = row.get("version", 1)
    # Save current version before overwrite
    await _save_site_version(site_id, uid, current_ver, current_html, f"Before: {req.instructions[:100]}")
    try: new_html=await _update_website_html(current_html,req.instructions,req.model_type)
    except Exception as exc: raise HTTPException(500,f"Update failed: {exc}")
    site_path.write_text(new_html,encoding="utf-8"); now=datetime.utcnow().isoformat()
    new_ver = current_ver + 1
    await db_execute("UPDATE websites SET html_size=?,version=?,updated_at=? WHERE id=?",(len(new_html),new_ver,now,site_id))
    return {"success":True,"id":site_id,"preview_url":f"/preview/{row['filename']}","html_size":len(new_html),"version":new_ver,"updated_at":now}

@app.get("/websites/{site_id}/versions")
async def list_website_versions(site_id: str, user: Dict = Depends(get_current_user)):
    row = await db_fetchone("SELECT id FROM websites WHERE id=? AND user_id=?", (site_id, user["sub"]))
    if not row: raise HTTPException(404, "Website not found")
    versions = await db_fetchall(
        "SELECT id,version,html_size,note,created_at FROM website_versions WHERE site_id=? ORDER BY version DESC",
        (site_id,),
    )
    return {"site_id": site_id, "versions": versions}

@app.post("/websites/{site_id}/restore")
async def restore_website_version(site_id: str, req: WebsiteRestoreRequest, user: Dict = Depends(get_current_user)):
    uid = user["sub"]
    site_row = await db_fetchone("SELECT * FROM websites WHERE id=? AND user_id=?", (site_id, uid))
    if not site_row: raise HTTPException(404, "Website not found")
    ver_row = await db_fetchone(
        "SELECT * FROM website_versions WHERE site_id=? AND version=?",
        (site_id, req.version),
    )
    if not ver_row: raise HTTPException(404, f"Version {req.version} not found")
    ver_path = SITE_VERSIONS_DIR / ver_row["filename"]
    if not ver_path.exists(): raise HTTPException(500, "Version file missing")
    archived_html = ver_path.read_text(encoding="utf-8")
    # Save current before restoring
    current_path = SITES_DIR / site_row["filename"]
    if current_path.exists():
        await _save_site_version(site_id, uid, site_row.get("version",1), current_path.read_text(encoding="utf-8"), "Before restore")
    current_path.write_text(archived_html, encoding="utf-8")
    now = datetime.utcnow().isoformat()
    await db_execute(
        "UPDATE websites SET html_size=?,version=?,updated_at=? WHERE id=?",
        (len(archived_html), site_row.get("version",1)+1, now, site_id),
    )
    return {"success": True, "restored_version": req.version, "preview_url": f"/preview/{site_row['filename']}"}

@app.get("/websites")
async def list_websites(user: Dict = Depends(get_current_user)):
    rows=await db_fetchall("SELECT id,title,description,filename,html_size,version,created_at,updated_at FROM websites WHERE user_id=? ORDER BY created_at DESC",(user["sub"],))
    for r in rows: r["preview_url"]=f"/preview/{r['filename']}"
    return {"websites":rows}

@app.get("/websites/{site_id}")
async def get_website(site_id: str, user: Dict = Depends(get_current_user)):
    row=await db_fetchone("SELECT * FROM websites WHERE id=? AND user_id=?",(site_id,user["sub"]))
    if not row: raise HTTPException(404,"Website not found")
    row["preview_url"]=f"/preview/{row['filename']}"
    site_path=SITES_DIR/row["filename"]
    row["html"]=site_path.read_text(encoding="utf-8") if site_path.exists() else ""
    return row

@app.delete("/websites/{site_id}")
async def delete_website(site_id: str, user: Dict = Depends(get_current_user)):
    uid=user["sub"]
    row=await db_fetchone("SELECT filename FROM websites WHERE id=? AND user_id=?",(site_id,uid))
    if not row: raise HTTPException(404,"Website not found")
    site_path=SITES_DIR/row["filename"]
    if site_path.exists(): site_path.unlink()
    # Clean up versions
    ver_rows = await db_fetchall("SELECT filename FROM website_versions WHERE site_id=?", (site_id,))
    for vr in ver_rows:
        vp = SITE_VERSIONS_DIR / vr["filename"]
        if vp.exists(): vp.unlink(missing_ok=True)
    await db_execute("DELETE FROM website_versions WHERE site_id=?", (site_id,))
    await db_execute("DELETE FROM websites WHERE id=?",(site_id,))
    return {"success":True}

# ══════════════════════════════════════════════════════════════════════════════
#  CODE RUNNER
# ══════════════════════════════════════════════════════════════════════════════

@app.post("/code/run")
async def run_code(req: CodeRunRequest, user: Dict = Depends(get_current_user)):
    uid=user["sub"]
    await check_account_active(uid)
    tier=await get_user_subscription(uid)
    if not rate_limiter.check(uid,tier,"code_runs_per_day"): raise HTTPException(429,"Daily code run limit reached")
    if req.language=="python":
        safe,reason=_is_code_safe(req.code,req.language)
        if not safe: raise HTTPException(400,f"Code blocked: {reason}")
    loop=asyncio.get_event_loop(); result=await loop.run_in_executor(_executor,_run_code_sync,req.code,req.language)
    run_id=str(uuid.uuid4())
    await db_execute("INSERT INTO code_runs (id,user_id,language,code,stdout,stderr,exit_code,duration_ms) VALUES (?,?,?,?,?,?,?,?)",(run_id,uid,req.language,req.code[:5000],result["stdout"],result["stderr"],result["exit_code"],result["duration_ms"]))
    return {"run_id":run_id,"language":req.language,"stdout":result["stdout"],"stderr":result["stderr"],"exit_code":result["exit_code"],"duration_ms":result["duration_ms"],"success":result["exit_code"]==0}

@app.get("/code/history")
async def code_history(limit: int = 20, user: Dict = Depends(get_current_user)):
    return {"runs":await db_fetchall("SELECT id,language,code,stdout,stderr,exit_code,duration_ms,created_at FROM code_runs WHERE user_id=? ORDER BY created_at DESC LIMIT ?",(user["sub"],limit))}

# ══════════════════════════════════════════════════════════════════════════════
#  SQLITE QUERY (admin only)
# ══════════════════════════════════════════════════════════════════════════════

@app.post("/sqlite/query")
async def sqlite_nl_query(req: SQLiteQueryRequest, user: Dict = Depends(require_admin)):
    return await run_sqlite_nl_query(req.question)

@app.post("/sqlite/raw")
async def sqlite_raw_query(req: SQLiteRawRequest, user: Dict = Depends(require_admin)):
    sql_upper=req.sql.upper().strip()
    if req.readonly and not sql_upper.startswith("SELECT") and not sql_upper.startswith("PRAGMA"):
        raise HTTPException(400,"Only SELECT/PRAGMA allowed in readonly mode")
    return await _run_sqlite_raw(req.sql)

@app.get("/sqlite/tables")
async def sqlite_tables(user: Dict = Depends(require_admin)):
    import sqlite3 as _s3
    try:
        with _s3.connect(DB_PATH) as conn:
            tables=[r[0] for r in conn.execute("SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%'").fetchall()]
        result=[]
        for tbl in tables:
            count=await db_count(f"SELECT COUNT(*) FROM `{tbl}`")
            result.append({"table":tbl,"rows":count})
        return {"tables":result}
    except Exception as exc: raise HTTPException(500,str(exc))

# ══════════════════════════════════════════════════════════════════════════════
#  PROMPT LIBRARY
# ══════════════════════════════════════════════════════════════════════════════

@app.post("/prompts")
async def create_prompt(req: PromptCreate, user: Dict = Depends(get_current_user)):
    pid=str(uuid.uuid4()); now=datetime.utcnow().isoformat()
    await db_execute("INSERT INTO prompt_templates (id,user_id,title,content,category,is_public,created_at) VALUES (?,?,?,?,?,?,?)",(pid,user["sub"],req.title,req.content,req.category or "general",1 if req.is_public else 0,now))
    return {"success":True,"prompt":await db_fetchone("SELECT * FROM prompt_templates WHERE id=?",(pid,))}

@app.get("/prompts")
async def list_prompts(user: Dict = Depends(get_current_user)):
    uid=user["sub"]
    rows=await db_fetchall("SELECT p.*,u.full_name as author FROM prompt_templates p LEFT JOIN users u ON p.user_id=u.id WHERE p.user_id=? OR p.is_public=1 ORDER BY p.use_count DESC,p.created_at DESC",(uid,))
    return {"prompts":rows}

@app.put("/prompts/{prompt_id}")
async def update_prompt(prompt_id: str, req: PromptUpdate, user: Dict = Depends(get_current_user)):
    row=await db_fetchone("SELECT id FROM prompt_templates WHERE id=? AND user_id=?",(prompt_id,user["sub"]))
    if not row: raise HTTPException(404,"Prompt not found")
    updates={}
    if req.title     is not None: updates["title"]     = req.title
    if req.content   is not None: updates["content"]   = req.content
    if req.category  is not None: updates["category"]  = req.category
    if req.is_public is not None: updates["is_public"] = 1 if req.is_public else 0
    if not updates: raise HTTPException(400,"Nothing to update")
    sets=", ".join(f"{k}=?" for k in updates)
    await db_execute(f"UPDATE prompt_templates SET {sets} WHERE id=?",tuple(updates.values())+(prompt_id,))
    return {"success":True,"prompt":await db_fetchone("SELECT * FROM prompt_templates WHERE id=?",(prompt_id,))}

@app.delete("/prompts/{prompt_id}")
async def delete_prompt(prompt_id: str, user: Dict = Depends(get_current_user)):
    await db_execute("DELETE FROM prompt_templates WHERE id=? AND user_id=?",(prompt_id,user["sub"])); return {"success":True}

@app.post("/prompts/{prompt_id}/use")
async def use_prompt(prompt_id: str, user: Dict = Depends(get_current_user)):
    row=await db_fetchone("SELECT * FROM prompt_templates WHERE id=?",(prompt_id,))
    if not row: raise HTTPException(404,"Prompt not found")
    await db_execute("UPDATE prompt_templates SET use_count=use_count+1 WHERE id=?",(prompt_id,))
    return {"content":row["content"],"title":row["title"]}

# ══════════════════════════════════════════════════════════════════════════════
#  AGENT JOBS
# ══════════════════════════════════════════════════════════════════════════════

@app.post("/agent-jobs")
async def create_agent_job(req: AgentJobCreate, user: Dict = Depends(get_current_user)):
    uid=user["sub"]
    await check_account_active(uid)
    tier=await get_user_subscription(uid); limits=await _get_effective_limits_async(tier); job_limit=limits.get("agent_jobs",1)
    if job_limit!=-1:
        count=await db_count("SELECT COUNT(*) FROM agent_jobs WHERE user_id=?",(uid,))
        if count>=job_limit: raise HTTPException(403,f"Agent job limit ({job_limit}) reached for {tier} plan")
    if len(req.cron_schedule.strip().split())!=5: raise HTTPException(400,"cron_schedule must be 5 fields")
    job_id=str(uuid.uuid4()); now_iso=datetime.utcnow().isoformat()
    creds_enc=encrypt_creds(req.credentials); params_js=json.dumps(req.parameters or {})
    await db_execute("INSERT INTO agent_jobs (id,user_id,name,description,job_type,cron_schedule,time_window_start,time_window_end,credentials_enc,parameters,status,enabled,created_at,updated_at) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)",(job_id,uid,req.name,req.description or "",req.job_type,req.cron_schedule,req.time_window_start or "11:00",req.time_window_end or "12:00",creds_enc,params_js,"idle",1,now_iso,now_iso))
    job=await db_fetchone("SELECT * FROM agent_jobs WHERE id=?",(job_id,)); job.pop("credentials_enc",None)
    _schedule_job({**job,"credentials_enc":creds_enc})
    return {"success":True,"job":job}

@app.get("/agent-jobs")
async def list_agent_jobs(user: Dict = Depends(get_current_user)):
    return {"jobs":await db_fetchall("SELECT id,user_id,name,description,job_type,cron_schedule,time_window_start,time_window_end,parameters,status,enabled,last_run_at,last_run_status,next_run_at,total_runs,total_applied,created_at,updated_at FROM agent_jobs WHERE user_id=? ORDER BY created_at DESC",(user["sub"],))}

@app.get("/agent-jobs/{job_id}")
async def get_agent_job(job_id: str, user: Dict = Depends(get_current_user)):
    row=await db_fetchone("SELECT id,user_id,name,description,job_type,cron_schedule,time_window_start,time_window_end,parameters,status,enabled,last_run_at,last_run_status,next_run_at,total_runs,total_applied,created_at,updated_at FROM agent_jobs WHERE id=? AND user_id=?",(job_id,user["sub"]))
    if not row: raise HTTPException(404,"Job not found")
    return row

@app.put("/agent-jobs/{job_id}")
async def update_agent_job(job_id: str, req: AgentJobUpdate, user: Dict = Depends(get_current_user)):
    existing=await db_fetchone("SELECT * FROM agent_jobs WHERE id=? AND user_id=?",(job_id,user["sub"]))
    if not existing: raise HTTPException(404,"Job not found")
    updates={}
    if req.name              is not None: updates["name"]=req.name
    if req.description       is not None: updates["description"]=req.description
    if req.cron_schedule     is not None:
        if len(req.cron_schedule.split())!=5: raise HTTPException(400,"cron_schedule must be 5 fields")
        updates["cron_schedule"]=req.cron_schedule
    if req.time_window_start is not None: updates["time_window_start"]=req.time_window_start
    if req.time_window_end   is not None: updates["time_window_end"]=req.time_window_end
    if req.parameters        is not None: updates["parameters"]=json.dumps(req.parameters)
    if req.enabled           is not None: updates["enabled"]=1 if req.enabled else 0
    if req.credentials       is not None: updates["credentials_enc"]=encrypt_creds(req.credentials)
    if not updates: raise HTTPException(400,"Nothing to update")
    updates["updated_at"]=datetime.utcnow().isoformat()
    sets=", ".join(f"{k}=?" for k in updates)
    await db_execute(f"UPDATE agent_jobs SET {sets} WHERE id=?",tuple(updates.values())+(job_id,))
    updated=await db_fetchone("SELECT * FROM agent_jobs WHERE id=?",(job_id,))
    if updated.get("enabled"): _schedule_job(updated)
    else: _unschedule_job(job_id)
    updated.pop("credentials_enc",None)
    return {"success":True,"job":updated}

@app.delete("/agent-jobs/{job_id}")
async def delete_agent_job(job_id: str, user: Dict = Depends(get_current_user)):
    existing=await db_fetchone("SELECT id FROM agent_jobs WHERE id=? AND user_id=?",(job_id,user["sub"]))
    if not existing: raise HTTPException(404,"Job not found")
    _unschedule_job(job_id); await db_execute("DELETE FROM agent_jobs WHERE id=?",(job_id,)); return {"success":True}

@app.post("/agent-jobs/{job_id}/run")
async def run_agent_job_now(job_id: str, user: Dict = Depends(get_current_user)):
    existing=await db_fetchone("SELECT id FROM agent_jobs WHERE id=? AND user_id=?",(job_id,user["sub"]))
    if not existing: raise HTTPException(404,"Job not found")
    task=asyncio.create_task(execute_job(job_id,bypass_time_window=True))
    def _on_done(t):
        if t.cancelled(): logger.error("[JOB] CANCELLED %s",job_id)
        elif t.exception(): logger.error("[JOB] CRASHED %s: %s",job_id,t.exception())
    task.add_done_callback(_on_done)
    return {"success":True,"message":"Job triggered — check logs"}

@app.post("/agent-jobs/{job_id}/toggle")
async def toggle_agent_job(job_id: str, user: Dict = Depends(get_current_user)):
    row=await db_fetchone("SELECT * FROM agent_jobs WHERE id=? AND user_id=?",(job_id,user["sub"]))
    if not row: raise HTTPException(404,"Job not found")
    new_enabled=0 if row["enabled"] else 1
    await db_execute("UPDATE agent_jobs SET enabled=?,updated_at=? WHERE id=?",(new_enabled,datetime.utcnow().isoformat(),job_id))
    if new_enabled: _schedule_job({**row,"enabled":new_enabled})
    else: _unschedule_job(job_id)
    return {"success":True,"enabled":bool(new_enabled)}

@app.get("/agent-jobs/{job_id}/logs")
async def get_job_logs(job_id: str, limit: int = 20, user: Dict = Depends(get_current_user)):
    existing=await db_fetchone("SELECT id FROM agent_jobs WHERE id=? AND user_id=?",(job_id,user["sub"]))
    if not existing: raise HTTPException(404,"Job not found")
    rows=await db_fetchall("SELECT id,job_id,status,output,applied_jobs,error,started_at,completed_at,duration_sec FROM agent_job_logs WHERE job_id=? ORDER BY started_at DESC LIMIT ?",(job_id,limit))
    for r in rows:
        try: r["applied_jobs"]=json.loads(r.get("applied_jobs") or "[]")
        except Exception: r["applied_jobs"]=[]
    return {"logs":rows}

# ══════════════════════════════════════════════════════════════════════════════
#  FILE WORKSPACE
# ══════════════════════════════════════════════════════════════════════════════

@app.post("/files")
async def file_operation(req: FileOperationRequest):
    ops={"create":lambda:file_mgr.create(req.filename,req.content or ""),"read":lambda:file_mgr.read(req.filename),"update":lambda:file_mgr.update(req.filename,req.content or ""),"delete":lambda:file_mgr.delete(req.filename),"list":lambda:{"success":True,"files":file_mgr.list_files()}}
    fn=ops.get(req.operation)
    if not fn: raise HTTPException(400,f"Unknown operation: {req.operation}")
    return fn()

# ══════════════════════════════════════════════════════════════════════════════
#  HEALTH
# ══════════════════════════════════════════════════════════════════════════════

@app.get("/health")
async def health():
    db_ok=False
    try: await db_fetchone("SELECT 1"); db_ok=True
    except Exception: pass
    sites_size_mb=sum(f.stat().st_size for f in SITES_DIR.rglob("*") if f.is_file())/(1024*1024)
    return {
        "status":         "healthy" if db_ok else "degraded",
        "timestamp":      datetime.utcnow().isoformat(),
        "uptime_seconds": int((datetime.utcnow()-_SERVER_START).total_seconds()),
        "services":       {"sqlite":db_ok,"chromadb":chroma_client is not None,"mysql":_mysql_pool is not None,"playwright":_PLAYWRIGHT_AVAILABLE},
        "storage":        {"sites_mb":round(sites_size_mb,2),"db_path":DB_PATH},
        "agents":         [a.name for a in AGENTS],
        "version":        "7.1.0",
        "features":       [
            "memory","vision","rag","streaming","agent_jobs","website_builder","website_versioning",
            "code_runner","sqlite_query","prompt_library","chat_export","subscription_management",
            "chat_sessions","chat_search","api_keys","notifications","announcements","activity_log",
            "password_reset","gdpr_export","account_self_delete","admin_impersonation","bulk_doc_delete",
        ],
    }

@app.get("/health/history")
async def health_history(limit: int = Query(48, le=288)):
    """Last N health snapshots (default = last 4 hours at 5-min intervals)."""
    rows = await db_fetchall(
        "SELECT ts,sqlite_ok,chroma_ok,mysql_ok,uptime_sec FROM health_history ORDER BY created_at DESC LIMIT ?",
        (limit,),
    )
    return {"history": rows, "interval_minutes": 5}

@app.get("/model-info")
async def model_info():
    return {
        "models":[
            {"id":"censored",   "name":"Groq Llama 3.3 70B",                   "speed":"ultra-fast","filtered":True},
            {"id":"uncensored", "name":"Dolphin Mistral 24B Venice",            "speed":"variable",  "filtered":False},
        ],
        "vision_model":     "Qwen2.5-VL-7B-Instruct",
        "website_styles":   list(_STYLE_HINTS.keys()),
        "code_languages":   ["python","javascript","bash"],
        "version":          "7.1.0",
    }

# ══════════════════════════════════════════════════════════════════════════════
#  Entry point
# ══════════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    import uvicorn
    _module_name = Path(__file__).stem
    uvicorn.run(
        f"{_module_name}:app",
        host      = "0.0.0.0",
        port      = int(os.getenv("PORT", "8000")),
        reload    = False,
        workers   = 1,
        log_level = "info",
        access_log= True,
        timeout_keep_alive = 75,
    )
