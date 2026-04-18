"""
JAZZ AI — v7.0 Production Edition
==================================
Built on v6.0 Obsidian Terminal Edition.

v7.0 Additions:
  ✅ API Key authentication (SHA-256, jz_ prefix)
  ✅ Chat Sessions (group conversations)
  ✅ Notifications system (per-user + broadcasts)
  ✅ System Announcements (admin broadcasts)
  ✅ Coupon / Subscription redemption codes
  ✅ Password change endpoint
  ✅ Account deletion + GDPR full export
  ✅ Admin per-user detail views (chats, memories, docs, jobs, sites)
  ✅ Admin user impersonation
  ✅ Admin targeted notifications
  ✅ Fixed /files endpoint (auth required)
  ✅ Enhanced rate limiter with per-resource tracking
  ✅ Auto-title chat sessions from first message
"""

from __future__ import annotations

import asyncio
import base64
import csv
import hashlib
import json
import logging
import os
import re
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
from fastapi import Depends, FastAPI, File, HTTPException, Request, UploadFile
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

# ──────────────────────────────────────────────────────────────────────────────
# Logging
# ──────────────────────────────────────────────────────────────────────────────

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s — %(message)s",
    handlers=[logging.FileHandler("jazz.log"), logging.StreamHandler()],
)
logger = logging.getLogger("jazz")


# ──────────────────────────────────────────────────────────────────────────────
# Configuration
# ──────────────────────────────────────────────────────────────────────────────

def _env(key: str, default: str = "") -> str:
    val = os.getenv(key, default)
    if not val:
        logger.warning("[CONFIG] %s not set", key)
    return val


SECRET_KEY: str = _env("JWT_SECRET_KEY")
if not SECRET_KEY:
    import secrets as _sec
    SECRET_KEY = _sec.token_urlsafe(64)
    logger.warning("[CONFIG] JWT_SECRET_KEY not set — using ephemeral key (sessions will reset on restart)")

ALGORITHM:                    str  = "HS256"
ACCESS_TOKEN_EXPIRE_MINUTES:  int  = 60 * 24 * 7   # 7 days

_FERNET_KEY_RAW: str = os.getenv("FERNET_KEY", "")
if not _FERNET_KEY_RAW:
    _FERNET_KEY_RAW = Fernet.generate_key().decode()
    logger.warning("[CONFIG] FERNET_KEY not set — add to .env: FERNET_KEY=%s", _FERNET_KEY_RAW)
_fernet = Fernet(_FERNET_KEY_RAW.encode() if isinstance(_FERNET_KEY_RAW, str) else _FERNET_KEY_RAW)


def encrypt_creds(data: dict) -> str:
    return _fernet.encrypt(json.dumps(data).encode()).decode()


def decrypt_creds(token: str) -> dict:
    return json.loads(_fernet.decrypt(token.encode()).decode())


ADMIN_EMAIL:   str = os.getenv("ADMIN_EMAIL", "jasmeet.15069@gmail.com")
ADMIN_PASSWORD: str = _env("ADMIN_PASSWORD")
GROQ_API_KEY:  str = _env("GROQ_API_KEY")
HF_TOKEN:      str = _env("HF_TOKEN")
DB_PATH:       str = str(Path(os.getenv("DB_PATH", "./jazz.db")).resolve())

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

CHUNK_SIZE:      int = 500
CHUNK_OVERLAP:   int = 50
TOP_K_RETRIEVAL: int = 4

WORKSPACE_DIR = Path("workspace")
WORKSPACE_DIR.mkdir(exist_ok=True)
SITES_DIR = WORKSPACE_DIR / "sites"
SITES_DIR.mkdir(exist_ok=True)

ALLOWED_EXTENSIONS: set = {
    ".html", ".css", ".js", ".py", ".txt", ".md", ".json", ".xml",
    ".yaml", ".yml", ".csv", ".sql", ".sh", ".conf", ".ini", ".log",
    ".svg", ".ts", ".jsx", ".tsx", ".bat",
}

MAX_FILE_SIZE:        int  = 10 * 1024 * 1024
MAX_IMAGE_SIZE:       int  = 20 * 1024 * 1024
ALLOWED_IMAGE_TYPES: set  = {"image/jpeg", "image/png", "image/gif", "image/webp"}
CODE_EXEC_TIMEOUT:   int  = 10
CODE_EXEC_MAX_OUT:   int  = 50_000
MAX_MEMORIES:        int  = 100
_SERVER_START: datetime   = datetime.utcnow()
_executor: ThreadPoolExecutor = ThreadPoolExecutor(max_workers=8)

SUBSCRIPTION_LIMITS: Dict[str, Dict[str, Any]] = {
    "free": {
        "messages_per_day": 50,
        "documents": 5,
        "max_file_mb": 10,
        "mysql_access": False,
        "image_analysis": True,
        "agent_jobs": 1,
        "websites": 3,
        "code_runs_per_day": 20,
        "api_keys": 2,
        "chat_sessions": 10,
    },
    "pro": {
        "messages_per_day": 500,
        "documents": 50,
        "max_file_mb": 50,
        "mysql_access": True,
        "image_analysis": True,
        "agent_jobs": 10,
        "websites": 50,
        "code_runs_per_day": 200,
        "api_keys": 10,
        "chat_sessions": -1,
    },
    "enterprise": {
        "messages_per_day": -1,
        "documents": -1,
        "max_file_mb": 100,
        "mysql_access": True,
        "image_analysis": True,
        "agent_jobs": -1,
        "websites": -1,
        "code_runs_per_day": -1,
        "api_keys": -1,
        "chat_sessions": -1,
    },
}

_PLAYWRIGHT_AVAILABLE: bool = False
try:
    from playwright.sync_api import sync_playwright as _pw_check  # noqa: F401
    _PLAYWRIGHT_AVAILABLE = True
    logger.info("[PLAYWRIGHT] Available ✅")
except ImportError:
    logger.warning("[PLAYWRIGHT] Not installed ❌")


# ──────────────────────────────────────────────────────────────────────────────
# Security
# ──────────────────────────────────────────────────────────────────────────────

pwd_context = CryptContext(schemes=["bcrypt"], deprecated="auto")
security    = HTTPBearer(auto_error=False)


# ──────────────────────────────────────────────────────────────────────────────
# SQLite async helpers
# ──────────────────────────────────────────────────────────────────────────────

def db() -> aiosqlite.Connection:
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


# ──────────────────────────────────────────────────────────────────────────────
# Database schema
# ──────────────────────────────────────────────────────────────────────────────

SCHEMA_SQL = """
PRAGMA journal_mode=WAL;
PRAGMA foreign_keys=ON;

CREATE TABLE IF NOT EXISTS users (
    id               TEXT PRIMARY KEY,
    email            TEXT UNIQUE NOT NULL,
    password_hash    TEXT NOT NULL,
    full_name        TEXT DEFAULT '',
    role             TEXT DEFAULT 'client',
    subscription     TEXT DEFAULT 'free',
    memory_enabled   INTEGER DEFAULT 1,
    avatar_url       TEXT DEFAULT '',
    preferred_model  TEXT DEFAULT 'uncensored',
    created_at       TEXT DEFAULT (datetime('now')),
    updated_at       TEXT DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS chat_sessions (
    id            TEXT PRIMARY KEY,
    user_id       TEXT NOT NULL,
    title         TEXT NOT NULL DEFAULT 'New Chat',
    model_type    TEXT DEFAULT 'uncensored',
    message_count INTEGER DEFAULT 0,
    created_at    TEXT DEFAULT (datetime('now')),
    updated_at    TEXT DEFAULT (datetime('now')),
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
    FOREIGN KEY (user_id)    REFERENCES users(id)         ON DELETE CASCADE,
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
    created_at  TEXT DEFAULT (datetime('now')),
    updated_at  TEXT DEFAULT (datetime('now')),
    FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE
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

-- v7.0: API Keys
CREATE TABLE IF NOT EXISTS api_keys (
    id           TEXT PRIMARY KEY,
    user_id      TEXT NOT NULL,
    name         TEXT NOT NULL,
    key_hash     TEXT NOT NULL UNIQUE,
    key_prefix   TEXT NOT NULL,
    last_used_at TEXT DEFAULT NULL,
    expires_at   TEXT DEFAULT NULL,
    is_active    INTEGER DEFAULT 1,
    created_at   TEXT DEFAULT (datetime('now')),
    FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE
);

-- v7.0: Notifications
CREATE TABLE IF NOT EXISTS notifications (
    id         TEXT PRIMARY KEY,
    user_id    TEXT NOT NULL,
    title      TEXT NOT NULL,
    message    TEXT NOT NULL,
    type       TEXT DEFAULT 'info',
    link       TEXT DEFAULT '',
    is_read    INTEGER DEFAULT 0,
    created_at TEXT DEFAULT (datetime('now')),
    FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE
);

-- v7.0: System Announcements
CREATE TABLE IF NOT EXISTS announcements (
    id         TEXT PRIMARY KEY,
    title      TEXT NOT NULL,
    message    TEXT NOT NULL,
    type       TEXT DEFAULT 'info',
    is_active  INTEGER DEFAULT 1,
    created_by TEXT NOT NULL,
    expires_at TEXT DEFAULT NULL,
    created_at TEXT DEFAULT (datetime('now'))
);

-- v7.0: Coupons
CREATE TABLE IF NOT EXISTS coupons (
    id                TEXT PRIMARY KEY,
    code              TEXT UNIQUE NOT NULL COLLATE NOCASE,
    description       TEXT DEFAULT '',
    subscription_tier TEXT NOT NULL,
    max_uses          INTEGER DEFAULT 1,
    uses              INTEGER DEFAULT 0,
    is_active         INTEGER DEFAULT 1,
    expires_at        TEXT DEFAULT NULL,
    created_at        TEXT DEFAULT (datetime('now'))
);

-- v7.0: Coupon Redemptions
CREATE TABLE IF NOT EXISTS coupon_redemptions (
    id          TEXT PRIMARY KEY,
    coupon_id   TEXT NOT NULL,
    user_id     TEXT NOT NULL,
    redeemed_at TEXT DEFAULT (datetime('now')),
    UNIQUE(coupon_id, user_id),
    FOREIGN KEY (coupon_id) REFERENCES coupons(id),
    FOREIGN KEY (user_id)   REFERENCES users(id) ON DELETE CASCADE
);

-- v7.1: Message Feedback (thumbs up/down, ratings)
CREATE TABLE IF NOT EXISTS message_feedback (
    id          TEXT PRIMARY KEY,
    chat_id     TEXT NOT NULL,
    user_id     TEXT NOT NULL,
    rating      INTEGER DEFAULT 0,
    feedback    TEXT DEFAULT '',
    created_at  TEXT DEFAULT (datetime('now')),
    UNIQUE(chat_id),
    FOREIGN KEY (chat_id)   REFERENCES chat_history(id) ON DELETE CASCADE,
    FOREIGN KEY (user_id)   REFERENCES users(id) ON DELETE CASCADE
);

-- v7.1: Conversation Topics/Tags
CREATE TABLE IF NOT EXISTS conversation_topics (
    id             TEXT PRIMARY KEY,
    session_id     TEXT NOT NULL,
    user_id        TEXT NOT NULL,
    topic          TEXT NOT NULL,
    confidence     REAL DEFAULT 0.8,
    auto_detected  INTEGER DEFAULT 1,
    created_at     TEXT DEFAULT (datetime('now')),
    FOREIGN KEY (session_id) REFERENCES chat_sessions(id) ON DELETE CASCADE,
    FOREIGN KEY (user_id)    REFERENCES users(id) ON DELETE CASCADE
);

-- v7.1: Conversation Summaries
CREATE TABLE IF NOT EXISTS conversation_summaries (
    id             TEXT PRIMARY KEY,
    session_id     TEXT NOT NULL,
    user_id        TEXT NOT NULL,
    summary        TEXT NOT NULL,
    message_range  TEXT DEFAULT '{}',
    created_at     TEXT DEFAULT (datetime('now')),
    updated_at     TEXT DEFAULT (datetime('now')),
    UNIQUE(session_id),
    FOREIGN KEY (session_id) REFERENCES chat_sessions(id) ON DELETE CASCADE,
    FOREIGN KEY (user_id)    REFERENCES users(id) ON DELETE CASCADE
);

-- v7.1: Draft Messages
CREATE TABLE IF NOT EXISTS draft_messages (
    id             TEXT PRIMARY KEY,
    session_id     TEXT NOT NULL,
    user_id        TEXT NOT NULL,
    content        TEXT NOT NULL,
    metadata       TEXT DEFAULT '{}',
    created_at     TEXT DEFAULT (datetime('now')),
    updated_at     TEXT DEFAULT (datetime('now')),
    UNIQUE(session_id),
    FOREIGN KEY (session_id) REFERENCES chat_sessions(id) ON DELETE CASCADE,
    FOREIGN KEY (user_id)    REFERENCES users(id) ON DELETE CASCADE
);

-- v7.1: Conversation Branches (fork conversations)
CREATE TABLE IF NOT EXISTS conversation_branches (
    id                TEXT PRIMARY KEY,
    parent_session_id TEXT NOT NULL,
    child_session_id  TEXT NOT NULL,
    branch_from_msg   TEXT NOT NULL,
    user_id           TEXT NOT NULL,
    created_at        TEXT DEFAULT (datetime('now')),
    FOREIGN KEY (parent_session_id) REFERENCES chat_sessions(id) ON DELETE CASCADE,
    FOREIGN KEY (child_session_id)  REFERENCES chat_sessions(id) ON DELETE CASCADE,
    FOREIGN KEY (user_id)           REFERENCES users(id) ON DELETE CASCADE
);

-- v7.1: System Prompts/Presets
CREATE TABLE IF NOT EXISTS system_prompts (
    id             TEXT PRIMARY KEY,
    user_id        TEXT NOT NULL,
    name           TEXT NOT NULL,
    description    TEXT DEFAULT '',
    system_message TEXT NOT NULL,
    preset_type    TEXT DEFAULT 'custom',
    is_public      INTEGER DEFAULT 0,
    use_count      INTEGER DEFAULT 0,
    created_at     TEXT DEFAULT (datetime('now')),
    FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE
);

-- v7.1: Message Edit History
CREATE TABLE IF NOT EXISTS message_edits (
    id             TEXT PRIMARY KEY,
    message_id     TEXT NOT NULL,
    old_content    TEXT NOT NULL,
    new_content    TEXT NOT NULL,
    edited_at      TEXT DEFAULT (datetime('now')),
    FOREIGN KEY (message_id) REFERENCES chat_history(id) ON DELETE CASCADE
);

-- v7.1: Conversation Analytics
CREATE TABLE IF NOT EXISTS conversation_analytics (
    id                  TEXT PRIMARY KEY,
    session_id          TEXT NOT NULL,
    user_id             TEXT NOT NULL,
    message_count       INTEGER DEFAULT 0,
    average_response_time INTEGER DEFAULT 0,
    models_used         TEXT DEFAULT '[]',
    agents_used         TEXT DEFAULT '[]',
    total_tokens_in     INTEGER DEFAULT 0,
    total_tokens_out    INTEGER DEFAULT 0,
    created_at          TEXT DEFAULT (datetime('now')),
    UNIQUE(session_id),
    FOREIGN KEY (session_id) REFERENCES chat_sessions(id) ON DELETE CASCADE,
    FOREIGN KEY (user_id)    REFERENCES users(id) ON DELETE CASCADE
);

-- v7.1: Regeneration History
CREATE TABLE IF NOT EXISTS message_regenerations (
    id                TEXT PRIMARY KEY,
    original_msg_id   TEXT NOT NULL,
    regenerated_msg_id TEXT DEFAULT NULL,
    regeneration_count INTEGER DEFAULT 1,
    created_at        TEXT DEFAULT (datetime('now')),
    FOREIGN KEY (original_msg_id)   REFERENCES chat_history(id),
    FOREIGN KEY (regenerated_msg_id) REFERENCES chat_history(id)
);

-- v7.1: User Sessions Config
CREATE TABLE IF NOT EXISTS user_session_config (
    id                TEXT PRIMARY KEY,
    user_id           TEXT NOT NULL UNIQUE,
    system_prompt     TEXT DEFAULT '',
    preferred_model   TEXT DEFAULT 'uncensored',
    max_context_msgs  INTEGER DEFAULT 20,
    auto_summarize    INTEGER DEFAULT 1,
    auto_title        INTEGER DEFAULT 1,
    created_at        TEXT DEFAULT (datetime('now')),
    updated_at        TEXT DEFAULT (datetime('now')),
    FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE
);

CREATE INDEX IF NOT EXISTS idx_chat_user     ON chat_history(user_id, created_at);
CREATE INDEX IF NOT EXISTS idx_chat_session  ON chat_history(session_id);
CREATE INDEX IF NOT EXISTS idx_sessions_user ON chat_sessions(user_id, updated_at);
CREATE INDEX IF NOT EXISTS idx_docs_user     ON documents(user_id);
CREATE INDEX IF NOT EXISTS idx_mem_user      ON user_memories(user_id);
CREATE INDEX IF NOT EXISTS idx_ajob_user     ON agent_jobs(user_id);
CREATE INDEX IF NOT EXISTS idx_ajlog_job     ON agent_job_logs(job_id, started_at);
CREATE INDEX IF NOT EXISTS idx_sites_user    ON websites(user_id);
CREATE INDEX IF NOT EXISTS idx_prompts_user  ON prompt_templates(user_id);
CREATE INDEX IF NOT EXISTS idx_code_user     ON code_runs(user_id);
CREATE INDEX IF NOT EXISTS idx_apikeys_user  ON api_keys(user_id);
CREATE INDEX IF NOT EXISTS idx_apikeys_hash  ON api_keys(key_hash);
CREATE INDEX IF NOT EXISTS idx_notif_user    ON notifications(user_id, is_read, created_at);
"""


async def init_db() -> None:
    async with db() as conn:
        await conn.executescript(SCHEMA_SQL)
        await conn.commit()

    # Reset stuck running jobs
    await db_execute(
        "UPDATE agent_jobs SET status='idle', updated_at=? WHERE status='running'",
        (datetime.utcnow().isoformat(),),
    )

    # Bootstrap admin user
    existing = await db_fetchone("SELECT id FROM users WHERE email=?", (ADMIN_EMAIL,))
    if not existing and ADMIN_PASSWORD:
        uid   = str(uuid.uuid4())
        phash = pwd_context.hash(ADMIN_PASSWORD)
        await db_execute(
            "INSERT INTO users (id,email,password_hash,full_name,role,subscription) VALUES (?,?,?,?,?,?)",
            (uid, ADMIN_EMAIL, phash, "Admin", "admin", "enterprise"),
        )
        logger.info("[DB] Admin user created: %s", ADMIN_EMAIL)

    logger.info("[DB] Initialized ✅  path=%s", DB_PATH)


# ──────────────────────────────────────────────────────────────────────────────
# ChromaDB
# ──────────────────────────────────────────────────────────────────────────────

chroma_client: Optional[chromadb.PersistentClient] = None
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


# ──────────────────────────────────────────────────────────────────────────────
# MySQL pool (optional)
# ──────────────────────────────────────────────────────────────────────────────

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


# ──────────────────────────────────────────────────────────────────────────────
# Rate limiter
# ──────────────────────────────────────────────────────────────────────────────

class RateLimiter:
    _WINDOW: int = 86_400  # 24h

    def __init__(self) -> None:
        self._window: Dict[str, List[float]] = defaultdict(list)

    def _clean(self, key: str) -> None:
        now = time.time()
        self._window[key] = [t for t in self._window[key] if now - t < self._WINDOW]

    def check(self, uid: str, tier: str, resource: str = "messages_per_day") -> bool:
        limit = SUBSCRIPTION_LIMITS.get(tier, SUBSCRIPTION_LIMITS["free"]).get(resource, 50)
        if limit == -1:
            return True
        key = f"{uid}:{resource}"
        self._clean(key)
        if len(self._window[key]) >= limit:
            return False
        self._window[key].append(time.time())
        return True

    def usage_today(self, uid: str, resource: str = "messages_per_day") -> int:
        key = f"{uid}:{resource}"
        self._clean(key)
        return len(self._window.get(key, []))


rate_limiter = RateLimiter()


# ──────────────────────────────────────────────────────────────────────────────
# Pydantic models
# ──────────────────────────────────────────────────────────────────────────────

class ChatRequest(BaseModel):
    message:    str           = Field(..., min_length=1, max_length=8000)
    model_type: str           = Field("uncensored", pattern="^(censored|uncensored)$")
    use_rag:    bool          = True
    image_url:  Optional[str] = None
    session_id: Optional[str] = None


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


class ChangePasswordRequest(BaseModel):
    current_password: str = Field(..., min_length=1)
    new_password:     str = Field(..., min_length=8)


class ProfileUpdate(BaseModel):
    full_name:       Optional[str]  = None
    avatar_url:      Optional[str]  = None
    preferred_model: Optional[str]  = Field(None, pattern="^(censored|uncensored)$")
    memory_enabled:  Optional[bool] = None


class MemoryCreate(BaseModel):
    content:  str           = Field(..., min_length=1, max_length=500)
    category: Optional[str] = "general"


class MemoryUpdate(BaseModel):
    content: str = Field(..., min_length=1, max_length=500)


class MemorySettingsUpdate(BaseModel):
    enabled: bool


class DocumentResponse(BaseModel):
    id:            str
    filename:      str
    original_name: str
    file_type:     str
    file_size:     int
    chunk_count:   int
    uploaded_at:   str


class MySQLQueryRequest(BaseModel):
    question: str = Field(..., min_length=1, max_length=1000)


class MySQLRawQueryRequest(BaseModel):
    sql: str = Field(..., min_length=1)


class SubscriptionUpdate(BaseModel):
    user_id: str
    tier:    str = Field(..., pattern="^(free|pro|enterprise)$")


class FileOperationRequest(BaseModel):
    filename:  str
    content:   Optional[str] = None
    operation: str           = Field(..., pattern="^(create|read|update|delete|list)$")


class AgentJobCreate(BaseModel):
    name:              str                      = Field(..., min_length=1, max_length=100)
    description:       Optional[str]            = ""
    job_type:          str                      = Field(..., pattern="^(naukri_apply|linkedin_apply|custom)$")
    cron_schedule:     str
    time_window_start: Optional[str]            = "11:00"
    time_window_end:   Optional[str]            = "12:00"
    credentials:       Dict[str, Any]
    parameters:        Optional[Dict[str, Any]] = {}


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
    instructions: str = Field(..., min_length=5, max_length=1000)
    model_type:   str = Field("censored", pattern="^(censored|uncensored)$")


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


# v7.0 models
class APIKeyCreate(BaseModel):
    name:       str           = Field(..., min_length=1, max_length=80)
    expires_at: Optional[str] = None


class ChatSessionCreate(BaseModel):
    title: str = Field("New Chat", min_length=1, max_length=100)


class ChatSessionUpdate(BaseModel):
    title: str = Field(..., min_length=1, max_length=100)


class NotificationCreate(BaseModel):
    user_id: str
    title:   str = Field(..., min_length=1, max_length=100)
    message: str = Field(..., min_length=1, max_length=500)
    type:    str = Field("info", pattern="^(info|success|warning|error)$")
    link:    str = ""


class AnnouncementCreate(BaseModel):
    title:      str           = Field(..., min_length=1, max_length=100)
    message:    str           = Field(..., min_length=1, max_length=1000)
    type:       str           = Field("info", pattern="^(info|success|warning|error)$")
    expires_at: Optional[str] = None


class CouponCreate(BaseModel):
    code:              str           = Field(..., min_length=3, max_length=30)
    description:       str           = ""
    subscription_tier: str           = Field(..., pattern="^(free|pro|enterprise)$")
    max_uses:          int           = Field(1, ge=1)
    expires_at:        Optional[str] = None


class CouponRedeem(BaseModel):
    code: str = Field(..., min_length=1)


class BroadcastCreate(BaseModel):
    title:   str = Field(..., min_length=1, max_length=100)
    message: str = Field(..., min_length=1, max_length=500)
    type:    str = Field("info", pattern="^(info|success|warning|error)$")


# ──────────────────────────────────────────────────────────────────────────────
# Auth helpers
# ──────────────────────────────────────────────────────────────────────────────

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


async def _auth_via_api_key(key: str) -> Optional[Dict]:
    """Authenticate using a raw API key. Returns payload dict or None."""
    key_hash = hashlib.sha256(key.encode()).hexdigest()
    row = await db_fetchone(
        """SELECT ak.id, ak.user_id, ak.expires_at,
                  u.role, u.subscription, u.email
           FROM api_keys ak
           JOIN users u ON ak.user_id = u.id
           WHERE ak.key_hash=? AND ak.is_active=1""",
        (key_hash,),
    )
    if not row:
        return None
    if row.get("expires_at") and datetime.fromisoformat(row["expires_at"]) < datetime.utcnow():
        return None
    # Update last_used_at (fire-and-forget)
    asyncio.create_task(
        db_execute("UPDATE api_keys SET last_used_at=? WHERE id=?",
                   (datetime.utcnow().isoformat(), row["id"]))
    )
    return {
        "sub":         row["user_id"],
        "email":       row["email"],
        "role":        row["role"],
        "subscription": row["subscription"],
        "auth_method": "api_key",
    }


async def _resolve_auth(
    request: Request,
    credentials: Optional[HTTPAuthorizationCredentials],
) -> Optional[Dict]:
    """Try JWT Bearer, then API-key Bearer, then X-API-Key header."""
    if credentials and credentials.credentials:
        token = credentials.credentials
        if token.startswith("jz_"):
            return await _auth_via_api_key(token)
        payload = verify_token(token)
        if payload:
            return payload
    ak = request.headers.get("X-API-Key", "")
    if ak.startswith("jz_"):
        return await _auth_via_api_key(ak)
    return None


async def get_current_user(
    request: Request,
    credentials: Optional[HTTPAuthorizationCredentials] = Depends(security),
) -> Dict:
    payload = await _resolve_auth(request, credentials)
    if not payload:
        raise HTTPException(401, "Authorization required", headers={"WWW-Authenticate": "Bearer"})
    return payload


async def require_admin(
    request: Request,
    credentials: Optional[HTTPAuthorizationCredentials] = Depends(security),
) -> Dict:
    payload = await _resolve_auth(request, credentials)
    if not payload:
        raise HTTPException(401, "Authorization required")
    if payload.get("role") != "admin":
        raise HTTPException(403, "Admin access required")
    return payload


async def get_user_subscription(user_id: str) -> str:
    row = await db_fetchone("SELECT subscription FROM users WHERE id=?", (user_id,))
    return (row or {}).get("subscription", "free")


async def is_memory_enabled(user_id: str) -> bool:
    row = await db_fetchone("SELECT memory_enabled FROM users WHERE id=?", (user_id,))
    return bool((row or {}).get("memory_enabled", 1))


# ──────────────────────────────────────────────────────────────────────────────
# Notification helper
# ──────────────────────────────────────────────────────────────────────────────

async def push_notification(
    user_id: str, title: str, message: str,
    ntype: str = "info", link: str = ""
) -> None:
    try:
        await db_execute(
            "INSERT INTO notifications (id,user_id,title,message,type,link) VALUES (?,?,?,?,?,?)",
            (str(uuid.uuid4()), user_id, title, message, ntype, link),
        )
    except Exception as exc:
        logger.warning("[NOTIF] Could not send notification: %s", exc)


# ──────────────────────────────────────────────────────────────────────────────
# Memory system
# ──────────────────────────────────────────────────────────────────────────────

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

    def _count() -> int:
        with _s3.connect(DB_PATH) as c:
            r = c.execute("SELECT COUNT(*) FROM user_memories WHERE user_id=?", (user_id,)).fetchone()
            return r[0] if r else 0

    def _existing() -> List[str]:
        with _s3.connect(DB_PATH) as c:
            return [r[0].lower() for r in c.execute(
                "SELECT content FROM user_memories WHERE user_id=?", (user_id,)
            ).fetchall()]

    def _insert(mid: str, content: str, category: str) -> None:
        with _s3.connect(DB_PATH) as c:
            c.execute(
                "INSERT INTO user_memories (id,user_id,content,category,source) VALUES (?,?,?,?,?)",
                (mid, user_id, content, category, "auto"),
            )
            c.commit()

    def _mem_enabled() -> bool:
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
        comp = client.chat.completions.create(
            model=model,
            messages=[{"role": "user", "content": prompt}],
            max_tokens=300,
            temperature=0.1,
        )
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
        logger.error("[MEMORY] Extraction error: %s", exc)
        return []


async def extract_and_save_memories(
    user_id: str, user_message: str, ai_response: str
) -> List[str]:
    loop = asyncio.get_event_loop()
    return await loop.run_in_executor(
        _executor, _sync_extract_memories, user_id, user_message, ai_response
    )


# ──────────────────────────────────────────────────────────────────────────────
# v7.1: Enhanced Conversation Context & History
# ──────────────────────────────────────────────────────────────────────────────

async def get_conversation_context(
    session_id: str, user_id: str, max_messages: int = 20
) -> List[Dict]:
    """Retrieve previous messages in a conversation for context."""
    if not session_id:
        return []
    rows = await db_fetchall(
        """SELECT id, message, response, agent_used, model_used, created_at 
           FROM chat_history 
           WHERE session_id=? AND user_id=? 
           ORDER BY created_at DESC LIMIT ?""",
        (session_id, user_id, max_messages),
    )
    return list(reversed(rows))


async def build_conversation_prompt(
    user_message: str, agent, context: str = "", memory_block: str = "",
    session_id: Optional[str] = None, user_id: Optional[str] = None,
    include_history: bool = True
) -> str:
    """Build enhanced prompt including conversation history."""
    history_text = ""
    if include_history and session_id and user_id:
        history = await get_conversation_context(session_id, user_id, max_messages=5)
        if history:
            history_lines = ["[Recent Conversation History:]"]
            for msg_pair in history:
                if msg_pair["message"]:
                    history_lines.append(f"User: {msg_pair['message'][:200]}")
                if msg_pair["response"]:
                    history_lines.append(f"Assistant: {msg_pair['response'][:200]}")
            history_text = "\n".join(history_lines[:15]) + "\n\n"
    
    full_context = f"{history_text}{context}" if history_text or context else ""
    return agent.build_prompt(user_message, full_context, memory_block)


async def auto_detect_topics(session_id: str, user_id: str, messages: List[str]) -> List[str]:
    """Auto-detect conversation topics using AI."""
    if not messages or len(messages) < 2:
        return []
    try:
        client, model = get_model_client("censored")
        combined = "\n".join(messages[:5])
        prompt = (
            f"Extract 1-3 keywords/topics that best describe this conversation:\n{combined}\n"
            'Return ONLY a JSON array of strings: ["topic1", "topic2"]'
        )
        comp = client.chat.completions.create(
            model=model, messages=[{"role": "user", "content": prompt}],
            max_tokens=100, temperature=0.2
        )
        raw = re.sub(r"```json|```", "", comp.choices[0].message.content.strip()).strip()
        topics = json.loads(raw)
        if isinstance(topics, list):
            for topic in topics[:3]:
                await db_execute(
                    "INSERT INTO conversation_topics (id,session_id,user_id,topic) VALUES (?,?,?,?)",
                    (str(uuid.uuid4()), session_id, user_id, topic),
                )
            return topics
    except Exception as exc:
        logger.warning("[TOPICS] Detection failed: %s", exc)
    return []


async def summarize_conversation(session_id: str, user_id: str) -> Optional[str]:
    """Generate a summary of the conversation."""
    messages = await db_fetchall(
        "SELECT message, response FROM chat_history WHERE session_id=? ORDER BY created_at",
        (session_id,),
    )
    if len(messages) < 3:
        return None
    try:
        client, model = get_model_client("censored")
        conv_text = "\n".join(
            [f"User: {m['message']}\nAssistant: {m['response']}"
             for m in messages[:10]]
        )
        prompt = f"Summarize this conversation in 2-3 sentences:\n{conv_text}"
        comp = client.chat.completions.create(
            model=model, messages=[{"role": "user", "content": prompt}],
            max_tokens=150, temperature=0.3
        )
        summary = comp.choices[0].message.content or ""
        await db_execute(
            "INSERT OR REPLACE INTO conversation_summaries (id,session_id,user_id,summary) VALUES (?,?,?,?)",
            (str(uuid.uuid4()), session_id, user_id, summary),
        )
        return summary
    except Exception as exc:
        logger.warning("[SUMMARY] Generation failed: %s", exc)
    return None


async def save_message_feedback(chat_id: str, user_id: str, rating: int, feedback: str = "") -> bool:
    """Save thumbs up/down feedback for a message."""
    try:
        await db_execute(
            "INSERT OR REPLACE INTO message_feedback (id,chat_id,user_id,rating,feedback) VALUES (?,?,?,?,?)",
            (str(uuid.uuid4()), chat_id, user_id, rating, feedback),
        )
        return True
    except Exception as exc:
        logger.error("[FEEDBACK] Save failed: %s", exc)
    return False


async def regenerate_message(session_id: str, user_id: str, orig_msg_id: str) -> Optional[str]:
    """Regenerate AI response for a message."""
    msg_row = await db_fetchone(
        "SELECT message FROM chat_history WHERE id=? AND user_id=?",
        (orig_msg_id, user_id),
    )
    if not msg_row:
        return None
    try:
        mem_on = await is_memory_enabled(user_id)
        memories = await get_user_memories(user_id) if mem_on else []
        memory_block = format_memories_for_prompt(memories)
        context = get_relevant_context(user_id, msg_row["message"])
        agent = route_message(msg_row["message"], False)
        prompt = await build_conversation_prompt(
            msg_row["message"], agent, context, memory_block, session_id, user_id
        )
        ai, model = get_model_client("uncensored")
        comp = ai.chat.completions.create(
            model=model, messages=[{"role": "user", "content": prompt}],
            max_tokens=1024, temperature=0.8
        )
        new_response = comp.choices[0].message.content or ""
        new_msg_id = str(uuid.uuid4())
        await db_insert(
            "INSERT INTO chat_history (id,user_id,session_id,message,response,query_type,agent_used,model_used) VALUES (?,?,?,?,?,?,?,?)",
            (new_msg_id, user_id, session_id, msg_row["message"], new_response,
             "regeneration", agent.name, "uncensored"),
        )
        await db_execute(
            "INSERT INTO message_regenerations (id,original_msg_id,regenerated_msg_id,regeneration_count) VALUES (?,?,?,?)",
            (str(uuid.uuid4()), orig_msg_id, new_msg_id, 1),
        )
        return new_response
    except Exception as exc:
        logger.error("[REGEN] Failed: %s", exc)
    return None


async def branch_conversation(parent_session_id: str, user_id: str, branch_from_msg: str) -> Optional[str]:
    """Create a new conversation branch from a message."""
    parent = await db_fetchone(
        "SELECT title FROM chat_sessions WHERE id=?", (parent_session_id,)
    )
    if not parent:
        return None
    try:
        new_session_id = str(uuid.uuid4())
        now = datetime.utcnow().isoformat()
        new_title = f"{parent['title']} (Branch)"
        await db_execute(
            "INSERT INTO chat_sessions (id,user_id,title,created_at,updated_at) VALUES (?,?,?,?,?)",
            (new_session_id, user_id, new_title, now, now),
        )
        await db_execute(
            "INSERT INTO conversation_branches (id,parent_session_id,child_session_id,branch_from_msg,user_id) VALUES (?,?,?,?,?)",
            (str(uuid.uuid4()), parent_session_id, new_session_id, branch_from_msg, user_id),
        )
        return new_session_id
    except Exception as exc:
        logger.error("[BRANCH] Creation failed: %s", exc)
    return None


async def export_conversation(session_id: str, user_id: str, format_type: str = "json") -> Optional[str]:
    """Export conversation in various formats."""
    messages = await db_fetchall(
        "SELECT message, response, created_at, agent_used FROM chat_history WHERE session_id=? ORDER BY created_at",
        (session_id,),
    )
    if not messages:
        return None
    try:
        if format_type == "json":
            return json.dumps({"messages": messages}, indent=2, default=str)
        elif format_type == "markdown":
            lines = ["# Conversation Export\n"]
            for msg in messages:
                lines.append(f"**User:** {msg['message']}\n")
                lines.append(f"**Assistant:** {msg['response']}\n")
                lines.append(f"*({msg['created_at']})*\n\n")
            return "\n".join(lines)
        elif format_type == "csv":
            import csv
            import io
            output = io.StringIO()
            writer = csv.writer(output)
            writer.writerow(["User Message", "AI Response", "Timestamp", "Agent"])
            for msg in messages:
                writer.writerow([msg["message"], msg["response"], msg["created_at"], msg["agent_used"]])
            return output.getvalue()
    except Exception as exc:
        logger.error("[EXPORT] Failed: %s", exc)
    return None


# ──────────────────────────────────────────────────────────────────────────────
# Document processing
# ──────────────────────────────────────────────────────────────────────────────

def chunk_text(text: str, size: int = CHUNK_SIZE, overlap: int = CHUNK_OVERLAP) -> List[str]:
    chunks, start = [], 0
    while start < len(text):
        chunk = text[start:start + size]
        if chunk.strip():
            chunks.append(chunk)
        start += size - overlap
    return chunks


def extract_text_from_pdf(path: str) -> str:
    try:
        with fitz.open(path) as pdf:
            return "\n".join(page.get_text() for page in pdf)
    except Exception as exc:
        logger.error("[PDF] %s", exc)
        return ""


def extract_text_from_docx(path: str) -> str:
    try:
        doc = DocxDocument(path)
        return "\n".join(p.text for p in doc.paragraphs)
    except Exception as exc:
        logger.error("[DOCX] %s", exc)
        return ""


def extract_text_from_excel(path: str) -> str:
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
        parts: List[str] = []
        for sheet in wb.worksheets:
            parts.append(f"\nSheet: {sheet.title}")
            for row in sheet.iter_rows(values_only=True):
                row_str = " | ".join(str(c) for c in row if c is not None)
                if row_str.strip():
                    parts.append(row_str)
        return "\n".join(parts)
    except Exception as exc:
        logger.error("[EXCEL] %s", exc)
        return ""


def process_document(path: str, mime: str) -> str:
    if "pdf" in mime:
        return extract_text_from_pdf(path)
    if "word" in mime or "docx" in mime:
        return extract_text_from_docx(path)
    if "excel" in mime or "sheet" in mime:
        return extract_text_from_excel(path)
    if "csv" in mime:
        try:
            with open(path, "r", encoding="utf-8") as f:
                return "\n".join(" | ".join(row) for row in csv.reader(f))
        except Exception:
            return ""
    try:
        return Path(path).read_text(encoding="utf-8")
    except Exception:
        return ""


# ──────────────────────────────────────────────────────────────────────────────
# RAG
# ──────────────────────────────────────────────────────────────────────────────

def get_relevant_context(user_id: str, query: str, top_k: int = TOP_K_RETRIEVAL) -> str:
    if documents_collection is None:
        return ""
    try:
        results = documents_collection.query(
            query_texts=[query], n_results=top_k, where={"user_id": user_id}
        )
        if not results or not results["documents"] or not results["documents"][0]:
            return ""
        parts: List[str] = []
        for i, doc in enumerate(results["documents"][0]):
            meta = results["metadatas"][0][i] if results.get("metadatas") else {}
            parts.append(f"[Source: {meta.get('filename','Unknown')}]\n{doc}")
        return "\n\n---\n\n".join(parts)
    except Exception as exc:
        logger.error("[RAG] %s", exc)
        return ""


# ──────────────────────────────────────────────────────────────────────────────
# Schema indexing
# ──────────────────────────────────────────────────────────────────────────────

_last_reindex_at: Optional[str] = None


def index_sqlite_schema() -> bool:
    global _last_reindex_at
    if sqlite_schema_collection is None:
        return False
    import sqlite3 as _s3
    try:
        with _s3.connect(DB_PATH) as conn:
            tables = [r[0] for r in conn.execute(
                "SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%'"
            ).fetchall()]
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
                    chunks.append(text)
                    ids.append(f"tbl_{tbl}")
                    metas.append({"table_name": tbl, "type": "schema"})
                except Exception as exc:
                    logger.warning("[SQLITE-SCHEMA] Skipping %s: %s", tbl, exc)
        if chunks:
            sqlite_schema_collection.add(documents=chunks, ids=ids, metadatas=metas)
        _last_reindex_at = datetime.utcnow().isoformat()
        logger.info("[SQLITE] Schema indexed — %d tables ✅", len(chunks))
        return True
    except Exception as exc:
        logger.error("[SQLITE-SCHEMA] %s", exc)
        return False


def get_schema_context(query: str, top_k: int = 3) -> str:
    if sqlite_schema_collection is None:
        return ""
    try:
        results = sqlite_schema_collection.query(query_texts=[query], n_results=top_k)
        if results and results["documents"] and results["documents"][0]:
            return "\n\n".join(results["documents"][0])
        return ""
    except Exception as exc:
        logger.error("[SQLITE-SCHEMA] %s", exc)
        return ""


def index_mysql_schema() -> bool:
    if not MYSQL_ENABLED or sqlite_schema_collection is None:
        return False
    conn = get_mysql_connection()
    if not conn:
        return False
    try:
        cur = conn.cursor(dictionary=True)
        cur.execute("SHOW TABLES")
        tables = [list(r.values())[0] for r in cur.fetchall()]
        chunks, ids, metas = [], [], []
        for tbl in tables:
            try:
                cur.execute(f"DESCRIBE `{tbl}`")
                cols = cur.fetchall()
                cur.execute(f"SELECT * FROM `{tbl}` LIMIT 3")
                samples = cur.fetchall()
                col_info = ", ".join(
                    f"{c['Field']}({c['Type']})" + (",PK" if c["Key"] == "PRI" else "") for c in cols
                )
                text = f"[MySQL] Table: {tbl} | Columns: {col_info}"
                if samples:
                    text += f" | Samples: {json.dumps(samples, default=str)}"
                chunks.append(text)
                ids.append(f"mysql_tbl_{tbl}")
                metas.append({"table_name": tbl, "type": "mysql_schema"})
            except Exception as exc:
                logger.warning("[MYSQL] Skipping %s: %s", tbl, exc)
        if chunks:
            sqlite_schema_collection.add(documents=chunks, ids=ids, metadatas=metas)
        cur.close()
        conn.close()
        logger.info("[MYSQL] Schema indexed — %d tables ✅", len(chunks))
        return True
    except Exception as exc:
        logger.error("[MYSQL] %s", exc)
        return False


_DB_KEYWORDS: set = {
    "table", "column", "row", "select", "count", "show", "describe",
    "database", "schema", "query", "fetch", "list", "find", "get",
    "records", "how many", "what is in", "show me", "tell me about",
    "get me", "find all", "data from",
}


def is_db_intent(message: str) -> bool:
    low = message.lower()
    return any(kw in low for kw in _DB_KEYWORDS)


# ──────────────────────────────────────────────────────────────────────────────
# AI model clients
# ──────────────────────────────────────────────────────────────────────────────

def get_model_client(model_type: str) -> Tuple[OpenAI, str]:
    if model_type == "censored":
        return (
            OpenAI(base_url="https://api.groq.com/openai/v1", api_key=GROQ_API_KEY),
            "llama-3.3-70b-versatile",
        )
    return (
        OpenAI(base_url="https://router.huggingface.co/v1", api_key=HF_TOKEN),
        "dphn/Dolphin-Mistral-24B-Venice-Edition:featherless-ai",
    )


def get_vision_client() -> Tuple[OpenAI, str]:
    return (
        OpenAI(base_url="https://router.huggingface.co/v1", api_key=HF_TOKEN),
        "Qwen/Qwen2.5-VL-7B-Instruct:fastest",
    )


# ──────────────────────────────────────────────────────────────────────────────
# Image analysis
# ──────────────────────────────────────────────────────────────────────────────

async def analyze_image_pipeline(
    image_data: str, user_message: str, model_type: str,
    user_id: str, memories: List[Dict]
) -> Tuple[str, str]:
    t0 = time.time()
    vc, vm = get_vision_client()
    vision_prompt = (
        'Analyze this image and return JSON:\n'
        '{"description":"...","objects":[],"text_in_image":null,'
        '"scene_type":"photo/diagram/screenshot/art/chart/other",'
        f'"technical_details":"...","user_question_relevance":"re: {user_message}"}}'
    )
    vision_json: Dict = {}
    try:
        vcomp = vc.chat.completions.create(
            model=vm,
            messages=[{"role": "user", "content": [
                {"type": "text", "text": vision_prompt},
                {"type": "image_url", "image_url": {"url": image_data}},
            ]}],
            max_tokens=600, temperature=0.1,
        )
        raw = re.sub(r"```json|```", "", vcomp.choices[0].message.content.strip()).strip()
        vision_json = json.loads(raw)
    except Exception as exc:
        logger.error("[VISION] %s", exc)
        vision_json = {"description": "Image provided", "scene_type": "unknown",
                       "user_question_relevance": user_message}
    ai_client, ai_model = get_model_client(model_type)
    memory_block = format_memories_for_prompt(memories)
    combined = (
        f"{memory_block}\n## Image Analysis:\n{json.dumps(vision_json, indent=2)}\n"
        f"## User:\n{user_message}\nProvide a thorough response."
    )
    comp  = ai_client.chat.completions.create(
        model=ai_model, messages=[{"role": "user", "content": combined}],
        max_tokens=1024, temperature=0.7,
    )
    final = comp.choices[0].message.content or ""
    logger.info("[VISION] Pipeline %.2fs", time.time() - t0)
    return final, json.dumps(vision_json)


# ──────────────────────────────────────────────────────────────────────────────
# Agent routing
# ──────────────────────────────────────────────────────────────────────────────

class Agent:
    def __init__(self, name: str, system_prompt: str, keywords: List[str]) -> None:
        self.name          = name
        self.system_prompt = system_prompt
        self.keywords      = keywords

    def build_prompt(self, message: str, context: str = "", memory_block: str = "") -> str:
        parts = [self.system_prompt]
        if memory_block:
            parts.append(f"\n{memory_block}\n")
        if context:
            parts.append(f"\nRelevant context:\n{context}\n")
        parts.append(f"\nUser: {message}")
        if context:
            parts.append("\nAnswer using the context where relevant.")
        return "\n".join(parts)


AGENTS: List[Agent] = [
    Agent("coding_agent",
          "You are a senior software engineer. Write clean, efficient, well-documented code with error handling.",
          ["code", "python", "script", "program", "function", "class", "debug", "algorithm",
           "api", "bug", "error", "javascript", "typescript", "sql", "refactor", "implement"]),
    Agent("research_agent",
          "You are an expert researcher. Provide comprehensive, structured explanations with bullet points and clear headings.",
          ["research", "explain", "analyze", "investigate", "compare", "review", "report",
           "study", "summarize", "difference", "how does", "why does", "what is", "overview"]),
    Agent("vision_agent",
          "You are an expert visual analyst. Analyze images with precision and detail.",
          ["image", "photo", "picture", "screenshot", "diagram", "chart", "see", "look",
           "show", "visual", "analyze image", "what is this"]),
    Agent("data_agent",
          "You are a data analyst. Interpret data and generate actionable insights with clear formatting.",
          ["data", "chart", "graph", "plot", "statistics", "average", "sum", "trend",
           "insight", "analysis", "metrics", "kpi", "dashboard"]),
    Agent("website_agent",
          "You are an expert web developer and UI/UX designer. Help design, build, and improve websites.",
          ["website", "webpage", "landing page", "portfolio", "html", "css", "web app",
           "frontend", "design", "layout", "ui", "ux", "responsive"]),
    Agent("general_agent",
          "You are JAZZ — a sharp, uncensored AI assistant. Be direct, witty, and genuinely helpful. Use markdown.",
          ["hello", "hi", "help", "chat", "talk", "general"]),
]
DEFAULT_AGENT: Agent = AGENTS[-1]


def route_message(message: str, has_image: bool = False) -> Agent:
    if has_image:
        return next(a for a in AGENTS if a.name == "vision_agent")
    low = message.lower()
    best, best_score = DEFAULT_AGENT, 0
    for agent in AGENTS[:-1]:
        score = sum(1 for kw in agent.keywords if kw in low)
        if score > best_score:
            best, best_score = agent, score
    return best


# ──────────────────────────────────────────────────────────────────────────────
# NL-to-SQL (MySQL)
# ──────────────────────────────────────────────────────────────────────────────

_FORBIDDEN_SQL: List[str] = ["INSERT", "UPDATE", "DELETE", "DROP", "ALTER", "CREATE", "TRUNCATE"]


async def run_nl_to_sql(question: str, user_id: str) -> Dict:
    t0         = time.time()
    schema_ctx = get_schema_context(question)
    client, model = get_model_client("censored")
    prompt     = f"MySQL schema:\n{schema_ctx}\n\nConvert to safe SELECT:\n{question}\nReturn ONLY SQL."
    comp       = client.chat.completions.create(
        model=model, messages=[{"role": "user", "content": prompt}], max_tokens=250, temperature=0.05
    )
    sql = re.sub(r"```sql|```", "", comp.choices[0].message.content.strip()).strip()
    if any(kw in sql.upper() for kw in _FORBIDDEN_SQL) or not sql.upper().startswith("SELECT"):
        return {"success": False, "error": "Only SELECT queries allowed"}
    if "LIMIT" not in sql.upper():
        sql += " LIMIT 200"
    conn = get_mysql_connection()
    if not conn:
        return {"success": False, "error": "MySQL unavailable"}
    try:
        cur = conn.cursor(dictionary=True)
        cur.execute(sql)
        rows    = cur.fetchall()
        elapsed = int((time.time() - t0) * 1000)
        cur.close()
        conn.close()
        return {"success": True, "sql": sql, "results": rows, "row_count": len(rows), "duration_ms": elapsed}
    except Exception as exc:
        try: conn.close()
        except Exception: pass
        return {"success": False, "sql": sql, "error": str(exc)}


# ──────────────────────────────────────────────────────────────────────────────
# File workspace
# ──────────────────────────────────────────────────────────────────────────────

class FileManager:
    def __init__(self, base: Path) -> None:
        self.base = base.resolve()
        self.base.mkdir(parents=True, exist_ok=True)

    def _safe_path(self, filename: str) -> Path:
        name = os.path.basename(filename)
        if Path(name).suffix.lower() not in ALLOWED_EXTENSIONS:
            raise ValueError("Extension not allowed")
        resolved = (self.base / name).resolve()
        if not str(resolved).startswith(str(self.base)):
            raise ValueError("Path traversal denied")
        return resolved

    def create(self, filename: str, content: str) -> Dict:
        p = self._safe_path(filename)
        if p.exists():
            return {"success": False, "message": "File already exists"}
        p.write_text(content, encoding="utf-8")
        return {"success": True, "filename": filename}

    def read(self, filename: str) -> Dict:
        p = self._safe_path(filename)
        if not p.exists():
            return {"success": False, "message": "File not found"}
        return {"success": True, "filename": filename, "content": p.read_text(encoding="utf-8")}

    def update(self, filename: str, content: str) -> Dict:
        p = self._safe_path(filename)
        if not p.exists():
            return {"success": False, "message": "File not found"}
        p.write_text(content, encoding="utf-8")
        return {"success": True}

    def delete(self, filename: str) -> Dict:
        p = self._safe_path(filename)
        if not p.exists():
            return {"success": False, "message": "File not found"}
        p.unlink()
        return {"success": True}

    def list_files(self) -> List[Dict]:
        return sorted(
            [{"name": f.name, "size": f.stat().st_size,
              "modified": datetime.fromtimestamp(f.stat().st_mtime).isoformat()}
             for f in self.base.iterdir() if f.is_file()],
            key=lambda x: x["modified"], reverse=True,
        )


file_mgr = FileManager(WORKSPACE_DIR)


# ──────────────────────────────────────────────────────────────────────────────
# Website builder
# ──────────────────────────────────────────────────────────────────────────────

_STYLE_HINTS = {
    "modern":        "Clean, bold typography, white space, subtle shadows, sans-serif, CSS Grid",
    "retro":         "80s/90s aesthetic, pixel fonts, neon colors on dark bg, CRT scanline effects",
    "minimal":       "Extreme whitespace, single accent color, fine typography, zero decoration",
    "glassmorphism": "Frosted glass panels, backdrop-filter blur, translucency, light gradients",
    "brutalist":     "Raw HTML feel, bold borders, high contrast, asymmetric layout",
}

_WEBSITE_SYSTEM_PROMPT = """You are an elite frontend developer. Generate COMPLETE, SELF-CONTAINED single-file HTML websites.
- All CSS in <style> tag, all JS in <script> tag
- CDN links allowed: Google Fonts, cdnjs.cloudflare.com only
- Fully functional, visually impressive, responsive
- Smooth animations, no Lorem Ipsum
- Return ONLY the HTML code — no explanations, no markdown fences"""


async def _build_website_html(description: str, title: str, style: str, model_type: str) -> str:
    client, model = get_model_client(model_type)
    hint   = _STYLE_HINTS.get(style, _STYLE_HINTS["modern"])
    prompt = (f"Build a complete website for: {description}\nTitle: {title}\n"
              f"Style: {style} — {hint}\nReturn ONLY raw HTML.")
    comp = client.chat.completions.create(
        model=model,
        messages=[{"role": "system", "content": _WEBSITE_SYSTEM_PROMPT},
                  {"role": "user", "content": prompt}],
        max_tokens=4096, temperature=0.8,
    )
    html = comp.choices[0].message.content or ""
    html = re.sub(r"^```html\s*", "", html.strip(), flags=re.IGNORECASE)
    html = re.sub(r"```\s*$", "", html.strip())
    return html.strip()


async def _update_website_html(current_html: str, instructions: str, model_type: str) -> str:
    client, model = get_model_client(model_type)
    prompt = (f"Existing HTML:\n{current_html[:8000]}\n\nApply: {instructions}\n"
              "Return COMPLETE updated HTML. No markdown.")
    comp = client.chat.completions.create(
        model=model,
        messages=[{"role": "system", "content": _WEBSITE_SYSTEM_PROMPT},
                  {"role": "user", "content": prompt}],
        max_tokens=4096, temperature=0.7,
    )
    html = comp.choices[0].message.content or ""
    html = re.sub(r"^```html\s*", "", html.strip(), flags=re.IGNORECASE)
    html = re.sub(r"```\s*$", "", html.strip())
    return html.strip()


# ──────────────────────────────────────────────────────────────────────────────
# Code runner
# ──────────────────────────────────────────────────────────────────────────────

_PYTHON_BLOCKED = [
    r"import\s+os", r"import\s+subprocess", r"import\s+sys", r"__import__",
    r"open\s*\(", r"exec\s*\(", r"eval\s*\(", r"compile\s*\(",
    r"importlib", r"shutil", r"socket", r"requests\.", r"urllib",
]


def _is_code_safe(code: str, language: str) -> Tuple[bool, str]:
    if language == "python":
        for pattern in _PYTHON_BLOCKED:
            if re.search(pattern, code, re.IGNORECASE):
                return False, f"Blocked: {pattern}"
    return True, ""


def _run_code_sync(code: str, language: str) -> Dict:
    t0 = time.time()
    if language == "python":
        cmd = [sys.executable, "-c", code]
    elif language == "javascript":
        cmd = ["node", "-e", code]
    elif language == "bash":
        cmd = ["bash", "-c", code]
    else:
        return {"stdout": "", "stderr": "Unsupported language", "exit_code": 1, "duration_ms": 0}
    try:
        result = subprocess.run(cmd, capture_output=True, text=True,
                                timeout=CODE_EXEC_TIMEOUT, cwd=tempfile.gettempdir())
        return {"stdout": result.stdout[:CODE_EXEC_MAX_OUT],
                "stderr": result.stderr[:CODE_EXEC_MAX_OUT],
                "exit_code": result.returncode,
                "duration_ms": int((time.time() - t0) * 1000)}
    except subprocess.TimeoutExpired:
        return {"stdout": "", "stderr": f"⏱ Timed out ({CODE_EXEC_TIMEOUT}s)",
                "exit_code": -1, "duration_ms": CODE_EXEC_TIMEOUT * 1000}
    except FileNotFoundError as e:
        return {"stdout": "", "stderr": f"Runtime not found: {e}", "exit_code": -1, "duration_ms": 0}
    except Exception as e:
        return {"stdout": "", "stderr": str(e), "exit_code": -1, "duration_ms": 0}


# ──────────────────────────────────────────────────────────────────────────────
# SQLite query interface
# ──────────────────────────────────────────────────────────────────────────────

async def run_sqlite_nl_query(question: str) -> Dict:
    schema_ctx    = get_schema_context(question, top_k=5)
    client, model = get_model_client("censored")
    prompt = (f"SQLite schema:\n{schema_ctx}\n\nWrite safe SELECT for: {question}\n"
              "Rules: SELECT only. Add LIMIT 100. Return ONLY SQL.")
    comp = client.chat.completions.create(
        model=model, messages=[{"role": "user", "content": prompt}],
        max_tokens=300, temperature=0.05
    )
    sql = re.sub(r"```sql|```", "", comp.choices[0].message.content.strip()).strip()
    if not sql.upper().startswith("SELECT"):
        return {"success": False, "error": "Must be a SELECT statement", "sql": sql}
    return await _run_sqlite_raw(sql)


async def _run_sqlite_raw(sql: str) -> Dict:
    t0 = time.time()
    try:
        rows    = await db_fetchall(sql)
        elapsed = int((time.time() - t0) * 1000)
        return {"success": True, "sql": sql, "rows": rows,
                "row_count": len(rows), "duration_ms": elapsed}
    except Exception as exc:
        return {"success": False, "sql": sql, "error": str(exc), "rows": [], "row_count": 0}


# ──────────────────────────────────────────────────────────────────────────────
# Agent Job engine
# ──────────────────────────────────────────────────────────────────────────────

class JobLogger:
    def __init__(self) -> None:
        self.lines:   List[str]  = []
        self.applied: List[Dict] = []

    def log(self, msg: str) -> None:
        ts    = datetime.now().strftime("%H:%M:%S")
        entry = f"[{ts}] {msg}"
        self.lines.append(entry)
        logger.info("[JOB] %s", msg)

    def add_applied(self, job: Dict) -> None:
        self.applied.append(job)

    def output(self) -> str:
        return "\n".join(self.lines)


def _is_within_time_window(start_str: str, end_str: str) -> bool:
    try:
        now   = datetime.now()
        fmt   = "%H:%M"
        start = datetime.strptime(start_str, fmt).replace(
            year=now.year, month=now.month, day=now.day)
        end   = datetime.strptime(end_str, fmt).replace(
            year=now.year, month=now.month, day=now.day)
        return start <= now <= end
    except Exception:
        return True


def _run_naukri_agent(creds: Dict, params: Dict, jlog: JobLogger) -> List[Dict]:
    if not _PLAYWRIGHT_AVAILABLE:
        jlog.log("❌ Playwright not installed.")
        return []
    from playwright.sync_api import sync_playwright
    email      = creds.get("email", "")
    password   = creds.get("password", "")
    roles      = params.get("roles", ["Data Analyst"])
    location   = params.get("location", "")
    exp_min    = params.get("experience_min", 0)
    exp_max    = params.get("experience_max", 10)
    max_apply  = min(params.get("max_apply", 10), 50)
    kw_exclude = [k.lower() for k in params.get("keywords_exclude", [])]
    applied: List[Dict] = []
    jlog.log(f"🚀 Naukri agent | roles={roles} location={location} exp={exp_min}-{exp_max}")
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True, args=["--no-sandbox", "--disable-dev-shm-usage"])
        ctx     = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36",
            viewport={"width": 1280, "height": 800},
        )
        page = ctx.new_page()
        page.set_default_timeout(30_000)
        try:
            jlog.log("🔐 Logging in to Naukri...")
            page.goto("https://www.naukri.com/nlogin/login", wait_until="domcontentloaded")
            page.wait_for_timeout(2000)
            email_sel = "input[placeholder*='Email'], input[type='email'], #usernameField"
            page.wait_for_selector(email_sel, timeout=10_000)
            page.fill(email_sel, email)
            page.wait_for_timeout(400)
            page.fill("input[type='password'], #passwordField", password)
            page.wait_for_timeout(400)
            page.click("button[type='submit']")
            page.wait_for_timeout(4000)
            if "login" in page.url.lower():
                jlog.log("❌ Login failed")
                browser.close()
                return []
            jlog.log("✅ Logged in to Naukri")
        except Exception as exc:
            jlog.log(f"❌ Login error: {exc}")
            browser.close()
            return []
        total_applied = 0
        for role in roles:
            if total_applied >= max_apply:
                break
            jlog.log(f"\n🔍 Searching: '{role}'")
            kw_enc  = role.replace(" ", "%20")
            loc_enc = location.replace(" ", "%20") if location else ""
            search_url = (
                f"https://www.naukri.com/jobs?k={kw_enc}&l={loc_enc}&experience={exp_min}&to={exp_max}"
                if loc_enc
                else f"https://www.naukri.com/jobs?k={kw_enc}&experience={exp_min}&to={exp_max}"
            )
            try:
                page.goto(search_url, wait_until="domcontentloaded")
                page.wait_for_timeout(3000)
            except Exception as exc:
                jlog.log(f"⚠ Search error: {exc}")
                continue
            job_cards = page.query_selector_all(
                "article.jobTuple, div.srp-jobtuple-wrapper, div[class*='jobTuple'], li[class*='jobTupleHeader']"
            )
            jlog.log(f"   Found {len(job_cards)} cards")
            for card in job_cards:
                if total_applied >= max_apply:
                    break
                try:
                    title_el   = card.query_selector("a.title, a[class*='title'], a[class*='jobTitle']")
                    company_el = card.query_selector("a.subTitle, a[class*='company'], span[class*='company']")
                    title   = title_el.inner_text().strip()   if title_el   else "Unknown"
                    company = company_el.inner_text().strip() if company_el else "Unknown"
                    job_url = title_el.get_attribute("href")  if title_el   else ""
                    if any(kw in title.lower() for kw in kw_exclude):
                        continue
                    if not job_url:
                        continue
                    jlog.log(f"   📋 {title} @ {company}")
                    job_page = ctx.new_page()
                    try:
                        job_page.goto(job_url, wait_until="domcontentloaded")
                        job_page.wait_for_timeout(2000)
                        apply_btn = job_page.query_selector(
                            "button:has-text('Apply'), button:has-text('Easy Apply'), a:has-text('Apply')"
                        )
                        if not apply_btn:
                            job_page.close()
                            continue
                        apply_btn.click()
                        job_page.wait_for_timeout(2500)
                        for _ in range(4):
                            step_btn = job_page.query_selector(
                                "button:has-text('Apply'),button:has-text('Submit'),button:has-text('Continue')"
                            )
                            if not step_btn:
                                break
                            try:
                                step_btn.click()
                                job_page.wait_for_timeout(1500)
                            except Exception:
                                break
                        applied.append({"title": title, "company": company, "url": job_url,
                                        "applied_at": datetime.utcnow().isoformat(), "status": "applied"})
                        jlog.add_applied(applied[-1])
                        total_applied += 1
                        jlog.log(f"      ✅ Applied! ({total_applied}/{max_apply})")
                    except Exception as exc:
                        jlog.log(f"      ⚠ {exc}")
                    finally:
                        try: job_page.close()
                        except Exception: pass
                except Exception as exc:
                    jlog.log(f"   ⚠ Card error: {exc}")
        browser.close()
        jlog.log(f"\n🏁 Naukri done. Applied: {total_applied}")
    return applied


def _run_linkedin_agent(creds: Dict, params: Dict, jlog: JobLogger) -> List[Dict]:
    if not _PLAYWRIGHT_AVAILABLE:
        jlog.log("❌ Playwright not installed.")
        return []
    from playwright.sync_api import sync_playwright
    email     = creds.get("email", "")
    password  = creds.get("password", "")
    roles     = params.get("roles", ["Data Analyst"])
    location  = params.get("location", "India")
    max_apply = min(params.get("max_apply", 5), 25)
    applied: List[Dict] = []
    jlog.log(f"🚀 LinkedIn agent | roles={roles}")
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True, args=["--no-sandbox"])
        ctx     = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36"
        )
        page = ctx.new_page()
        page.set_default_timeout(30_000)
        try:
            page.goto("https://www.linkedin.com/login", wait_until="domcontentloaded")
            page.fill("#username", email)
            page.fill("#password", password)
            page.click("button[type='submit']")
            page.wait_for_timeout(4000)
            if "checkpoint" in page.url or "login" in page.url:
                jlog.log("❌ LinkedIn login failed")
                browser.close()
                return []
            jlog.log("✅ LinkedIn logged in")
        except Exception as exc:
            jlog.log(f"❌ {exc}")
            browser.close()
            return []
        total_applied = 0
        for role in roles:
            if total_applied >= max_apply:
                break
            try:
                kw_enc  = role.replace(" ", "%20")
                loc_enc = location.replace(" ", "%20")
                page.goto(f"https://www.linkedin.com/jobs/search/?keywords={kw_enc}&location={loc_enc}&f_LF=f_AL",
                          wait_until="domcontentloaded")
                page.wait_for_timeout(3000)
                cards = page.query_selector_all("li.jobs-search-results__list-item")
                jlog.log(f"   Found {len(cards)} listings")
                for card in cards:
                    if total_applied >= max_apply:
                        break
                    try:
                        card.click()
                        page.wait_for_timeout(2000)
                        title_el   = page.query_selector("h1.job-details-jobs-unified-top-card__job-title")
                        company_el = page.query_selector("div.job-details-jobs-unified-top-card__company-name")
                        title   = title_el.inner_text().strip()   if title_el   else "Unknown"
                        company = company_el.inner_text().strip() if company_el else "Unknown"
                        easy_btn = page.query_selector("button:has-text('Easy Apply')")
                        if not easy_btn:
                            continue
                        easy_btn.click()
                        page.wait_for_timeout(2000)
                        for _ in range(5):
                            next_btn = page.query_selector(
                                "button:has-text('Next'),button:has-text('Review'),button:has-text('Submit application')"
                            )
                            if not next_btn:
                                break
                            next_btn.click()
                            page.wait_for_timeout(1500)
                        applied.append({"title": title, "company": company,
                                        "applied_at": datetime.utcnow().isoformat(), "status": "applied"})
                        jlog.add_applied(applied[-1])
                        total_applied += 1
                        jlog.log(f"   ✅ Applied: {title} @ {company}")
                    except Exception as exc:
                        jlog.log(f"   ⚠ {exc}")
            except Exception as exc:
                jlog.log(f"⚠ Search error: {exc}")
        browser.close()
        jlog.log(f"\n🏁 LinkedIn done. Applied: {total_applied}")
    return applied


def _execute_job_sync(job_id: str, bypass_time_window: bool = False) -> None:
    import sqlite3 as _s3

    def _db(sql: str, p: tuple = ()):
        with _s3.connect(DB_PATH) as c:
            c.row_factory = _s3.Row
            return c.execute(sql, p).fetchone()

    def _dbw(sql: str, p: tuple = ()) -> None:
        with _s3.connect(DB_PATH) as c:
            c.execute(sql, p)
            c.commit()

    row = _db("SELECT * FROM agent_jobs WHERE id=?", (job_id,))
    if not row:
        logger.error("[JOB] Job %s not found", job_id)
        return
    job = dict(row)

    if not bypass_time_window:
        tw_start = job.get("time_window_start") or "00:00"
        tw_end   = job.get("time_window_end")   or "23:59"
        if not _is_within_time_window(tw_start, tw_end):
            logger.info("[JOB] '%s' outside window, skipping", job["name"])
            return

    now_iso = datetime.utcnow().isoformat()
    log_id  = str(uuid.uuid4())
    _dbw("UPDATE agent_jobs SET status='running', last_run_at=?, updated_at=? WHERE id=?",
         (now_iso, now_iso, job_id))
    _dbw("INSERT INTO agent_job_logs (id,job_id,user_id,status,output,started_at) VALUES (?,?,?,?,?,?)",
         (log_id, job_id, job["user_id"], "running",
          f"[{datetime.now().strftime('%H:%M:%S')}] ⚡ Starting '{job['name']}'\n", now_iso))

    jlog = JobLogger()
    jlog.log(f"⚡ Job '{job['name']}' starting (type={job['job_type']})")
    t0            = time.time()
    applied:      List[Dict]     = []
    error:        Optional[str]  = None
    final_status: str            = "completed"

    try:
        creds  = decrypt_creds(job["credentials_enc"]) if job.get("credentials_enc") else {}
        params = json.loads(job.get("parameters") or "{}")
        if job["job_type"] == "naukri_apply":
            applied = _run_naukri_agent(creds, params, jlog)
        elif job["job_type"] == "linkedin_apply":
            applied = _run_linkedin_agent(creds, params, jlog)
        else:
            jlog.log(f"⚠ Unknown job type: {job['job_type']}")
        jlog.log(f"\n✅ Done. Applied to {len(applied)} positions.")
    except Exception as exc:
        tb           = traceback.format_exc()
        error        = f"{exc}\n\nTraceback:\n{tb}"
        final_status = "failed"
        jlog.log(f"\n❌ Failed: {exc}")
        logger.error("[JOB] '%s' failed: %s\n%s", job["name"], exc, tb)

    elapsed  = int(time.time() - t0)
    done_iso = datetime.utcnow().isoformat()
    _dbw("UPDATE agent_job_logs SET status=?,output=?,applied_jobs=?,error=?,completed_at=?,duration_sec=? WHERE id=?",
         (final_status, jlog.output(), json.dumps(applied, default=str), error, done_iso, elapsed, log_id))
    _dbw("UPDATE agent_jobs SET status='idle',last_run_at=?,last_run_status=?,total_runs=total_runs+1,total_applied=total_applied+?,updated_at=? WHERE id=?",
         (done_iso, final_status, len(applied), done_iso, job_id))

    # Send notification to user
    notification_title   = f"Job '{job['name']}' {final_status}"
    notification_message = f"Applied to {len(applied)} positions in {elapsed}s."
    if final_status == "failed":
        notification_message = f"Job failed: {str(error)[:100]}"
    with _s3.connect(DB_PATH) as c:
        c.execute(
            "INSERT INTO notifications (id,user_id,title,message,type,link) VALUES (?,?,?,?,?,?)",
            (str(uuid.uuid4()), job["user_id"], notification_title, notification_message,
             "success" if final_status == "completed" else "error", "/panel-jobs-logs"),
        )
        c.commit()
    logger.info("[JOB] '%s' %s in %ds — applied: %d", job["name"], final_status, elapsed, len(applied))


async def execute_job(job_id: str, bypass_time_window: bool = False) -> None:
    loop = asyncio.get_event_loop()
    await loop.run_in_executor(_executor, _execute_job_sync, job_id, bypass_time_window)


# ──────────────────────────────────────────────────────────────────────────────
# APScheduler
# ──────────────────────────────────────────────────────────────────────────────

scheduler = BackgroundScheduler(daemon=True)


def _schedule_job(job: Dict) -> None:
    sid = f"agent_{job['id']}"
    try:
        parts = job["cron_schedule"].split()
        if len(parts) != 5:
            raise ValueError(f"Invalid cron: {job['cron_schedule']}")
        minute, hour, day, month, dow = parts
        trigger = CronTrigger(minute=minute, hour=hour, day=day, month=month, day_of_week=dow)

        def _wrapper(jid: str = job["id"]) -> None:
            import asyncio as _aio
            loop = _aio.new_event_loop()
            try:
                loop.run_until_complete(execute_job(jid, bypass_time_window=False))
            finally:
                loop.close()

        if scheduler.get_job(sid):
            scheduler.remove_job(sid)
        scheduler.add_job(_wrapper, trigger=trigger, id=sid, replace_existing=True)
        logger.info("[SCHEDULER] Registered '%s' (%s)", job["name"], job["cron_schedule"])
    except Exception as exc:
        logger.error("[SCHEDULER] Failed to register %s: %s", job["id"], exc)


def _unschedule_job(job_id: str) -> None:
    sid = f"agent_{job_id}"
    if scheduler.get_job(sid):
        scheduler.remove_job(sid)


def _load_all_jobs_into_scheduler() -> None:
    import sqlite3 as _s3
    try:
        with _s3.connect(DB_PATH) as conn:
            conn.row_factory = _s3.Row
            rows = conn.execute("SELECT * FROM agent_jobs WHERE enabled=1").fetchall()
        for row in rows:
            _schedule_job(dict(row))
        logger.info("[SCHEDULER] Loaded %d agent jobs ✅", len(rows))
    except Exception as exc:
        logger.error("[SCHEDULER] Failed to load jobs: %s", exc)


# ──────────────────────────────────────────────────────────────────────────────
# FastAPI app
# ──────────────────────────────────────────────────────────────────────────────

@asynccontextmanager
async def lifespan(app: FastAPI) -> AsyncGenerator:
    logger.info("=" * 60)
    logger.info("[STARTUP] JAZZ AI v7.0 Production Edition")
    logger.info("  DB:         %s", DB_PATH)
    logger.info("  ADMIN:      %s", ADMIN_EMAIL)
    logger.info("  CHROMADB:   %s", "✅" if chroma_client else "❌")
    logger.info("  GROQ:       %s", "✅" if GROQ_API_KEY else "❌")
    logger.info("  HF_TOKEN:   %s", "✅" if HF_TOKEN else "❌")
    logger.info("  PLAYWRIGHT: %s", "✅" if _PLAYWRIGHT_AVAILABLE else "❌")
    logger.info("=" * 60)
    await init_db()
    index_sqlite_schema()
    if MYSQL_ENABLED:
        _init_mysql_pool()
        index_mysql_schema()
    scheduler.add_job(index_sqlite_schema, trigger=IntervalTrigger(hours=1),
                      id="index_sqlite", replace_existing=True)
    if MYSQL_ENABLED:
        scheduler.add_job(index_mysql_schema, trigger=IntervalTrigger(minutes=30),
                          id="mysql_reindex", replace_existing=True)
    scheduler.start()
    _load_all_jobs_into_scheduler()
    logger.info("[STARTUP] All services ready ✅")
    yield
    logger.info("[SHUTDOWN] Stopping...")
    scheduler.shutdown(wait=False)
    _executor.shutdown(wait=False)


app = FastAPI(
    title="JAZZ AI",
    version="7.0.0",
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
    response = await call_next(request)
    response.headers["X-Request-ID"] = rid
    response.headers["X-Powered-By"] = "JAZZ-AI-v7.0"
    return response


app.mount("/preview", StaticFiles(directory=str(SITES_DIR), html=True), name="preview")

_static_dir = Path("static")
if _static_dir.exists():
    app.mount("/static", StaticFiles(directory="static"), name="static")

    @app.get("/")
    async def root():
        return FileResponse("static/index.html")
else:
    @app.get("/")
    async def root():
        return {"status": "JAZZ AI API v7.0", "docs": "/api/docs"}


# ══════════════════════════════════════════════════════════════════════════════
# AUTH routes
# ══════════════════════════════════════════════════════════════════════════════

@app.post("/auth/signup", response_model=TokenResponse)
async def signup(req: UserSignup):
    if await db_fetchone("SELECT id FROM users WHERE email=?", (req.email,)):
        raise HTTPException(400, "Email already registered")
    uid   = str(uuid.uuid4())
    phash = pwd_context.hash(req.password)
    role  = "admin" if req.email == ADMIN_EMAIL else "client"
    sub   = "enterprise" if req.email == ADMIN_EMAIL else "free"
    await db_execute(
        "INSERT INTO users (id,email,password_hash,full_name,role,subscription) VALUES (?,?,?,?,?,?)",
        (uid, req.email, phash, req.full_name or "", role, sub),
    )
    # Welcome notification
    await push_notification(uid, "Welcome to JAZZ AI! 🎉",
                            "Your account is set up. Start chatting or explore the features.",
                            "success")
    token = create_access_token({"sub": uid, "email": req.email, "role": role})
    return TokenResponse(access_token=token,
                         user={"id": uid, "email": req.email, "role": role,
                               "subscription": sub, "full_name": req.full_name or ""})


@app.post("/auth/login", response_model=TokenResponse)
async def login(req: UserLogin):
    if req.email == ADMIN_EMAIL and ADMIN_PASSWORD and req.password == ADMIN_PASSWORD:
        user = await db_fetchone("SELECT * FROM users WHERE email=?", (ADMIN_EMAIL,))
        if not user:
            uid   = str(uuid.uuid4())
            phash = pwd_context.hash(ADMIN_PASSWORD)
            await db_execute(
                "INSERT INTO users (id,email,password_hash,full_name,role,subscription) VALUES (?,?,?,?,?,?)",
                (uid, ADMIN_EMAIL, phash, "Admin", "admin", "enterprise"),
            )
            user = {"id": uid, "email": ADMIN_EMAIL, "role": "admin",
                    "subscription": "enterprise", "full_name": "Admin", "memory_enabled": 1}
        token = create_access_token({"sub": user["id"], "email": req.email, "role": "admin"})
        return TokenResponse(access_token=token,
                             user={"id": user["id"], "email": req.email, "role": "admin",
                                   "subscription": "enterprise", "full_name": "Admin",
                                   "memory_enabled": True})
    user = await db_fetchone("SELECT * FROM users WHERE email=?", (req.email,))
    if not user:
        raise HTTPException(401, "Invalid credentials")
    try:
        ok = pwd_context.verify(req.password, user["password_hash"])
    except Exception:
        ok = False
    if not ok:
        raise HTTPException(401, "Invalid credentials")
    token = create_access_token({"sub": user["id"], "email": user["email"], "role": user["role"]})
    return TokenResponse(access_token=token,
                         user={"id": user["id"], "email": user["email"], "role": user["role"],
                               "subscription": user["subscription"],
                               "full_name": user.get("full_name", ""),
                               "memory_enabled": bool(user.get("memory_enabled", 1))})


@app.post("/auth/admin-login", response_model=TokenResponse)
async def admin_login(req: UserLogin):
    if req.email != ADMIN_EMAIL:
        raise HTTPException(403, "Admin access only")
    return await login(req)


@app.get("/auth/me")
async def me(user: Dict = Depends(get_current_user)):
    row = await db_fetchone(
        "SELECT id,email,full_name,role,subscription,memory_enabled,avatar_url,preferred_model,created_at FROM users WHERE id=?",
        (user["sub"],),
    )
    if not row:
        raise HTTPException(404, "User not found")
    row["memory_enabled"] = bool(row.get("memory_enabled", 1))
    return row


@app.post("/auth/logout")
async def logout(user: Dict = Depends(get_current_user)):
    return {"message": "Logged out"}


@app.post("/auth/change-password")
async def change_password(req: ChangePasswordRequest, user: Dict = Depends(get_current_user)):
    row = await db_fetchone("SELECT password_hash FROM users WHERE id=?", (user["sub"],))
    if not row:
        raise HTTPException(404, "User not found")
    try:
        if not pwd_context.verify(req.current_password, row["password_hash"]):
            raise HTTPException(400, "Current password is incorrect")
    except Exception:
        raise HTTPException(400, "Current password is incorrect")
    new_hash = pwd_context.hash(req.new_password)
    await db_execute("UPDATE users SET password_hash=?,updated_at=? WHERE id=?",
                     (new_hash, datetime.utcnow().isoformat(), user["sub"]))
    await push_notification(user["sub"], "Password changed", "Your password was changed successfully.", "info")
    return {"success": True}


@app.delete("/auth/delete-account")
async def delete_account(user: Dict = Depends(get_current_user)):
    uid = user["sub"]
    # Clean up ChromaDB docs
    if documents_collection:
        try:
            existing = documents_collection.get(where={"user_id": uid})
            if existing["ids"]:
                documents_collection.delete(ids=existing["ids"])
        except Exception:
            pass
    # Delete site files
    sites = await db_fetchall("SELECT filename FROM websites WHERE user_id=?", (uid,))
    for site in sites:
        sp = SITES_DIR / site["filename"]
        if sp.exists():
            sp.unlink()
    # Cascade deletes handle DB records
    await db_execute("DELETE FROM users WHERE id=?", (uid,))
    return {"success": True, "message": "Account and all data deleted."}


# ══════════════════════════════════════════════════════════════════════════════
# PROFILE / USER routes
# ══════════════════════════════════════════════════════════════════════════════

@app.get("/profile")
async def get_profile(user: Dict = Depends(get_current_user)):
    row = await db_fetchone(
        "SELECT id,email,full_name,role,subscription,memory_enabled,avatar_url,preferred_model,created_at FROM users WHERE id=?",
        (user["sub"],),
    )
    if not row:
        raise HTTPException(404, "Profile not found")
    row["memory_enabled"] = bool(row.get("memory_enabled", 1))
    return row


@app.put("/profile")
async def update_profile(req: ProfileUpdate, user: Dict = Depends(get_current_user)):
    data = req.model_dump(exclude_none=True)
    if not data:
        raise HTTPException(400, "Nothing to update")
    sets = ", ".join(f"{k}=?" for k in data)
    await db_execute(f"UPDATE users SET {sets}, updated_at=? WHERE id=?",
                     tuple(data.values()) + (datetime.utcnow().isoformat(), user["sub"]))
    return {"success": True}


@app.get("/user/stats")
async def user_stats(user: Dict = Depends(get_current_user)):
    uid      = user["sub"]
    sub_tier = await get_user_subscription(uid)
    limits   = SUBSCRIPTION_LIMITS.get(sub_tier, SUBSCRIPTION_LIMITS["free"])
    return {
        "subscription":          sub_tier,
        "messages_used_today":   rate_limiter.usage_today(uid),
        "messages_limit":        limits["messages_per_day"],
        "documents":             await db_count("SELECT COUNT(*) FROM documents WHERE user_id=?", (uid,)),
        "document_limit":        limits["documents"],
        "websites":              await db_count("SELECT COUNT(*) FROM websites WHERE user_id=?", (uid,)),
        "websites_limit":        limits.get("websites", 3),
        "mysql_access":          limits["mysql_access"],
        "image_analysis":        limits["image_analysis"],
        "memory_count":          await db_count("SELECT COUNT(*) FROM user_memories WHERE user_id=?", (uid,)),
        "prompt_templates":      await db_count("SELECT COUNT(*) FROM prompt_templates WHERE user_id=?", (uid,)),
        "agent_jobs":            await db_count("SELECT COUNT(*) FROM agent_jobs WHERE user_id=?", (uid,)),
        "agent_jobs_limit":      limits.get("agent_jobs", 1),
        "api_keys":              await db_count("SELECT COUNT(*) FROM api_keys WHERE user_id=? AND is_active=1", (uid,)),
        "api_keys_limit":        limits.get("api_keys", 2),
        "code_runs_per_day":     limits.get("code_runs_per_day", 20),
        "playwright_available":  _PLAYWRIGHT_AVAILABLE,
        "version":               "7.0.0",
    }


@app.get("/user/export")
async def gdpr_export(user: Dict = Depends(get_current_user)):
    """Full GDPR data export for the current user."""
    uid = user["sub"]
    profile   = await db_fetchone(
        "SELECT id,email,full_name,role,subscription,created_at FROM users WHERE id=?", (uid,)
    )
    chats     = await db_fetchall(
        "SELECT id,message,response,agent_used,model_used,created_at FROM chat_history WHERE user_id=? ORDER BY created_at DESC",
        (uid,),
    )
    memories  = await db_fetchall("SELECT content,category,source,created_at FROM user_memories WHERE user_id=?", (uid,))
    documents = await db_fetchall("SELECT filename,file_type,file_size,chunk_count,uploaded_at FROM documents WHERE user_id=?", (uid,))
    websites  = await db_fetchall("SELECT title,description,created_at FROM websites WHERE user_id=?", (uid,))
    jobs      = await db_fetchall("SELECT name,job_type,cron_schedule,total_runs,total_applied,created_at FROM agent_jobs WHERE user_id=?", (uid,))
    prompts   = await db_fetchall("SELECT title,category,use_count,created_at FROM prompt_templates WHERE user_id=?", (uid,))

    export_data = {
        "export_date": datetime.utcnow().isoformat(),
        "profile":     profile,
        "chat_history": chats,
        "memories":    memories,
        "documents":   documents,
        "websites":    websites,
        "agent_jobs":  jobs,
        "prompt_library": prompts,
    }
    content  = json.dumps(export_data, indent=2, default=str)
    filename = f"jazz_export_{uid[:8]}_{datetime.utcnow().strftime('%Y%m%d')}.json"
    return StreamingResponse(
        iter([content.encode()]),
        media_type="application/json",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


# ══════════════════════════════════════════════════════════════════════════════
# NOTIFICATIONS routes
# ══════════════════════════════════════════════════════════════════════════════

@app.get("/user/notifications")
async def list_notifications(
    unread: bool = False,
    limit: int   = 30,
    user: Dict   = Depends(get_current_user),
):
    uid   = user["sub"]
    sql   = "SELECT * FROM notifications WHERE user_id=?"
    args: list = [uid]
    if unread:
        sql  += " AND is_read=0"
    sql += " ORDER BY created_at DESC LIMIT ?"
    args.append(limit)
    rows  = await db_fetchall(sql, tuple(args))
    count = await db_count("SELECT COUNT(*) FROM notifications WHERE user_id=? AND is_read=0", (uid,))
    return {"notifications": rows, "unread_count": count}


@app.post("/user/notifications/{nid}/read")
async def mark_notification_read(nid: str, user: Dict = Depends(get_current_user)):
    await db_execute("UPDATE notifications SET is_read=1 WHERE id=? AND user_id=?",
                     (nid, user["sub"]))
    return {"success": True}


@app.post("/user/notifications/read-all")
async def mark_all_read(user: Dict = Depends(get_current_user)):
    await db_execute("UPDATE notifications SET is_read=1 WHERE user_id=?", (user["sub"],))
    return {"success": True}


@app.delete("/user/notifications/{nid}")
async def delete_notification(nid: str, user: Dict = Depends(get_current_user)):
    await db_execute("DELETE FROM notifications WHERE id=? AND user_id=?", (nid, user["sub"]))
    return {"success": True}


# ══════════════════════════════════════════════════════════════════════════════
# API KEYS routes
# ══════════════════════════════════════════════════════════════════════════════

@app.get("/user/api-keys")
async def list_api_keys(user: Dict = Depends(get_current_user)):
    uid  = user["sub"]
    rows = await db_fetchall(
        "SELECT id,name,key_prefix,last_used_at,expires_at,is_active,created_at FROM api_keys WHERE user_id=? ORDER BY created_at DESC",
        (uid,),
    )
    return {"api_keys": rows}


@app.post("/user/api-keys")
async def create_api_key(req: APIKeyCreate, user: Dict = Depends(get_current_user)):
    uid   = user["sub"]
    tier  = await get_user_subscription(uid)
    limit = SUBSCRIPTION_LIMITS.get(tier, SUBSCRIPTION_LIMITS["free"]).get("api_keys", 2)
    if limit != -1:
        count = await db_count("SELECT COUNT(*) FROM api_keys WHERE user_id=? AND is_active=1", (uid,))
        if count >= limit:
            raise HTTPException(403, f"API key limit ({limit}) reached for {tier} plan")
    # Generate key
    raw_key    = "jz_" + uuid.uuid4().hex + uuid.uuid4().hex[:8]
    key_hash   = hashlib.sha256(raw_key.encode()).hexdigest()
    key_prefix = raw_key[:12]
    kid        = str(uuid.uuid4())
    now        = datetime.utcnow().isoformat()
    await db_execute(
        "INSERT INTO api_keys (id,user_id,name,key_hash,key_prefix,expires_at,created_at) VALUES (?,?,?,?,?,?,?)",
        (kid, uid, req.name, key_hash, key_prefix, req.expires_at, now),
    )
    return {
        "success":    True,
        "id":         kid,
        "name":       req.name,
        "key":        raw_key,      # ONLY returned once
        "key_prefix": key_prefix,
        "created_at": now,
        "note":       "⚠ Save this key now — it will NOT be shown again.",
    }


@app.delete("/user/api-keys/{kid}")
async def delete_api_key(kid: str, user: Dict = Depends(get_current_user)):
    await db_execute("UPDATE api_keys SET is_active=0 WHERE id=? AND user_id=?",
                     (kid, user["sub"]))
    return {"success": True}


# ══════════════════════════════════════════════════════════════════════════════
# MEMORY routes
# ══════════════════════════════════════════════════════════════════════════════

@app.get("/memory")
async def list_memories(user: Dict = Depends(get_current_user)):
    uid = user["sub"]
    return {
        "memories":       await get_user_memories(uid),
        "count":          await db_count("SELECT COUNT(*) FROM user_memories WHERE user_id=?", (uid,)),
        "memory_enabled": await is_memory_enabled(uid),
        "max_memories":   MAX_MEMORIES,
    }


@app.post("/memory")
async def create_memory(req: MemoryCreate, user: Dict = Depends(get_current_user)):
    uid   = user["sub"]
    count = await db_count("SELECT COUNT(*) FROM user_memories WHERE user_id=?", (uid,))
    if count >= MAX_MEMORIES:
        raise HTTPException(400, f"Memory limit ({MAX_MEMORIES}) reached")
    mid = str(uuid.uuid4())
    await db_execute(
        "INSERT INTO user_memories (id,user_id,content,category,source) VALUES (?,?,?,?,?)",
        (mid, uid, req.content.strip(), req.category or "general", "manual"),
    )
    return {"success": True, "memory": await db_fetchone("SELECT * FROM user_memories WHERE id=?", (mid,))}


@app.put("/memory/{memory_id}")
async def update_memory(memory_id: str, req: MemoryUpdate, user: Dict = Depends(get_current_user)):
    row = await db_fetchone("SELECT id FROM user_memories WHERE id=? AND user_id=?",
                            (memory_id, user["sub"]))
    if not row:
        raise HTTPException(404, "Memory not found")
    await db_execute("UPDATE user_memories SET content=?,updated_at=? WHERE id=? AND user_id=?",
                     (req.content.strip(), datetime.utcnow().isoformat(), memory_id, user["sub"]))
    return {"success": True, "memory": await db_fetchone("SELECT * FROM user_memories WHERE id=?", (memory_id,))}


@app.delete("/memory/{memory_id}")
async def delete_memory(memory_id: str, user: Dict = Depends(get_current_user)):
    await db_execute("DELETE FROM user_memories WHERE id=? AND user_id=?", (memory_id, user["sub"]))
    return {"success": True}


@app.delete("/memory")
async def clear_all_memories(user: Dict = Depends(get_current_user)):
    await db_execute("DELETE FROM user_memories WHERE user_id=?", (user["sub"],))
    return {"success": True}


@app.put("/memory/settings/toggle")
async def toggle_memory(req: MemorySettingsUpdate, user: Dict = Depends(get_current_user)):
    await db_execute("UPDATE users SET memory_enabled=?,updated_at=? WHERE id=?",
                     (1 if req.enabled else 0, datetime.utcnow().isoformat(), user["sub"]))
    return {"success": True, "memory_enabled": req.enabled}


# ══════════════════════════════════════════════════════════════════════════════
# IMAGES routes
# ══════════════════════════════════════════════════════════════════════════════

@app.post("/images/upload")
async def upload_image(file: UploadFile = File(...), user: Dict = Depends(get_current_user)):
    if file.content_type not in ALLOWED_IMAGE_TYPES:
        raise HTTPException(400, f"Unsupported image type: {file.content_type}")
    content = await file.read()
    if len(content) > MAX_IMAGE_SIZE:
        raise HTTPException(413, "Image too large")
    b64 = base64.b64encode(content).decode("utf-8")
    return {"success": True, "data_url": f"data:{file.content_type};base64,{b64}",
            "filename": file.filename, "size": len(content), "mime_type": file.content_type}


# ══════════════════════════════════════════════════════════════════════════════
# CHAT SESSIONS routes
# ══════════════════════════════════════════════════════════════════════════════

@app.get("/chat/sessions")
async def list_sessions(user: Dict = Depends(get_current_user)):
    rows = await db_fetchall(
        "SELECT s.*, (SELECT COUNT(*) FROM chat_history WHERE session_id=s.id) AS message_count FROM chat_sessions s WHERE s.user_id=? ORDER BY s.updated_at DESC",
        (user["sub"],),
    )
    return {"sessions": rows}


@app.post("/chat/sessions")
async def create_session(req: ChatSessionCreate, user: Dict = Depends(get_current_user)):
    uid  = user["sub"]
    tier = await get_user_subscription(uid)
    lim  = SUBSCRIPTION_LIMITS.get(tier, SUBSCRIPTION_LIMITS["free"]).get("chat_sessions", 10)
    if lim != -1:
        count = await db_count("SELECT COUNT(*) FROM chat_sessions WHERE user_id=?", (uid,))
        if count >= lim:
            raise HTTPException(403, f"Chat session limit ({lim}) reached")
    sid = str(uuid.uuid4())
    now = datetime.utcnow().isoformat()
    await db_execute("INSERT INTO chat_sessions (id,user_id,title,created_at,updated_at) VALUES (?,?,?,?,?)",
                     (sid, uid, req.title, now, now))
    return {"success": True, "session": await db_fetchone("SELECT * FROM chat_sessions WHERE id=?", (sid,))}


@app.put("/chat/sessions/{sid}")
async def rename_session(sid: str, req: ChatSessionUpdate, user: Dict = Depends(get_current_user)):
    row = await db_fetchone("SELECT id FROM chat_sessions WHERE id=? AND user_id=?", (sid, user["sub"]))
    if not row:
        raise HTTPException(404, "Session not found")
    await db_execute("UPDATE chat_sessions SET title=?,updated_at=? WHERE id=?",
                     (req.title, datetime.utcnow().isoformat(), sid))
    return {"success": True}


@app.delete("/chat/sessions/{sid}")
async def delete_session(sid: str, user: Dict = Depends(get_current_user)):
    row = await db_fetchone("SELECT id FROM chat_sessions WHERE id=? AND user_id=?", (sid, user["sub"]))
    if not row:
        raise HTTPException(404, "Session not found")
    await db_execute("DELETE FROM chat_sessions WHERE id=?", (sid,))
    return {"success": True}


@app.get("/chat/sessions/{sid}/messages")
async def get_session_messages(sid: str, page: int = 1, per_page: int = 30,
                               user: Dict = Depends(get_current_user)):
    row = await db_fetchone("SELECT id FROM chat_sessions WHERE id=? AND user_id=?", (sid, user["sub"]))
    if not row:
        raise HTTPException(404, "Session not found")
    offset = (page - 1) * per_page
    rows   = await db_fetchall(
        "SELECT * FROM chat_history WHERE session_id=? ORDER BY created_at ASC LIMIT ? OFFSET ?",
        (sid, per_page, offset),
    )
    return {"messages": rows, "page": page, "per_page": per_page}


# ══════════════════════════════════════════════════════════════════════════════
# v7.1: Enhanced Chat Features
# ══════════════════════════════════════════════════════════════════════════════

@app.post("/chat/{chat_id}/feedback")
async def save_feedback(chat_id: str, rating: int = 1, feedback: str = "", 
                       user: Dict = Depends(get_current_user)):
    """Save thumbs up/down feedback for a message (1=good, -1=bad, 0=neutral)."""
    msg = await db_fetchone("SELECT id FROM chat_history WHERE id=? AND user_id=?", (chat_id, user["sub"]))
    if not msg:
        raise HTTPException(404, "Message not found")
    success = await save_message_feedback(chat_id, user["sub"], rating, feedback)
    return {"success": success}


@app.post("/chat/{chat_id}/regenerate")
async def regenerate_response(chat_id: str, user: Dict = Depends(get_current_user)):
    """Regenerate AI response for a user's message."""
    msg = await db_fetchone(
        "SELECT session_id FROM chat_history WHERE id=? AND user_id=?", (chat_id, user["sub"])
    )
    if not msg:
        raise HTTPException(404, "Message not found")
    new_response = await regenerate_message(msg["session_id"], user["sub"], chat_id)
    if not new_response:
        raise HTTPException(500, "Regeneration failed")
    return {"success": True, "new_response": new_response}


@app.post("/chat/sessions/{sid}/branch")
async def create_branch(sid: str, branch_from_msg: str = "", user: Dict = Depends(get_current_user)):
    """Branch conversation from a specific message."""
    session = await db_fetchone("SELECT id FROM chat_sessions WHERE id=? AND user_id=?", (sid, user["sub"]))
    if not session:
        raise HTTPException(404, "Session not found")
    new_sid = await branch_conversation(sid, user["sub"], branch_from_msg)
    if not new_sid:
        raise HTTPException(500, "Branching failed")
    return {"success": True, "new_session_id": new_sid}


@app.get("/chat/sessions/{sid}/summary")
async def get_summary(sid: str, user: Dict = Depends(get_current_user)):
    """Get or generate conversation summary."""
    summary = await db_fetchone(
        "SELECT summary FROM conversation_summaries WHERE session_id=? AND user_id=?",
        (sid, user["sub"]),
    )
    if summary:
        return {"summary": summary["summary"], "generated": False}
    new_summary = await summarize_conversation(sid, user["sub"])
    if new_summary:
        return {"summary": new_summary, "generated": True}
    raise HTTPException(404, "Could not generate summary")


@app.get("/chat/sessions/{sid}/topics")
async def get_topics(sid: str, user: Dict = Depends(get_current_user)):
    """Get conversation topics/tags."""
    rows = await db_fetchall(
        "SELECT topic, confidence FROM conversation_topics WHERE session_id=? AND user_id=?",
        (sid, user["sub"]),
    )
    return {"topics": rows}


@app.get("/chat/sessions/{sid}/export")
async def export_conv(sid: str, format_type: str = "json", user: Dict = Depends(get_current_user)):
    """Export conversation in various formats (json, markdown, csv)."""
    session = await db_fetchone("SELECT id FROM chat_sessions WHERE id=? AND user_id=?", (sid, user["sub"]))
    if not session:
        raise HTTPException(404, "Session not found")
    content = await export_conversation(sid, user["sub"], format_type)
    if not content:
        raise HTTPException(500, "Export failed")
    
    media_types = {
        "json": "application/json",
        "markdown": "text/markdown",
        "csv": "text/csv",
    }
    ext = {"json": "json", "markdown": "md", "csv": "csv"}
    
    return StreamingResponse(
        iter([content.encode()]),
        media_type=media_types.get(format_type, "text/plain"),
        headers={"Content-Disposition": f'attachment; filename="chat_{sid[:8]}.{ext.get(format_type, "txt")}"'},
    )


@app.post("/chat/{chat_id}/edit")
async def edit_message(chat_id: str, new_content: str = "", user: Dict = Depends(get_current_user)):
    """Edit a user's message (creates new version)."""
    msg = await db_fetchone(
        "SELECT message, session_id FROM chat_history WHERE id=? AND user_id=?", 
        (chat_id, user["sub"])
    )
    if not msg:
        raise HTTPException(404, "Message not found")
    
    # Save edit history
    await db_execute(
        "INSERT INTO message_edits (id,message_id,old_content,new_content) VALUES (?,?,?,?)",
        (str(uuid.uuid4()), chat_id, msg["message"], new_content),
    )
    # Update message
    await db_execute(
        "UPDATE chat_history SET message=? WHERE id=?", (new_content, chat_id)
    )
    return {"success": True}


@app.get("/chat/search")
async def search_messages(q: str = "", session_id: str = "", user: Dict = Depends(get_current_user)):
    """Search through user's chat messages."""
    uid = user["sub"]
    if not q:
        raise HTTPException(400, "Search query required")
    
    sql = "SELECT id,message,response,created_at FROM chat_history WHERE user_id=? AND (message LIKE ? OR response LIKE ?)"
    params: list = [uid, f"%{q}%", f"%{q}%"]
    
    if session_id:
        sql += " AND session_id=?"
        params.append(session_id)
    
    sql += " ORDER BY created_at DESC LIMIT 20"
    rows = await db_fetchall(sql, tuple(params))
    return {"results": rows, "count": len(rows)}


# ══════════════════════════════════════════════════════════════════════════════
# CHAT routes
# ══════════════════════════════════════════════════════════════════════════════

async def _auto_title_session(session_id: str, message: str) -> None:
    """Generate a short title for a session from the first message."""
    row = await db_fetchone("SELECT message_count FROM chat_sessions WHERE id=?", (session_id,))
    if not row or row["message_count"] > 0:
        return
    title = message[:50].strip()
    if len(message) > 50:
        title += "…"
    await db_execute("UPDATE chat_sessions SET title=?,updated_at=? WHERE id=?",
                     (title, datetime.utcnow().isoformat(), session_id))


@app.post("/chat")
async def chat(req: ChatRequest, user: Dict = Depends(get_current_user)):
    """Enhanced chat endpoint with full conversation context."""
    t0   = time.time()
    uid  = user["sub"]
    tier = await get_user_subscription(uid)
    if not rate_limiter.check(uid, tier):
        raise HTTPException(429, "Daily limit reached. Upgrade your plan.")
    mem_on       = await is_memory_enabled(uid)
    memories     = await get_user_memories(uid) if mem_on else []
    memory_block = format_memories_for_prompt(memories)
    has_image    = bool(req.image_url)

    if has_image:
        try:
            text, _ = await analyze_image_pipeline(req.image_url, req.message, req.model_type, uid, memories)
            new_mems = await extract_and_save_memories(uid, req.message, text)
            chat_id = str(uuid.uuid4())
            await db_insert(
                "INSERT INTO chat_history (id,user_id,session_id,message,response,query_type,agent_used,model_used,response_time_ms) VALUES (?,?,?,?,?,?,?,?,?)",
                (chat_id, uid, req.session_id, req.message, text, "image",
                 "vision_agent", req.model_type, int((time.time() - t0) * 1000)),
            )
            if req.session_id:
                await auto_detect_topics(req.session_id, uid, [req.message, text])
            return {"response": text, "response_time": time.time() - t0, "agent_used": "vision_agent",
                    "memory_updated": bool(new_mems), "new_memories": new_mems, "chat_id": chat_id}
        except Exception as exc:
            raise HTTPException(500, f"Image analysis failed: {exc}")

    if is_db_intent(req.message) and MYSQL_ENABLED and SUBSCRIPTION_LIMITS.get(tier, {}).get("mysql_access"):
        result = await run_nl_to_sql(req.message, uid)
        if result.get("success"):
            text = f"**SQL:**\n```sql\n{result['sql']}\n```\n\n**Results** ({result['row_count']} rows):\n" + \
                   json.dumps(result["results"][:20], indent=2, default=str)
            return {"response": text, "response_time": time.time() - t0, "agent_used": "data_agent"}

    agent   = route_message(req.message, has_image)
    context = get_relevant_context(uid, req.message) if req.use_rag else ""
    # ✨ NEW: Include full conversation history in prompt
    prompt  = await build_conversation_prompt(
        req.message, agent, context, memory_block, 
        req.session_id, uid, include_history=True
    )
    ai, model = get_model_client(req.model_type)
    completion   = ai.chat.completions.create(
        model=model, messages=[{"role": "user", "content": prompt}], max_tokens=1024, temperature=0.7
    )
    text         = completion.choices[0].message.content or ""
    new_memories = await extract_and_save_memories(uid, req.message, text)

    if req.session_id:
        await _auto_title_session(req.session_id, req.message)
        await db_execute("UPDATE chat_sessions SET message_count=message_count+1,updated_at=? WHERE id=?",
                         (datetime.utcnow().isoformat(), req.session_id))
        await auto_detect_topics(req.session_id, uid, [req.message, text])

    chat_id = str(uuid.uuid4())
    await db_insert(
        "INSERT INTO chat_history (id,user_id,session_id,message,response,query_type,agent_used,model_used,response_time_ms) VALUES (?,?,?,?,?,?,?,?,?)",
        (chat_id, uid, req.session_id, req.message, text,
         "rag" if context else "chat", agent.name, req.model_type, int((time.time() - t0) * 1000)),
    )
    return {"response": text, "response_time": time.time() - t0, "agent_used": agent.name,
            "memory_updated": bool(new_memories), "new_memories": new_memories, "chat_id": chat_id}


@app.post("/chat-stream")
async def chat_stream(req: ChatRequest, user: Dict = Depends(get_current_user)):
    t0   = time.time()
    uid  = user["sub"]
    tier = await get_user_subscription(uid)
    if not rate_limiter.check(uid, tier):
        raise HTTPException(429, "Daily limit reached.")
    mem_on       = await is_memory_enabled(uid)
    memories     = await get_user_memories(uid) if mem_on else []
    memory_block = format_memories_for_prompt(memories)
    has_image    = bool(req.image_url)
    _sse_headers = {"Cache-Control": "no-cache", "X-Accel-Buffering": "no"}

    if has_image:
        async def _image_stream():
            yield f"data: {json.dumps({'type': 'agent', 'agent': 'vision_agent'})}\n\n"
            yield f"data: {json.dumps({'type': 'status', 'message': '🔍 Analyzing image...'})}\n\n"
            try:
                text, _ = await analyze_image_pipeline(req.image_url, req.message, req.model_type, uid, memories)
                words = text.split()
                for i in range(0, len(words), 8):
                    chunk = " ".join(words[i:i + 8]) + (" " if i + 8 < len(words) else "")
                    yield f"data: {json.dumps({'type': 'chunk', 'content': chunk})}\n\n"
                    await asyncio.sleep(0.01)
                new_mems = await extract_and_save_memories(uid, req.message, text)
                await db_insert(
                    "INSERT INTO chat_history (id,user_id,session_id,message,response,query_type,agent_used,model_used,response_time_ms) VALUES (?,?,?,?,?,?,?,?,?)",
                    (str(uuid.uuid4()), uid, req.session_id, req.message, text, "image",
                     "vision_agent", req.model_type, int((time.time() - t0) * 1000)),
                )
                yield f"data: {json.dumps({'type': 'done', 'response_time': time.time()-t0, 'agent': 'vision_agent', 'full_response': text, 'memory_updated': bool(new_mems), 'new_memories': new_mems})}\n\n"
            except Exception as exc:
                yield f"data: {json.dumps({'type': 'error', 'message': str(exc)})}\n\n"
        return StreamingResponse(_image_stream(), media_type="text/event-stream", headers=_sse_headers)

    agent   = route_message(req.message, has_image)
    context = get_relevant_context(uid, req.message) if req.use_rag else ""
    prompt  = agent.build_prompt(req.message, context, memory_block)
    ai, model = get_model_client(req.model_type)
    try:
        stream = ai.chat.completions.create(
            model=model, messages=[{"role": "user", "content": prompt}],
            max_tokens=1024, temperature=0.7, stream=True,
        )
    except Exception as exc:
        raise HTTPException(503, f"AI service error: {exc}")

    async def _generate():
        full = ""
        try:
            yield f"data: {json.dumps({'type': 'agent', 'agent': agent.name})}\n\n"
            for chunk in stream:
                if chunk.choices and chunk.choices[0].delta.content:
                    content = chunk.choices[0].delta.content
                    full   += content
                    yield f"data: {json.dumps({'type': 'chunk', 'content': content})}\n\n"
            elapsed      = time.time() - t0
            new_memories = await extract_and_save_memories(uid, req.message, full)
            if req.session_id:
                await _auto_title_session(req.session_id, req.message)
                await db_execute("UPDATE chat_sessions SET message_count=message_count+1,updated_at=? WHERE id=?",
                                 (datetime.utcnow().isoformat(), req.session_id))
            await db_insert(
                "INSERT INTO chat_history (id,user_id,session_id,message,response,query_type,agent_used,model_used,response_time_ms) VALUES (?,?,?,?,?,?,?,?,?)",
                (str(uuid.uuid4()), uid, req.session_id, req.message, full,
                 "rag" if context else "chat", agent.name, req.model_type, int(elapsed * 1000)),
            )
            if new_memories:
                notif_msg = f"Saved {len(new_memories)} memory item(s) from your conversation."
                await push_notification(uid, "🧠 Memory Updated", notif_msg, "info")
            yield f"data: {json.dumps({'type': 'done', 'response_time': elapsed, 'agent': agent.name, 'full_response': full, 'memory_updated': bool(new_memories), 'new_memories': new_memories})}\n\n"
        except Exception as exc:
            yield f"data: {json.dumps({'type': 'error', 'message': str(exc)})}\n\n"

    return StreamingResponse(_generate(), media_type="text/event-stream", headers=_sse_headers)


@app.get("/chat/history")
async def chat_history(page: int = 1, per_page: int = 20,
                       user: Dict = Depends(get_current_user)):
    uid    = user["sub"]
    offset = (page - 1) * per_page
    rows   = await db_fetchall(
        "SELECT * FROM chat_history WHERE user_id=? AND session_id IS NULL ORDER BY created_at DESC LIMIT ? OFFSET ?",
        (uid, per_page, offset),
    )
    return {"history": rows, "page": page, "per_page": per_page}


@app.get("/chat/export")
async def export_chat(fmt: str = "markdown", user: Dict = Depends(get_current_user)):
    uid  = user["sub"]
    rows = await db_fetchall("SELECT * FROM chat_history WHERE user_id=? ORDER BY created_at ASC", (uid,))
    if fmt == "json":
        content    = json.dumps(rows, indent=2, default=str)
        media_type = "application/json"
        filename   = f"jazz_chat_{uid[:8]}.json"
    else:
        lines = [f"# JAZZ AI Chat Export\n_Exported: {datetime.utcnow().isoformat()}_\n"]
        for r in rows:
            lines.append(f"\n---\n**[{r.get('created_at','')}]** `{r.get('agent_used','')}`\n")
            lines.append(f"**You:** {r['message']}\n\n**JAZZ:** {r['response']}\n")
        content    = "\n".join(lines)
        media_type = "text/markdown"
        filename   = f"jazz_chat_{uid[:8]}.md"
    return StreamingResponse(
        iter([content.encode()]),
        media_type=media_type,
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


# ══════════════════════════════════════════════════════════════════════════════
# DOCUMENTS routes
# ══════════════════════════════════════════════════════════════════════════════

@app.post("/documents/upload")
async def upload_document(file: UploadFile = File(...), user: Dict = Depends(get_current_user)):
    if not documents_collection:
        raise HTTPException(503, "Vector DB unavailable")
    uid    = user["sub"]
    tier   = await get_user_subscription(uid)
    limits = SUBSCRIPTION_LIMITS.get(tier, SUBSCRIPTION_LIMITS["free"])
    if limits["documents"] != -1:
        count = await db_count("SELECT COUNT(*) FROM documents WHERE user_id=?", (uid,))
        if count >= limits["documents"]:
            raise HTTPException(403, f"Document limit reached for {tier} plan")
    content = await file.read()
    if len(content) > limits["max_file_mb"] * 1024 * 1024:
        raise HTTPException(413, "File too large")
    tmp = tempfile.mkdtemp()
    try:
        fpath = os.path.join(tmp, file.filename or "upload")
        with open(fpath, "wb") as f:
            f.write(content)
        mime   = file.content_type or "application/octet-stream"
        text   = process_document(fpath, mime)
        if not text.strip():
            raise HTTPException(400, "Could not extract text from document")
        chunks = chunk_text(text)
        doc_id = str(uuid.uuid4())
        documents_collection.add(
            documents=chunks,
            ids=[f"{doc_id}_{i}" for i in range(len(chunks))],
            metadatas=[{"user_id": uid, "filename": file.filename, "doc_id": doc_id} for _ in chunks],
        )
        now = datetime.utcnow().isoformat()
        await db_execute(
            "INSERT INTO documents (id,user_id,filename,original_name,file_type,file_size,chunk_count,chroma_collection,uploaded_at) VALUES (?,?,?,?,?,?,?,?,?)",
            (doc_id, uid, file.filename or "", file.filename or "", mime, len(content), len(chunks), "user_documents", now),
        )
        return DocumentResponse(id=doc_id, filename=file.filename or "", original_name=file.filename or "",
                                file_type=mime, file_size=len(content), chunk_count=len(chunks), uploaded_at=now)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


@app.get("/documents")
async def list_documents(user: Dict = Depends(get_current_user)):
    return {"documents": await db_fetchall(
        "SELECT * FROM documents WHERE user_id=? ORDER BY uploaded_at DESC", (user["sub"],)
    )}


@app.delete("/documents/{doc_id}")
async def delete_document(doc_id: str, user: Dict = Depends(get_current_user)):
    uid = user["sub"]
    if documents_collection:
        try:
            existing = documents_collection.get(where={"doc_id": doc_id, "user_id": uid})
            if existing["ids"]:
                documents_collection.delete(ids=existing["ids"])
        except Exception as exc:
            logger.warning("[CHROMA] %s", exc)
    await db_execute("DELETE FROM documents WHERE id=? AND user_id=?", (doc_id, uid))
    return {"success": True}


# ══════════════════════════════════════════════════════════════════════════════
# WEBSITE BUILDER routes
# ══════════════════════════════════════════════════════════════════════════════

@app.post("/websites/build")
async def build_website(req: WebsiteBuildRequest, user: Dict = Depends(get_current_user)):
    uid    = user["sub"]
    tier   = await get_user_subscription(uid)
    limits = SUBSCRIPTION_LIMITS.get(tier, SUBSCRIPTION_LIMITS["free"])
    wlimit = limits.get("websites", 3)
    if wlimit != -1:
        count = await db_count("SELECT COUNT(*) FROM websites WHERE user_id=?", (uid,))
        if count >= wlimit:
            raise HTTPException(403, f"Website limit ({wlimit}) reached for {tier} plan")
    title    = req.title or f"Site {datetime.utcnow().strftime('%Y%m%d_%H%M%S')}"
    site_id  = str(uuid.uuid4())
    filename = f"{site_id}.html"
    try:
        html = await _build_website_html(req.description, title, req.style or "modern", req.model_type)
    except Exception as exc:
        raise HTTPException(500, f"Website generation failed: {exc}")
    site_path = SITES_DIR / filename
    site_path.write_text(html, encoding="utf-8")
    now = datetime.utcnow().isoformat()
    await db_execute(
        "INSERT INTO websites (id,user_id,title,description,filename,html_size,prompt_used,created_at,updated_at) VALUES (?,?,?,?,?,?,?,?,?)",
        (site_id, uid, title, req.description[:500], filename, len(html), req.description, now, now),
    )
    return {"success": True, "id": site_id, "title": title, "filename": filename,
            "preview_url": f"/preview/{filename}", "html_size": len(html), "created_at": now}


@app.put("/websites/{site_id}")
async def update_website(site_id: str, req: WebsiteUpdateRequest, user: Dict = Depends(get_current_user)):
    uid = user["sub"]
    row = await db_fetchone("SELECT * FROM websites WHERE id=? AND user_id=?", (site_id, uid))
    if not row:
        raise HTTPException(404, "Website not found")
    site_path = SITES_DIR / row["filename"]
    if not site_path.exists():
        raise HTTPException(404, "Site file missing")
    current_html = site_path.read_text(encoding="utf-8")
    try:
        new_html = await _update_website_html(current_html, req.instructions, req.model_type)
    except Exception as exc:
        raise HTTPException(500, f"Update failed: {exc}")
    site_path.write_text(new_html, encoding="utf-8")
    now = datetime.utcnow().isoformat()
    await db_execute("UPDATE websites SET html_size=?,updated_at=? WHERE id=?", (len(new_html), now, site_id))
    return {"success": True, "id": site_id, "preview_url": f"/preview/{row['filename']}",
            "html_size": len(new_html), "updated_at": now}


@app.get("/websites")
async def list_websites(user: Dict = Depends(get_current_user)):
    rows = await db_fetchall(
        "SELECT id,title,description,filename,html_size,created_at,updated_at FROM websites WHERE user_id=? ORDER BY created_at DESC",
        (user["sub"],),
    )
    for r in rows:
        r["preview_url"] = f"/preview/{r['filename']}"
    return {"websites": rows}


@app.get("/websites/{site_id}")
async def get_website(site_id: str, user: Dict = Depends(get_current_user)):
    row = await db_fetchone("SELECT * FROM websites WHERE id=? AND user_id=?", (site_id, user["sub"]))
    if not row:
        raise HTTPException(404, "Website not found")
    row["preview_url"] = f"/preview/{row['filename']}"
    site_path = SITES_DIR / row["filename"]
    row["html"] = site_path.read_text(encoding="utf-8") if site_path.exists() else ""
    return row


@app.delete("/websites/{site_id}")
async def delete_website(site_id: str, user: Dict = Depends(get_current_user)):
    uid = user["sub"]
    row = await db_fetchone("SELECT filename FROM websites WHERE id=? AND user_id=?", (site_id, uid))
    if not row:
        raise HTTPException(404, "Website not found")
    sp = SITES_DIR / row["filename"]
    if sp.exists():
        sp.unlink()
    await db_execute("DELETE FROM websites WHERE id=?", (site_id,))
    return {"success": True}


# ══════════════════════════════════════════════════════════════════════════════
# CODE RUNNER routes
# ══════════════════════════════════════════════════════════════════════════════

@app.post("/code/run")
async def run_code(req: CodeRunRequest, user: Dict = Depends(get_current_user)):
    uid  = user["sub"]
    tier = await get_user_subscription(uid)
    if not rate_limiter.check(uid, tier, "code_runs_per_day"):
        raise HTTPException(429, "Daily code run limit reached")
    if req.language == "python":
        safe, reason = _is_code_safe(req.code, req.language)
        if not safe:
            raise HTTPException(400, f"Code blocked for safety: {reason}")
    loop   = asyncio.get_event_loop()
    result = await loop.run_in_executor(_executor, _run_code_sync, req.code, req.language)
    run_id = str(uuid.uuid4())
    await db_execute(
        "INSERT INTO code_runs (id,user_id,language,code,stdout,stderr,exit_code,duration_ms) VALUES (?,?,?,?,?,?,?,?)",
        (run_id, uid, req.language, req.code[:5000], result["stdout"], result["stderr"],
         result["exit_code"], result["duration_ms"]),
    )
    return {"run_id": run_id, "language": req.language, "stdout": result["stdout"],
            "stderr": result["stderr"], "exit_code": result["exit_code"],
            "duration_ms": result["duration_ms"], "success": result["exit_code"] == 0}


@app.get("/code/history")
async def code_history(limit: int = 20, user: Dict = Depends(get_current_user)):
    rows = await db_fetchall(
        "SELECT id,language,code,stdout,stderr,exit_code,duration_ms,created_at FROM code_runs WHERE user_id=? ORDER BY created_at DESC LIMIT ?",
        (user["sub"], limit),
    )
    return {"runs": rows}


# ══════════════════════════════════════════════════════════════════════════════
# SQLITE QUERY routes (admin only)
# ══════════════════════════════════════════════════════════════════════════════

@app.post("/sqlite/query")
async def sqlite_nl_query(req: SQLiteQueryRequest, user: Dict = Depends(require_admin)):
    return await run_sqlite_nl_query(req.question)


@app.post("/sqlite/raw")
async def sqlite_raw_query(req: SQLiteRawRequest, user: Dict = Depends(require_admin)):
    sql_upper = req.sql.upper().strip()
    if req.readonly and not sql_upper.startswith("SELECT") and not sql_upper.startswith("PRAGMA"):
        raise HTTPException(400, "Only SELECT/PRAGMA allowed in readonly mode")
    return await _run_sqlite_raw(req.sql)


@app.get("/sqlite/tables")
async def sqlite_tables(user: Dict = Depends(require_admin)):
    import sqlite3 as _s3
    try:
        with _s3.connect(DB_PATH) as conn:
            tables = [r[0] for r in conn.execute(
                "SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%'"
            ).fetchall()]
        result = []
        for tbl in tables:
            count = await db_count(f"SELECT COUNT(*) FROM `{tbl}`")
            result.append({"table": tbl, "rows": count})
        return {"tables": result}
    except Exception as exc:
        raise HTTPException(500, str(exc))


# ══════════════════════════════════════════════════════════════════════════════
# PROMPT LIBRARY routes
# ══════════════════════════════════════════════════════════════════════════════

@app.post("/prompts")
async def create_prompt(req: PromptCreate, user: Dict = Depends(get_current_user)):
    pid = str(uuid.uuid4())
    now = datetime.utcnow().isoformat()
    await db_execute(
        "INSERT INTO prompt_templates (id,user_id,title,content,category,is_public,created_at) VALUES (?,?,?,?,?,?,?)",
        (pid, user["sub"], req.title, req.content, req.category or "general", 1 if req.is_public else 0, now),
    )
    return {"success": True, "prompt": await db_fetchone("SELECT * FROM prompt_templates WHERE id=?", (pid,))}


@app.get("/prompts")
async def list_prompts(user: Dict = Depends(get_current_user)):
    uid  = user["sub"]
    rows = await db_fetchall(
        "SELECT p.*,u.full_name as author FROM prompt_templates p LEFT JOIN users u ON p.user_id=u.id WHERE p.user_id=? OR p.is_public=1 ORDER BY p.use_count DESC, p.created_at DESC",
        (uid,),
    )
    return {"prompts": rows}


@app.put("/prompts/{prompt_id}")
async def update_prompt(prompt_id: str, req: PromptUpdate, user: Dict = Depends(get_current_user)):
    row = await db_fetchone("SELECT id FROM prompt_templates WHERE id=? AND user_id=?",
                            (prompt_id, user["sub"]))
    if not row:
        raise HTTPException(404, "Prompt not found")
    updates: Dict[str, Any] = {}
    if req.title     is not None: updates["title"]     = req.title
    if req.content   is not None: updates["content"]   = req.content
    if req.category  is not None: updates["category"]  = req.category
    if req.is_public is not None: updates["is_public"] = 1 if req.is_public else 0
    if not updates:
        raise HTTPException(400, "Nothing to update")
    sets = ", ".join(f"{k}=?" for k in updates)
    await db_execute(f"UPDATE prompt_templates SET {sets} WHERE id=?",
                     tuple(updates.values()) + (prompt_id,))
    return {"success": True, "prompt": await db_fetchone("SELECT * FROM prompt_templates WHERE id=?", (prompt_id,))}


@app.delete("/prompts/{prompt_id}")
async def delete_prompt(prompt_id: str, user: Dict = Depends(get_current_user)):
    await db_execute("DELETE FROM prompt_templates WHERE id=? AND user_id=?",
                     (prompt_id, user["sub"]))
    return {"success": True}


@app.post("/prompts/{prompt_id}/use")
async def use_prompt(prompt_id: str, user: Dict = Depends(get_current_user)):
    row = await db_fetchone("SELECT * FROM prompt_templates WHERE id=?", (prompt_id,))
    if not row:
        raise HTTPException(404, "Prompt not found")
    await db_execute("UPDATE prompt_templates SET use_count=use_count+1 WHERE id=?", (prompt_id,))
    return {"content": row["content"], "title": row["title"]}


# ══════════════════════════════════════════════════════════════════════════════
# AGENT JOB routes
# ══════════════════════════════════════════════════════════════════════════════

@app.post("/agent-jobs")
async def create_agent_job(req: AgentJobCreate, user: Dict = Depends(get_current_user)):
    uid       = user["sub"]
    tier      = await get_user_subscription(uid)
    limits    = SUBSCRIPTION_LIMITS.get(tier, SUBSCRIPTION_LIMITS["free"])
    job_limit = limits.get("agent_jobs", 1)
    if job_limit != -1:
        count = await db_count("SELECT COUNT(*) FROM agent_jobs WHERE user_id=?", (uid,))
        if count >= job_limit:
            raise HTTPException(403, f"Agent job limit ({job_limit}) reached for {tier} plan")
    if len(req.cron_schedule.strip().split()) != 5:
        raise HTTPException(400, "cron_schedule must be 5 fields e.g. '0 11 * * *'")
    job_id    = str(uuid.uuid4())
    now_iso   = datetime.utcnow().isoformat()
    creds_enc = encrypt_creds(req.credentials)
    await db_execute(
        "INSERT INTO agent_jobs (id,user_id,name,description,job_type,cron_schedule,time_window_start,time_window_end,credentials_enc,parameters,status,enabled,created_at,updated_at) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
        (job_id, uid, req.name, req.description or "", req.job_type, req.cron_schedule,
         req.time_window_start or "11:00", req.time_window_end or "12:00",
         creds_enc, json.dumps(req.parameters or {}), "idle", 1, now_iso, now_iso),
    )
    job = await db_fetchone("SELECT * FROM agent_jobs WHERE id=?", (job_id,))
    job.pop("credentials_enc", None)
    _schedule_job({**job, "credentials_enc": creds_enc})
    return {"success": True, "job": job}


@app.get("/agent-jobs")
async def list_agent_jobs(user: Dict = Depends(get_current_user)):
    rows = await db_fetchall(
        "SELECT id,user_id,name,description,job_type,cron_schedule,time_window_start,time_window_end,parameters,status,enabled,last_run_at,last_run_status,next_run_at,total_runs,total_applied,created_at,updated_at FROM agent_jobs WHERE user_id=? ORDER BY created_at DESC",
        (user["sub"],),
    )
    return {"jobs": rows}


@app.get("/agent-jobs/{job_id}")
async def get_agent_job(job_id: str, user: Dict = Depends(get_current_user)):
    row = await db_fetchone(
        "SELECT id,user_id,name,description,job_type,cron_schedule,time_window_start,time_window_end,parameters,status,enabled,last_run_at,last_run_status,next_run_at,total_runs,total_applied,created_at,updated_at FROM agent_jobs WHERE id=? AND user_id=?",
        (job_id, user["sub"]),
    )
    if not row:
        raise HTTPException(404, "Job not found")
    return row


@app.put("/agent-jobs/{job_id}")
async def update_agent_job(job_id: str, req: AgentJobUpdate, user: Dict = Depends(get_current_user)):
    existing = await db_fetchone("SELECT * FROM agent_jobs WHERE id=? AND user_id=?",
                                 (job_id, user["sub"]))
    if not existing:
        raise HTTPException(404, "Job not found")
    updates: Dict[str, Any] = {}
    if req.name              is not None: updates["name"]              = req.name
    if req.description       is not None: updates["description"]       = req.description
    if req.cron_schedule     is not None:
        if len(req.cron_schedule.split()) != 5:
            raise HTTPException(400, "cron_schedule must be 5 fields")
        updates["cron_schedule"] = req.cron_schedule
    if req.time_window_start is not None: updates["time_window_start"] = req.time_window_start
    if req.time_window_end   is not None: updates["time_window_end"]   = req.time_window_end
    if req.parameters        is not None: updates["parameters"]        = json.dumps(req.parameters)
    if req.enabled           is not None: updates["enabled"]           = 1 if req.enabled else 0
    if req.credentials       is not None: updates["credentials_enc"]   = encrypt_creds(req.credentials)
    if not updates:
        raise HTTPException(400, "Nothing to update")
    updates["updated_at"] = datetime.utcnow().isoformat()
    sets = ", ".join(f"{k}=?" for k in updates)
    await db_execute(f"UPDATE agent_jobs SET {sets} WHERE id=?", tuple(updates.values()) + (job_id,))
    updated = await db_fetchone("SELECT * FROM agent_jobs WHERE id=?", (job_id,))
    if updated.get("enabled"):
        _schedule_job(updated)
    else:
        _unschedule_job(job_id)
    updated.pop("credentials_enc", None)
    return {"success": True, "job": updated}


@app.delete("/agent-jobs/{job_id}")
async def delete_agent_job(job_id: str, user: Dict = Depends(get_current_user)):
    existing = await db_fetchone("SELECT id FROM agent_jobs WHERE id=? AND user_id=?",
                                 (job_id, user["sub"]))
    if not existing:
        raise HTTPException(404, "Job not found")
    _unschedule_job(job_id)
    await db_execute("DELETE FROM agent_jobs WHERE id=?", (job_id,))
    return {"success": True}


@app.post("/agent-jobs/{job_id}/run")
async def run_agent_job_now(job_id: str, user: Dict = Depends(get_current_user)):
    existing = await db_fetchone("SELECT id FROM agent_jobs WHERE id=? AND user_id=?",
                                 (job_id, user["sub"]))
    if not existing:
        raise HTTPException(404, "Job not found")
    task = asyncio.create_task(execute_job(job_id, bypass_time_window=True))
    def _on_done(t: asyncio.Task) -> None:
        if t.cancelled():
            logger.error("[JOB] Task CANCELLED for %s", job_id)
        elif t.exception():
            logger.error("[JOB] Task CRASHED for %s: %s", job_id, t.exception())
    task.add_done_callback(_on_done)
    return {"success": True, "message": "Job triggered — check logs for progress"}


@app.post("/agent-jobs/{job_id}/run-debug")
async def run_agent_job_debug(job_id: str, user: Dict = Depends(get_current_user)):
    existing = await db_fetchone("SELECT id FROM agent_jobs WHERE id=? AND user_id=?",
                                 (job_id, user["sub"]))
    if not existing:
        raise HTTPException(404, "Job not found")
    try:
        loop = asyncio.get_event_loop()
        await loop.run_in_executor(_executor, _execute_job_sync, job_id, True)
        logs   = await db_fetchall(
            "SELECT id,status,output,applied_jobs,error,started_at,completed_at,duration_sec FROM agent_job_logs WHERE job_id=? ORDER BY started_at DESC LIMIT 1",
            (job_id,),
        )
        latest = logs[0] if logs else None
        if latest:
            try: latest["applied_jobs"] = json.loads(latest.get("applied_jobs") or "[]")
            except Exception: latest["applied_jobs"] = []
        return {"success": True, "playwright_available": _PLAYWRIGHT_AVAILABLE, "latest_log": latest}
    except Exception as exc:
        return {"success": False, "error": str(exc), "traceback": traceback.format_exc(),
                "playwright_available": _PLAYWRIGHT_AVAILABLE}


@app.post("/agent-jobs/{job_id}/toggle")
async def toggle_agent_job(job_id: str, user: Dict = Depends(get_current_user)):
    row = await db_fetchone("SELECT * FROM agent_jobs WHERE id=? AND user_id=?",
                            (job_id, user["sub"]))
    if not row:
        raise HTTPException(404, "Job not found")
    new_enabled = 0 if row["enabled"] else 1
    await db_execute("UPDATE agent_jobs SET enabled=?,updated_at=? WHERE id=?",
                     (new_enabled, datetime.utcnow().isoformat(), job_id))
    if new_enabled:
        _schedule_job({**row, "enabled": new_enabled})
    else:
        _unschedule_job(job_id)
    return {"success": True, "enabled": bool(new_enabled)}


@app.get("/agent-jobs/{job_id}/logs")
async def get_job_logs(job_id: str, limit: int = 20, user: Dict = Depends(get_current_user)):
    existing = await db_fetchone("SELECT id FROM agent_jobs WHERE id=? AND user_id=?",
                                 (job_id, user["sub"]))
    if not existing:
        raise HTTPException(404, "Job not found")
    rows = await db_fetchall(
        "SELECT id,job_id,status,output,applied_jobs,error,started_at,completed_at,duration_sec FROM agent_job_logs WHERE job_id=? ORDER BY started_at DESC LIMIT ?",
        (job_id, limit),
    )
    for r in rows:
        try: r["applied_jobs"] = json.loads(r.get("applied_jobs") or "[]")
        except Exception: r["applied_jobs"] = []
    return {"logs": rows}


# ══════════════════════════════════════════════════════════════════════════════
# FILE WORKSPACE routes (auth required)
# ══════════════════════════════════════════════════════════════════════════════

@app.post("/files")
async def file_operation(req: FileOperationRequest, user: Dict = Depends(get_current_user)):
    ops = {
        "create": lambda: file_mgr.create(req.filename, req.content or ""),
        "read":   lambda: file_mgr.read(req.filename),
        "update": lambda: file_mgr.update(req.filename, req.content or ""),
        "delete": lambda: file_mgr.delete(req.filename),
        "list":   lambda: {"success": True, "files": file_mgr.list_files()},
    }
    fn = ops.get(req.operation)
    if not fn:
        raise HTTPException(400, f"Unknown operation: {req.operation}")
    return fn()


# ══════════════════════════════════════════════════════════════════════════════
# ANNOUNCEMENTS routes
# ══════════════════════════════════════════════════════════════════════════════

@app.get("/announcements")
async def get_announcements(user: Dict = Depends(get_current_user)):
    """Get active announcements (authenticated users)."""
    now = datetime.utcnow().isoformat()
    rows = await db_fetchall(
        "SELECT * FROM announcements WHERE is_active=1 AND (expires_at IS NULL OR expires_at > ?) ORDER BY created_at DESC LIMIT 10",
        (now,),
    )
    return {"announcements": rows}


@app.post("/admin/announcements")
async def create_announcement(req: AnnouncementCreate, user: Dict = Depends(require_admin)):
    aid = str(uuid.uuid4())
    now = datetime.utcnow().isoformat()
    await db_execute(
        "INSERT INTO announcements (id,title,message,type,is_active,created_by,expires_at,created_at) VALUES (?,?,?,?,?,?,?,?)",
        (aid, req.title, req.message, req.type, 1, user["sub"], req.expires_at, now),
    )
    return {"success": True, "announcement": await db_fetchone("SELECT * FROM announcements WHERE id=?", (aid,))}


@app.get("/admin/announcements")
async def admin_list_announcements(user: Dict = Depends(require_admin)):
    rows = await db_fetchall("SELECT * FROM announcements ORDER BY created_at DESC")
    return {"announcements": rows}


@app.delete("/admin/announcements/{aid}")
async def delete_announcement(aid: str, user: Dict = Depends(require_admin)):
    await db_execute("DELETE FROM announcements WHERE id=?", (aid,))
    return {"success": True}


@app.post("/admin/broadcast")
async def broadcast_notification(req: BroadcastCreate, user: Dict = Depends(require_admin)):
    """Send a notification to ALL users."""
    users = await db_fetchall("SELECT id FROM users")
    now   = datetime.utcnow().isoformat()
    for u in users:
        nid = str(uuid.uuid4())
        await db_execute(
            "INSERT INTO notifications (id,user_id,title,message,type,created_at) VALUES (?,?,?,?,?,?)",
            (nid, u["id"], req.title, req.message, req.type, now),
        )
    return {"success": True, "sent_to": len(users)}


# ══════════════════════════════════════════════════════════════════════════════
# COUPONS routes
# ══════════════════════════════════════════════════════════════════════════════

@app.get("/admin/coupons")
async def admin_list_coupons(user: Dict = Depends(require_admin)):
    rows = await db_fetchall("SELECT * FROM coupons ORDER BY created_at DESC")
    return {"coupons": rows}


@app.post("/admin/coupons")
async def admin_create_coupon(req: CouponCreate, user: Dict = Depends(require_admin)):
    cid = str(uuid.uuid4())
    now = datetime.utcnow().isoformat()
    try:
        await db_execute(
            "INSERT INTO coupons (id,code,description,subscription_tier,max_uses,is_active,expires_at,created_at) VALUES (?,?,?,?,?,?,?,?)",
            (cid, req.code.upper().strip(), req.description, req.subscription_tier,
             req.max_uses, 1, req.expires_at, now),
        )
    except Exception:
        raise HTTPException(400, "Coupon code already exists")
    return {"success": True, "coupon": await db_fetchone("SELECT * FROM coupons WHERE id=?", (cid,))}


@app.delete("/admin/coupons/{cid}")
async def admin_delete_coupon(cid: str, user: Dict = Depends(require_admin)):
    await db_execute("DELETE FROM coupons WHERE id=?", (cid,))
    return {"success": True}


@app.post("/coupons/redeem")
async def redeem_coupon(req: CouponRedeem, user: Dict = Depends(get_current_user)):
    uid  = user["sub"]
    code = req.code.upper().strip()
    now  = datetime.utcnow().isoformat()

    coupon = await db_fetchone("SELECT * FROM coupons WHERE code=? COLLATE NOCASE AND is_active=1", (code,))
    if not coupon:
        raise HTTPException(404, "Invalid or inactive coupon code")
    if coupon.get("expires_at") and coupon["expires_at"] < now:
        raise HTTPException(400, "Coupon has expired")
    if coupon["max_uses"] != -1 and coupon["uses"] >= coupon["max_uses"]:
        raise HTTPException(400, "Coupon usage limit reached")

    # Check already redeemed
    already = await db_fetchone(
        "SELECT id FROM coupon_redemptions WHERE coupon_id=? AND user_id=?",
        (coupon["id"], uid),
    )
    if already:
        raise HTTPException(400, "You have already used this coupon")

    # Apply
    new_tier = coupon["subscription_tier"]
    await db_execute("UPDATE users SET subscription=?,updated_at=? WHERE id=?", (new_tier, now, uid))
    await db_execute("UPDATE coupons SET uses=uses+1 WHERE id=?", (coupon["id"],))
    await db_insert(
        "INSERT INTO coupon_redemptions (id,coupon_id,user_id,redeemed_at) VALUES (?,?,?,?)",
        (str(uuid.uuid4()), coupon["id"], uid, now),
    )
    await push_notification(uid, "🎉 Subscription Upgraded!",
                            f"Coupon applied. Your plan is now: {new_tier.upper()}",
                            "success")
    return {"success": True, "new_subscription": new_tier,
            "message": f"Coupon applied! Subscription upgraded to {new_tier}."}


# ══════════════════════════════════════════════════════════════════════════════
# ADMIN routes
# ══════════════════════════════════════════════════════════════════════════════

@app.get("/admin/stats")
async def admin_stats(user: Dict = Depends(require_admin)):
    today = datetime.utcnow().strftime("%Y-%m-%d")
    return {
        "total_users":       await db_count("SELECT COUNT(*) FROM users"),
        "new_users_today":   await db_count("SELECT COUNT(*) FROM users WHERE created_at >= ?", (today,)),
        "total_documents":   await db_count("SELECT COUNT(*) FROM documents"),
        "chats_today":       await db_count("SELECT COUNT(*) FROM chat_history WHERE created_at >= ?", (today,)),
        "total_memories":    await db_count("SELECT COUNT(*) FROM user_memories"),
        "total_agent_jobs":  await db_count("SELECT COUNT(*) FROM agent_jobs"),
        "agent_runs_today":  await db_count("SELECT COUNT(*) FROM agent_job_logs WHERE started_at >= ?", (today,)),
        "total_websites":    await db_count("SELECT COUNT(*) FROM websites"),
        "code_runs_today":   await db_count("SELECT COUNT(*) FROM code_runs WHERE created_at >= ?", (today,)),
        "prompt_templates":  await db_count("SELECT COUNT(*) FROM prompt_templates"),
        "active_coupons":    await db_count("SELECT COUNT(*) FROM coupons WHERE is_active=1"),
        "total_api_keys":    await db_count("SELECT COUNT(*) FROM api_keys WHERE is_active=1"),
        "unread_notifications": await db_count("SELECT COUNT(*) FROM notifications WHERE is_read=0"),
        "mysql_enabled":     MYSQL_ENABLED,
        "playwright_available": _PLAYWRIGHT_AVAILABLE,
        "system_status":     "healthy",
        "version":           "7.0.0",
        "uptime_since":      _SERVER_START.isoformat(),
    }


@app.get("/admin/users")
async def admin_users(user: Dict = Depends(require_admin)):
    rows = await db_fetchall(
        "SELECT id,email,full_name,role,subscription,memory_enabled,created_at FROM users ORDER BY created_at DESC"
    )
    for r in rows:
        r["memory_enabled"] = bool(r.get("memory_enabled", 1))
    return {"users": rows}


@app.get("/admin/users/{uid}/overview")
async def admin_user_overview(uid: str, user: Dict = Depends(require_admin)):
    profile = await db_fetchone(
        "SELECT id,email,full_name,role,subscription,memory_enabled,created_at FROM users WHERE id=?", (uid,)
    )
    if not profile:
        raise HTTPException(404, "User not found")
    return {
        "profile":    profile,
        "chats":      await db_count("SELECT COUNT(*) FROM chat_history WHERE user_id=?", (uid,)),
        "documents":  await db_count("SELECT COUNT(*) FROM documents WHERE user_id=?", (uid,)),
        "memories":   await db_count("SELECT COUNT(*) FROM user_memories WHERE user_id=?", (uid,)),
        "agent_jobs": await db_count("SELECT COUNT(*) FROM agent_jobs WHERE user_id=?", (uid,)),
        "websites":   await db_count("SELECT COUNT(*) FROM websites WHERE user_id=?", (uid,)),
        "api_keys":   await db_count("SELECT COUNT(*) FROM api_keys WHERE user_id=? AND is_active=1", (uid,)),
    }


@app.get("/admin/users/{uid}/chats")
async def admin_user_chats(uid: str, limit: int = 50, user: Dict = Depends(require_admin)):
    rows = await db_fetchall(
        "SELECT id,session_id,message,response,agent_used,model_used,created_at FROM chat_history WHERE user_id=? ORDER BY created_at DESC LIMIT ?",
        (uid, limit),
    )
    return {"chats": rows}


@app.get("/admin/users/{uid}/memories")
async def admin_user_memories(uid: str, user: Dict = Depends(require_admin)):
    rows = await db_fetchall("SELECT * FROM user_memories WHERE user_id=? ORDER BY created_at DESC", (uid,))
    return {"memories": rows}


@app.get("/admin/users/{uid}/documents")
async def admin_user_documents(uid: str, user: Dict = Depends(require_admin)):
    rows = await db_fetchall("SELECT * FROM documents WHERE user_id=? ORDER BY uploaded_at DESC", (uid,))
    return {"documents": rows}


@app.get("/admin/users/{uid}/jobs")
async def admin_user_jobs(uid: str, user: Dict = Depends(require_admin)):
    rows = await db_fetchall(
        "SELECT id,name,job_type,cron_schedule,status,enabled,total_runs,total_applied,last_run_at,last_run_status FROM agent_jobs WHERE user_id=? ORDER BY created_at DESC",
        (uid,),
    )
    return {"jobs": rows}


@app.get("/admin/users/{uid}/websites")
async def admin_user_websites(uid: str, user: Dict = Depends(require_admin)):
    rows = await db_fetchall("SELECT id,title,description,html_size,created_at FROM websites WHERE user_id=? ORDER BY created_at DESC", (uid,))
    for r in rows:
        r["preview_url"] = f"/preview/{r.get('filename','')}"
    return {"websites": rows}


@app.post("/admin/users/{uid}/notify")
async def admin_notify_user(uid: str, req: NotificationCreate, user: Dict = Depends(require_admin)):
    target = await db_fetchone("SELECT id FROM users WHERE id=?", (uid,))
    if not target:
        raise HTTPException(404, "User not found")
    await push_notification(uid, req.title, req.message, req.type, req.link)
    return {"success": True}


@app.post("/admin/users/{uid}/impersonate")
async def admin_impersonate(uid: str, user: Dict = Depends(require_admin)):
    """Generate a short-lived token to impersonate a user (admin auditing)."""
    target = await db_fetchone("SELECT id,email,role,subscription FROM users WHERE id=?", (uid,))
    if not target:
        raise HTTPException(404, "User not found")
    token = create_access_token(
        {"sub": uid, "email": target["email"], "role": target["role"],
         "impersonated_by": user["sub"]},
        expires_delta=timedelta(hours=1),
    )
    return {"access_token": token, "token_type": "bearer",
            "user": {"id": uid, "email": target["email"], "role": target["role"],
                     "subscription": target["subscription"]},
            "warning": "This is a 1-hour impersonation token. Use responsibly."}


@app.put("/admin/subscription")
async def admin_update_subscription(req: SubscriptionUpdate, user: Dict = Depends(require_admin)):
    await db_execute("UPDATE users SET subscription=?,updated_at=? WHERE id=?",
                     (req.tier, datetime.utcnow().isoformat(), req.user_id))
    target = await db_fetchone("SELECT id FROM users WHERE id=?", (req.user_id,))
    if target:
        await push_notification(req.user_id, "Subscription Updated",
                                f"Your plan has been updated to {req.tier.upper()} by an admin.", "success")
    return {"success": True, "user_id": req.user_id, "subscription": req.tier}


@app.put("/admin/user-role")
async def admin_update_role(user_id: str, role: str, user: Dict = Depends(require_admin)):
    if role not in ("admin", "client"):
        raise HTTPException(400, "Role must be 'admin' or 'client'")
    await db_execute("UPDATE users SET role=?,updated_at=? WHERE id=?",
                     (role, datetime.utcnow().isoformat(), user_id))
    return {"success": True}


@app.delete("/admin/user/{user_id}")
async def admin_delete_user(user_id: str, user: Dict = Depends(require_admin)):
    if user_id == user["sub"]:
        raise HTTPException(400, "Cannot delete your own account via admin panel")
    sites = await db_fetchall("SELECT filename FROM websites WHERE user_id=?", (user_id,))
    for site in sites:
        sp = SITES_DIR / site["filename"]
        if sp.exists():
            sp.unlink()
    if documents_collection:
        try:
            existing = documents_collection.get(where={"user_id": user_id})
            if existing["ids"]:
                documents_collection.delete(ids=existing["ids"])
        except Exception:
            pass
    await db_execute("DELETE FROM users WHERE id=?", (user_id,))
    return {"success": True}


@app.get("/admin/agent-jobs")
async def admin_all_jobs(user: Dict = Depends(require_admin)):
    rows = await db_fetchall(
        "SELECT j.id,j.name,j.job_type,j.cron_schedule,j.status,j.enabled,j.last_run_at,j.last_run_status,j.total_runs,j.total_applied,u.email FROM agent_jobs j LEFT JOIN users u ON j.user_id=u.id ORDER BY j.created_at DESC"
    )
    return {"jobs": rows}


@app.get("/admin/all-chats")
async def admin_all_chats(limit: int = 100, user: Dict = Depends(require_admin)):
    rows = await db_fetchall(
        "SELECT h.*,u.email FROM chat_history h LEFT JOIN users u ON h.user_id=u.id ORDER BY h.created_at DESC LIMIT ?",
        (limit,),
    )
    return {"chats": rows}


@app.get("/logs")
async def get_logs(user: Dict = Depends(require_admin)):
    try:
        return {"logs": Path("jazz.log").read_text().splitlines()[-300:]}
    except FileNotFoundError:
        return {"logs": []}


# ══════════════════════════════════════════════════════════════════════════════
# HEALTH / INFO routes
# ══════════════════════════════════════════════════════════════════════════════

@app.get("/health")
async def health():
    db_ok = False
    try:
        await db_fetchone("SELECT 1")
        db_ok = True
    except Exception:
        pass
    sites_size_mb = sum(
        f.stat().st_size for f in SITES_DIR.rglob("*") if f.is_file()
    ) / (1024 * 1024)
    return {
        "status":         "healthy" if db_ok else "degraded",
        "timestamp":      datetime.utcnow().isoformat(),
        "uptime_seconds": int((datetime.utcnow() - _SERVER_START).total_seconds()),
        "services":       {
            "sqlite":    db_ok,
            "chromadb":  chroma_client is not None,
            "mysql":     _mysql_pool is not None,
            "playwright": _PLAYWRIGHT_AVAILABLE,
        },
        "storage":        {"sites_mb": round(sites_size_mb, 2), "db_path": DB_PATH},
        "agents":         [a.name for a in AGENTS],
        "version":        "7.0.0",
        "features":       ["memory", "vision", "rag", "streaming", "agent_jobs",
                           "website_builder", "code_runner", "sqlite_query",
                           "prompt_library", "chat_export", "chat_sessions",
                           "api_keys", "notifications", "announcements", "coupons"],
    }


@app.get("/model-info")
async def model_info():
    return {
        "models": [
            {"id": "censored",   "name": "Groq Llama 3.3 70B",        "speed": "ultra-fast", "filtered": True},
            {"id": "uncensored", "name": "Dolphin Mistral 24B Venice", "speed": "variable",   "filtered": False},
        ],
        "vision_model":   "Qwen2.5-VL-7B-Instruct",
        "website_styles": list(_STYLE_HINTS.keys()),
        "code_languages": ["python", "javascript", "bash"],
    }


# ══════════════════════════════════════════════════════════════════════════════
# Entry point
# ══════════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    import uvicorn
    _module_name = Path(__file__).stem
    uvicorn.run(
        f"{_module_name}:app",
        host="0.0.0.0",
        port=int(os.getenv("PORT", "8000")),
        reload=False,
        workers=1,
        log_level="info",
        access_log=True,
        timeout_keep_alive=75,
    )

