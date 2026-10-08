"""
Database and data access layer for Pharmacy & Chemistry Lab Automation Dashboard
SUPABASE (PostgreSQL) EDITION - permanent cloud storage.

Same functions, same return shapes as the SQLite edition, so app.py keeps working.
  * All data lives in Supabase Postgres, never on the app server's disk, so app
    sleep / reboot / redeploy can not erase stock, ledger or logs.
  * Stock deductions run in ONE transaction with row locks (safe for many users).
  * Snapshots (backup / restore) are stored in the database and can be downloaded.
  * Attachments are stored in the database (size limit per file, see below).

Connection: set DATABASE_URL (Streamlit: Settings -> Secrets, or an environment
variable). Use the Supabase "Session pooler" connection string.
"""

import os
import re
import io
import json
import gzip
import time
import uuid
import shutil
import tempfile
import threading
import functools
from datetime import datetime, date
from urllib.parse import quote, unquote
from zoneinfo import ZoneInfo

import pandas as pd
import psycopg2
import psycopg2.extras
import psycopg2.extensions
import psycopg2.pool

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Seed files travel with the code (only used to fill an EMPTY database)
EXCEL_PATH = os.path.join(BASE_DIR, "Total chemicals and lab ware list.xlsx")
PDF_PATH = os.path.join(BASE_DIR, "Lab_Timetable_135B_138B_FA26.pdf")

# Time zone used for all timestamps written by the database (audit trail)
_TZ_NAME = os.environ.get("PHARMALAB_TZ", "Asia/Karachi")
try:
    TZ = ZoneInfo(_TZ_NAME)
except Exception:  # tzdata missing -> fall back to UTC rather than crash
    _TZ_NAME, TZ = "UTC", ZoneInfo("UTC")

# SQL expression that produces 'YYYY-MM-DD HH:MM:SS' in local time (stored as text,
# exactly like the previous SQLite edition, so sorting and display do not change)
_TS_SQL = f"to_char(timezone('{_TZ_NAME}', now()), 'YYYY-MM-DD HH24:MI:SS')"

# Temporary folder used ONLY for generated Excel exports (downloaded right away)
TMP_DIR = os.path.join(tempfile.gettempdir(), "pharmalab_exports")
os.makedirs(TMP_DIR, exist_ok=True)

MAX_ATTACHMENT_MB = 20
MAX_ATTACHMENT_BYTES = MAX_ATTACHMENT_MB * 1024 * 1024
MAX_SNAPSHOTS_KEPT = 60
DB_PLAN_LIMIT_GB = float(os.environ.get("PHARMALAB_DB_LIMIT_GB", "0.5"))  # Supabase Free = 500 MB

_LOCK_KEY = 727001  # advisory lock id used while creating tables / seeding


def _now():
    return datetime.now(TZ)


# numpy numbers (coming from pandas) are not understood by psycopg2 by default
try:
    import numpy as _np
    for _t in (_np.int8, _np.int16, _np.int32, _np.int64, _np.uint8, _np.uint16, _np.uint32, _np.uint64):
        psycopg2.extensions.register_adapter(_t, lambda v: psycopg2.extensions.AsIs(int(v)))
    for _t in (_np.float32, _np.float64):
        psycopg2.extensions.register_adapter(_t, lambda v: psycopg2.extensions.AsIs(repr(float(v))))
    psycopg2.extensions.register_adapter(_np.bool_, lambda v: psycopg2.extensions.AsIs("TRUE" if bool(v) else "FALSE"))
except Exception:
    pass


# ==============================================================================
# CONNECTION HANDLING (pooled, so a click does not open a new internet connection)
# ==============================================================================

def _setting(name):
    """Reads a setting from an environment variable or from Streamlit secrets."""
    val = os.environ.get(name, "").strip()
    if val:
        return val
    try:
        import streamlit as st
        if name in st.secrets:
            return str(st.secrets[name]).strip()
    except Exception:
        pass
    return ""


def _get_database_url():
    """Finds the database address. Three ways to provide it (first one found wins):
      A) DATABASE_URL            - the whole Supabase 'Session pooler' string
      B) [connections.db] url    - same string, Streamlit connection style
      C) DB_HOST, DB_USER, DB_PASSWORD (+ optional DB_PORT, DB_NAME) - separate fields,
         no URL-encoding worries even if the password has symbols like @ # / :
    Each can be a Streamlit secret or a normal environment variable."""
    url = _setting("DATABASE_URL")
    if not url:
        try:
            import streamlit as st
            if "connections" in st.secrets and "db" in st.secrets["connections"]:
                url = str(st.secrets["connections"]["db"]["url"]).strip()
        except Exception:
            pass
    if not url:
        host, user, pwd = _setting("DB_HOST"), _setting("DB_USER"), _setting("DB_PASSWORD")
        if host and user and pwd:
            port = _setting("DB_PORT") or "5432"
            name = _setting("DB_NAME") or "postgres"
            host_part = quote(host, safe="") if host.startswith("/") else host
            url = f"postgresql://{quote(user, safe='')}:{quote(pwd, safe='')}@{host_part}:{port}/{quote(name, safe='')}"
    if not url:
        raise RuntimeError(
            "No database address found. Add DATABASE_URL (the Supabase 'Session pooler' string) "
            "in Streamlit -> Settings -> Secrets, OR add DB_HOST, DB_USER and DB_PASSWORD there."
        )
    if "YOUR-PASSWORD" in url or "YOUR-REAL-PASSWORD" in url or pwd_placeholder(url):
        raise RuntimeError("The database address still contains a placeholder. Replace it with your real database password.")
    return _normalize_url(url)


def pwd_placeholder(url):
    return "[" in url and "]" in url and "PASSWORD" in url.upper()


def _normalize_url(url):
    """Accepts postgres:// or postgresql://, and URL-encodes passwords that contain
    special characters such as @ # / : so a pasted string still works."""
    if url.startswith("postgres://"):
        url = "postgresql://" + url[len("postgres://"):]
    m = re.match(r"^(postgresql://)(.*)$", url)
    if not m:
        return url
    rest = m.group(2)
    if "@" not in rest:
        return url
    userinfo, hostpart = rest.rsplit("@", 1)
    if ":" in userinfo:
        user, pwd = userinfo.split(":", 1)
        userinfo = f"{user}:{quote(unquote(pwd), safe='')}"
    return f"{m.group(1)}{userinfo}@{hostpart}"


_POOL = None
_POOL_LOCK = threading.Lock()
_LAST_USED = {}


def _get_pool():
    global _POOL
    with _POOL_LOCK:
        if _POOL is None:
            url = _get_database_url()
            kwargs = dict(
                connect_timeout=15,
                keepalives=1, keepalives_idle=30, keepalives_interval=10, keepalives_count=5,
                cursor_factory=psycopg2.extras.DictCursor,
                application_name="pharmalab_os",
            )
            if "sslmode=" not in url and ("supabase" in url or "pooler" in url):
                kwargs["sslmode"] = "require"
            _POOL = psycopg2.pool.ThreadedConnectionPool(1, 8, url, **kwargs)
        return _POOL


def _acquire():
    pool = _get_pool()
    last_err = None
    for _ in range(3):
        raw = None
        for _wait in range(100):  # wait up to ~10 s if every connection is busy
            try:
                raw = pool.getconn()
                break
            except psycopg2.pool.PoolError:
                time.sleep(0.1)
        if raw is None:
            raise RuntimeError("Database is busy, please try again in a moment.")
        try:
            if raw.closed:
                raise psycopg2.OperationalError("connection closed")
            if time.time() - _LAST_USED.get(id(raw), 0) > 45:  # idle for a while: make sure it is alive
                with raw.cursor() as c:
                    c.execute("SELECT 1")
                raw.rollback()
            return pool, raw
        except Exception as e:  # stale connection -> throw it away and try another
            last_err = e
            try:
                pool.putconn(raw, close=True)
            except Exception:
                pass
    raise RuntimeError(f"Could not connect to the database: {last_err}")


def _translate(sql):
    """Small SQLite -> Postgres translator so existing query text keeps working."""
    head = sql.lstrip().upper()
    if head.startswith("BEGIN") or head.startswith("PRAGMA"):
        return None  # Postgres starts transactions automatically
    sql = sql.replace("CURRENT_TIMESTAMP", f"({_TS_SQL})")
    return sql.replace("?", "%s")


class _Cursor:
    def __init__(self, raw_cursor):
        self.raw = raw_cursor

    def execute(self, sql, params=None):
        sql_t = _translate(sql)
        if sql_t is not None:
            self.raw.execute(sql_t, list(params) if params is not None else None)
        return self

    def executemany(self, sql, seq):
        sql_t = _translate(sql)
        if sql_t is not None:
            psycopg2.extras.execute_batch(self.raw, sql_t, [list(p) for p in seq], page_size=500)
        return self

    def fetchone(self):
        return self.raw.fetchone()

    def fetchall(self):
        return self.raw.fetchall()

    @property
    def description(self):
        return self.raw.description

    @property
    def rowcount(self):
        return self.raw.rowcount

    def close(self):
        try:
            self.raw.close()
        except Exception:
            pass


class _Connection:
    def __init__(self, pool, raw):
        self._pool, self._raw, self._closed = pool, raw, False

    def cursor(self):
        return _Cursor(self._raw.cursor())

    def execute(self, sql, params=None):
        return self.cursor().execute(sql, params)

    def commit(self):
        self._raw.commit()

    def rollback(self):
        try:
            self._raw.rollback()
        except Exception:
            pass

    def close(self):
        if self._closed:
            return
        self._closed = True
        _LAST_USED[id(self._raw)] = time.time()
        try:
            self._pool.putconn(self._raw)  # putconn rolls back anything left uncommitted
        except Exception:
            pass


def get_db_connection():
    """Returns a pooled Supabase/Postgres connection (use .close() to give it back)."""
    pool, raw = _acquire()
    return _Connection(pool, raw)


def _read_df(conn, sql, params=()):
    cur = conn.cursor()
    cur.execute(sql, list(params))
    cols = [d[0] for d in cur.description]
    rows = [list(r) for r in cur.fetchall()]
    cur.close()
    return pd.DataFrame(rows, columns=cols)


# ==============================================================================
# SMALL IN-MEMORY READ CACHE (cleared automatically after every write)
# ==============================================================================
_CACHE = {}
_CACHE_LOCK = threading.Lock()


def _invalidate():
    with _CACHE_LOCK:
        _CACHE.clear()


def _copy(val):
    return val.copy() if hasattr(val, "copy") else val


def _cached(ttl=20):
    def deco(fn):
        @functools.wraps(fn)
        def wrapper(*args, **kwargs):
            key = (fn.__name__, repr(args), repr(sorted(kwargs.items())))
            now = time.time()
            with _CACHE_LOCK:
                hit = _CACHE.get(key)
            if hit and now - hit[0] < ttl:
                return _copy(hit[1])
            val = fn(*args, **kwargs)
            with _CACHE_LOCK:
                _CACHE[key] = (now, val)
            return _copy(val)
        return wrapper
    return deco


# ==============================================================================
# SCHEMA
# ==============================================================================

def _create_schema(cursor):
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS inventory (
        id BIGSERIAL PRIMARY KEY,
        item_id TEXT UNIQUE NOT NULL,
        lab TEXT,
        inventory_group TEXT,
        fr_code TEXT,
        item_name TEXT NOT NULL,
        item_type TEXT,
        unit TEXT,
        expiry TEXT,
        stock_location TEXT,
        current_stock DOUBLE PRECISION DEFAULT 0.0,
        reorder_level DOUBLE PRECISION DEFAULT 0.0,
        status TEXT,
        updated_at TEXT DEFAULT CURRENT_TIMESTAMP
    )
    """)

    cursor.execute("""
    CREATE TABLE IF NOT EXISTS timetable_schedule (
        id BIGSERIAL PRIMARY KEY,
        lab TEXT NOT NULL,
        day_of_week TEXT NOT NULL,
        time_slot TEXT NOT NULL,
        course_name TEXT NOT NULL,
        room TEXT,
        instructor TEXT,
        experiment_name TEXT DEFAULT '',
        apparatus_required TEXT DEFAULT '',
        chemicals_planned TEXT DEFAULT '',
        expected_students INTEGER DEFAULT 0,
        status TEXT DEFAULT 'Scheduled',
        notes TEXT DEFAULT '',
        date_logged TEXT DEFAULT '',
        semester TEXT DEFAULT 'Fall 2026',
        week_number INTEGER DEFAULT 1
    )
    """)

    cursor.execute("""
    CREATE TABLE IF NOT EXISTS experiment_logs (
        id BIGSERIAL PRIMARY KEY,
        log_id TEXT UNIQUE NOT NULL,
        timestamp TEXT DEFAULT CURRENT_TIMESTAMP,
        lab TEXT NOT NULL,
        course_name TEXT,
        instructor TEXT,
        experiment_name TEXT NOT NULL,
        apparatus_used TEXT,
        chemicals_used_json TEXT NOT NULL,
        logged_by TEXT,
        observations TEXT,
        status TEXT DEFAULT 'Completed',
        schedule_id INTEGER,
        semester TEXT DEFAULT 'Fall 2026',
        week_number INTEGER DEFAULT 1
    )
    """)

    cursor.execute("""
    CREATE TABLE IF NOT EXISTS stock_transactions (
        id BIGSERIAL PRIMARY KEY,
        timestamp TEXT DEFAULT CURRENT_TIMESTAMP,
        transaction_type TEXT NOT NULL,
        item_id TEXT NOT NULL,
        item_name TEXT NOT NULL,
        quantity_change DOUBLE PRECISION NOT NULL,
        balance_after DOUBLE PRECISION NOT NULL,
        unit TEXT,
        experiment_log_id TEXT,
        reference_reason TEXT,
        user_name TEXT
    )
    """)

    cursor.execute("""
    CREATE TABLE IF NOT EXISTS notes_and_tasks (
        id BIGSERIAL PRIMARY KEY,
        title TEXT NOT NULL,
        category TEXT NOT NULL,
        priority TEXT DEFAULT 'Medium',
        status TEXT DEFAULT 'Pending',
        due_date TEXT,
        content TEXT,
        tags TEXT DEFAULT '',
        created_at TEXT DEFAULT CURRENT_TIMESTAMP,
        updated_at TEXT DEFAULT CURRENT_TIMESTAMP
    )
    """)

    # Attachment metadata (file_path is kept for compatibility; the bytes live in attachment_files)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS attachments (
        id BIGSERIAL PRIMARY KEY,
        file_uuid TEXT UNIQUE NOT NULL,
        entity_type TEXT NOT NULL,
        entity_id TEXT NOT NULL,
        file_name TEXT NOT NULL,
        file_path TEXT NOT NULL,
        file_size_bytes BIGINT NOT NULL,
        mime_type TEXT,
        uploaded_by TEXT,
        uploaded_at TEXT DEFAULT CURRENT_TIMESTAMP
    )
    """)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS attachment_files (
        attachment_id BIGINT PRIMARY KEY REFERENCES attachments(id) ON DELETE CASCADE,
        data BYTEA NOT NULL
    )
    """)

    # Snapshots (backups) stored inside the database as compressed JSON
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS db_snapshots (
        id BIGSERIAL PRIMARY KEY,
        filename TEXT UNIQUE NOT NULL,
        tag TEXT,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP,
        size_bytes BIGINT NOT NULL,
        data BYTEA NOT NULL
    )
    """)

    # Column migrations (harmless if the columns already exist)
    for col, table, ctype in [
        ('semester', 'timetable_schedule', "TEXT DEFAULT 'Fall 2026'"),
        ('week_number', 'timetable_schedule', "INTEGER DEFAULT 1"),
        ('semester', 'experiment_logs', "TEXT DEFAULT 'Fall 2026'"),
        ('week_number', 'experiment_logs', "INTEGER DEFAULT 1"),
    ]:
        cursor.execute(f"ALTER TABLE {table} ADD COLUMN IF NOT EXISTS {col} {ctype}")

    for stmt in [
        "CREATE INDEX IF NOT EXISTS idx_inv_item_id ON inventory(item_id)",
        "CREATE INDEX IF NOT EXISTS idx_inv_lab ON inventory(lab)",
        "CREATE INDEX IF NOT EXISTS idx_inv_name ON inventory(item_name)",
        "CREATE INDEX IF NOT EXISTS idx_sched_lab_day ON timetable_schedule(lab, day_of_week)",
        "CREATE INDEX IF NOT EXISTS idx_sched_sem_week ON timetable_schedule(semester, week_number)",
        "CREATE INDEX IF NOT EXISTS idx_exp_log_id ON experiment_logs(log_id)",
        "CREATE INDEX IF NOT EXISTS idx_exp_timestamp ON experiment_logs(timestamp)",
        "CREATE INDEX IF NOT EXISTS idx_exp_sem_week ON experiment_logs(semester, week_number)",
        "CREATE INDEX IF NOT EXISTS idx_tx_item_id ON stock_transactions(item_id)",
        "CREATE INDEX IF NOT EXISTS idx_tx_timestamp ON stock_transactions(timestamp)",
        "CREATE INDEX IF NOT EXISTS idx_attach_entity ON attachments(entity_type, entity_id)",
    ]:
        cursor.execute(stmt)


_SCHEMA_READY = False


def init_schema():
    """Creates all tables (no seeding). Safe to call any number of times."""
    global _SCHEMA_READY
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.raw.execute("SELECT pg_advisory_lock(%s)", [_LOCK_KEY])
        _create_schema(cursor)
        conn.commit()
        _SCHEMA_READY = True
    finally:
        try:
            conn.rollback()
            cursor.raw.execute("SELECT pg_advisory_unlock(%s)", [_LOCK_KEY])
            conn.commit()
        except Exception:
            pass
        conn.close()


def init_db(force_reseed=False):
    """Creates tables once per app process, then seeds ONLY empty tables
    (a normal restart can never reset live stock)."""
    global _SCHEMA_READY
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        # one starter at a time (two people opening the app together must not double-seed)
        cursor.raw.execute("SELECT pg_advisory_lock(%s)", [_LOCK_KEY])

        if not _SCHEMA_READY:
            _create_schema(cursor)
            conn.commit()
            _SCHEMA_READY = True

        # SAFETY: a forced reseed replaces live inventory with the Excel file.
        # Always take a restorable snapshot first so no stock/ledger state is lost.
        if force_reseed:
            create_database_backup(tag_name="pre_reseed_safety")

        cursor.execute("SELECT COUNT(*) FROM inventory")
        if cursor.fetchone()[0] == 0 or force_reseed:
            seed_inventory_from_excel(conn)

        cursor.execute("SELECT COUNT(*) FROM timetable_schedule")
        if cursor.fetchone()[0] == 0 or force_reseed:
            seed_timetable_schedule(conn)

        cursor.execute("SELECT COUNT(*) FROM notes_and_tasks")
        if cursor.fetchone()[0] == 0:
            seed_default_notes(conn)
        conn.commit()
    finally:
        try:
            conn.rollback()
            cursor.raw.execute("SELECT pg_advisory_unlock(%s)", [_LOCK_KEY])
            conn.commit()
        except Exception:
            pass
        conn.close()
        _invalidate()


def seed_inventory_from_excel(conn):
    """Seed inventory from the provided Excel file with batch execution."""
    if not os.path.exists(EXCEL_PATH):
        print(f"Excel file not found at {EXCEL_PATH}")
        return

    try:
        print("Reading Excel file for initial database seed...")
        df = pd.read_excel(EXCEL_PATH)
        cursor = conn.cursor()
        cursor.execute("DELETE FROM inventory")

        batch_records = []
        for _, row in df.iterrows():
            item_id = str(row.get('Item ID', '')).strip()
            if not item_id or item_id.lower() == 'nan':
                continue

            lab = str(row.get('Lab', '')).strip() if pd.notna(row.get('Lab')) else ''
            inv_group = str(row.get('Inventory Group', '')).strip() if pd.notna(row.get('Inventory Group')) else ''
            fr_code = str(row.get('FR Code', '')).strip() if pd.notna(row.get('FR Code')) else ''
            item_name = str(row.get('Item Name', '')).strip() if pd.notna(row.get('Item Name')) else ''
            item_type = str(row.get('Type', '')).strip() if pd.notna(row.get('Type')) else ''
            unit = str(row.get('Unit', '')).strip() if pd.notna(row.get('Unit')) else ''

            expiry_val = row.get('Expiry')
            expiry = ''
            if pd.notna(expiry_val):
                if isinstance(expiry_val, (datetime, pd.Timestamp)):
                    expiry = expiry_val.strftime('%Y-%m-%d')
                else:
                    expiry = str(expiry_val).strip()

            stock_loc = str(row.get('Stock Location', '')).strip() if pd.notna(row.get('Stock Location')) else ''

            try:
                curr_stock = float(row.get('Current Stock', 0.0)) if pd.notna(row.get('Current Stock')) else 0.0
            except Exception:
                curr_stock = 0.0

            try:
                reorder_lvl = float(row.get('Reorder Level', 0.0)) if pd.notna(row.get('Reorder Level')) else 0.0
            except Exception:
                reorder_lvl = 0.0

            status = str(row.get('Status (as per fresh audit)', 'Available')).strip() if pd.notna(row.get('Status (as per fresh audit)')) else 'Available'

            batch_records.append((item_id, lab, inv_group, fr_code, item_name, item_type, unit, expiry, stock_loc, curr_stock, reorder_lvl, status))

        cursor.executemany("""
        INSERT INTO inventory
        (item_id, lab, inventory_group, fr_code, item_name, item_type, unit, expiry, stock_location, current_stock, reorder_level, status)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT (item_id) DO UPDATE SET
            lab = EXCLUDED.lab, inventory_group = EXCLUDED.inventory_group, fr_code = EXCLUDED.fr_code,
            item_name = EXCLUDED.item_name, item_type = EXCLUDED.item_type, unit = EXCLUDED.unit,
            expiry = EXCLUDED.expiry, stock_location = EXCLUDED.stock_location,
            current_stock = EXCLUDED.current_stock, reorder_level = EXCLUDED.reorder_level,
            status = EXCLUDED.status
        """, batch_records)

        conn.commit()
        print(f"Successfully seeded {len(batch_records)} inventory items into the Supabase database.")
    except Exception as e:
        conn.rollback()
        print(f"Error seeding inventory: {e}")


def seed_timetable_schedule(conn):
    """Seed the default timetable schedule extracted from Lab_Timetable_135B_138B_FA26.pdf."""
    cursor = conn.cursor()
    cursor.execute("DELETE FROM timetable_schedule")

    schedule_data = [
        # Lab 135-B
        {"lab": "135-B", "day_of_week": "Monday", "time_slot": "08:00 - 11:00", "course_name": "Pharmaceutics 3A", "room": "Room 211C", "instructor": "Ms. Saman A", "experiment_name": "Preparation of Emulsions & Suspensions", "apparatus_required": "Mortar & pestle, Homogenizer, Measuring cylinder", "chemicals_planned": "Liquid Paraffin: 200 mL, Acacia powder: 50 gms", "week_number": 1},
        {"lab": "135-B", "day_of_week": "Monday", "time_slot": "11:00 - 14:00", "course_name": "Practice 2", "room": "Room 315A", "instructor": "Ms. Muryam AR", "experiment_name": "Hospital Pharmacy Prescription Compounding", "apparatus_required": "Glass beakers 250ml, Stirrer rods, Droppers", "chemicals_planned": "Glycerin: 100 mL, Purified Water: 500 mL", "week_number": 1},
        {"lab": "135-B", "day_of_week": "Monday", "time_slot": "14:00 - 17:00", "course_name": "Pharmaceutics 1A", "room": "Room 121C", "instructor": "Ms. Muryam", "experiment_name": "Introduction to Pharmaceutical Calculations & Weighing", "apparatus_required": "Analytical balance, Watch glass, Spatulas", "chemicals_planned": "Sodium Chloride: 50 gms, Lactose: 50 gms", "week_number": 1},
        {"lab": "135-B", "day_of_week": "Monday", "time_slot": "08:00 - 08:50 (Pre-lab)", "course_name": "Pharmaceutics 2A", "room": "Room 210A", "instructor": "Dr. Abdul Haleem Khan", "experiment_name": "Physical Pharmacy Surface Tension Determination", "apparatus_required": "Stalagometer, Specific gravity bottle", "chemicals_planned": "Ethanol 96%: 150 mL, Distilled water", "week_number": 1},
        
        {"lab": "135-B", "day_of_week": "Tuesday", "time_slot": "08:00 - 11:00", "course_name": "Practice", "room": "Room 319A", "instructor": "Mr. Omaid K", "experiment_name": "Dosage Form Dispensing & Labeling Protocol", "apparatus_required": "Dispensing amber bottles 60ml, Pipette pumps", "chemicals_planned": "Simple Syrup BP: 250 mL", "week_number": 1},
        {"lab": "135-B", "day_of_week": "Tuesday", "time_slot": "14:00 - 17:00", "course_name": "Pharmaceutics 3A", "room": "Room 211A", "instructor": "Ms. Saman A", "experiment_name": "Ointment Base Formulation & Fusion Method", "apparatus_required": "Water bath, Porcelain evaporating dish, Spatulas", "chemicals_planned": "White Soft Paraffin: 250 gms, Cetostearyl Alcohol: 50 gms", "week_number": 1},
        
        {"lab": "135-B", "day_of_week": "Wednesday", "time_slot": "08:00 - 11:00", "course_name": "Pharmaceutics 3A", "room": "Room 211C", "instructor": "Ms. Saman A", "experiment_name": "Suppositories Preparation (Mould Calibration & Displacement Value)", "apparatus_required": "Suppository moulds (1g/2g), Water bath, Lubricant", "chemicals_planned": "Theobroma oil (Cocoa Butter): 150 gms, Paracetamol powder: 20 gms", "week_number": 1},
        {"lab": "135-B", "day_of_week": "Wednesday", "time_slot": "11:00 - 14:00", "course_name": "Practice 2", "room": "Room 315A", "instructor": "Ms. Muryam AR", "experiment_name": "Incompatibilities in Liquid Preparations", "apparatus_required": "Conical flasks 100ml, Filter funnels, Filter paper", "chemicals_planned": "Quinine sulphate: 10 gms, Dilute Sulphuric Acid: 50 mL", "week_number": 1},
        {"lab": "135-B", "day_of_week": "Wednesday", "time_slot": "14:00 - 17:00", "course_name": "Pharmaceutics 2A", "room": "Room 210B", "instructor": "Dr. Abdul Haleem Khan", "experiment_name": "Viscosity Determination using Ostwald Viscometer", "apparatus_required": "Ostwald Viscometer, Stop-watch, Constant temp bath", "chemicals_planned": "Glycerol solutions (10%, 20%, 30%): 300 mL", "week_number": 1},

        {"lab": "135-B", "day_of_week": "Thursday", "time_slot": "11:00 - 14:00", "course_name": "Practice 2A", "room": "Room 310A", "instructor": "Mr. Sufyan J", "experiment_name": "Dispensing of Topical Powders & Dusting Powders", "apparatus_required": "Sieves #80/#100, Tile & Spatula", "chemicals_planned": "Talc powder: 200 gms, Zinc Oxide: 100 gms, Starch: 100 gms", "week_number": 1},
        {"lab": "135-B", "day_of_week": "Thursday", "time_slot": "14:00 - 17:00", "course_name": "Pharmaceutics 1A", "room": "Room 121B", "instructor": "Ms. Muryam", "experiment_name": "Preparation of Lugol's Iodine Solution", "apparatus_required": "Volumetric flask 250ml, Glass funnel, Amber bottle", "chemicals_planned": "Iodine resublimed: 15 gms, Potassium Iodide: 25 gms", "week_number": 1},

        {"lab": "135-B", "day_of_week": "Friday", "time_slot": "08:00 - 11:00", "course_name": "Pharmaceutics 3A", "room": "Room 211C", "instructor": "Ms. Saman A", "experiment_name": "Quality Evaluation of Tablets (Hardness, Friability, Disintegration)", "apparatus_required": "Monsanto hardness tester, Roche friabilator, Disintegration tester", "chemicals_planned": "Standard Paracetamol tablets (Batch sample)", "week_number": 1},
        {"lab": "135-B", "day_of_week": "Friday", "time_slot": "11:00 - 14:00", "course_name": "Practice 2", "room": "Room 315A", "instructor": "Ms. Muryam AR", "experiment_name": "Pediatric Elixirs Formulation & Stability", "apparatus_required": "Magnetic stirrer, Beakers, Pipettes", "chemicals_planned": "Propylene glycol: 150 mL, Alcohol 95%: 100 mL", "week_number": 1},
        {"lab": "135-B", "day_of_week": "Friday", "time_slot": "14:00 - 17:00", "course_name": "Practice 5B", "room": "Room 356B", "instructor": "Mr. Omaid K", "experiment_name": "Clinical Pharmacokinetics Simulation & TDM Case Study", "apparatus_required": "Scientific calculators, Graph sheets, Formulary books", "chemicals_planned": "None (Dry lab / computational)", "week_number": 1},
        {"lab": "135-B", "day_of_week": "Friday", "time_slot": "08:00 - 08:50 (Pre-lab)", "course_name": "Pharmaceutics 2A", "room": "Room 210C", "instructor": "Dr. Abdul Haleem Khan", "experiment_name": "Partition Coefficient of Benzoic Acid in Oil/Water", "apparatus_required": "Separating funnel 250ml, Burette 50ml", "chemicals_planned": "Benzoic Acid: 20 gms, Benzene/Chloroform: 100 mL, 0.1M NaOH: 250 mL", "week_number": 1},

        # Lab 138-B
        {"lab": "138-B", "day_of_week": "Monday", "time_slot": "08:00 - 11:00", "course_name": "Pharmacognosy", "room": "Room 213A", "instructor": "Mr. Sabi UR", "experiment_name": "Morphological & Microscopic Identification of Senna Leaf", "apparatus_required": "Compound microscope, Razor blades, Glass slides, Cover slips", "chemicals_planned": "Chloral hydrate solution: 50 mL, Phloroglucinol + HCl: 30 mL", "week_number": 1},
        {"lab": "138-B", "day_of_week": "Monday", "time_slot": "11:00 - 14:00", "course_name": "Pharmacognosy 2A", "room": "Room 313A", "instructor": "Mr. Sohaib P", "experiment_name": "Extraction of Alkaloids from Cinchona Bark (Stas-Otto method)", "apparatus_required": "Soxhlet extraction apparatus, Heating mantle, Rotary evaporator", "chemicals_planned": "Methanol: 500 mL, Chloroform: 200 mL, Dilute Ammonia: 50 mL", "week_number": 1},
        {"lab": "138-B", "day_of_week": "Monday", "time_slot": "14:00 - 17:00", "course_name": "Pharm. Chem 3B", "room": "Room 316A", "instructor": "Mr. Nasir A", "experiment_name": "Synthesis and Purification of Aspirin (Acetylsalicylic Acid)", "apparatus_required": "Reflux condenser, Buchner funnel, Suction flask, Melting point apparatus", "chemicals_planned": "Salicylic acid: 100 gms, Acetic anhydride: 150 mL, Concentrated H2SO4: 20 mL", "week_number": 1},

        {"lab": "138-B", "day_of_week": "Tuesday", "time_slot": "08:00 - 11:00", "course_name": "Pharmaceutics 3A", "room": "Room 211C", "instructor": "Ms. Saman A", "experiment_name": "Microencapsulation of Drugs by Coacervation Phase Separation", "apparatus_required": "Overhead stirrer, Temp controlled water bath", "chemicals_planned": "Gelatin: 50 gms, Sodium sulphate solution: 200 mL", "week_number": 1},
        {"lab": "138-B", "day_of_week": "Tuesday", "time_slot": "11:00 - 12:15", "course_name": "Pharmaceutics 5B", "room": "Room 358A", "instructor": "Dr. Omer Salman Q", "experiment_name": "Biopharmaceutics Dissolution Profile Testing (USP Apparatus II)", "apparatus_required": "USP Dissolution Tester, Syringe filters, UV spectrophotometer", "chemicals_planned": "0.1N Hydrochloric Acid: 2000 mL, Phosphate Buffer pH 6.8: 2000 mL", "week_number": 1},
        {"lab": "138-B", "day_of_week": "Tuesday", "time_slot": "14:00 - 17:00", "course_name": "Pharmacognosy", "room": "Room 213B", "instructor": "Mr. Sohaib P", "experiment_name": "Isolation of Volatile Oil from Clove by Clevenger Apparatus", "apparatus_required": "Clevenger distillation apparatus, Round bottom flask 1000ml", "chemicals_planned": "Clove buds (crude drug): 100 gms, Anhydrous sodium sulfate: 30 gms", "week_number": 1},

        {"lab": "138-B", "day_of_week": "Wednesday", "time_slot": "08:00 - 11:00", "course_name": "Pharmacognosy", "room": "Room 213A", "instructor": "Mr. Sabi UR", "experiment_name": "Phytochemical Screening: Tests for Tannins & Flavonoids", "apparatus_required": "Test tubes & rack, Bunsen burner, Pipettes", "chemicals_planned": "Ferric Chloride 5%: 50 mL, Gelatin solution 1%: 50 mL, Lead acetate 10%: 50 mL", "week_number": 1},
        {"lab": "138-B", "day_of_week": "Wednesday", "time_slot": "11:00 - 14:00", "course_name": "Pharmacognosy 2B", "room": "Room 318A", "instructor": "Mr. Sabi R", "experiment_name": "Thin Layer Chromatography (TLC) of Plant Pigments", "apparatus_required": "TLC Silica gel plates, Developing chamber, UV viewing cabinet", "chemicals_planned": "Petroleum ether: 100 mL, Acetone: 100 mL, Ninhydrin spray: 20 mL", "week_number": 1},
        {"lab": "138-B", "day_of_week": "Wednesday", "time_slot": "14:00 - 17:00", "course_name": "Pharmacognosy", "room": "Room 213C", "instructor": "Mr. Sohaib P", "experiment_name": "Quantitative Microscopy: Lycopodium Spore Method", "apparatus_required": "Microscope, Camera lucida/Stage micrometer, Hemocytometer", "chemicals_planned": "Lycopodium powder: 10 gms, Fixed oil/Glycerol: 50 mL", "week_number": 1},

        {"lab": "138-B", "day_of_week": "Thursday", "time_slot": "08:00 - 11:00", "course_name": "Pharmaceutics 3A", "room": "Room 211C", "instructor": "Ms. Saman A", "experiment_name": "Granulation & Tablet Compression Studies", "apparatus_required": "Single punch tablet press, Granulator, Sieves #16/#20", "chemicals_planned": "Starch paste 10%: 200 gms, Magnesium stearate: 20 gms, Talc: 30 gms", "week_number": 1},
        {"lab": "138-B", "day_of_week": "Thursday", "time_slot": "11:00 - 12:15", "course_name": "Pharmaceutics 5B", "room": "Room 358A", "instructor": "Dr. Omer Salman Q", "experiment_name": "In Vitro Drug Permeation Study Using Franz Diffusion Cell", "apparatus_required": "Franz diffusion cell, Dialysis membrane, Micro pipettes", "chemicals_planned": "Phosphate buffered saline pH 7.4: 500 mL", "week_number": 1},
        {"lab": "138-B", "day_of_week": "Thursday", "time_slot": "14:00 - 17:00", "course_name": "Pharmaceutics 3A", "room": "Room 211C", "instructor": "Ms. Saman A", "experiment_name": "Stability Testing of Liquid Dosage Forms at Elevated Temperatures", "apparatus_required": "Stability oven 40°C, pH meter, Viscometer", "chemicals_planned": "Buffered ascorbic acid syrup solution: 300 mL", "week_number": 1},

        {"lab": "138-B", "day_of_week": "Friday", "time_slot": "08:00 - 11:00", "course_name": "Pharmacognosy", "room": "Room 213A", "instructor": "Mr. Sabi UR", "experiment_name": "Histochemical Staining & Cell Wall Components Analysis", "apparatus_required": "Slide warmers, Microscopic reagents set", "chemicals_planned": "Ruthenium red: 20 mL, Iodine/potassium iodide solution: 50 mL", "week_number": 1},
        {"lab": "138-B", "day_of_week": "Friday", "time_slot": "11:00 - 14:00", "course_name": "Pharmacognosy 2B", "room": "Room 318B", "instructor": "Mr. Sabi R", "experiment_name": "Extraction of Curcumin from Curcuma longa", "apparatus_required": "Reflux condenser, Filter funnel, Vacuum desiccator", "chemicals_planned": "Acetone: 250 mL, Hexane: 150 mL, Turmeric powder: 100 gms", "week_number": 1},
        {"lab": "138-B", "day_of_week": "Friday", "time_slot": "14:00 - 17:00", "course_name": "Pharmacognosy", "room": "Room 213A", "instructor": "Mr. Sabi UR", "experiment_name": "Chromatographic Separation of Anthraquinone Glycosides", "apparatus_required": "Column chromatography glass column, UV lamp 365nm", "chemicals_planned": "Silica gel for column: 150 gms, Ethyl acetate: 200 mL, Methanol: 100 mL", "week_number": 1}
    ]

    cursor.executemany("""
    INSERT INTO timetable_schedule
    (lab, day_of_week, time_slot, course_name, room, instructor, experiment_name, apparatus_required, chemicals_planned, expected_students, status, semester, week_number)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'Scheduled', 'Fall 2026', ?)
    """, [(
        item["lab"], item["day_of_week"], item["time_slot"], item["course_name"],
        item["room"], item["instructor"], item["experiment_name"],
        item["apparatus_required"], item["chemicals_planned"], 35, item.get("week_number", 1)
    ) for item in schedule_data])

    conn.commit()
    print("Timetable schedule successfully seeded.")


def seed_default_notes(conn):
    """Seed helpful default notes and productivity tasks."""
    cursor = conn.cursor()
    default_notes = [
        ("Safety Audit & Eye-Wash Inspection", "Safety & Compliance", "High", "In Progress", "2026-10-05", "Inspect emergency eyewash stations and chemical spill kits in Lab 135-B and 138-B. Verify neutralization absorbent availability.", "safety, audit, compliance"),
        ("Requisition for Hydrochloric Acid & Acetone", "Procurement & Requisition", "High", "Pending", "2026-10-08", "Coordinate with Central Chemical Store for batch replenishment of 5L Analytical Grade Acetone and 2.5L Conc. HCl for FA26 labs.", "chemicals, purchase, order"),
        ("Annual Calibration of Analytical Balances", "Equipment Maintenance", "Medium", "Pending", "2026-10-12", "Service technician from Shimadzu scheduled to calibrate precision electronic balances in Room 211C and Room 213A.", "calibration, hardware, balance"),
        ("Pharmacognosy Herbarium Specimen Follow-up", "Correspondences", "Medium", "In Progress", "2026-10-04", "Follow up with botanical garden curator regarding delivery of fresh Digitalis purpurea and Senna folium specimens for week 4.", "botany, specimens, pharmacognosy"),
        ("Glassware Breakage Ledger Reconciliation", "Follow-ups", "Low", "Pending", "2026-10-15", "Reconcile student breakage slips from Pharmaceutics 3A and update inventory stock for 100ml measuring cylinders.", "glassware, breakage, audit")
    ]
    cursor.executemany("""
    INSERT INTO notes_and_tasks (title, category, priority, status, due_date, content, tags)
    VALUES (?, ?, ?, ?, ?, ?, ?)
    """, default_notes)
    conn.commit()


# --- Inventory Operations ---

@_cached()
def get_inventory_items(search_query="", lab_filter="All", group_filter="All", type_filter="All", status_filter="All", limit=2000):
    """Queries inventory items with flexible filters."""
    conn = get_db_connection()
    try:
        query = "SELECT * FROM inventory WHERE 1=1"
        params = []

        if lab_filter and lab_filter != "All":
            query += " AND lab = ?"
            params.append(lab_filter)

        if group_filter and group_filter != "All":
            query += " AND inventory_group = ?"
            params.append(group_filter)

        if type_filter and type_filter != "All":
            query += " AND item_type = ?"
            params.append(type_filter)

        if status_filter == "Low Stock":
            query += " AND current_stock <= reorder_level AND current_stock > 0"
        elif status_filter == "Out of Stock":
            query += " AND current_stock <= 0"
        elif status_filter == "Available":
            query += " AND current_stock > 0"

        if search_query:
            query += " AND (item_name ILIKE ? OR item_id ILIKE ? OR fr_code ILIKE ? OR stock_location ILIKE ?)"
            q = f"%{search_query.strip()}%"
            params.extend([q, q, q, q])

        query += " ORDER BY item_name ASC LIMIT ?"
        params.append(int(limit))

        return _read_df(conn, query, params)
    finally:
        conn.close()


def get_inventory_item_by_id(item_id):
    """Fetch single item details."""
    conn = get_db_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM inventory WHERE item_id = ?", (item_id,))
        row = cursor.fetchone()
        return dict(row) if row else None
    finally:
        conn.close()


# --- Stock Deduction & Pharmacy Style Billing / POS Cart ---

def deduct_inventory_for_experiment(experiment_info, items_used, user_name="Lab Technician"):
    """
    Deducts quantities for multiple items in a single ACID transaction,
    creates experiment log, and records audit trail transactions.
    Rows are locked while deducting, so two users can never overwrite each other's stock.
    """
    conn = get_db_connection()
    cursor = conn.cursor()

    try:
        # Generate unique Log ID
        log_id = f"EXP-{_now().strftime('%Y%m%d-%H%M%S')}"
        cursor.execute("SELECT 1 FROM experiment_logs WHERE log_id = ?", (log_id,))
        if cursor.fetchone():
            log_id = f"{log_id}-{uuid.uuid4().hex[:4].upper()}"

        # Lock every item involved (sorted order, so two users can not deadlock)
        ids = sorted({str(i['item_id']) for i in items_used})
        if ids:
            cursor.execute("SELECT item_id FROM inventory WHERE item_id = ANY(?) ORDER BY item_id FOR UPDATE", (ids,))
            cursor.fetchall()

        # 1. Deduct stock for each item & record transaction
        for item in items_used:
            item_id = item['item_id']
            qty_deduct = float(item['quantity'])

            cursor.execute("SELECT current_stock, item_name, unit FROM inventory WHERE item_id = ?", (item_id,))
            row = cursor.fetchone()
            if not row:
                raise ValueError(f"Item ID {item_id} not found in inventory.")

            curr_stock = float(row['current_stock'] or 0.0)
            item_name = row['item_name']
            unit = row['unit']

            new_stock = max(0.0, curr_stock - qty_deduct)

            # Update inventory table
            cursor.execute("""
            UPDATE inventory
            SET current_stock = ?, updated_at = CURRENT_TIMESTAMP
            WHERE item_id = ?
            """, (new_stock, item_id))

            # Record in stock_transactions ledger
            cursor.execute("""
            INSERT INTO stock_transactions
            (transaction_type, item_id, item_name, quantity_change, balance_after, unit, experiment_log_id, reference_reason, user_name)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                "DISPENSE / EXPERIMENT USAGE",
                item_id,
                item_name,
                -qty_deduct,
                new_stock,
                unit,
                log_id,
                f"Used in Experiment: {experiment_info.get('experiment_name', 'General Lab')}",
                user_name
            ))

        # 2. Insert into experiment_logs with semester and week_number
        cursor.execute("""
        INSERT INTO experiment_logs
        (log_id, lab, course_name, instructor, experiment_name, apparatus_used, chemicals_used_json, logged_by, observations, schedule_id, semester, week_number)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            log_id,
            experiment_info.get('lab', '135-B'),
            experiment_info.get('course_name', ''),
            experiment_info.get('instructor', ''),
            experiment_info.get('experiment_name', ''),
            experiment_info.get('apparatus_used', ''),
            json.dumps(items_used, default=str),
            user_name,
            experiment_info.get('observations', ''),
            experiment_info.get('schedule_id', None),
            experiment_info.get('semester', 'Fall 2026'),
            int(experiment_info.get('week_number', 1))
        ))

        # 3. Update schedule status if linked
        if experiment_info.get('schedule_id'):
            cursor.execute("""
            UPDATE timetable_schedule
            SET status = 'Completed'
            WHERE id = ?
            """, (experiment_info['schedule_id'],))

        conn.commit()
        conn.close()
        _invalidate()
        return True, log_id, "Stock successfully deducted and experiment logged."
    except Exception as e:
        conn.rollback()
        conn.close()
        return False, None, str(e)


def restock_or_adjust_item(item_id, quantity_change, adjustment_type="RESTOCK", reason="Stock Inward / Delivery", user_name="Inventory Manager"):
    """Adds stock or manually adjusts quantity with audit trail."""
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT current_stock, item_name, unit FROM inventory WHERE item_id = ? FOR UPDATE", (item_id,))
        row = cursor.fetchone()
        if not row:
            raise ValueError(f"Item ID {item_id} not found.")

        current = float(row['current_stock'] or 0.0)
        new_balance = max(0.0, current + quantity_change) if adjustment_type == "RESTOCK" else max(0.0, quantity_change)
        actual_change = new_balance - current

        cursor.execute("UPDATE inventory SET current_stock = ?, updated_at = CURRENT_TIMESTAMP WHERE item_id = ?", (new_balance, item_id))

        cursor.execute("""
        INSERT INTO stock_transactions
        (transaction_type, item_id, item_name, quantity_change, balance_after, unit, experiment_log_id, reference_reason, user_name)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            adjustment_type,
            item_id,
            row['item_name'],
            actual_change,
            new_balance,
            row['unit'],
            None,
            reason,
            user_name
        ))

        conn.commit()
        conn.close()
        _invalidate()
        return True, f"Item {item_id} updated. Previous: {current}, New Balance: {new_balance}"
    except Exception as e:
        conn.rollback()
        conn.close()
        return False, str(e)


# --- Corrections: return unused chemicals / void a wrongly logged experiment ---
# Nothing is ever deleted from the ledger. A correction ADDS new lines
# ("RETURN TO STOCK" / "VOID REVERSAL") that point at the original experiment log.

_DISPENSE_TYPE = "DISPENSE / EXPERIMENT USAGE"
_RETURN_TYPE = "RETURN TO STOCK"
_VOID_TYPE = "VOID REVERSAL"


def get_experiment_usage(log_id):
    """Per item for one experiment log: dispensed, returned so far, and still counted as used."""
    conn = get_db_connection()
    try:
        df = _read_df(conn, """
        SELECT item_id,
               MAX(item_name) AS item_name,
               MAX(unit) AS unit,
               COALESCE(SUM(CASE WHEN transaction_type = ? THEN -quantity_change ELSE 0 END), 0) AS dispensed,
               COALESCE(SUM(CASE WHEN transaction_type IN (?, ?) THEN quantity_change ELSE 0 END), 0) AS returned
        FROM stock_transactions
        WHERE experiment_log_id = ?
        GROUP BY item_id
        ORDER BY MAX(item_name)
        """, [_DISPENSE_TYPE, _RETURN_TYPE, _VOID_TYPE, log_id])
    finally:
        conn.close()
    if df.empty:
        return pd.DataFrame(columns=["item_id", "item_name", "unit", "dispensed", "returned", "net_used"])
    df["dispensed"] = df["dispensed"].astype(float)
    df["returned"] = df["returned"].astype(float)
    df["net_used"] = (df["dispensed"] - df["returned"]).round(6)
    return df


def _apply_return(cursor, log_id, item_id, qty, ttype, reason, user_name):
    """Core of a return (inside the caller's transaction). Locks the stock row, checks the
    quantity against what the log still counts as used, adds stock and writes the ledger line."""
    cursor.execute("SELECT current_stock, item_name, unit FROM inventory WHERE item_id = ? FOR UPDATE", (item_id,))
    inv = cursor.fetchone()
    if not inv:
        raise ValueError(f"Item {item_id} no longer exists in the inventory.")

    cursor.execute("""
    SELECT COALESCE(SUM(CASE WHEN transaction_type = ? THEN -quantity_change ELSE 0 END), 0)
         - COALESCE(SUM(CASE WHEN transaction_type IN (?, ?) THEN quantity_change ELSE 0 END), 0)
    FROM stock_transactions WHERE experiment_log_id = ? AND item_id = ?
    """, (_DISPENSE_TYPE, _RETURN_TYPE, _VOID_TYPE, log_id, item_id))
    net_used = float(cursor.fetchone()[0] or 0.0)
    if qty > net_used + 1e-9:
        raise ValueError(f"Cannot return {qty:g}: only {max(net_used, 0):g} of this item is still counted as used in {log_id}.")

    new_balance = float(inv["current_stock"] or 0.0) + qty
    cursor.execute("UPDATE inventory SET current_stock = ?, updated_at = CURRENT_TIMESTAMP WHERE item_id = ?", (new_balance, item_id))
    cursor.execute("""
    INSERT INTO stock_transactions
    (transaction_type, item_id, item_name, quantity_change, balance_after, unit, experiment_log_id, reference_reason, user_name)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (ttype, item_id, inv["item_name"], qty, new_balance, inv["unit"], log_id, reason, user_name))
    return new_balance, inv["item_name"], inv["unit"]


def return_to_stock(log_id, item_id, quantity, reason="Unused chemical returned", user_name="Lab Officer"):
    """Returns part (or all) of one item from an experiment back to stock.
    Returns (True, message) or (False, error message)."""
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        qty = float(quantity)
        if qty <= 0:
            raise ValueError("The quantity to return must be greater than zero.")
        cursor.execute("SELECT status FROM experiment_logs WHERE log_id = ? FOR UPDATE", (log_id,))
        row = cursor.fetchone()
        if not row:
            raise ValueError(f"Experiment log {log_id} not found.")
        if str(row["status"]) == "Voided":
            raise ValueError("This experiment log is already voided; its chemicals were returned to stock.")
        new_balance, name, unit = _apply_return(cursor, log_id, item_id, qty, _RETURN_TYPE,
                                                (reason or "Unused chemical returned").strip(), user_name)
        conn.commit()
        conn.close()
        _invalidate()
        return True, f"Returned {qty:g} {unit or ''} of {name} to stock. New balance: {new_balance:g}"
    except Exception as e:
        conn.rollback()
        conn.close()
        return False, str(e)


def void_experiment_log(log_id, reason, user_name="Lab Officer"):
    """Cancels a wrongly logged experiment: everything still counted as used is returned to stock
    and the log is marked 'Voided' (it stays visible, with the reason, for the audit trail)."""
    reason = (reason or "").strip()
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        if not reason:
            raise ValueError("Please give a reason for voiding this log.")
        cursor.execute("SELECT status, schedule_id FROM experiment_logs WHERE log_id = ? FOR UPDATE", (log_id,))
        row = cursor.fetchone()
        if not row:
            raise ValueError(f"Experiment log {log_id} not found.")
        if str(row["status"]) == "Voided":
            raise ValueError("This experiment log is already voided.")
        schedule_id = row["schedule_id"]

        cursor.execute("""
        SELECT item_id,
               COALESCE(SUM(CASE WHEN transaction_type = ? THEN -quantity_change ELSE 0 END), 0)
             - COALESCE(SUM(CASE WHEN transaction_type IN (?, ?) THEN quantity_change ELSE 0 END), 0) AS net_used
        FROM stock_transactions WHERE experiment_log_id = ?
        GROUP BY item_id ORDER BY item_id
        """, (_DISPENSE_TYPE, _RETURN_TYPE, _VOID_TYPE, log_id))
        to_return = [(r["item_id"], float(r["net_used"])) for r in cursor.fetchall() if float(r["net_used"]) > 1e-9]

        for item_id, qty in to_return:  # sorted by item_id: two users can not deadlock
            _apply_return(cursor, log_id, item_id, qty, _VOID_TYPE, f"Log voided: {reason}", user_name)

        note = f"\n[VOIDED {_now().strftime('%Y-%m-%d %H:%M')} by {user_name}: {reason}]"
        cursor.execute("UPDATE experiment_logs SET status = 'Voided', observations = COALESCE(observations, '') || ? WHERE log_id = ?",
                       (note, log_id))

        if schedule_id:  # un-complete the timetable slot if no other live log uses it
            cursor.execute("SELECT COUNT(*) FROM experiment_logs WHERE schedule_id = ? AND status <> 'Voided'", (schedule_id,))
            if cursor.fetchone()[0] == 0:
                cursor.execute("UPDATE timetable_schedule SET status = 'Scheduled' WHERE id = ?", (schedule_id,))

        conn.commit()
        conn.close()
        _invalidate()
        return True, f"Log {log_id} voided. {len(to_return)} item(s) returned to stock."
    except Exception as e:
        conn.rollback()
        conn.close()
        return False, str(e)


# --- Schedule Operations ---

@_cached()
def get_schedule(lab_filter="All", day_filter="All", semester_filter="All", week_filter="All"):
    """Fetch timetable entries with flexible filters."""
    conn = get_db_connection()
    try:
        query = "SELECT * FROM timetable_schedule WHERE 1=1"
        params = []
        if lab_filter and lab_filter != "All":
            query += " AND lab = ?"
            params.append(lab_filter)
        if day_filter and day_filter != "All":
            query += " AND day_of_week = ?"
            params.append(day_filter)
        if semester_filter and semester_filter != "All":
            query += " AND semester = ?"
            params.append(semester_filter)
        if week_filter and week_filter != "All":
            query += " AND week_number = ?"
            params.append(int(week_filter))

        query += " ORDER BY CASE day_of_week WHEN 'Monday' THEN 1 WHEN 'Tuesday' THEN 2 WHEN 'Wednesday' THEN 3 WHEN 'Thursday' THEN 4 WHEN 'Friday' THEN 5 ELSE 6 END, time_slot"
        return _read_df(conn, query, params)
    finally:
        conn.close()


def update_schedule_entry(entry_id, experiment_name, apparatus_required, chemicals_planned, expected_students, status, notes, semester="Fall 2026", week_number=1):
    """Updates pre-planned experiment details for a schedule slot."""
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""
        UPDATE timetable_schedule
        SET experiment_name = ?, apparatus_required = ?, chemicals_planned = ?, expected_students = ?, status = ?, notes = ?, semester = ?, week_number = ?
        WHERE id = ?
        """, (experiment_name, apparatus_required, chemicals_planned, int(expected_students), status, notes, semester, int(week_number), int(entry_id)))
        conn.commit()
    finally:
        conn.close()
    _invalidate()
    return True


def add_schedule_entry(lab, day_of_week, time_slot, course_name, room, instructor, experiment_name="", apparatus="", chemicals="", students=35, notes="", semester="Fall 2026", week_number=1):
    """Adds a new schedule slot."""
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""
        INSERT INTO timetable_schedule
        (lab, day_of_week, time_slot, course_name, room, instructor, experiment_name, apparatus_required, chemicals_planned, expected_students, status, notes, semester, week_number)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'Scheduled', ?, ?, ?)
        """, (lab, day_of_week, time_slot, course_name, room, instructor, experiment_name, apparatus, chemicals, int(students), notes, semester, int(week_number)))
        conn.commit()
    finally:
        conn.close()
    _invalidate()
    return True


# --- Experiment Logs & Audit Trail Operations ---

@_cached()
def get_experiment_logs(lab_filter="All", semester_filter="All", week_filter="All", limit=1000):
    conn = get_db_connection()
    try:
        query = "SELECT * FROM experiment_logs WHERE 1=1"
        params = []
        if lab_filter and lab_filter != "All":
            query += " AND lab = ?"
            params.append(lab_filter)
        if semester_filter and semester_filter != "All":
            query += " AND semester = ?"
            params.append(semester_filter)
        if week_filter and week_filter != "All":
            query += " AND week_number = ?"
            params.append(int(week_filter))
        query += " ORDER BY timestamp DESC LIMIT ?"
        params.append(int(limit))
        return _read_df(conn, query, params)
    finally:
        conn.close()


@_cached()
def get_stock_transactions(item_id=None, limit=2000):
    conn = get_db_connection()
    try:
        query = "SELECT * FROM stock_transactions WHERE 1=1"
        params = []
        if item_id:
            query += " AND item_id = ?"
            params.append(item_id)
        query += " ORDER BY timestamp DESC LIMIT ?"
        params.append(int(limit))
        return _read_df(conn, query, params)
    finally:
        conn.close()


# --- Notes & Productivity Tasks Operations ---

@_cached()
def get_notes(category="All", status="All", search=""):
    conn = get_db_connection()
    try:
        query = "SELECT * FROM notes_and_tasks WHERE 1=1"
        params = []
        if category and category != "All":
            query += " AND category = ?"
            params.append(category)
        if status and status != "All":
            query += " AND status = ?"
            params.append(status)
        if search:
            query += " AND (title ILIKE ? OR content ILIKE ? OR tags ILIKE ?)"
            q = f"%{search.strip()}%"
            params.extend([q, q, q])
        query += " ORDER BY CASE priority WHEN 'High' THEN 1 WHEN 'Medium' THEN 2 WHEN 'Low' THEN 3 END, created_at DESC"
        return _read_df(conn, query, params)
    finally:
        conn.close()


def add_note(title, category, priority, status, due_date, content, tags=""):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""
        INSERT INTO notes_and_tasks (title, category, priority, status, due_date, content, tags)
        VALUES (?, ?, ?, ?, ?, ?, ?)
        """, (title, category, priority, status, due_date, content, tags))
        conn.commit()
    finally:
        conn.close()
    _invalidate()
    return True


def update_note_status(note_id, new_status):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("UPDATE notes_and_tasks SET status = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?", (new_status, int(note_id)))
        conn.commit()
    finally:
        conn.close()
    _invalidate()
    return True


def delete_note(note_id):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("DELETE FROM notes_and_tasks WHERE id = ?", (int(note_id),))
        conn.commit()
    finally:
        conn.close()
    _invalidate()
    return True


# --- Attachments Repository (stored inside the database) ---

def save_attachment(entity_type, entity_id, file_bytes, original_filename, uploaded_by="Lab Incharge"):
    """
    Saves an uploaded file inside the database (table attachment_files) and records
    its metadata. Files above MAX_ATTACHMENT_MB are refused (free database space is limited).
    """
    file_size = len(file_bytes)
    if file_size > MAX_ATTACHMENT_BYTES:
        raise ValueError(f"File is too large ({file_size / (1024 * 1024):.1f} MB). The limit is {MAX_ATTACHMENT_MB} MB per file.")

    file_uuid = str(uuid.uuid4())
    ext = os.path.splitext(original_filename)[1]
    virtual_path = f"db://attachments/{file_uuid}{ext}"

    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""
        INSERT INTO attachments
        (file_uuid, entity_type, entity_id, file_name, file_path, file_size_bytes, mime_type, uploaded_by)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        RETURNING id
        """, (file_uuid, entity_type, str(entity_id), original_filename, virtual_path, file_size, ext.lower(), uploaded_by))
        new_id = cursor.fetchone()[0]
        cursor.execute("INSERT INTO attachment_files (attachment_id, data) VALUES (?, ?)", (new_id, psycopg2.Binary(file_bytes)))
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()
    _invalidate()
    return file_uuid, virtual_path


@_cached()
def get_attachments(entity_type=None, entity_id=None):
    """Fetches attachments metadata filtered by entity (file contents are loaded only on demand)."""
    conn = get_db_connection()
    try:
        query = "SELECT * FROM attachments WHERE 1=1"
        params = []
        if entity_type:
            query += " AND entity_type = ?"
            params.append(entity_type)
        if entity_id:
            query += " AND entity_id = ?"
            params.append(str(entity_id))
        query += " ORDER BY uploaded_at DESC"
        return _read_df(conn, query, params)
    finally:
        conn.close()


def get_attachment_bytes(attachment_id):
    """Returns the stored file content (bytes) or None."""
    conn = get_db_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("SELECT data FROM attachment_files WHERE attachment_id = ?", (int(attachment_id),))
        row = cursor.fetchone()
        return bytes(row[0]) if row else None
    finally:
        conn.close()


def delete_attachment(attachment_id):
    """Deletes an attachment (metadata and file content)."""
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("DELETE FROM attachments WHERE id = ?", (int(attachment_id),))  # file content is removed by cascade
        conn.commit()
    finally:
        conn.close()
    _invalidate()
    return True


# --- Semester Backup & Snapshot System (stored in the database, downloadable) ---

_SNAPSHOT_TABLES = ["inventory", "timetable_schedule", "experiment_logs", "stock_transactions", "notes_and_tasks"]


def _build_snapshot_bytes(conn):
    """Consistent point-in-time copy of all main tables as compressed JSON."""
    cursor = conn.cursor()
    cursor.raw.execute("SET TRANSACTION ISOLATION LEVEL REPEATABLE READ READ ONLY")
    payload = {
        "format": "pharmalab-snapshot", "version": 1,
        "created_at": _now().isoformat(timespec="seconds"), "tables": {},
    }
    for t in _SNAPSHOT_TABLES:
        cursor.execute(f"SELECT * FROM {t} ORDER BY id")
        cols = [d[0] for d in cursor.description]
        payload["tables"][t] = {"columns": cols, "rows": [list(r) for r in cursor.fetchall()]}
    conn.rollback()  # ends the read-only transaction
    return gzip.compress(json.dumps(payload, separators=(",", ":"), default=str).encode("utf-8"))


def create_database_backup(tag_name="manual"):
    """
    Creates a consistent point-in-time snapshot of inventory, ledger, logs,
    schedule and notes, stored in the database (and downloadable from the app).
    """
    safe_tag = re.sub(r"[^A-Za-z0-9_-]+", "_", str(tag_name or "manual")).strip("_")[:40] or "manual"
    base_name = f"lab_backup_{_now().strftime('%Y%m%d_%H%M%S')}_{safe_tag}"
    conn = get_db_connection()
    try:
        data = _build_snapshot_bytes(conn)
        cursor = conn.cursor()
        filename, n = f"{base_name}.json.gz", 1
        while True:
            cursor.execute("SELECT 1 FROM db_snapshots WHERE filename = ?", (filename,))
            if not cursor.fetchone():
                break
            n += 1
            filename = f"{base_name}-{n}.json.gz"
        cursor.execute(
            "INSERT INTO db_snapshots (filename, tag, size_bytes, data) VALUES (?, ?, ?, ?)",
            (filename, safe_tag, len(data), psycopg2.Binary(data)),
        )
        cursor.execute(
            "DELETE FROM db_snapshots WHERE id NOT IN (SELECT id FROM db_snapshots ORDER BY id DESC LIMIT ?)",
            (MAX_SNAPSHOTS_KEPT,),
        )
        conn.commit()
        _invalidate()
        return True, filename, f"db://snapshots/{filename}", round(len(data) / 1024, 2)
    except Exception as e:
        conn.rollback()
        return False, None, str(e), 0
    finally:
        conn.close()


@_cached()
def list_backups():
    """Lists all available database snapshots (newest first)."""
    conn = get_db_connection()
    try:
        df = _read_df(conn, "SELECT filename, size_bytes, created_at FROM db_snapshots ORDER BY id DESC")
    finally:
        conn.close()
    if df.empty:
        return pd.DataFrame()
    return pd.DataFrame({
        "filename": df["filename"],
        "path": "db://snapshots/" + df["filename"],
        "size_kb": (df["size_bytes"].astype(float) / 1024).round(2),
        "created_at": df["created_at"],
    })


def get_backup_bytes(filename):
    """Returns the snapshot file content (bytes, .json.gz) for download, or None."""
    conn = get_db_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("SELECT data FROM db_snapshots WHERE filename = ?", (filename,))
        row = cursor.fetchone()
        return bytes(row[0]) if row else None
    finally:
        conn.close()


def _restore_from_payload_bytes(raw_bytes):
    try:
        payload = json.loads(gzip.decompress(raw_bytes).decode("utf-8"))
        if payload.get("format") != "pharmalab-snapshot" or "tables" not in payload:
            raise ValueError("not a PharmaLab snapshot")
    except Exception:
        return False, "This file is not a valid PharmaLab snapshot (.json.gz)."

    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("LOCK TABLE " + ", ".join(_SNAPSHOT_TABLES) + " IN ACCESS EXCLUSIVE MODE")
        cursor.execute("TRUNCATE TABLE " + ", ".join(_SNAPSHOT_TABLES) + " RESTART IDENTITY")
        restored = {}
        for t in _SNAPSHOT_TABLES:
            info = payload["tables"].get(t)
            if not info:
                continue
            cursor.execute("SELECT column_name FROM information_schema.columns WHERE table_schema = current_schema() AND table_name = ?", (t,))
            existing = {r[0] for r in cursor.fetchall()}
            keep = [i for i, c in enumerate(info["columns"]) if c in existing and re.fullmatch(r"[a-z_][a-z0-9_]*", c)]
            cols = [info["columns"][i] for i in keep]
            rows = [[r[i] for i in keep] for r in info["rows"]]
            if rows and cols:
                psycopg2.extras.execute_values(
                    cursor.raw, f"INSERT INTO {t} ({', '.join(cols)}) VALUES %s", rows, page_size=500)
            cursor.execute(
                f"SELECT setval(pg_get_serial_sequence('{t}', 'id'), COALESCE((SELECT MAX(id) FROM {t}), 1), "
                f"(SELECT MAX(id) IS NOT NULL FROM {t}))")
            restored[t] = len(rows)
        conn.commit()
        _invalidate()
        return True, "Restore complete: " + ", ".join(f"{k}: {v} rows" for k, v in restored.items())
    except Exception as e:
        conn.rollback()
        return False, f"Restore failed, nothing was changed: {e}"
    finally:
        conn.close()


def restore_database_backup(backup_filename):
    """Restores the database from a selected snapshot (a safety snapshot of the current state is taken first)."""
    data = get_backup_bytes(backup_filename)
    if data is None:
        return False, f"Backup file {backup_filename} not found."
    ok, name, _info, _sz = create_database_backup(tag_name="pre_restore_safety")
    if not ok:
        return False, f"Could not create the safety snapshot, restore cancelled: {_info}"
    ok, msg = _restore_from_payload_bytes(data)
    if ok:
        return True, f"Successfully restored database from {backup_filename}. {msg}"
    return False, msg


def restore_from_snapshot_bytes(file_bytes):
    """Restores from a snapshot file the user downloaded earlier (disaster recovery)."""
    ok, name, info, _sz = create_database_backup(tag_name="pre_restore_safety")
    if not ok:
        return False, f"Could not create the safety snapshot, restore cancelled: {info}"
    return _restore_from_payload_bytes(file_bytes)


# --- Full Semester Multi-Sheet Excel Exporter ---

def export_full_semester_archive(semester_name="Fall 2026"):
    """
    Exports a comprehensive, multi-sheet Excel workbook containing:
    1. Inventory (Current Stock & Audit)
    2. Timetable Schedule (All 18 Weeks)
    3. Experiment Logs (All semester experiments)
    4. Stock Transaction Ledger (Complete audit trail)
    5. Notes & Correspondences
    The file is created in a temporary folder and is meant to be downloaded right away.
    """
    timestamp = _now().strftime("%Y%m%d_%H%M%S")
    archive_filename = f"Semester_Archive_{semester_name.replace(' ', '_')}_{timestamp}.xlsx"
    archive_path = os.path.join(TMP_DIR, archive_filename)

    conn = get_db_connection()
    try:
        df_inv = _read_df(conn, "SELECT * FROM inventory ORDER BY item_id")
        df_sched = _read_df(conn, "SELECT * FROM timetable_schedule ORDER BY week_number, day_of_week")
        df_exp = _read_df(conn, "SELECT * FROM experiment_logs ORDER BY timestamp DESC")
        df_tx = _read_df(conn, "SELECT * FROM stock_transactions ORDER BY timestamp DESC")
        df_notes = _read_df(conn, "SELECT * FROM notes_and_tasks ORDER BY created_at DESC")
    finally:
        conn.close()

    with pd.ExcelWriter(archive_path, engine='openpyxl') as writer:
        df_inv.to_excel(writer, sheet_name='Inventory_Master', index=False)
        df_sched.to_excel(writer, sheet_name='Timetable_Preplan', index=False)
        df_exp.to_excel(writer, sheet_name='Experiment_Logs', index=False)
        df_tx.to_excel(writer, sheet_name='Stock_Transactions', index=False)
        df_notes.to_excel(writer, sheet_name='Notes_Tasks', index=False)

    return archive_path, archive_filename


def export_inventory_to_excel(export_file_path=None):
    """Exports current database inventory back to Excel format."""
    conn = get_db_connection()
    try:
        df = _read_df(conn, """
        SELECT item_id AS "Item ID", lab AS "Lab", inventory_group AS "Inventory Group",
               fr_code AS "FR Code", item_name AS "Item Name", item_type AS "Type",
               unit AS "Unit", expiry AS "Expiry", stock_location AS "Stock Location",
               current_stock AS "Current Stock", reorder_level AS "Reorder Level",
               status AS "Status (as per fresh audit)"
        FROM inventory ORDER BY item_id
        """)
    finally:
        conn.close()

    if export_file_path is None:
        export_file_path = os.path.join(TMP_DIR, "Updated_Chemicals_and_Labware.xlsx")

    df.to_excel(export_file_path, index=False)
    return export_file_path, len(df)


# --- Storage Telemetry & Capacity Diagnostics ---

@_cached()
def get_storage_diagnostics():
    """Returns real-time usage figures of the cloud database (one single query)."""
    conn = get_db_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("""
        SELECT
          (SELECT COUNT(*) FROM inventory)          AS total_inventory,
          (SELECT COUNT(*) FROM timetable_schedule) AS total_schedule_slots,
          (SELECT COUNT(*) FROM experiment_logs)    AS total_experiments,
          (SELECT COUNT(*) FROM stock_transactions) AS total_transactions,
          (SELECT COUNT(*) FROM notes_and_tasks)    AS total_notes,
          (SELECT COUNT(*) FROM attachments)        AS total_attachments,
          (SELECT COALESCE(SUM(file_size_bytes), 0) FROM attachments) AS attachments_bytes,
          pg_database_size(current_database())      AS db_bytes
        """)
        r = cursor.fetchone()
    finally:
        conn.close()

    db_bytes = float(r["db_bytes"] or 0)
    attachments_bytes = float(r["attachments_bytes"] or 0)
    limit_bytes = DB_PLAN_LIMIT_GB * (1024 ** 3)
    free_gb = max(0.0, (limit_bytes - db_bytes) / (1024 ** 3))

    return {
        "db_size_kb": round(db_bytes / 1024, 2),
        "db_size_mb": round(db_bytes / (1024 ** 2), 2),
        "attachments_count": int(r["total_attachments"]),
        "attachments_mb": round(attachments_bytes / (1024 ** 2), 2),
        "total_inventory": int(r["total_inventory"]),
        "total_schedule_slots": int(r["total_schedule_slots"]),
        "total_experiments": int(r["total_experiments"]),
        "total_transactions": int(r["total_transactions"]),
        "total_notes": int(r["total_notes"]),
        "free_disk_gb": round(free_gb, 2),
        "total_disk_gb": round(DB_PLAN_LIMIT_GB, 2),
        "journal_mode": "Postgres",
        "theoretical_limit": f"{int(DB_PLAN_LIMIT_GB * 1024)} MB (Supabase Free plan)",
    }
