"""Transitional durable snapshot store for the legacy SQLite CRM.

Not a native PostgreSQL schema migration. Single instance/single worker only.
"""
import gzip
import os
import sqlite3
import tempfile
from pathlib import Path

import psycopg

URL = (os.environ.get('DATABASE_URL') or '').strip()
TABLE_SQL = '''CREATE TABLE IF NOT EXISTS alrawasi_crm_snapshots (
  singleton INTEGER PRIMARY KEY CHECK(singleton = 1),
  data BYTEA NOT NULL,
  updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  revision BIGINT NOT NULL DEFAULT 1
)'''


def _connection():
    if not URL:
        raise RuntimeError('DATABASE_URL required for Neon snapshot mode')
    return psycopg.connect(URL, connect_timeout=15)


def _healthy(path):
    if not os.path.isfile(path):
        return False
    with sqlite3.connect(path) as con:
        tables = {r[0] for r in con.execute("SELECT name FROM sqlite_master WHERE type='table'")}
        return {'users', 'companies', 'interactions', 'templates'} <= tables and con.execute('PRAGMA quick_check').fetchone()[0] == 'ok'


def restore(path, seed_path=None):
    """Fail closed rather than silently falling back to the seed on loss of Neon."""
    if not URL:
        raise RuntimeError('DATABASE_URL is required to start this deployment')
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with _connection() as pg:
        with pg.cursor() as cur:
            cur.execute(TABLE_SQL)
            cur.execute('SELECT data FROM alrawasi_crm_snapshots WHERE singleton=1')
            row = cur.fetchone()
    if row is None:
        if os.getenv('ALLOW_SEED_INITIALIZATION') != 'true':
            raise RuntimeError('Neon has no CRM snapshot. Restore/migrate the live database first; set ALLOW_SEED_INITIALIZATION=true ONLY for first-time seed initialization.')
        if not seed_path or not _healthy(seed_path):
            raise RuntimeError('Cannot initialize Neon: seed database is missing or corrupt')
        with sqlite3.connect(seed_path) as seed:
            publish(seed, allow_initial=True)
        row = None
        with _connection() as pg:
            with pg.cursor() as cur:
                cur.execute('SELECT data FROM alrawasi_crm_snapshots WHERE singleton=1')
                row = cur.fetchone()
        if row is None:
            raise RuntimeError('Neon initialization failed')
    raw = gzip.decompress(bytes(row[0]))
    fd, tmp = tempfile.mkstemp(prefix='.crm-neon-', suffix='.db', dir=os.path.dirname(path))
    try:
        with os.fdopen(fd, 'wb') as f:
            f.write(raw)
            f.flush()
            os.fsync(f.fileno())
        if not _healthy(tmp):
            raise RuntimeError('Neon snapshot failed SQLite integrity check')
        os.replace(tmp, path)
    finally:
        if os.path.exists(tmp):
            os.remove(tmp)
    print('CRM: restored complete SQLite snapshot from Neon', flush=True)


def publish(db_conn, allow_initial=False):
    """Publish a committed, consistent image of the complete database."""
    fd, tmp = tempfile.mkstemp(prefix='crm-checkpoint-', suffix='.db')
    os.close(fd)
    try:
        target = sqlite3.connect(tmp)
        try:
            db_conn.backup(target)
        finally:
            target.close()
        if not _healthy(tmp):
            raise RuntimeError('Refusing to persist an invalid database snapshot')
        data = gzip.compress(Path(tmp).read_bytes(), compresslevel=5)
        with _connection() as pg:
            with pg.cursor() as cur:
                cur.execute(TABLE_SQL)
                if allow_initial:
                    cur.execute('''INSERT INTO alrawasi_crm_snapshots(singleton,data)
                                   VALUES(1,%s) ON CONFLICT(singleton) DO NOTHING''', (data,))
                    if cur.rowcount != 1:
                        raise RuntimeError('Neon already contains CRM data. Import refused; no overwrite.')
                else:
                    cur.execute('SELECT revision FROM alrawasi_crm_snapshots WHERE singleton=1 FOR UPDATE')
                    if cur.fetchone() is None:
                        raise RuntimeError('Neon snapshot disappeared; refusing to overwrite')
                    cur.execute('''UPDATE alrawasi_crm_snapshots SET data=%s,
                                   updated_at=NOW(), revision=revision+1 WHERE singleton=1''', (data,))
    finally:
        if os.path.exists(tmp):
            os.remove(tmp)
