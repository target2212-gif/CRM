"""Run locally with DATABASE_URL set, using a vetted live SQLite backup.
Never overwrites an existing Neon snapshot.
"""
import os, sqlite3, sys
import psycopg
import neon_snapshot

if len(sys.argv) != 2:
    raise SystemExit('Usage: python initialize_neon_from_live.py /path/to/live-crm.db')
source = sys.argv[1]
if not neon_snapshot._healthy(source):
    raise SystemExit('STOP: invalid/incomplete SQLite backup')
with neon_snapshot._connection() as pg:
    with pg.cursor() as cur:
        cur.execute(neon_snapshot.TABLE_SQL)
        cur.execute('SELECT 1 FROM alrawasi_crm_snapshots WHERE singleton=1')
        if cur.fetchone():
            raise SystemExit('STOP: Neon already has a snapshot. No overwrite performed.')
with sqlite3.connect(source) as sqlite_db:
    neon_snapshot.publish(sqlite_db, allow_initial=True)
print('Full CRM database imported to Neon snapshot store. Verify before deployment.')
