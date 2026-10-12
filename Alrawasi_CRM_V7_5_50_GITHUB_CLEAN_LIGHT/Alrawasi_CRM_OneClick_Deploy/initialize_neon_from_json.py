"""Safely initialize an EMPTY Neon snapshot store from a vetted CRM JSON export.

Never replace an existing snapshot. The file is converted locally; neither its
path nor its contents should be committed to GitHub or logged in Render.
"""
import argparse
import os
import sqlite3
import tempfile
from pathlib import Path


def main():
    parser=argparse.ArgumentParser(description='One-time first import from live CRM JSON into empty Neon')
    parser.add_argument('input_json',help='Downloaded recent JSON snapshot; keep outside the Git repository')
    args=parser.parse_args()
    if not os.environ.get('DATABASE_URL'):
        raise SystemExit('STOP: DATABASE_URL not set. Use a private EMPTY Neon test database.')
    import neon_snapshot
    from snapshot_to_sqlite import convert_snapshot
    # Check existence before expensive local work. The publish operation ALSO
    # uses atomic INSERT ON CONFLICT DO NOTHING to prevent a race overwrite.
    with neon_snapshot._connection() as pg:
        with pg.cursor() as cur:
            cur.execute(neon_snapshot.TABLE_SQL)
            cur.execute('SELECT 1 FROM alrawasi_crm_snapshots WHERE singleton=1')
            if cur.fetchone():
                raise SystemExit('STOP: Neon already contains a CRM snapshot. No data changed.')
    with tempfile.TemporaryDirectory(prefix='alrawasi-seed-') as temp:
        db_path=Path(temp)/'latest_crm.db'
        counts=convert_snapshot(args.input_json,db_path)
        if not neon_snapshot._healthy(str(db_path)):
            raise SystemExit('STOP: Converted database failed safety check.')
        with sqlite3.connect(db_path) as sqlite_db:
            neon_snapshot.publish(sqlite_db, allow_initial=True)
    print('Imported to previously empty Neon snapshot store:',counts)
    print('IMPORTANT: Test deployment and exported record counts before production cutover.')

if __name__=='__main__':
    main()
