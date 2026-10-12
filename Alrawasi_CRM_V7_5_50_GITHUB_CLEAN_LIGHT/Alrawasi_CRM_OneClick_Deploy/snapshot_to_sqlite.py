"""Offline conversion of an Alrawasi CRM JSON export into a checked SQLite image.

No PostgreSQL connection is used by this module. It refuses to overwrite an existing
file and preserves row identifiers, password hashes, and contact metadata.
"""
import argparse
import json
import os
import sqlite3
import sys
import tempfile
from collections import Counter
from pathlib import Path

BASE = Path(__file__).resolve().parent
TABLES = ('companies', 'interactions', 'templates', 'users')
# Export has these two historic interaction columns that the bundled early seed
# does not have; they must never silently disappear during restore.
EXTRA_COLUMNS = {'interactions': {'specialty_at_contact':'TEXT', 'actor_email_at_contact':'TEXT'}}


def convert_snapshot(json_path, dest, min_companies=15000):
    json_path, dest = Path(json_path).resolve(), Path(dest).resolve()
    if not json_path.is_file():
        raise ValueError('Cannot find supplied JSON snapshot')
    if dest.exists():
        raise FileExistsError('Destination already exists. Refusing to overwrite: ' + str(dest))
    with json_path.open('r', encoding='utf-8') as f:
        snapshot = json.load(f)
    if not isinstance(snapshot,dict) or snapshot.get('format') != 'alrawasi_crm_json_snapshot' or snapshot.get('format_version') != 1:
        raise ValueError('Unexpected snapshot format/version')
    data = snapshot.get('tables')
    if not isinstance(data,dict) or set(data) != set(TABLES):
        raise ValueError('Missing or unexpected snapshot tables')
    if not all(isinstance(data[k],list) for k in TABLES):
        raise ValueError('Snapshot contains invalid table rows')
    if len(data['companies']) < min_companies or len(data['users']) < 1:
        raise ValueError('Snapshot counts too small. Not safe to initialize')
    for table in TABLES:
        values=data[table]
        if any(not isinstance(r,dict) or not isinstance(r.get('id'),int) for r in values):
            raise ValueError('Invalid row identifier in ' + table)
        ids=[r['id'] for r in values]
        if len(set(ids)) != len(ids):
            raise ValueError('Duplicate row identifier in ' + table)
        declared=snapshot.get('counts',{}).get(table)
        if declared != len(values):
            raise ValueError('Count mismatch in ' + table)
    companies={r['id'] for r in data['companies']}
    users={r['id'] for r in data['users']}
    if any(r.get('company_id') not in companies for r in data['interactions']):
        raise ValueError('Interaction with a missing company ID')
    if any(r.get('user_id') is not None and r.get('user_id') not in users for r in data['interactions']):
        raise ValueError('Interaction with a missing user ID')
    if any(r.get('assigned_user_id') is not None and r.get('assigned_user_id') not in users for r in data['companies']):
        raise ValueError('Company assigned to a missing user ID')
    for table,field in (('interactions','remote_event_id'),('companies','source_id'),('companies','muqawil_key'),('companies','crm_number')):
        vals=[r.get(field) for r in data[table] if r.get(field) not in (None,'')]
        if len(vals) != len(set(vals)):
            raise ValueError('Duplicate unique value for ' + table + '.' + field)

    dest.parent.mkdir(parents=True,exist_ok=True)
    fd,tmp_name=tempfile.mkstemp(prefix='.alrawasi-json-import-',suffix='.db',dir=str(dest.parent))
    os.close(fd)
    con = None
    output_opened = False
    completed = False
    try:
        con=sqlite3.connect(tmp_name)
        con.executescript((BASE/'crm_schema.sql').read_text(encoding='utf8'))
        colmaps={}
        for t in TABLES:
            columns=[r[1] for r in con.execute(f'PRAGMA table_info("{t}")')]
            incoming=set().union(*(set(r) for r in data[t])) if data[t] else set()
            unknown=incoming-set(columns)
            for col in sorted(unknown):
                typename=EXTRA_COLUMNS.get(t,{}).get(col)
                if not typename:
                    raise ValueError(f'Unknown column in {t}: {col}. Nothing imported.')
                con.execute(f'ALTER TABLE "{t}" ADD COLUMN "{col}" {typename}')
            colmaps[t]=[r[1] for r in con.execute(f'PRAGMA table_info("{t}")')]
            if data[t] and 'id' not in incoming:
                raise ValueError('Missing primary ID in ' + t)

        con.execute('BEGIN IMMEDIATE')
        for t in TABLES:
            cols=[c for c in colmaps[t] if data[t] and c in data[t][0]]
            if any(set(row)-set(cols) for row in data[t]):
                raise ValueError('Inconsistent fields between rows in ' + t)
            if not cols:
                continue
            cols_sql=','.join('"'+c+'"' for c in cols)
            placeholders=','.join('?' for _ in cols)
            con.executemany(f'INSERT INTO "{t}" ({cols_sql}) VALUES ({placeholders})',
                            ([r.get(c) for c in cols] for r in data[t]))
            mx=max((r['id'] for r in data[t]),default=0)
            con.execute('DELETE FROM sqlite_sequence WHERE name=?',(t,))
            con.execute('INSERT INTO sqlite_sequence(name,seq) VALUES(?,?)',(t,mx))
        con.commit()
        if con.execute('PRAGMA integrity_check').fetchone()[0] != 'ok':
            raise RuntimeError('SQLite integrity_check failed')
        for t in TABLES:
            count=con.execute(f'SELECT count(*) FROM "{t}"').fetchone()[0]
            if count!=len(data[t]):
                raise RuntimeError(f'Conversion lost rows from {t}')
        # Verify essential extra interaction metadata is present and complete.
        for col in EXTRA_COLUMNS['interactions']:
            if col in data['interactions'][0]:
                n=con.execute(f'SELECT count(*) FROM interactions WHERE "{col}" IS NOT NULL').fetchone()[0]
                want=sum(r.get(col) is not None for r in data['interactions'])
                if n != want:
                    raise RuntimeError('Contact metadata changed unexpectedly')
        con.close();con=None
        # Destination must not exist, even from concurrent operations.
        with open(tmp_name,'rb') as src, open(dest,'xb') as dst:
            output_opened = True
            while True:
                block=src.read(1024*1024)
                if not block: break
                dst.write(block)
            dst.flush()
            os.fsync(dst.fileno())
        completed = True
        return {t:len(data[t]) for t in TABLES}
    except Exception:
        if con: con.close()
        if output_opened and not completed and dest.exists():
            dest.unlink()
        raise
    finally:
        if os.path.exists(tmp_name): os.remove(tmp_name)


def main():
    parser=argparse.ArgumentParser(description='Safely convert a trusted Alrawasi JSON CRM snapshot to SQLite')
    parser.add_argument('input_json')
    parser.add_argument('output_db')
    args=parser.parse_args()
    counts=convert_snapshot(args.input_json,args.output_db)
    print('VERIFIED:', ', '.join(f'{t}={counts[t]}' for t in TABLES), '| output:',args.output_db)

if __name__=='__main__':
    main()
