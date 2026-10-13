"""Read-only preview or cautious append-only merge of CRM contact events into Neon.

Use from your OWN Windows computer with the TEST Neon DATABASE_URL environment variable.
No local/customer data should be committed to GitHub. No credentials are printed.
"""
import argparse
import gzip
import json
import os
import sqlite3
import sys
import tempfile
from collections import Counter
from contextlib import closing
from datetime import datetime, timezone
from pathlib import Path

import psycopg

CORE = ('company_id', 'user_id', 'channel', 'outcome', 'notes', 'contacted_at', 'next_followup_at')
TABLE = 'alrawasi_crm_snapshots'


def canonical(r):
    return tuple(r.get(k) or '' for k in CORE)


def rows_for_table(c, table):
    return [dict(r) for r in c.execute(f'SELECT * FROM "{table}"')]


def checked_export(path):
    with open(path, 'r', encoding='utf-8-sig') as f:
        snap = json.load(f)
    if snap.get('format') != 'alrawasi_crm_json_snapshot' or snap.get('format_version') != 1:
        raise ValueError('Invalid CRM export format/version')
    rows = snap.get('tables', {})
    if not {'companies', 'interactions', 'users', 'templates'} <= set(rows):
        raise ValueError('Missing CRM tables')
    for name in ('companies', 'interactions', 'users', 'templates'):
        if not isinstance(rows[name], list) or len({r['id'] for r in rows[name]}) != len(rows[name]):
            raise ValueError(f'Duplicate/invalid IDs in export: {name}')
        if snap.get('counts', {}).get(name) != len(rows[name]):
            raise ValueError(f'Declared count is incorrect: {name}')
    company_ids = {r['id'] for r in rows['companies']}
    user_ids = {r['id'] for r in rows['users']}
    events = rows['interactions']
    if any(r['company_id'] not in company_ids or (r.get('user_id') is not None and r['user_id'] not in user_ids) for r in events):
        raise ValueError('Export has orphan contact events')
    eids = [str(r.get('remote_event_id') or '').strip() for r in events]
    eids = [x for x in eids if x]
    if len(eids) != len(set(eids)):
        raise ValueError('Duplicate remote_event_id in export')
    return snap


def plan_merge(db, snap):
    current = rows_for_table(db, 'interactions')
    existing_ids = {int(r['id']): r for r in current}
    existing_remote = {str(r['remote_event_id']).strip():r for r in current if r.get('remote_event_id')}
    existing_fingerprint = {canonical(r) for r in current}
    company_ids = {r['id'] for r in rows_for_table(db, 'companies')}
    user_ids = {r['id'] for r in rows_for_table(db, 'users')}
    additions, blocked = [], []
    counts = Counter()
    next_id = max([0]+list(existing_ids)+[int(r['id']) for r in snap['tables']['interactions']])
    proposed_ids = set(existing_ids)
    for row in snap['tables']['interactions']:
        event = dict(row)
        event_id = int(event['id'])
        remote = str(event.get('remote_event_id') or '').strip()
        if event['company_id'] not in company_ids or (event.get('user_id') is not None and event['user_id'] not in user_ids):
            blocked.append((event_id, 'Related company/user missing in Neon (NOT importing company or user records)'))
            continue
        if remote and remote in existing_remote:
            ex = existing_remote[remote]
            if canonical(ex) != canonical(event):
                blocked.append((event_id, 'Same remote_event_id has conflicting contact details'))
            else:
                counts['already_present'] += 1
            continue
        if canonical(event) in existing_fingerprint:
            counts['already_present'] += 1
            continue
        collision = existing_ids.get(event_id)
        if collision:
            # An identical ID in two independently edited SQLite DBs does not
            # necessarily represent the same contact. Never overwrite that row.
            if not remote or not collision.get('remote_event_id'):
                blocked.append((event_id, 'Numeric ID collision without two independent remote IDs'))
                continue
            if (event.get('company_id'),event.get('user_id'),event.get('contacted_at')) == (collision.get('company_id'),collision.get('user_id'),collision.get('contacted_at')):
                blocked.append((event_id, 'Ambiguous edited contact at the same time'))
                continue
            next_id = max(next_id, max(proposed_ids)) + 1
            event['id'] = next_id
            counts['id_remapped'] += 1
        elif event_id in proposed_ids:
            blocked.append((event_id, 'Duplicate incoming local ID'))
            continue
        # Attach department at the time of import when missing; never rewrite
        # department of an existing contact event.
        if not event.get('specialty_at_contact'):
            u = db.execute('SELECT specialty FROM users WHERE id=?', (event.get('user_id'),)).fetchone()
            event['specialty_at_contact'] = (u['specialty'] if u else None) or 'general'
        proposed_ids.add(event['id'])
        if remote:
            existing_remote[remote] = event
        existing_fingerprint.add(canonical(event))
        additions.append(event)
    return {'additions':additions, 'blocked':blocked, 'counts':counts,
            'previous_contacts':len(current), 'unique_pre':len({r['company_id'] for r in current})}


def add_to_db(db, plan):
    cols = {r['name'] for r in db.execute('PRAGMA table_info(interactions)')}
    for col in ('specialty_at_contact','actor_email_at_contact'):
        if col not in cols:
            db.execute(f'ALTER TABLE interactions ADD COLUMN {col} TEXT')
            cols.add(col)
    cols = list(cols)
    touched = set()
    added_ids = set()
    for item in plan['additions']:
        # Let SQLite choose a new safe ID if the exported ID is already in use
        # (the plan has already assigned a nonconflicting number).
        names = [x for x in item if x in cols]
        unknown = set(item) - set(cols)
        if unknown:
            raise RuntimeError(f'Unrecognized columns: {sorted(unknown)}')
        db.execute('INSERT INTO interactions ('+','.join('"'+x+'"' for x in names)+') VALUES ('+','.join('?' for _ in names)+')',
                   [item[x] for x in names])
        added_ids.add(int(item['id']))
        if item.get('outcome') in ('contacted','unreachable'):
            touched.add(item['company_id'])
    # Only update the contact status when the latest event for that company is
    # one we just imported, never for an older event or after a more advanced
    # CRM status (interested/contracted/quotation etc.).
    for company_id in touched:
        latest = db.execute("""SELECT id,outcome,contacted_at FROM interactions
            WHERE company_id=? AND outcome IN ('contacted','unreachable')
            ORDER BY contacted_at DESC,id DESC LIMIT 1""",(company_id,)).fetchone()
        if latest and latest['id'] in added_ids:
            db.execute("""UPDATE companies SET marketing_status=?,updated_at=COALESCE(?,updated_at)
                WHERE id=? AND (marketing_status IS NULL OR marketing_status IN ('new','contacted','unreachable'))""",
                (latest['outcome'],latest['contacted_at'],company_id))
    # sqlite_sequence must be >= the newest explicit ID
    max_id = db.execute('SELECT COALESCE(MAX(id),0) FROM interactions').fetchone()[0]
    db.execute('UPDATE sqlite_sequence SET seq=? WHERE name=? AND seq<?', (max_id,'interactions',max_id))
    db.commit()
    if db.execute('PRAGMA integrity_check').fetchone()[0] != 'ok':
        raise RuntimeError('Merged SQLite integrity check failed')
    if db.execute('SELECT COUNT(*) FROM interactions').fetchone()[0] != plan['previous_contacts'] + len(plan['additions']):
        raise RuntimeError('Merged contact event count incorrect')


def backup_path(directory, revision):
    p = Path(directory).expanduser().resolve()
    project = Path.cwd().resolve()
    if p == project or project in p.parents:
        raise RuntimeError('Backup directory must be OUTSIDE the source-code/GitHub folder')
    p.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now(timezone.utc).strftime('%Y%m%d_%H%M%S_UTC')
    file = p / f'crm-neon-BEFORE-contact-merge-revision-{revision}-{stamp}.db.gz'
    if file.exists():
        raise FileExistsError('Backup file already exists; refuse overwrite')
    return file


def fetch_state(pg):
    with pg.cursor() as cur:
        cur.execute(f'SELECT revision,data FROM {TABLE} WHERE singleton=1')
        row=cur.fetchone()
    if not row:
        raise RuntimeError('No CRM data in Neon. Stop: do not create/seed automatically.')
    return int(row[0]), bytes(row[1])


def run(args):
    if not os.environ.get('DATABASE_URL'):
        raise RuntimeError('DATABASE_URL missing; set the TEST Neon URL securely in PowerShell')
    snap = checked_export(args.export)
    with psycopg.connect(os.environ['DATABASE_URL'], connect_timeout=15) as pg:
        rev, original = fetch_state(pg)
    with tempfile.TemporaryDirectory(prefix='crm-neon-safe-merge-') as temp:
        local = Path(temp)/'live.db'
        local.write_bytes(gzip.decompress(original))
        with closing(sqlite3.connect(local)) as sql:
            sql.row_factory=sqlite3.Row
            if sql.execute('PRAGMA integrity_check').fetchone()[0] != 'ok':
                raise RuntimeError('Existing Neon SQLite image is corrupt; no changes allowed')
            plan = plan_merge(sql, snap)
            print('NEON REVISION:', rev)
            print('NEON CURRENT CONTACT EVENTS:', plan['previous_contacts'])
            print('JSON EXPORT CONTACT EVENTS:', len(snap['tables']['interactions']))
            print('ALREADY PRESENT (remote ID or exact contact):', plan['counts']['already_present'])
            print('MISSING CONTACT EVENTS TO ADD:', len(plan['additions']))
            print('LOCAL ID REMAPS (preserve both events):', plan['counts']['id_remapped'])
            print('BLOCKING CONFLICTS:', len(plan['blocked']))
            for event_id,reason in plan['blocked'][:20]:
                print('  Conflict, export ID',event_id,':',reason)
            print('NEON UPDATED BY THIS COMMAND:', 'NO' if not args.apply else 'PENDING CONFIRMATION')
            if not args.apply:
                return 0
            if plan['blocked']:
                print('STOP: resolve blocking conflicts before attempting write; nothing changed.')
                return 2
            if not plan['additions']:
                print('Nothing to merge. Live Neon was not changed.')
                return 0
            if not args.backup_dir:
                raise RuntimeError('--backup-dir is required when using --apply; use a PRIVATE folder outside GitHub')
            print('WARNING: Pause ALL edits in Render CRM (both environments) until restarted. Never merge while users are writing.')
            answer = input('Type MERGE TEST NEON to back up and apply: ').strip()
            if answer != 'MERGE TEST NEON':
                print('Canceled safely. No Neon changes.')
                return 2
            path = backup_path(args.backup_dir, rev)
            path.write_bytes(original)
            try: os.chmod(path,0o600)
            except OSError: pass
            print('PRIVATE BACKUP SAVED:',path)
            print('BACKUP SIZE BYTES:',len(original))
            add_to_db(sql, plan)
        merged = gzip.compress(local.read_bytes(), compresslevel=5)
        # Compare-and-swap revision prevents overwriting an intervening update.
        with psycopg.connect(os.environ['DATABASE_URL'], connect_timeout=15) as pg:
            with pg.cursor() as cur:
                cur.execute(f'UPDATE {TABLE} SET data=%s, revision=revision+1, updated_at=NOW() WHERE singleton=1 AND revision=%s', (merged,rev))
                if cur.rowcount!=1:
                    pg.rollback()
                    print('STOP: Neon changed during merge. Backup remains available. No write applied.')
                    return 2
            pg.commit()
        print('MERGE COMPLETE: +',len(plan['additions']),'contact events; all existing records preserved.')
        print('EXPECTED CONTACT EVENTS:',plan['previous_contacts']+len(plan['additions']))
        print('IMPORTANT: Restart Render service now, before allowing staff to make edits.')
    return 0


def main():
    parser=argparse.ArgumentParser(description='Preview or append-only merge of a CRM export into existing Neon snapshot')
    parser.add_argument('export',help='Local full JSON export downloaded from original CRM')
    parser.add_argument('--apply',action='store_true',help='Back up and merge only after an interactive confirmation')
    parser.add_argument('--backup-dir',help='Private local directory for Neon backup, REQUIRED with --apply')
    args=parser.parse_args()
    try: return run(args)
    except Exception as e:
        print('STOP (safe failure):',type(e).__name__, str(e).split('\n')[0],file=sys.stderr)
        return 2

if __name__=='__main__':
    sys.exit(main())
