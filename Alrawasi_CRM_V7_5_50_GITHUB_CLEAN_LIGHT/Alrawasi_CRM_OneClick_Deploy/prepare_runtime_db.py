import os, shutil, sqlite3, gzip, sys

BASE=os.path.dirname(os.path.abspath(__file__))
SEED_DB=os.path.join(BASE,"crm.db")
SEED_GZ=os.path.join(BASE,"crm.db.gz")
DATA_DIR=os.environ.get("DATA_DIR") or "/tmp/alrawasi-data"
LIVE_DB=os.path.join(DATA_DIR,"crm.db")
REQUIRED={"users","companies","interactions","templates"}

def tables(path):
    if not os.path.isfile(path) or os.path.getsize(path)==0:
        return set()
    try:
        c=sqlite3.connect(path)
        try:
            return {r[0] for r in c.execute(
                "SELECT name FROM sqlite_master WHERE type='table'"
            ).fetchall()}
        finally:
            c.close()
    except Exception:
        return set()

os.makedirs(DATA_DIR, exist_ok=True)
if (os.getenv('RENDER') or os.getenv('RENDER_SERVICE_ID')) and not os.environ.get('DATABASE_URL'):
    raise RuntimeError('DATABASE_URL missing: refusing ephemeral Render startup that risks data loss')

# Neon is authoritative. In cloud mode NO customer database or seed is shipped
# with application code. Initialize Neon securely from the supplied JSON snapshot
# using initialize_neon_from_json.py on an administrator machine BEFORE deploying.
if os.environ.get('DATABASE_URL'):
    import neon_snapshot
    # Never fall back to an older seed or silently create an empty database.
    if os.getenv('ALLOW_SEED_INITIALIZATION','false').lower() == 'true':
        raise RuntimeError('ALLOW_SEED_INITIALIZATION must be false for data-safe deployment')
    neon_snapshot.restore(LIVE_DB, None)
    print('Neon full-database restore OK', flush=True)
    sys.exit(0)

# Local-only operation without Neon requires the separately managed seed.
if not REQUIRED.issubset(tables(SEED_DB)):
    if not os.path.isfile(SEED_GZ):
        raise RuntimeError('Missing local crm.db: run snapshot_to_sqlite.py first')
    tmp=SEED_DB+'.init'
    try:
        if os.path.exists(tmp): os.remove(tmp)
        with gzip.open(SEED_GZ,'rb') as src, open(tmp,'wb') as dest:
            shutil.copyfileobj(src,dest)
        if not REQUIRED.issubset(tables(tmp)):
            raise RuntimeError('Invalid local seed database')
        os.replace(tmp,SEED_DB)
    finally:
        if os.path.exists(tmp): os.remove(tmp)

# Free Render runtime DB: initialize only when absent/incomplete.
if not REQUIRED.issubset(tables(LIVE_DB)):
    tmp=LIVE_DB+".init"
    try:
        if os.path.exists(tmp): os.remove(tmp)
        shutil.copy2(SEED_DB,tmp)
        if not REQUIRED.issubset(tables(tmp)):
            raise RuntimeError("Invalid runtime seed copy")
        os.replace(tmp,LIVE_DB)
    finally:
        if os.path.exists(tmp):
            os.remove(tmp)

missing=REQUIRED-tables(LIVE_DB)
if missing:
    raise RuntimeError("Runtime DB missing tables: "+", ".join(sorted(missing)))

print("Database initialization OK:", LIVE_DB, flush=True)
