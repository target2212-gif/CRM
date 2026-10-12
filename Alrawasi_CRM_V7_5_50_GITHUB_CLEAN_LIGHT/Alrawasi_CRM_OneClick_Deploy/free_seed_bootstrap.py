import os, shutil, sqlite3, gzip

BASE=os.path.dirname(os.path.abspath(__file__))
SEED_DB=os.path.join(BASE,"crm.db")
SEED_GZ=os.path.join(BASE,"crm.db.gz")
DATA_DIR=os.environ.get("DATA_DIR") or "/tmp/alrawasi-data"
LIVE_DB=os.path.join(DATA_DIR,"crm.db")
REQUIRED={"users","companies","interactions","templates"}

def healthy(path):
    if not os.path.isfile(path) or os.path.getsize(path) < 1024:
        return False
    try:
        c=sqlite3.connect(path)
        try:
            names={r[0] for r in c.execute("SELECT name FROM sqlite_master WHERE type='table'").fetchall()}
            return REQUIRED.issubset(names) and c.execute("PRAGMA quick_check").fetchone()[0] == "ok"
        finally: c.close()
    except Exception:
        return False

os.makedirs(DATA_DIR, exist_ok=True)

# In Neon mode, prepare_runtime_db.py has already restored the full database.
# The source ZIP intentionally contains no crm.db or crm.db.gz with customer data.
# Never fall back to an older packaged seed when Neon is configured.
if os.environ.get('DATABASE_URL'):
    if not healthy(LIVE_DB):
        raise RuntimeError('Neon restore missing/invalid; refusing to start without CRM data')
else:
    # Local development mode only: initialize from a separately supplied seed.
    if not healthy(SEED_DB):
        if not os.path.isfile(SEED_GZ):
            raise RuntimeError('Missing local crm.db.gz: supply a local seed outside GitHub')
        tmp=SEED_DB+'.init'
        if os.path.exists(tmp): os.remove(tmp)
        try:
            with gzip.open(SEED_GZ,'rb') as src, open(tmp,'wb') as out:
                shutil.copyfileobj(src,out)
            if not healthy(tmp): raise RuntimeError('Invalid decompressed seed database')
            os.replace(tmp,SEED_DB)
        finally:
            if os.path.exists(tmp): os.remove(tmp)

    # Critical rule: NEVER overwrite a healthy live DB. Seed only a missing/broken instance.
    if not healthy(LIVE_DB):
        tmp=LIVE_DB+'.init'
        if os.path.exists(tmp): os.remove(tmp)
        try:
            shutil.copy2(SEED_DB,tmp)
            if not healthy(tmp): raise RuntimeError('Invalid runtime seed copy')
            os.replace(tmp,LIVE_DB)
        finally:
            if os.path.exists(tmp): os.remove(tmp)

if not healthy(LIVE_DB):
    raise RuntimeError('Runtime database failed integrity/core-table validation')

print("V7.5.37 Free Data Guard OK:", LIVE_DB, flush=True)
# Import Flask only after the runtime DB is guaranteed valid.
import app as crm_app
flask_app=crm_app.app
