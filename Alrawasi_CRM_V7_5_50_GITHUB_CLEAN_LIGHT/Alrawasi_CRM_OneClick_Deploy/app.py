import urllib.request, urllib.error
import tempfile
from flask import Flask,request,redirect,url_for,render_template,session,flash,jsonify,abort,Response, send_file
import hashlib, sqlite3, os, urllib.parse, csv, io, json, re, secrets, uuid
from datetime import datetime
from openpyxl import load_workbook, Workbook
from functools import wraps
from urllib.parse import urljoin, quote_plus

APP_VERSION=os.environ.get('APP_VERSION') or open(os.path.join(os.path.dirname(__file__),'VERSION'),encoding='utf-8').read().strip()
import requests
from bs4 import BeautifulSoup
try:
 import psycopg
except Exception:
 psycopg=None

def check_password_hash(stored,pw):
 try:
  salt,hexd=stored.split('$',1); return hashlib.pbkdf2_hmac('sha256',pw.encode(),salt.encode(),200000).hex()==hexd
 except: return False
def generate_password_hash(pw):
 import secrets; salt=secrets.token_hex(16); return salt+'$'+hashlib.pbkdf2_hmac('sha256',pw.encode(),salt.encode(),200000).hex()

app=Flask(__name__); app.secret_key=os.environ.get('SECRET_KEY','change-this-secret')
ADMIN_RECOVERY_KEY=(os.environ.get('ADMIN_RECOVERY_KEY') or '').strip()
CONTACTS_DATABASE_URL=(os.environ.get('CONTACTS_DATABASE_URL') or '').strip()

@app.template_filter('clean_company_name')
def clean_company_name(value):
 value=str(value or '').strip()
 return re.sub(r'\s*#\d+\s*$','',value).strip()

BASE=os.path.dirname(__file__)
SEED_DB=os.path.join(BASE,'crm.db')
IS_RENDER=bool(os.environ.get('RENDER') or os.environ.get('RENDER_SERVICE_ID'))
DATA_DIR=os.environ.get('DATA_DIR') or ('/var/data' if IS_RENDER else BASE)
os.makedirs(DATA_DIR, exist_ok=True)
DB=os.path.join(DATA_DIR,'crm.db')
DATA_LOCK=os.path.join(DATA_DIR,'.alrawasi_data_initialized')
# V7.5.37: Free-Render data guard. The runtime DB always wins while it exists.
# Deploy/update code must never copy over a healthy runtime DB.
FREE_EPHEMERAL_RENDER = IS_RENDER and os.path.abspath(DATA_DIR) != '/var/data'
REQUIRED_CORE_TABLES={'users','companies','interactions','templates'}

def _db_has_core_tables(path):
 try:
  if not os.path.isfile(path) or os.path.getsize(path) < 1024: return False
  c=sqlite3.connect(path)
  try:
   names={r[0] for r in c.execute("SELECT name FROM sqlite_master WHERE type='table'").fetchall()}
   return REQUIRED_CORE_TABLES.issubset(names) and c.execute('PRAGMA quick_check').fetchone()[0]=='ok'
  finally: c.close()
 except Exception: return False
# Storage modes:
# - /var/data = paid Render persistent disk (strict Data-Lock).
# - any other DATA_DIR = free/ephemeral mode. The packaged crm.db is copied only
#   when the runtime database does not exist. This avoids importing app against
#   an empty SQLite file before the core tables exist.
PERSISTENT_RENDER = IS_RENDER and os.path.abspath(DATA_DIR) == '/var/data'
if DB != SEED_DB and not _db_has_core_tables(DB):
 if PERSISTENT_RENDER and os.path.exists(DATA_LOCK):
  raise RuntimeError('DATA-LOCK: live /var/data/crm.db is missing; refusing to recreate/replace it automatically')
 if not os.path.isfile(SEED_DB):
  raise RuntimeError('Seed database is missing: '+SEED_DB)
 import shutil
 shutil.copy2(SEED_DB, DB)
 with open(DATA_LOCK,'w',encoding='utf-8') as _f: _f.write(datetime.now().isoformat())
elif DB != SEED_DB and os.path.exists(DB) and not os.path.exists(DATA_LOCK):
 with open(DATA_LOCK,'w',encoding='utf-8') as _f: _f.write(datetime.now().isoformat())
STATUS={'new':'جديد','contacted':'تم التواصل','interested':'مهتم','quote_requested':'طلب عرض سعر','quote_sent':'تم إرسال عرض السعر','followup':'متابعة','contracted':'تم التعاقد','not_interested':'غير مهتم','unreachable':'لا يمكن التواصل','postponed':'مؤجل'}
STATUS_CLASS={k:'status-'+k.replace('_','-') for k in STATUS}
SPECIALTIES={'contracting':'خدمات المقاولين والشهادات','advertising':'الدعاية والإعلان','general':'خدمات عامة'}
DEFAULT_MESSAGES={
 'contracting':'السلام عليكم، معك {rep_name} من تلال الرواسي. نقدم خدمات تصنيف شركات المقاولات، شهادة المحتوى المحلي، وشهادات الأيزو. يسعدنا خدمة {company_name} وتزويدكم بالتفاصيل والعرض المناسب.',
 'advertising':'السلام عليكم، معك {rep_name} من تلال الرواسي للدعاية والإعلان. نقدم لوحات المحلات التجارية، الاستيكرات، البنرات، بوثات المعارض، الهوية والمطبوعات. يسعدنا خدمة {company_name} وتقديم عرض مناسب لاحتياجكم.',
 'general':'السلام عليكم، معك {rep_name} من تلال الرواسي. يسعدنا التعرف على احتياج {company_name} وتقديم خدماتنا والحلول المناسبة لكم.'}

# Transitional persistence mode: write-through complete SQLite snapshots to Neon.
# This is deliberately NOT advertised as a native PostgreSQL CRM migration.
_NEON_SNAPSHOT_MODE = bool(os.environ.get('DATABASE_URL'))
_NEON_WRITE_FAILED = False
if _NEON_SNAPSHOT_MODE:
 import neon_snapshot

class DurableConnection(sqlite3.Connection):
 def commit(self):
  global _NEON_WRITE_FAILED
  if _NEON_WRITE_FAILED:
   raise RuntimeError('Neon persistence previously failed; restart/reconcile before writing again')
  super().commit()
  if _NEON_SNAPSHOT_MODE and self.total_changes:
   try:
    neon_snapshot.publish(self)
   except Exception:
    _NEON_WRITE_FAILED = True
    app.logger.exception('CRITICAL: Neon full-data checkpoint failed; writes are now blocked')
    raise

def db():
 x=sqlite3.connect(DB, timeout=30, factory=DurableConnection)
 x.row_factory=sqlite3.Row
 x.execute('PRAGMA busy_timeout=30000')
 return x

def ensure_schema():
 # Gunicorn starts multiple workers at the same time. Serialize schema migration
 # so two workers cannot try to add the same SQLite column simultaneously.
 d=db()
 try:
  d.execute('BEGIN IMMEDIATE')
  cols={r['name'] for r in d.execute('PRAGMA table_info(users)').fetchall()}
  if 'specialty' not in cols: d.execute("ALTER TABLE users ADD COLUMN specialty TEXT DEFAULT 'general'")
  if 'whatsapp_template' not in cols: d.execute("ALTER TABLE users ADD COLUMN whatsapp_template TEXT")
  if 'appearance_default' not in cols: d.execute("ALTER TABLE users ADD COLUMN appearance_default TEXT")
  icols={r['name'] for r in d.execute('PRAGMA table_info(interactions)').fetchall()}
  if 'user_name_snapshot' not in icols: d.execute("ALTER TABLE interactions ADD COLUMN user_name_snapshot TEXT")
  if 'remote_event_id' not in icols: d.execute("ALTER TABLE interactions ADD COLUMN remote_event_id TEXT")
  # Preserve fields supplied by the latest 2026-10-12 live JSON export.
  if 'specialty_at_contact' not in icols: d.execute("ALTER TABLE interactions ADD COLUMN specialty_at_contact TEXT")
  if 'actor_email_at_contact' not in icols: d.execute("ALTER TABLE interactions ADD COLUMN actor_email_at_contact TEXT")
  d.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_interactions_remote_event ON interactions(remote_event_id) WHERE remote_event_id IS NOT NULL AND remote_event_id<>''")
  d.execute("UPDATE interactions SET user_name_snapshot=(SELECT name FROM users WHERE users.id=interactions.user_id) WHERE (user_name_snapshot IS NULL OR user_name_snapshot='') AND user_id IS NOT NULL")
  # Backfill only missing historical specialty; never rewrite a captured department when
  # staff later switch departments. Covers the 2026-10-12 JSON export's NULL values.
  d.execute("""UPDATE interactions SET specialty_at_contact=(SELECT u.specialty FROM users u WHERE u.id=interactions.user_id)
               WHERE (specialty_at_contact IS NULL OR TRIM(specialty_at_contact)='')
                 AND user_id IN (SELECT id FROM users WHERE specialty IS NOT NULL)""")
  ccols={r['name'] for r in d.execute('PRAGMA table_info(companies)').fetchall()}
  if 'business_area' not in ccols: d.execute("ALTER TABLE companies ADD COLUMN business_area TEXT DEFAULT 'contracting'")
  if 'muqawil_member_no' not in ccols: d.execute("ALTER TABLE companies ADD COLUMN muqawil_member_no TEXT")
  if 'source_url' not in ccols: d.execute("ALTER TABLE companies ADD COLUMN source_url TEXT")
  if 'address' not in ccols: d.execute("ALTER TABLE companies ADD COLUMN address TEXT")
  if 'data_source' not in ccols: d.execute("ALTER TABLE companies ADD COLUMN data_source TEXT")
  if 'imported_at' not in ccols: d.execute("ALTER TABLE companies ADD COLUMN imported_at TEXT")
  if 'classification_status' not in ccols: d.execute("ALTER TABLE companies ADD COLUMN classification_status TEXT")
  if 'classification_grade' not in ccols: d.execute("ALTER TABLE companies ADD COLUMN classification_grade TEXT")
  if 'latitude' not in ccols: d.execute("ALTER TABLE companies ADD COLUMN latitude REAL")
  if 'longitude' not in ccols: d.execute("ALTER TABLE companies ADD COLUMN longitude REAL")
  if 'logo_url' not in ccols: d.execute("ALTER TABLE companies ADD COLUMN logo_url TEXT")
  if 'google_place_id' not in ccols: d.execute("ALTER TABLE companies ADD COLUMN google_place_id TEXT")
  if 'phone_source' not in ccols: d.execute("ALTER TABLE companies ADD COLUMN phone_source TEXT")
  if 'phone_verified_at' not in ccols: d.execute("ALTER TABLE companies ADD COLUMN phone_verified_at TEXT")
  if 'phone_match_name' not in ccols: d.execute("ALTER TABLE companies ADD COLUMN phone_match_name TEXT")
  if 'phone_match_address' not in ccols: d.execute("ALTER TABLE companies ADD COLUMN phone_match_address TEXT")
  for col,typ in [('muqawil_key','TEXT'),('muqawil_internal_id','TEXT'),('muqawil_contractor_type','TEXT'),('company_size','TEXT'),('account_status','TEXT'),('training_hours','INTEGER'),('main_contractor_count','INTEGER'),('subcontractor_count','INTEGER'),('source_region_id','TEXT'),('membership_type','TEXT'),('membership_since','TEXT'),('muqawil_email','TEXT'),('muqawil_phone_landline','TEXT'),('muqawil_phone_mobile','TEXT'),('muqawil_address','TEXT'),('muqawil_city','TEXT'),('muqawil_region','TEXT'),('muqawil_enriched_at','TEXT')]:
   if col not in ccols: d.execute(f"ALTER TABLE companies ADD COLUMN {col} {typ}")
  d.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_companies_muqawil_key ON companies(muqawil_key) WHERE muqawil_key IS NOT NULL")
  d.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_companies_crm_number ON companies(crm_number) WHERE crm_number IS NOT NULL")
  d.execute("UPDATE companies SET business_area='contracting' WHERE business_area IS NULL OR business_area=''")
  d.execute('CREATE INDEX IF NOT EXISTS idx_companies_assigned ON companies(assigned_user_id)')
  d.execute('DROP INDEX IF EXISTS idx_companies_muqawil_member')
  d.execute('CREATE INDEX IF NOT EXISTS idx_companies_muqawil_member ON companies(muqawil_member_no)')
  d.execute('CREATE INDEX IF NOT EXISTS idx_interactions_company ON interactions(company_id, contacted_at)')
  d.commit()
 except Exception:
  d.rollback()
  raise
 finally:
  d.close()
ensure_schema()

# V7.5.43 — durable contact ledger.
# SQLite remains the fast local CRM database. On Render Free it is ephemeral,
# therefore contact events are mirrored to an external PostgreSQL database
# configured through CONTACTS_DATABASE_URL and restored on every startup.
def contact_ledger_configured():
 return bool(CONTACTS_DATABASE_URL and psycopg is not None)

def _ledger_conn():
 if not contact_ledger_configured():
  raise RuntimeError('CONTACTS_DATABASE_URL is not configured')
 return psycopg.connect(CONTACTS_DATABASE_URL,connect_timeout=10)

def _ledger_ensure_schema():
 with _ledger_conn() as p:
  with p.cursor() as cur:
   cur.execute("""
    CREATE TABLE IF NOT EXISTS alrawasi_contact_events(
      event_id TEXT PRIMARY KEY,
      company_key TEXT NOT NULL,
      company_name TEXT NOT NULL DEFAULT '',
      user_email TEXT NOT NULL DEFAULT '',
      user_name_snapshot TEXT NOT NULL DEFAULT '',
      channel TEXT NOT NULL DEFAULT '',
      outcome TEXT NOT NULL DEFAULT '',
      notes TEXT NOT NULL DEFAULT '',
      contacted_at TEXT NOT NULL,
      next_followup_at TEXT,
      deleted BOOLEAN NOT NULL DEFAULT FALSE,
      created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
      updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
    )
   """)
   cur.execute("CREATE INDEX IF NOT EXISTS idx_alrawasi_contact_company ON alrawasi_contact_events(company_key,contacted_at)")
  p.commit()

def _company_ledger_key(row):
 if not row: return ''
 try:
  if row['muqawil_key']: return str(row['muqawil_key']).strip()
 except Exception: pass
 try:
  if row['source_id']: return str(row['source_id']).strip()
 except Exception: pass
 try:
  return 'id:'+str(row['id'])
 except Exception:
  return ''

def _legacy_remote_event_id(row):
 raw='|'.join([
  str(row['id'] or ''),str(row['company_id'] or ''),str(row['user_id'] or ''),
  str(row['channel'] or ''),str(row['outcome'] or ''),str(row['notes'] or ''),
  str(row['contacted_at'] or ''),str(row['next_followup_at'] or ''),
  str(row['user_name_snapshot'] or '')
 ])
 return 'legacy-'+str(row['id'])+'-'+hashlib.sha256(raw.encode('utf-8')).hexdigest()[:24]

def _ledger_put(event,preserve_existing=False):
 with _ledger_conn() as p:
  with p.cursor() as cur:
   if preserve_existing:
    cur.execute("""
     INSERT INTO alrawasi_contact_events(
      event_id,company_key,company_name,user_email,user_name_snapshot,
      channel,outcome,notes,contacted_at,next_followup_at,deleted
     ) VALUES(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,FALSE)
     ON CONFLICT(event_id) DO NOTHING
    """,(
     event['event_id'],event['company_key'],event['company_name'],event['user_email'],
     event['user_name_snapshot'],event['channel'],event['outcome'],event['notes'],
     event['contacted_at'],event.get('next_followup_at')
    ))
   else:
    cur.execute("""
     INSERT INTO alrawasi_contact_events(
      event_id,company_key,company_name,user_email,user_name_snapshot,
      channel,outcome,notes,contacted_at,next_followup_at,deleted
     ) VALUES(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,FALSE)
     ON CONFLICT(event_id) DO UPDATE SET
      company_key=EXCLUDED.company_key,
      company_name=EXCLUDED.company_name,
      user_email=EXCLUDED.user_email,
      user_name_snapshot=EXCLUDED.user_name_snapshot,
      channel=EXCLUDED.channel,
      outcome=EXCLUDED.outcome,
      notes=EXCLUDED.notes,
      contacted_at=EXCLUDED.contacted_at,
      next_followup_at=EXCLUDED.next_followup_at,
      deleted=FALSE,
      updated_at=NOW()
    """,(
     event['event_id'],event['company_key'],event['company_name'],event['user_email'],
     event['user_name_snapshot'],event['channel'],event['outcome'],event['notes'],
     event['contacted_at'],event.get('next_followup_at')
    ))
  p.commit()

def _ledger_tombstone(event):
 with _ledger_conn() as p:
  with p.cursor() as cur:
   cur.execute("""
    INSERT INTO alrawasi_contact_events(
     event_id,company_key,company_name,user_email,user_name_snapshot,
     channel,outcome,notes,contacted_at,next_followup_at,deleted
    ) VALUES(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,TRUE)
    ON CONFLICT(event_id) DO UPDATE SET deleted=TRUE,updated_at=NOW()
   """,(
    event['event_id'],event['company_key'],event['company_name'],event['user_email'],
    event['user_name_snapshot'],event['channel'],event['outcome'],event['notes'],
    event['contacted_at'],event.get('next_followup_at')
   ))
  p.commit()

def _local_event_payload(d,row):
 c=d.execute("SELECT id,muqawil_key,source_id,name_ar FROM companies WHERE id=?",(row['company_id'],)).fetchone()
 u=d.execute("SELECT email,name FROM users WHERE id=?",(row['user_id'],)).fetchone() if row['user_id'] else None
 event_id=(row['remote_event_id'] or '').strip() if 'remote_event_id' in row.keys() else ''
 if not event_id: event_id=_legacy_remote_event_id(row)
 return {
  'event_id':event_id,
  'company_key':_company_ledger_key(c),
  'company_name':(c['name_ar'] or '') if c else '',
  'user_email':(u['email'] or '') if u else '',
  'user_name_snapshot':row['user_name_snapshot'] or ((u['name'] or '') if u else ''),
  'channel':row['channel'] or '',
  'outcome':row['outcome'] or '',
  'notes':row['notes'] or '',
  'contacted_at':row['contacted_at'] or '',
  'next_followup_at':row['next_followup_at']
 }

def sync_contact_ledger():
 if not contact_ledger_configured():
  print('CONTACT LEDGER: external PostgreSQL not configured; Render Free contacts remain ephemeral.',flush=True)
  return False
 _ledger_ensure_schema()
 d=db()
 try:
  d.execute('BEGIN IMMEDIATE')
  local=d.execute("SELECT * FROM interactions ORDER BY id").fetchall()
  # First adopt every existing local record into the remote ledger. ON CONFLICT
  # preserves remote tombstones, so an intentionally deleted event never returns.
  for row in local:
   ev=_local_event_payload(d,row)
   if not (row['remote_event_id'] or '').strip():
    d.execute("UPDATE interactions SET remote_event_id=? WHERE id=?",(ev['event_id'],row['id']))
   _ledger_put(ev,preserve_existing=True)
  d.commit()
 except Exception:
  d.rollback()
  d.close()
  raise
 finally:
  try:d.close()
  except:pass

 # Pull durable events back into the current ephemeral SQLite DB.
 with _ledger_conn() as p:
  with p.cursor() as cur:
   cur.execute("""
    SELECT event_id,company_key,company_name,user_email,user_name_snapshot,
           channel,outcome,notes,contacted_at,next_followup_at,deleted
    FROM alrawasi_contact_events
    ORDER BY contacted_at,event_id
   """)
   remote=cur.fetchall()

 d=db()
 try:
  d.execute('BEGIN IMMEDIATE')
  latest_by_company={}
  for r in remote:
   event_id,company_key,company_name,user_email,user_name,channel,outcome,notes,contacted_at,next_followup,deleted=r
   existing=d.execute("SELECT id FROM interactions WHERE remote_event_id=?",(event_id,)).fetchone()
   if deleted:
    if existing: d.execute("DELETE FROM interactions WHERE id=?",(existing['id'],))
    continue
   c=d.execute("""
    SELECT id FROM companies
    WHERE muqawil_key=? OR source_id=?
    ORDER BY CASE WHEN muqawil_key=? THEN 0 ELSE 1 END,id
    LIMIT 1
   """,(company_key,company_key,company_key)).fetchone()
   if not c and str(company_key).startswith('id:'):
    try:c=d.execute("SELECT id FROM companies WHERE id=?",(int(str(company_key)[3:]),)).fetchone()
    except: c=None
   if not c:
    continue
   u=None
   if user_email:
    u=d.execute("SELECT id,specialty,email FROM users WHERE lower(email)=lower(?) LIMIT 1",(user_email,)).fetchone()
   if not u and user_name:
    u=d.execute("SELECT id,specialty,email FROM users WHERE name=? LIMIT 1",(user_name,)).fetchone()
   uid=u['id'] if u else None
   if not existing:
    d.execute("""INSERT INTO interactions(
      company_id,user_id,user_name_snapshot,channel,outcome,notes,
      contacted_at,next_followup_at,remote_event_id,specialty_at_contact,actor_email_at_contact
     ) VALUES(?,?,?,?,?,?,?,?,?,?,?)""",
     (c['id'],uid,user_name or '',channel or '',outcome or '',notes or '',
      contacted_at,next_followup,event_id,(u['specialty'] if u else None),(u['email'] if u else user_email)))
   prev=latest_by_company.get(c['id'])
   if not prev or str(contacted_at or '')>=str(prev[0] or ''):
    latest_by_company[c['id']]=(contacted_at,outcome)
  for cid,(_,outcome) in latest_by_company.items():
   if outcome:
    d.execute("UPDATE companies SET marketing_status=?,updated_at=CURRENT_TIMESTAMP WHERE id=?",(outcome,cid))
  d.commit()
 finally:
  d.close()
 print('CONTACT LEDGER: durable PostgreSQL sync OK; remote events='+str(len(remote)),flush=True)
 return True

def _save_interaction_durable(d,c,cid,channel,outcome,notes,next_followup):
 user=d.execute("SELECT email,name,specialty FROM users WHERE id=?",(session['uid'],)).fetchone()
 contacted_at=d.execute("SELECT CURRENT_TIMESTAMP").fetchone()[0]
 event={
  'event_id':uuid.uuid4().hex,
  'company_key':_company_ledger_key(c),
  'company_name':c['name_ar'] or '',
  'user_email':(user['email'] or '') if user else '',
  'user_name_snapshot':session.get('name','') or ((user['name'] or '') if user else ''),
  'channel':channel or '',
  'outcome':outcome or '',
  'notes':notes or '',
  'contacted_at':contacted_at,
  'next_followup_at':next_followup
 }
 if contact_ledger_configured():
  _ledger_put(event)
 d.execute("""INSERT INTO interactions(
  company_id,user_id,user_name_snapshot,channel,outcome,notes,
  contacted_at,next_followup_at,remote_event_id,specialty_at_contact,actor_email_at_contact
 ) VALUES(?,?,?,?,?,?,?,?,?,?,?)""",
 (cid,session['uid'],event['user_name_snapshot'],event['channel'],event['outcome'],
  event['notes'],event['contacted_at'],event['next_followup_at'],event['event_id'],
  (user['specialty'] if user else 'general'),event['user_email']))
 d.execute("UPDATE companies SET marketing_status=?,updated_at=CURRENT_TIMESTAMP WHERE id=?",(event['outcome'],cid))
 d.commit()
 return event

try:
 sync_contact_ledger()
except Exception as _ledger_boot_error:
 print('CONTACT LEDGER WARNING:',repr(_ledger_boot_error),flush=True)

def refresh_muqawil_classification_index():
 """Populate Muqawil classification only; never touch CRM contact/status/follow-up data."""
 path=os.path.join(BASE,'muqawil_classification_index.json')
 if not os.path.exists(path): return
 try:
  with open(path,encoding='utf-8') as f: info=json.load(f)
  d=db()
  try:
   for source_id,v in info.items():
    if not source_id: continue
    d.execute("""UPDATE companies
                SET classification_status=?, classification_grade=?,
                    muqawil_member_no=COALESCE(NULLIF(muqawil_member_no,''),?)
                WHERE source_id=?""",
              (v.get('status') or 'غير مصنفة',v.get('grade') or '',v.get('member') or None,source_id))
   # Records from Muqawil that are not present in the index remain explicitly unknown.
   d.execute("""UPDATE companies SET classification_status='غير محدد'
                WHERE (data_source='Muqawil' OR source LIKE '%مقاول%' OR source_url LIKE '%muqawil.org%')
                  AND (classification_status IS NULL OR classification_status='')""")
   d.commit()
  finally: d.close()
 except Exception as e:
  print('classification migration warning:',e)
refresh_muqawil_classification_index()

def login_required(f):
 @wraps(f)
 def w(*a,**k):
  if 'uid' not in session:return redirect(url_for('login'))
  return f(*a,**k)
 return w

def admin_required(f):
 @wraps(f)
 def w(*a,**k):
  if 'uid' not in session:return redirect(url_for('login'))
  if session.get('role')!='admin': abort(403)
  return f(*a,**k)
 return w

def current_user(d): return d.execute('SELECT * FROM users WHERE id=?',(session['uid'],)).fetchone()
def specialty_scope(user, alias='c'):
 if user and (user['specialty'] or 'general')=='contracting': return f"{alias}.business_area='contracting'", []
 return '1=1', []
def scoped_company(d,cid):
 u=current_user(d); cond,args=specialty_scope(u,'c')
 return d.execute(f'SELECT c.*,u.name rep FROM companies c LEFT JOIN users u ON u.id=c.assigned_user_id WHERE c.id=? AND {cond}',[cid]+args).fetchone()
def can_add_companies(user):
 return True

def wa_text(user_row, company_name):
 specialty=(user_row['specialty'] or 'general') if user_row else 'general'; template=(user_row['whatsapp_template'] or '').strip() if user_row else ''
 if not template: template=DEFAULT_MESSAGES.get(specialty,DEFAULT_MESSAGES['general'])
 return template.replace('{{rep_name}}','{rep_name}').replace('{{company_name}}','{company_name}').format(rep_name=user_row['name'] if user_row else session.get('name',''),company_name=company_name or 'شركتكم')
def wa_link(phone,text):
 if not phone:return None
 p=''.join(ch for ch in phone if ch.isdigit())
 if p.startswith('05'): p='966'+p[1:]
 elif p.startswith('5') and len(p)==9: p='966'+p
 return 'https://wa.me/'+p+'?text='+urllib.parse.quote(text)

def email_link(email, company_name, text):
 if not email:return None
 subject='تواصل مع '+(company_name or 'الشركة')
 return 'mailto:'+urllib.parse.quote(email.strip(), safe='@,+')+'?subject='+urllib.parse.quote(subject)+'&body='+urllib.parse.quote(text)

@app.context_processor
def ctx():
 appearance={}
 if session.get('uid'):
  d=db(); u=d.execute('SELECT appearance_default FROM users WHERE id=?',(session['uid'],)).fetchone(); d.close()
  if u and u['appearance_default']:
   try: appearance=json.loads(u['appearance_default'])
   except: appearance={}
 return {'status':STATUS,'status_class':STATUS_CLASS,'specialties':SPECIALTIES,'user':session.get('name'),'appearance_default':appearance,'app_version':APP_VERSION,'contact_ledger_configured':contact_ledger_configured()}
@app.route('/login',methods=['GET','POST'])
def login():
 if request.method=='POST':
  d=db(); u=d.execute('SELECT * FROM users WHERE email=? AND active=1',(request.form['email'],)).fetchone(); d.close()
  if u and check_password_hash(u['password_hash'],request.form['password']): session.clear(); session.update(uid=u['id'],name=u['name'],role=u['role']); return redirect(url_for('dashboard'))
  flash('بيانات الدخول غير صحيحة')
 return render_template('login.html')

@app.route('/forgot-password',methods=['GET','POST'])
def forgot_password():
 if request.method=='POST':
  email=(request.form.get('email') or '').strip().lower()
  recovery=(request.form.get('recovery_key') or '').strip()
  password=request.form.get('password') or ''
  confirm=request.form.get('confirm_password') or ''
  if not ADMIN_RECOVERY_KEY:
   flash('استعادة كلمة المرور غير مفعلة. أضف ADMIN_RECOVERY_KEY في متغيرات Render ثم أعد المحاولة.')
   return render_template('forgot_password.html')
  if not secrets.compare_digest(recovery,ADMIN_RECOVERY_KEY):
   flash('رمز الاستعادة غير صحيح.')
   return render_template('forgot_password.html')
  if len(password) < 10:
   flash('كلمة المرور الجديدة يجب ألا تقل عن 10 أحرف.')
   return render_template('forgot_password.html')
  if password != confirm:
   flash('تأكيد كلمة المرور غير مطابق.')
   return render_template('forgot_password.html')
  d=db()
  try:
   d.execute('BEGIN IMMEDIATE')
   u=d.execute("SELECT id,name,email FROM users WHERE lower(email)=? AND role='admin' AND active=1",(email,)).fetchone()
   if not u:
    d.rollback(); flash('لا يوجد حساب مدير نشط بهذا البريد.'); return render_template('forgot_password.html')
   d.execute('UPDATE users SET password_hash=? WHERE id=?',(generate_password_hash(password),u['id']))
   d.commit()
  except Exception:
   d.rollback(); raise
  finally:
   d.close()
  flash('تم تعيين كلمة المرور الجديدة. يمكنك تسجيل الدخول الآن.')
  return redirect(url_for('login'))
 return render_template('forgot_password.html')

MUQAWIL_BASE='https://muqawil.org'
MUQAWIL_LIST='https://muqawil.org/ar/contractors'

def _norm_name(v):
 return re.sub(r'\s+',' ',(v or '').strip()).casefold()

def _muqawil_cards(html):
 soup=BeautifulSoup(html,'html.parser')
 out=[]
 # Contractor detail links have /ar/contractors/<id>/<id>
 for a in soup.select('a[href*="/ar/contractors/"]'):
  href=(a.get('href') or '').strip()
  name=' '.join(a.stripped_strings).strip()
  if not name or not re.search(r'/ar/contractors/\d+/\d+', href): continue
  url=urljoin(MUQAWIL_BASE,href)
  if not any(x['url']==url for x in out):
   out.append({'name':name,'url':url})
 return out

def _muqawil_detail(url):
 r=requests.get(url,timeout=25,headers={'User-Agent':'AlrawasiCRM/1.0 (+business CRM sync)'})
 r.raise_for_status()
 soup=BeautifulSoup(r.text,'html.parser')
 txt=' '.join(soup.stripped_strings)
 def grab(label, stop_labels):
  i=txt.find(label)
  if i<0:return ''
  val=txt[i+len(label):].strip()
  cuts=[val.find(x) for x in stop_labels if val.find(x)>=0]
  return val[:min(cuts)].strip(' :-') if cuts else val[:180].strip(' :-')
 member=grab('رقم العضويه',['العضوية','عضو منذ','حجم المنشأة'])
 phone=grab('رقم جوال المنشأة',['البريد الإلكتروني','المدينة','المنطقه'])
 city=grab('المدينة',['المنطقه','عنوان'])
 region=grab('المنطقه',['عنوان','طلب تعاقد'])
 address=grab('عنوان',['طلب تعاقد','التراخيص','الأنشطة'])
 title=(soup.find('h1').get_text(' ',strip=True) if soup.find('h1') else '')
 classification_status='غير مصنفة'; classification_grade=''
 if 'غير مصنف' not in txt and 'مصنف' in txt:
  classification_status='مصنفة'
  gm=re.search(r'الدرجة\s+([^\s،()]+)',txt)
  classification_grade=gm.group(1) if gm else 'غير محددة'
 return {'name':title,'member':member,'phone':phone,'city':city,'region':region,'address':address,'url':url,
         'classification_status':classification_status,'classification_grade':classification_grade}

def sync_muqawil(max_pages=8):
 d=db(); added=existing=seen=pages_ok=0; errors=[]
 try:
  for page in range(1,max_pages+1):
   try:
    r=requests.get(MUQAWIL_LIST,params={'page':page},timeout=25,headers={'User-Agent':'Mozilla/5.0 (compatible; AlrawasiCRM/1.0)'})
    if r.status_code in (401,403):
     errors.append(f'منصة مقاول رفضت الاتصال الآلي (HTTP {r.status_code}). لم يتم تعديل أي عميل حالي.')
     break
    if r.status_code==429:
     errors.append('منصة مقاول حدّت عدد الطلبات مؤقتاً (HTTP 429). لم يتم تعديل أي عميل حالي.')
     break
    if r.status_code!=200:
     errors.append(f'صفحة {page}: HTTP {r.status_code}'); continue
    cards=_muqawil_cards(r.text); pages_ok+=1
    if not cards: continue
   except requests.Timeout:
    errors.append(f'صفحة {page}: انتهت مهلة الاتصال'); continue
   except requests.RequestException as e:
    errors.append(f'صفحة {page}: خطأ اتصال {type(e).__name__}'); continue
   except Exception as e:
    errors.append(f'صفحة {page}: {type(e).__name__}: {str(e)[:100]}'); continue
   for card in cards:
    seen+=1
    try:
     info=_muqawil_detail(card['url']); name=(info.get('name') or card.get('name') or '').strip()
     if not name: errors.append(f"{card.get('url','')}: تعذر قراءة اسم الشركة"); continue
     member=(info.get('member') or '').strip() or None; row=None
     if member: row=d.execute('SELECT id FROM companies WHERE muqawil_member_no=?',(member,)).fetchone()
     if not row:
      row=d.execute("SELECT id FROM companies WHERE LOWER(TRIM(COALESCE(name_ar,'')))=LOWER(TRIM(?)) OR LOWER(TRIM(COALESCE(name_en,'')))=LOWER(TRIM(?)) LIMIT 1",(name,name)).fetchone()
     if row:
      existing+=1; continue
     cols={x['name'] for x in d.execute('PRAGMA table_info(companies)').fetchall()}
     vals={'name_ar':name,'name_en':'','activity':'مقاولات','city':info.get('city',''),'phone':info.get('phone',''),'whatsapp':info.get('phone',''),'email':'','website':'','cr_number':'','status':'جديد','business_area':'contracting','muqawil_member_no':member,'source_url':info.get('url',card['url']),'address':info.get('address',''),'data_source':'Muqawil','imported_at':datetime.now().strftime('%Y-%m-%d %H:%M:%S'),'classification_status':info.get('classification_status','غير مصنفة'),'classification_grade':info.get('classification_grade','')}
     keys=[k for k in vals if k in cols]
     d.execute(f"INSERT INTO companies ({','.join(keys)}) VALUES ({','.join('?' for _ in keys)})",tuple(vals[k] for k in keys))
     d.commit(); added+=1
    except requests.Timeout:
     d.rollback(); errors.append(f"{card.get('name','شركة')}: انتهت مهلة صفحة الشركة")
    except requests.RequestException as e:
     d.rollback(); errors.append(f"{card.get('name','شركة')}: خطأ اتصال {type(e).__name__}")
    except sqlite3.IntegrityError:
     d.rollback(); existing+=1
    except Exception as e:
     d.rollback(); errors.append(f"{card.get('name','شركة')}: {type(e).__name__}: {str(e)[:100]}")
  return {'added':added,'existing':existing,'seen':seen,'pages_ok':pages_ok,'errors':errors,'error_count':len(errors)}
 finally: d.close()

@app.post('/admin/sync-muqawil')
@admin_required
def admin_sync_muqawil():
 flash('المزامنة المباشرة متوقفة بسبب HTTP 403. استخدم صفحة تحديث بيانات مقاول JSON.')
 return redirect(url_for('import_muqawil_json'))

@app.post('/internal/sync-muqawil')
def sync_muqawil_internal():
 token=request.headers.get('X-Sync-Token','')
 expected=os.environ.get('MUQAWIL_SYNC_TOKEN','')
 if not expected or not secrets.compare_digest(token,expected):
  return {'ok':False},403
 return {'ok':True,**sync_muqawil()}

@app.get('/health')
def health(): return jsonify({'status':'ok'})
@app.get('/logout')
def logout(): session.clear(); return redirect(url_for('login'))

@app.get('/')
@login_required
def dashboard():
 d=db(); is_admin=session.get('role')=='admin'; me=current_user(d); cond,args=specialty_scope(me,'companies'); scope=' WHERE '+cond
 total=d.execute('SELECT COUNT(*) n FROM companies'+scope,args).fetchone()['n']
 counts={k:d.execute('SELECT COUNT(*) n FROM companies'+scope+' AND marketing_status=?',args+[k]).fetchone()['n'] for k in STATUS}
 if is_admin:
  follow=d.execute("SELECT COUNT(DISTINCT company_id) n FROM interactions WHERE next_followup_at IS NOT NULL AND datetime(next_followup_at)<=datetime('now','localtime')").fetchone()['n']
  reps=d.execute("SELECT u.id,u.name,u.role,COUNT(c.id) n,SUM(CASE WHEN c.marketing_status='contacted' THEN 1 ELSE 0 END) contacted,SUM(CASE WHEN c.marketing_status='interested' THEN 1 ELSE 0 END) interested,SUM(CASE WHEN c.marketing_status='contracted' THEN 1 ELSE 0 END) contracted,(SELECT MAX(i.contacted_at) FROM interactions i WHERE i.user_id=u.id) last_contact FROM users u LEFT JOIN companies c ON c.assigned_user_id=u.id WHERE u.active=1 GROUP BY u.id ORDER BY n DESC").fetchall()
 else:
  follow=d.execute("SELECT COUNT(DISTINCT i.company_id) n FROM interactions i JOIN companies c ON c.id=i.company_id WHERE c.assigned_user_id=? AND i.next_followup_at IS NOT NULL AND datetime(i.next_followup_at)<=datetime('now','localtime')",(session['uid'],)).fetchone()['n']; reps=[]
 recent=d.execute(f"SELECT c.id,c.name_ar,c.name_en,i.channel,i.outcome,i.contacted_at,i.next_followup_at FROM interactions i JOIN companies c ON c.id=i.company_id WHERE {specialty_scope(me,'c')[0]} ORDER BY i.contacted_at DESC LIMIT 10").fetchall(); d.close()
 return render_template('dashboard.html',counts=counts,total=total,follow=follow,reps=reps,recent=recent,is_admin=is_admin)

def company_directory_order_sql(alias='c'):
 a=alias
 # V7.5.41: crawler-only leads are always last while unverified.
 # After manual verification they join the trusted directory and are placed
 # after legacy companies inside the same current phone/classification bucket.
 return f"""CASE
 WHEN COALESCE({a}.source_verified,0)=0
      AND COALESCE({a}.data_source,'') LIKE 'Muqawil Platform Crawler V3%' THEN 1
 ELSE 0 END,
 CASE WHEN COALESCE(TRIM({a}.whatsapp),'')<>'' OR COALESCE(TRIM({a}.phone_mobile),'')<>'' THEN 0 ELSE 1 END,
 CASE
 WHEN {a}.classification_status='غير مصنفة' THEN 0
 WHEN {a}.classification_status IS NULL OR {a}.classification_status='' OR {a}.classification_status='غير محدد' OR {a}.classification_grade IS NULL OR {a}.classification_grade='' OR {a}.classification_grade LIKE '%غير محدد%' THEN 1
 WHEN {a}.classification_grade='6' OR {a}.classification_grade LIKE '%سادس%' THEN 2
 WHEN {a}.classification_grade='5' OR {a}.classification_grade LIKE '%خامس%' THEN 3
 WHEN {a}.classification_grade='4' OR {a}.classification_grade LIKE '%رابع%' THEN 4
 WHEN {a}.classification_grade='3' OR {a}.classification_grade LIKE '%ثالث%' THEN 5
 WHEN {a}.classification_grade='2' OR {a}.classification_grade LIKE '%ثان%' THEN 6
 WHEN {a}.classification_grade='1' OR {a}.classification_grade LIKE '%أول%' THEN 7
 ELSE 8 END,
 CASE
 WHEN COALESCE({a}.source_verified,0)=1
      AND COALESCE({a}.data_source,'') LIKE 'Muqawil Platform Crawler V3%' THEN 1
 ELSE 0 END,
 {a}.id"""

def contact_department_colors(d, company_ids):
 """Visual-only color state based on real contacts by the two service departments.

 A department counts after a recorded contact, not merely assigning a company,
 editing the marketing status, or creating an automatic quick follow-up.
 We prefer specialty_at_contact (stable history) over the current user's specialty.
 """
 if not company_ids:
  return {}
 params=','.join('?' for _ in company_ids)
 records=d.execute(f"""SELECT i.company_id,
    MAX(CASE WHEN COALESCE(NULLIF(TRIM(i.specialty_at_contact),''),u.specialty)='contracting' THEN 1 ELSE 0 END) as has_contracting,
    MAX(CASE WHEN COALESCE(NULLIF(TRIM(i.specialty_at_contact),''),u.specialty)='advertising' THEN 1 ELSE 0 END) as has_advertising
    FROM interactions i LEFT JOIN users u ON u.id=i.user_id
    WHERE i.company_id IN ({params})
      AND COALESCE(TRIM(i.channel),'') NOT IN ('','تحديث حالة العميل','متابعة')
      AND COALESCE(i.outcome,'') NOT IN ('new','unreachable')
    GROUP BY i.company_id""",company_ids).fetchall()
 states={}
 for record in records:
  contractor=bool(record['has_contracting']); advertiser=bool(record['has_advertising'])
  states[record['company_id']]=('both' if contractor and advertiser else
                                  'contracting' if contractor else
                                  'advertising' if advertiser else 'none')
 return states

@app.get('/companies')
@login_required
def companies():
 q=request.args.get('q','').strip(); city=request.args.get('city','').strip(); st=request.args.get('status','').strip(); rep=request.args.get('rep','').strip(); cls=request.args.get('classification','').strip()
 try: page=max(1,int(request.args.get('page',1)))
 except: page=1
 try: per=int(request.args.get('per',50))
 except: per=50
 if per not in (20,30,50,100): per=50
 where=[]; args=[]
 d0=db(); me0=current_user(d0); spec_cond,spec_args=specialty_scope(me0,'c'); d0.close(); where.append(spec_cond); args += spec_args
 if session.get('role')=='admin' and rep:
  if rep=='unassigned': where.append('c.assigned_user_id IS NULL')
  else: where.append('c.assigned_user_id=?'); args.append(rep)
 if q: where.append('(c.name_ar LIKE ? OR c.name_en LIKE ? OR c.phone_mobile LIKE ? OR c.phone_landline LIKE ? OR c.email LIKE ?)'); args += [f'%{q}%']*5
 if city: where.append('c.city=?'); args.append(city)
 if st: where.append('c.marketing_status=?'); args.append(st)
 if cls=='classified': where.append("c.classification_status='مصنفة'")
 elif cls=='unclassified': where.append("c.classification_status='غير مصنفة'")
 elif cls=='unspecified': where.append("(c.classification_status IS NULL OR c.classification_status='' OR c.classification_status='غير محدد' OR c.classification_grade IS NULL OR c.classification_grade='' OR c.classification_grade LIKE '%غير محدد%')")
 elif cls in ('6','5','4','3','2','1'):
  grade_words={'6':'سادس','5':'خامس','4':'رابع','3':'ثالث','2':'ثان','1':'أول'}
  where.append("(c.classification_grade=? OR c.classification_grade LIKE ?)"); args += [cls, '%'+grade_words[cls]+'%']
 w=(' WHERE '+' AND '.join(where)) if where else ''; d=db(); total=d.execute('SELECT COUNT(*) n FROM companies c'+w,args).fetchone()['n']
 total_pages=max(1,(total+per-1)//per)
 page=min(page,total_pages)
 sql="""SELECT c.*,u.name rep,(SELECT i.contacted_at FROM interactions i WHERE i.company_id=c.id ORDER BY i.contacted_at DESC LIMIT 1) last_contact,(SELECT i.next_followup_at FROM interactions i WHERE i.company_id=c.id AND i.next_followup_at IS NOT NULL ORDER BY i.contacted_at DESC LIMIT 1) next_followup FROM companies c LEFT JOIN users u ON u.id=c.assigned_user_id"""+w+" ORDER BY "+company_directory_order_sql('c')+" LIMIT ? OFFSET ?"
 rows=[dict(r) for r in d.execute(sql,args+[per,(page-1)*per]).fetchall()]; [r.update(display_number=((page-1)*per+i)) for i,r in enumerate(rows,1)]
 department_colors=contact_department_colors(d,[r['id'] for r in rows])
 for r in rows:
  r['contact_department_state']=department_colors.get(r['id'],'none')
 cities=d.execute("SELECT DISTINCT city FROM companies WHERE city IS NOT NULL AND city<>'' ORDER BY city").fetchall()
 me=d.execute('SELECT * FROM users WHERE id=?',(session['uid'],)).fetchone()
 users=d.execute("SELECT id,name FROM users WHERE active=1 ORDER BY name").fetchall() if session.get('role')=='admin' else []
 # V7.5.47: بطاقات إحصائية للمظهر الاحترافي فقط، بدون أي تعديل على البيانات.
 scope_cond,scope_args=specialty_scope(me,'c')
 summary_total=d.execute("SELECT COUNT(*) n FROM companies c WHERE "+scope_cond,scope_args).fetchone()['n']
 summary_verified=d.execute("SELECT COUNT(*) n FROM companies c WHERE "+scope_cond+" AND COALESCE(c.source_verified,0)=1",scope_args).fetchone()['n']
 summary_users=d.execute("SELECT COUNT(*) n FROM users WHERE active=1").fetchone()['n']
 summary_stats={'total':summary_total,'verified':summary_verified,'unverified':max(0,summary_total-summary_verified),'users':summary_users}
 d.close()
 for r in rows:
  msg=wa_text(me,r.get('name_ar') or r.get('name_en'))
  r['wa_link']=wa_link(r.get('whatsapp') or r.get('phone_mobile'),msg)
  r['email_link']=email_link(r.get('email'),r.get('name_ar') or r.get('name_en'),msg)
 return render_template('companies.html',rows=rows,total=total,total_pages=total_pages,page=page,per=per,q=q,city=city,status_filter=st,rep_filter=rep,classification_filter=cls,cities=cities,users=users,is_admin=session.get('role')=='admin',can_add=can_add_companies(me),my_specialty=(me['specialty'] or 'general'),summary_stats=summary_stats)

@app.post('/companies/assign')
@admin_required
def assign_companies():
 ids=[int(x) for x in request.form.getlist('company_ids') if x.isdigit()]; uid=request.form.get('assigned_user_id')
 if not ids: flash('حدد عميلاً واحداً على الأقل'); return redirect(request.referrer or url_for('companies'))
 d=db(); valid=d.execute('SELECT id FROM users WHERE id=? AND active=1',(uid,)).fetchone() if uid else None
 if uid and not valid: d.close(); abort(400)
 ph=','.join('?'*len(ids)); d.execute(f'UPDATE companies SET assigned_user_id=?,updated_at=CURRENT_TIMESTAMP WHERE id IN ({ph})',[uid or None]+ids); d.commit(); d.close(); flash(f'تم توزيع {len(ids)} عميل بنجاح'); return redirect(request.referrer or url_for('companies'))

@app.get('/company/<int:cid>')
@login_required
def company(cid):
 return_to=(request.args.get('return_to') or '').strip()
 if return_to.startswith('/companies') and not return_to.startswith('//'):
  session['companies_return_to']=return_to
 d=db(); c=scoped_company(d,cid)
 if not c: d.close(); abort(404)
 its=d.execute("SELECT i.*,COALESCE(u.name,i.user_name_snapshot,'مستخدم محذوف') rep FROM interactions i LEFT JOIN users u ON u.id=i.user_id WHERE company_id=? ORDER BY contacted_at DESC",(cid,)).fetchall(); users=d.execute('SELECT * FROM users WHERE active=1').fetchall() if session.get('role')=='admin' else []; me=d.execute('SELECT * FROM users WHERE id=?',(session['uid'],)).fetchone(); ordered_ids=[r['id'] for r in d.execute("SELECT c.id FROM companies c ORDER BY "+company_directory_order_sql('c')).fetchall()]
 display_number=(ordered_ids.index(cid)+1) if cid in ordered_ids else cid
 c=dict(c); c['display_number']=display_number; d.close(); msg=wa_text(me,c['name_ar'] or c['name_en']); link=wa_link(c['whatsapp'] or c['phone_mobile'],msg); mail_link=email_link(c['email'],c['name_ar'] or c['name_en'],msg)
 return render_template('company.html',c=c,its=its,users=users,wa_link=link,email_link=mail_link,is_admin=session.get('role')=='admin',return_to=session.get('companies_return_to') or url_for('companies'))

@app.post('/company/<int:cid>/update')
@login_required
def update_company(cid):
 d=db(); c=scoped_company(d,cid)
 if not c: d.close(); abort(404)
 # بيانات التواصل قابلة للتصحيح يدوياً من صفحة الشركة لجميع المستخدمين المصرح لهم برؤيتها.
 phone_mobile=(request.form.get('phone_mobile') or '').strip()
 whatsapp=(request.form.get('whatsapp') or '').strip()
 phone_landline=(request.form.get('phone_landline') or '').strip()
 email=(request.form.get('email') or '').strip()
 old_phone=(c['phone_mobile'] or '').strip()
 phone_source=c['phone_source']
 phone_verified_at=c['phone_verified_at']
 if phone_mobile != old_phone:
  phone_source='manual_crm' if phone_mobile else None
  phone_verified_at=datetime.now().strftime('%Y-%m-%d %H:%M:%S') if phone_mobile else None
 if session.get('role')=='admin':
  assigned=request.form.get('assigned_user_id') or None
  d.execute('UPDATE companies SET marketing_status=?,assigned_user_id=?,priority=?,marketing_notes=?,phone_mobile=?,whatsapp=?,phone_landline=?,email=?,phone_source=?,phone_verified_at=?,updated_at=CURRENT_TIMESTAMP WHERE id=?',(request.form['marketing_status'],assigned,request.form.get('priority','normal'),request.form.get('marketing_notes',''),phone_mobile,whatsapp,phone_landline,email,phone_source,phone_verified_at,cid))
 else:
  d.execute('UPDATE companies SET marketing_status=?,priority=?,marketing_notes=?,phone_mobile=?,whatsapp=?,phone_landline=?,email=?,phone_source=?,phone_verified_at=?,updated_at=CURRENT_TIMESTAMP WHERE id=?',(request.form['marketing_status'],request.form.get('priority','normal'),request.form.get('marketing_notes',''),phone_mobile,whatsapp,phone_landline,email,phone_source,phone_verified_at,cid)); # حفظ كل تحديث لحالة العميل في سجل التواصل الحالي
 new_status=(request.form.get('marketing_status') or '').strip()
 old_status=(c['marketing_status'] or 'new').strip()
 if new_status and new_status != old_status:
  note=f"تغيير الحالة: {STATUS.get(old_status,old_status)} ← {STATUS.get(new_status,new_status)}"
  extra=(request.form.get('marketing_notes') or '').strip()
  if extra: note += ' | ' + extra
  d.execute('INSERT INTO interactions(company_id,user_id,user_name_snapshot,channel,outcome,notes) VALUES(?,?,?,?,?,?)',
            (cid,session['uid'],session.get('name',''),'تحديث حالة العميل',new_status,note))
 d.commit(); d.close(); return redirect(url_for('company',cid=cid))

@app.post('/company/<int:cid>/crawler-details')
@admin_required
def update_crawler_company_details(cid):
 d=db()
 try:
  d.execute('BEGIN IMMEDIATE')
  c=d.execute("SELECT * FROM companies WHERE id=?",(cid,)).fetchone()
  if not c:
   d.rollback(); abort(404)
  if 'Muqawil Platform Crawler V3' not in (c['data_source'] or ''):
   d.rollback(); flash('تعديل البيانات الشامل مخصص للشركات الجديدة من مجموعة الزحف.'); return redirect(url_for('company',cid=cid))
  def txt(name,limit=500):
   return (request.form.get(name) or '').strip()[:limit]
  name_ar=txt('name_ar',300)
  if not name_ar:
   d.rollback(); flash('اسم الشركة مطلوب.'); return redirect(url_for('company',cid=cid))
  phone_mobile=''.join(ch for ch in txt('phone_mobile',30) if ch.isdigit())
  whatsapp=''.join(ch for ch in txt('whatsapp',30) if ch.isdigit())
  phone_landline=''.join(ch for ch in txt('phone_landline',30) if ch.isdigit())
  if phone_mobile and (len(phone_mobile)!=10 or not phone_mobile.startswith('05')):
   d.rollback(); flash('رقم الجوال يجب أن يكون بصيغة 05xxxxxxxx.'); return redirect(url_for('company',cid=cid))
  if whatsapp and (len(whatsapp)!=10 or not whatsapp.startswith('05')):
   d.rollback(); flash('رقم واتساب يجب أن يكون بصيغة 05xxxxxxxx.'); return redirect(url_for('company',cid=cid))
  training_raw=txt('training_hours',20)
  training_hours=None
  if training_raw:
   try:
    training_hours=max(0,int(training_raw))
   except ValueError:
    d.rollback(); flash('الساعات التدريبية يجب أن تكون رقماً صحيحاً.'); return redirect(url_for('company',cid=cid))
  old_mobile=(c['phone_mobile'] or '').strip()
  old_landline=(c['phone_landline'] or '').strip()
  old_whatsapp=(c['whatsapp'] or '').strip()
  phone_source=c['phone_source']
  phone_verified_at=c['phone_verified_at']
  if phone_mobile!=old_mobile or phone_landline!=old_landline or whatsapp!=old_whatsapp:
   phone_source='manual_crm'
   phone_verified_at=datetime.now().strftime('%Y-%m-%d %H:%M:%S') if (phone_mobile or phone_landline or whatsapp) else None
  city=txt('city',200)
  address=txt('address',1000)
  email=txt('email',300)
  sql=("UPDATE companies SET name_ar=?,cr_number=?,website=?,muqawil_member_no=?,membership_type=?,membership_since=?,"
       "company_size=?,account_status=?,training_hours=?,muqawil_region=?,muqawil_city=?,city=?,muqawil_address=?,address=?,"
       "sector=?,email=?,muqawil_email=?,phone_mobile=?,muqawil_phone_mobile=?,whatsapp=?,phone_landline=?,muqawil_phone_landline=?,"
       "phone_source=?,phone_verified_at=?,updated_at=CURRENT_TIMESTAMP WHERE id=?")
  vals=(name_ar,txt('cr_number',100),txt('website',500),txt('muqawil_member_no',100),txt('membership_type',200),
        txt('membership_since',100),txt('company_size',200),txt('account_status',200),training_hours,txt('muqawil_region',200),
        city,city,address,address,txt('sector',200) or 'مقاولات',email,email,phone_mobile,phone_mobile,whatsapp,phone_landline,
        phone_landline,phone_source,phone_verified_at,cid)
  d.execute(sql,vals)
  d.commit(); flash('تم حفظ بيانات الشركة. حالة التوثيق لم تتغير.')
 except Exception:
  d.rollback(); raise
 finally:
  d.close()
 return redirect(url_for('company',cid=cid))

@app.post('/company/<int:cid>/crawler-verification')
@admin_required
def verify_crawler_company(cid):
 d=db()
 try:
  d.execute('BEGIN IMMEDIATE')
  c=d.execute("SELECT * FROM companies WHERE id=?",(cid,)).fetchone()
  if not c:
   d.rollback(); abort(404)
  if 'Muqawil Platform Crawler V3' not in (c['data_source'] or ''):
   d.rollback(); flash('هذه الشركة ليست من مجموعة الشركات الجديدة غير الموثقة.'); return redirect(url_for('company',cid=cid))
  state=(request.form.get('verification_state') or 'unverified').strip()
  classification_status=(request.form.get('classification_status') or 'غير مصنفة').strip()
  grade=(request.form.get('classification_grade') or '').strip()
  if classification_status not in {'غير مصنفة','غير محدد','مصنفة'}:
   d.rollback(); flash('حالة التصنيف غير صحيحة.'); return redirect(url_for('company',cid=cid))
  if grade not in {'','سادسة','خامسة','رابعة','ثالثة','ثانية','أولى'}:
   d.rollback(); flash('درجة التصنيف غير صحيحة.'); return redirect(url_for('company',cid=cid))
  if classification_status=='مصنفة' and not grade:
   d.rollback(); flash('اختر درجة التصنيف قبل توثيق شركة مصنفة.'); return redirect(url_for('company',cid=cid))
  if classification_status!='مصنفة':
   grade=''
  verified=1 if state=='verified' else 0
  profile_status='verified_manual' if verified else 'pending_manual_verification'
  d.execute('''UPDATE companies
               SET source_verified=?,muqawil_profile_status=?,
                   classification_status=?,classification_grade=?,
                   updated_at=CURRENT_TIMESTAMP
               WHERE id=?''',
            (verified,profile_status,classification_status,grade,cid))
  d.commit()
  flash('تم توثيق الشركة ونقلها إلى مجموعة تصنيفها.' if verified else 'تم إبقاء الشركة غير موثقة في نهاية الدليل.')
 except Exception:
  d.rollback(); raise
 finally:
  d.close()
 return redirect(url_for('company',cid=cid))

@app.post('/company/<int:cid>/interaction')
@login_required
def interaction(cid):
 d=db(); c=scoped_company(d,cid)
 if not c: d.close(); abort(404)
 try:
  _save_interaction_durable(
   d,c,cid,request.form['channel'],request.form['outcome'],
   request.form.get('notes',''),request.form.get('next_followup_at') or None
  )
 except Exception as e:
  try:d.rollback()
  except:pass
  d.close()
  if contact_ledger_configured():
   flash('تعذر الحفظ الدائم لسجل التواصل. لم يتم اعتماد العملية حتى لا تضيع لاحقاً.')
   return redirect(url_for('company',cid=cid))
  raise
 d.close()
 if not contact_ledger_configured():
  flash('تم تسجيل التواصل محلياً، لكن التخزين الدائم غير مفعل. أضف CONTACTS_DATABASE_URL لحماية السجل من إعادة تشغيل Render.')
 return redirect(url_for('company',cid=cid))

@app.post('/company/<int:cid>/quick-followup')
@login_required
def quick_followup(cid):
 d=db(); c=scoped_company(d,cid)
 if not c: d.close(); abort(404)
 when=request.form.get('next_followup_at') or None
 notes=request.form.get('notes','متابعة سريعة')
 try:
  _save_interaction_durable(d,c,cid,'متابعة','followup',notes,when)
 except Exception:
  try:d.rollback()
  except:pass
  d.close()
  if contact_ledger_configured():
   flash('تعذر الحفظ الدائم للمتابعة. لم يتم اعتماد العملية حتى لا تضيع لاحقاً.')
   return redirect(request.referrer or url_for('companies'))
  raise
 d.close()
 if not contact_ledger_configured():
  flash('تم تسجيل المتابعة محلياً، لكن التخزين الدائم غير مفعل.')
 else:
  flash('تم تسجيل المتابعة وحفظها في السجل الدائم.')
 return redirect(request.referrer or url_for('companies'))

@app.get('/users')
@admin_required
def users():
 d=db(); rows=d.execute('SELECT id,name,email,role,active,specialty,whatsapp_template,appearance_default FROM users ORDER BY id').fetchall(); d.close(); appearance_map={}
 for u in rows:
  try: appearance_map[u['id']]=json.loads(u['appearance_default'] or '{}')
  except: appearance_map[u['id']]={}
 return render_template('users.html',rows=rows,default_messages=DEFAULT_MESSAGES,appearance_map=appearance_map)

def _appearance_from_form():
 def color(name, default):
  v=(request.form.get(name) or default).strip()
  return v if re.fullmatch(r'#[0-9a-fA-F]{6}',v) else default
 try: scale=max(85,min(160,int(request.form.get('appearance_scale','100'))))
 except: scale=100
 return json.dumps({'bg':color('appearance_bg','#f5f6f8'),'card':color('appearance_card','#ffffff'),'text':color('appearance_text','#212529'),'heading':color('appearance_heading','#111827'),'border':color('appearance_border','#dee2e6'),'nav':color('appearance_nav','#111827'),'scale':scale},ensure_ascii=False)

@app.post('/users/add')
@admin_required
def add_user():
 specialty=request.form.get('specialty','general'); template=request.form.get('whatsapp_template','').strip() or DEFAULT_MESSAGES.get(specialty,DEFAULT_MESSAGES['general']); appearance=_appearance_from_form(); d=db(); d.execute('INSERT INTO users(name,email,password_hash,role,specialty,whatsapp_template,appearance_default) VALUES(?,?,?,?,?,?,?)',(request.form['name'],request.form['email'],generate_password_hash(request.form['password']),request.form['role'],specialty,template,appearance)); d.commit(); d.close(); flash('تمت إضافة المستخدم مع مظهره الافتراضي'); return redirect(url_for('users'))
@app.post('/users/<int:uid>/update')
@admin_required
def update_user(uid):
 specialty=request.form.get('specialty','general'); template=request.form.get('whatsapp_template','').strip() or DEFAULT_MESSAGES.get(specialty,DEFAULT_MESSAGES['general']); appearance=_appearance_from_form(); d=db(); d.execute('UPDATE users SET specialty=?,whatsapp_template=?,appearance_default=? WHERE id=?',(specialty,template,appearance,uid)); d.commit(); d.close(); flash('تم تحديث المستخدم والمظهر الافتراضي'); return redirect(url_for('users'))

@app.post('/users/<int:uid>/reset-password')
@admin_required
def reset_user_password(uid):
 password=request.form.get('new_password') or ''
 confirm=request.form.get('confirm_password') or ''
 if len(password) < 10:
  flash('كلمة المرور الجديدة يجب ألا تقل عن 10 أحرف.'); return redirect(url_for('users'))
 if password != confirm:
  flash('تأكيد كلمة المرور غير مطابق.'); return redirect(url_for('users'))
 d=db()
 try:
  d.execute('BEGIN IMMEDIATE')
  u=d.execute('SELECT id,name FROM users WHERE id=? AND active=1',(uid,)).fetchone()
  if not u:
   d.rollback(); flash('المستخدم غير موجود أو غير نشط.'); return redirect(url_for('users'))
  d.execute('UPDATE users SET password_hash=? WHERE id=?',(generate_password_hash(password),uid))
  d.commit()
 except Exception:
  d.rollback(); raise
 finally:
  d.close()
 flash('تم تعيين كلمة مرور جديدة للمستخدم '+u['name'])
 return redirect(url_for('users'))

@app.post('/users/<int:uid>/delete')
@admin_required
def delete_user(uid):
 if uid==session.get('uid'):
  flash('لا يمكن حذف حساب المدير الذي تستخدمه حالياً.'); return redirect(url_for('users'))
 d=db(); u=d.execute('SELECT id,name FROM users WHERE id=?',(uid,)).fetchone()
 if not u: d.close(); abort(404)
 d.execute("UPDATE interactions SET user_name_snapshot=COALESCE(NULLIF(user_name_snapshot,''),?) WHERE user_id=?",(u['name'],uid))
 d.execute('UPDATE companies SET assigned_user_id=NULL WHERE assigned_user_id=?',(uid,))
 d.execute('DELETE FROM users WHERE id=?',(uid,))
 d.commit(); d.close(); flash(f"تم حذف المستخدم {u['name']} مع الاحتفاظ باسمه في سجل التواصل وإلغاء تعيين شركاته."); return redirect(url_for('users'))

@app.post('/company/<int:cid>/interaction/<int:iid>/delete')
@admin_required
def delete_interaction(cid,iid):
 d=db(); row=d.execute('SELECT * FROM interactions WHERE id=? AND company_id=?',(iid,cid)).fetchone()
 if not row: d.close(); abort(404)
 try:
  ev=_local_event_payload(d,row)
  if contact_ledger_configured():
   _ledger_tombstone(ev)
  d.execute('DELETE FROM interactions WHERE id=?',(iid,))
  d.commit()
 except Exception:
  d.rollback(); d.close()
  if contact_ledger_configured():
   flash('تعذر تسجيل الحذف في السجل الدائم، لذلك لم يتم حذف سجل التواصل.')
   return redirect(url_for('company',cid=cid))
  raise
 d.close()
 flash('تم حذف تسجيل التواصل الذي أضيف بالخطأ.')
 return redirect(url_for('company',cid=cid))

@app.route('/companies/add',methods=['GET','POST'])
@login_required
def add_company():
 d=db(); me=current_user(d)
 if not can_add_companies(me): d.close(); abort(403)
 if request.method=='POST':
  area=request.form.get('business_area','advertising'); area=area if area in ('contracting','advertising') else 'advertising'
  vals=[request.form.get(k,'').strip() for k in ('name_ar','name_en','sector','subcategory','city','district','address','phone_mobile','phone_landline','whatsapp','email','website','cr_number')]
  d.execute("INSERT INTO companies(name_ar,name_en,sector,subcategory,city,district,address,phone_mobile,phone_landline,whatsapp,email,website,cr_number,source,marketing_status,business_area,created_at,updated_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,CURRENT_TIMESTAMP,CURRENT_TIMESTAMP)",vals+['manual','new',area]); d.commit(); d.close(); flash('تمت إضافة الشركة بنجاح'); return redirect(url_for('companies'))
 d.close(); return render_template('company_add.html')

@app.route('/companies/import',methods=['GET','POST'])
@login_required
def import_companies():
 d=db(); me=current_user(d)
 if not can_add_companies(me): d.close(); abort(403)
 if request.method=='POST':
  f=request.files.get('file'); area=request.form.get('business_area','advertising'); area=area if area in ('contracting','advertising') else 'advertising'
  if not f or not f.filename: d.close(); flash('اختر ملفاً'); return redirect(request.url)
  ext=os.path.splitext(f.filename.lower())[1]; records=[]
  aliases={'اسم الشركة':'name_ar','الاسم':'name_ar','name_ar':'name_ar','name':'name_ar','company':'name_ar','name_en':'name_en','الاسم الانجليزي':'name_en','القطاع':'sector','sector':'sector','النشاط':'sector','التصنيف':'subcategory','subcategory':'subcategory','المدينة':'city','city':'city','الحي':'district','district':'district','العنوان':'address','address':'address','الجوال':'phone_mobile','الهاتف المحمول':'phone_mobile','phone_mobile':'phone_mobile','mobile':'phone_mobile','الهاتف':'phone_landline','phone':'phone_landline','phone_landline':'phone_landline','واتساب':'whatsapp','whatsapp':'whatsapp','البريد':'email','البريد الإلكتروني':'email','email':'email','الموقع':'website','website':'website','السجل التجاري':'cr_number','cr_number':'cr_number'}
  try:
   if ext=='.csv': records=list(csv.DictReader(io.StringIO(f.read().decode('utf-8-sig'))))
   elif ext in ('.xlsx','.xlsm'):
    ws=load_workbook(f,read_only=True,data_only=True).active; vals=list(ws.iter_rows(values_only=True)); headers=[str(x or '').strip() for x in vals[0]]; records=[dict(zip(headers,row)) for row in vals[1:]]
   else: raise ValueError('صيغة غير مدعومة. استخدم CSV أو XLSX')
   n=0
   for row in records:
    x={}
    for k,v in row.items():
     key=aliases.get(str(k or '').strip().lower()) or aliases.get(str(k or '').strip())
     if key: x[key]=str(v or '').strip()
    if not (x.get('name_ar') or x.get('name_en')): continue
    vals=[x.get(k,'') for k in ('name_ar','name_en','sector','subcategory','city','district','address','phone_mobile','phone_landline','whatsapp','email','website','cr_number')]
    d.execute("INSERT INTO companies(name_ar,name_en,sector,subcategory,city,district,address,phone_mobile,phone_landline,whatsapp,email,website,cr_number,source,marketing_status,business_area,created_at,updated_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,CURRENT_TIMESTAMP,CURRENT_TIMESTAMP)",vals+['import','new',area]); n+=1
   d.commit(); d.close(); flash(f'تم استيراد {n} شركة بنجاح'); return redirect(url_for('companies'))
  except Exception as e: d.rollback(); d.close(); flash('تعذر الاستيراد: '+str(e)); return redirect(request.url)
 d.close(); return render_template('company_import.html')



def _extract_logo_url(x):
 """Return a company logo/image URL only when the Muqawil JSON actually supplies one."""
 for key in ('logoUrl','logo_url','logo','imageUrl','image_url','companyLogo','company_logo','avatar','photo'):
  v=x.get(key)
  if isinstance(v,str) and v.strip().lower().startswith(('http://','https://')):
   return v.strip()
 # Some exporters nest media/image fields.
 for parent in ('media','image','companyImage','company_image'):
  obj=x.get(parent)
  if isinstance(obj,dict):
   for key in ('url','src','logoUrl','imageUrl'):
    v=obj.get(key)
    if isinstance(v,str) and v.strip().lower().startswith(('http://','https://')):
     return v.strip()
 return ''

def _muqawil_catalog_records(payload):
 if isinstance(payload,list): return payload
 if isinstance(payload,dict) and isinstance(payload.get('companies'),list): return payload['companies']
 raise ValueError('ملف مقاول يجب أن يحتوي companies أو قائمة شركات')

def _muqawil_status(v):
 v=str(v or '').strip()
 if v in ('مصنف','مصنفة'): return 'مصنفة'
 if v in ('غير مصنف','غير مصنفة'): return 'غير مصنفة'
 return 'غير محدد'

def _next_crm_number(d):
 return (d.execute('SELECT COALESCE(MAX(crm_number),0)+1 FROM companies').fetchone()[0] or 1)

def _upsert_muqawil_v42(d,x,next_no):
 """Update Muqawil directory facts only. Never import phone/WhatsApp from collector files."""
 key=str(x.get('muqawilKey') or x.get('muqawil_key') or '').strip()
 iid=str(x.get('muqawilInternalId') or x.get('muqawil_internal_id') or '').strip()
 ctype=str(x.get('muqawilContractorType') or x.get('muqawil_contractor_type') or '').strip()
 profile=str(x.get('muqawilProfileUrl') or x.get('sourceUrl') or x.get('source_url') or '').strip()
 if not key and iid and ctype: key=f'{iid}:{ctype}'
 if not key:
  m=re.search(r'/ar/contractors/(\d+)/(\d+)',profile)
  if m: iid,ctype=m.group(1),m.group(2); key=f'{iid}:{ctype}'
 name=str(x.get('nameAr') or x.get('name_ar') or '').strip()
 if not key or not name: return 'rejected',next_no
 member=str(x.get('membershipNo') or x.get('muqawil_member_no') or '').strip() or None
 row=d.execute('SELECT id FROM companies WHERE muqawil_key=? LIMIT 1',(key,)).fetchone()
 if not row and profile: row=d.execute('SELECT id FROM companies WHERE muqawil_profile_url=? LIMIT 1',(profile,)).fetchone()
 vals=(name,str(x.get('city') or ''),str(x.get('address') or ''),_muqawil_status(x.get('classificationStatus') or x.get('classification_status')),str(x.get('classificationGrade') or x.get('classification_grade') or ''),member,str(x.get('logoUrl') or x.get('logo_url') or ''),profile,profile,key,iid,ctype,str(x.get('companySize') or x.get('company_size') or ''),str(x.get('accountStatus') or x.get('account_status') or ''),x.get('trainingHours'),x.get('mainContractorCount'),x.get('subcontractorCount'),str(x.get('sourceRegionId') or x.get('source_region_id') or ''))
 if row:
  d.execute('''UPDATE companies SET name_ar=?,city=?,address=?,classification_status=?,classification_grade=?,muqawil_member_no=?,logo_url=?,source_url=?,muqawil_profile_url=?,muqawil_key=?,muqawil_internal_id=?,muqawil_contractor_type=?,company_size=?,account_status=?,training_hours=?,main_contractor_count=?,subcontractor_count=?,source_region_id=?,source_verified=1,data_source='Muqawil',updated_at=CURRENT_TIMESTAMP WHERE id=?''',vals+(row['id'],))
  return 'existing',next_no
 d.execute('''INSERT INTO companies(source_id,name_ar,sector,city,address,source,source_url,source_verified,marketing_status,priority,classification_status,classification_grade,muqawil_member_no,business_area,data_source,imported_at,logo_url,muqawil_profile_url,muqawil_profile_status,crm_number,muqawil_key,muqawil_internal_id,muqawil_contractor_type,company_size,account_status,training_hours,main_contractor_count,subcontractor_count,source_region_id,phone_mobile,phone_landline,whatsapp,phone_source,created_at,updated_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,CURRENT_TIMESTAMP,CURRENT_TIMESTAMP)''',
  (key,name,'مقاولات',str(x.get('city') or ''),str(x.get('address') or ''),'الهيئة السعودية للمقاولين (منصة مقاول)',profile,1,'new','normal',_muqawil_status(x.get('classificationStatus') or x.get('classification_status')),str(x.get('classificationGrade') or x.get('classification_grade') or ''),member,'contracting','Muqawil',datetime.now().strftime('%Y-%m-%d %H:%M:%S'),str(x.get('logoUrl') or x.get('logo_url') or ''),profile,'verified' if profile else 'missing',next_no,key,iid,ctype,str(x.get('companySize') or x.get('company_size') or ''),str(x.get('accountStatus') or x.get('account_status') or ''),x.get('trainingHours'),x.get('mainContractorCount'),x.get('subcontractorCount'),str(x.get('sourceRegionId') or x.get('source_region_id') or ''),'','','',''))
 return 'added',next_no+1

@app.route('/companies/import-muqawil',methods=['GET','POST'])
@admin_required
def import_muqawil_json():
 if request.method=='GET': return render_template('muqawil_import.html')
 f=request.files.get('file')
 if not f or not f.filename.lower().endswith('.json'):
  flash('اختر ملف JSON الخاص بمنصة مقاول.'); return redirect(request.url)
 d=None
 try:
  payload=json.load(f.stream); records=_muqawil_catalog_records(payload)
  d=db(); next_no=_next_crm_number(d); added=existing=rejected=0
  for x in records:
   if not isinstance(x,dict): rejected+=1; continue
   profile=str(x.get('muqawilProfileUrl') or x.get('sourceUrl') or '')
   if not (x.get('muqawilKey') or 'muqawil.org' in profile.lower()): rejected+=1; continue
   result,next_no=_upsert_muqawil_v42(d,x,next_no)
   if result=='added': added+=1
   elif result=='existing': existing+=1
   else: rejected+=1
  d.commit(); d.close(); d=None
  flash(f'تم تحديث دليل مقاول: جديد {added}، موجود/محدّث {existing}، مرفوض {rejected}. لم يتم استيراد أي أرقام هاتف من الملف.')
  return redirect(url_for('companies'))
 except Exception as e:
  if d: d.rollback(); d.close()
  flash('تعذر استيراد ملف مقاول: '+str(e)); return redirect(request.url)

@app.post('/admin/rebuild-muqawil-v42')
@admin_required
def rebuild_muqawil_v42():
 if request.form.get('confirm')!='REBUILD':
  flash('لم يتم التنفيذ. اكتب REBUILD للتأكيد.'); return redirect(url_for('import_muqawil_json'))
 path=os.path.join(BASE,'muqawil_v4_2_catalog.json')
 if not os.path.exists(path):
  flash('ملف دليل V4.2 المدمج غير موجود.'); return redirect(url_for('import_muqawil_json'))
 d=db()
 try:
  payload=json.load(open(path,encoding='utf-8')); records=_muqawil_catalog_records(payload)
  d.execute('DELETE FROM interactions'); d.execute('DELETE FROM companies')
  next_no=1; added=rejected=0
  for x in records:
   result,next_no=_upsert_muqawil_v42(d,x,next_no)
   if result=='added': added+=1
   elif result=='rejected': rejected+=1
  d.commit(); flash(f'تم إنشاء دليل مقاول V4.2: {added} شركة. المستخدمون والرسائل الجاهزة محفوظة، وأرقام الهواتف فارغة.')
 except Exception as e:
  d.rollback(); flash('تعذر إنشاء الدليل: '+str(e))
 finally: d.close()
 return redirect(url_for('companies'))



@app.post('/admin/reset-import-verified-muqawil')
@admin_required
def reset_import_verified_muqawil():
 if request.form.get('confirm')!='RESET15966':
  flash('لم يتم التنفيذ. اكتب RESET15966 للتأكيد.'); return redirect(url_for('import_muqawil_json'))
 f=request.files.get('file')
 if not f or not f.filename.lower().endswith('.json'):
  flash('اختر ملف JSON الموثق 15,966 شركة.'); return redirect(url_for('import_muqawil_json'))
 try:
  payload=json.load(f.stream); items=payload.get('companies',payload if isinstance(payload,list) else [])
 except Exception:
  flash('ملف JSON غير صالح.'); return redirect(url_for('import_muqawil_json'))
 if not isinstance(items,list) or len(items)!=15966:
  flash(f'تم إيقاف العملية: الملف يجب أن يحتوي بالضبط 15,966 شركة، والموجود {len(items) if isinstance(items,list) else 0}.'); return redirect(url_for('import_muqawil_json'))
 prepared=[]; seen=set(); duplicates=[]; rejected=[]
 for i,x in enumerate(items,1):
  if not isinstance(x,dict): rejected.append(i); continue
  key=str(x.get('muqawilKey') or '').strip(); name=str(x.get('nameAr') or '').strip(); url=str(x.get('profileUrl') or '').strip()
  m=re.search(r'/ar/contractors/(\d+)/(\d+)',url)
  if not key and m: key=f'{m.group(1)}:{m.group(2)}'
  if not key or not name or not m: rejected.append(i); continue
  if key in seen: duplicates.append(key); continue
  seen.add(key); prepared.append((x,key,name,url,m.group(1),m.group(2)))
 if duplicates or rejected or len(prepared)!=15966:
  flash(f'تم إيقاف العملية قبل الحذف: تكرار {len(duplicates)}، سجلات غير صالحة {len(rejected)}، فريدة صالحة {len(prepared)}.'); return redirect(url_for('import_muqawil_json'))
 d=db()
 try:
  d.execute('BEGIN IMMEDIATE')
  tables=[r['name'] for r in d.execute("SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%'").fetchall()]
  for table in tables:
   if table=='companies': continue
   cols={r['name'] for r in d.execute(f'PRAGMA table_info("{table}")').fetchall()}
   if 'company_id' in cols: d.execute(f'DELETE FROM "{table}"')
  d.execute('DELETE FROM companies')
  try:
   d.execute("DELETE FROM sqlite_sequence WHERE name IN ('companies','interactions')")
  except Exception: pass
  for no,(x,key,name,url,iid,ctype) in enumerate(prepared,1):
   mobile=str(x.get('phoneMobile') or '').strip(); landline=str(x.get('phoneLandline') or '').strip(); email=str(x.get('email') or '').strip()
   member=str(x.get('membershipNumber') or '').strip()
   d.execute("""INSERT INTO companies(source_id,name_ar,sector,city,address,source,source_url,source_verified,marketing_status,priority,
   classification_status,classification_grade,muqawil_member_no,business_area,data_source,imported_at,muqawil_profile_url,muqawil_profile_status,
   crm_number,muqawil_key,muqawil_internal_id,muqawil_contractor_type,company_size,account_status,training_hours,phone_mobile,phone_landline,whatsapp,
   email,phone_source,phone_verified_at,membership_type,membership_since,muqawil_region,muqawil_city,muqawil_address,muqawil_email,
   muqawil_phone_landline,muqawil_phone_mobile,muqawil_enriched_at,created_at,updated_at)
   VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,CURRENT_TIMESTAMP,CURRENT_TIMESTAMP,CURRENT_TIMESTAMP)""",
   (key,name,'مقاولات',str(x.get('city') or '').strip(),str(x.get('address') or '').strip(),'الهيئة السعودية للمقاولين (منصة مقاول)',url,1,
    'new','normal',_muqawil_status(x.get('classificationStatus')),str(x.get('classificationGrade') or '').strip(),member,
    'contracting','Muqawil',datetime.now().strftime('%Y-%m-%d %H:%M:%S'),url,'verified',no,key,iid,ctype,str(x.get('companySize') or '').strip(),
    str(x.get('accountStatus') or '').strip(),str(x.get('trainingHours') or '').strip(),mobile,landline,mobile,email,'Muqawil verified profile',
    str(x.get('fetchedAt') or '').strip() or None,str(x.get('membershipType') or '').strip(),str(x.get('membershipSince') or '').strip(),
    str(x.get('region') or '').strip(),str(x.get('city') or '').strip(),str(x.get('address') or '').strip(),email,landline,mobile))
  count=d.execute('SELECT COUNT(*) FROM companies').fetchone()[0]; uniq=d.execute('SELECT COUNT(DISTINCT muqawil_key) FROM companies').fetchone()[0]
  if count!=15966 or uniq!=15966: raise ValueError(f'فشل فحص ما بعد الاستيراد: الإجمالي {count} والفريد {uniq}')
  d.commit(); flash('تم حذف دليل الشركات السابق وسجل تواصله، ثم استيراد 15,966 شركة موثقة بدون تكرار. المستخدمون والقوالب محفوظة.')
 except Exception as e:
  d.rollback(); flash('تم التراجع عن العملية بالكامل ولم يتغير الدليل: '+str(e))
 finally: d.close()
 return redirect(url_for('companies'))


@app.get('/admin/muqawil-enrichment-queue')
@admin_required
def muqawil_enrichment_queue():
 d=db()
 rows=d.execute("""SELECT muqawil_key,muqawil_profile_url,source_url,name_ar,muqawil_member_no
                   FROM companies
                   WHERE COALESCE(muqawil_profile_url,source_url,'') LIKE '%muqawil.org/ar/contractors/%'
                   ORDER BY id""").fetchall(); d.close()
 payload={'version':'V'+APP_VERSION,'type':'muqawil_enrichment_queue','count':len(rows),'companies':[dict(x) for x in rows]}
 return app.response_class(json.dumps(payload,ensure_ascii=False,indent=2),mimetype='application/json',headers={'Content-Disposition':'attachment; filename=muqawil_enrichment_queue.json'})

@app.post('/admin/import-muqawil-enrichment')
@admin_required
def import_muqawil_enrichment():
 f=request.files.get('file')
 if not f: flash('اختر ملف إثراء JSON أولاً.'); return redirect(url_for('companies'))
 try: payload=json.load(f)
 except Exception: flash('ملف JSON غير صالح.'); return redirect(url_for('companies'))
 items=payload.get('companies',payload if isinstance(payload,list) else [])
 if not isinstance(items,list): flash('صيغة ملف الإثراء غير مدعومة.'); return redirect(url_for('companies'))
 d=db(); updated=missing=0
 for x in items:
  if not isinstance(x,dict): continue
  key=str(x.get('muqawilKey') or x.get('muqawil_key') or '').strip()
  url=str(x.get('profileUrl') or x.get('muqawil_profile_url') or '').strip()
  row=d.execute('SELECT id FROM companies WHERE muqawil_key=?',(key,)).fetchone() if key else None
  if not row and url: row=d.execute('SELECT id FROM companies WHERE muqawil_profile_url=? OR source_url=?',(url,url)).fetchone()
  if not row: missing+=1; continue
  # حقول مقاول فقط: لا نلمس حالة CRM أو المندوب أو سجل التواصل.
  vals={
   'membership_type':x.get('membershipType'),'membership_since':x.get('membershipSince'),
   'company_size':x.get('companySize'),'account_status':x.get('accountStatus'),'training_hours':x.get('trainingHours'),
   'muqawil_region':x.get('region'),'muqawil_city':x.get('city'),'muqawil_address':x.get('address'),
   'muqawil_email':x.get('email'),'muqawil_phone_landline':x.get('phoneLandline'),'muqawil_phone_mobile':x.get('phoneMobile'),
   'classification_status':x.get('classificationStatus'),'classification_grade':x.get('classificationGrade')}
  clean={k:(str(v).strip() if v is not None else '') for k,v in vals.items()}
  # املأ الحقول العامة فقط إذا كانت فارغة، مع إبقاء المصدر التفصيلي منفصلاً.
  sets=[]; params=[]
  for k,v in clean.items():
   if v!='': sets.append(f"{k}=?"); params.append(v)
  if clean.get('muqawil_city'): sets.append("city=CASE WHEN COALESCE(TRIM(city),'')='' THEN ? ELSE city END"); params.append(clean['muqawil_city'])
  if clean.get('muqawil_address'): sets.append("address=CASE WHEN COALESCE(TRIM(address),'')='' THEN ? ELSE address END"); params.append(clean['muqawil_address'])
  if clean.get('muqawil_email'): sets.append("email=CASE WHEN COALESCE(TRIM(email),'')='' THEN ? ELSE email END"); params.append(clean['muqawil_email'])
  if clean.get('muqawil_phone_landline'): sets.append("phone_landline=CASE WHEN COALESCE(TRIM(phone_landline),'')='' THEN ? ELSE phone_landline END"); params.append(clean['muqawil_phone_landline'])
  if clean.get('muqawil_phone_mobile'): sets.append("phone_mobile=CASE WHEN COALESCE(TRIM(phone_mobile),'')='' THEN ? ELSE phone_mobile END"); params.append(clean['muqawil_phone_mobile'])
  if sets:
   sets += ["muqawil_enriched_at=CURRENT_TIMESTAMP","updated_at=CURRENT_TIMESTAMP"]
   d.execute('UPDATE companies SET '+','.join(sets)+' WHERE id=?',params+[row['id']]); updated+=1
 d.commit(); d.close(); flash(f'تم إثراء {updated} شركة من صفحات مقاول. غير موجودة في الدليل: {missing}.'); return redirect(url_for('companies'))



@app.get('/admin/membership-audit')
@admin_required
def membership_audit():
 d=db()
 groups=d.execute("""SELECT muqawil_member_no,COUNT(*) n
 FROM companies WHERE COALESCE(TRIM(muqawil_member_no),'')<>''
 GROUP BY muqawil_member_no HAVING COUNT(*)>1 ORDER BY n DESC,muqawil_member_no""").fetchall()
 rows=[]
 for g in groups:
  cs=d.execute("""SELECT id,name_ar,muqawil_key,muqawil_profile_url,source_url
                  FROM companies WHERE muqawil_member_no=? ORDER BY id""",(g['muqawil_member_no'],)).fetchall()
  rows.append({'member':g['muqawil_member_no'],'count':g['n'],'companies':[dict(x) for x in cs]})
 d.close()
 return render_template('membership_audit.html',groups=rows,total_groups=len(rows),total_companies=sum(x['count'] for x in rows),is_admin=True)

@app.get('/admin/membership-audit/queue')
@admin_required
def membership_audit_queue():
 d=db()
 rows=d.execute("""SELECT c.muqawil_key,c.name_ar,c.muqawil_member_no,c.muqawil_profile_url,c.source_url
 FROM companies c JOIN (
  SELECT muqawil_member_no FROM companies WHERE COALESCE(TRIM(muqawil_member_no),'')<>''
  GROUP BY muqawil_member_no HAVING COUNT(*)>1
 ) x ON x.muqawil_member_no=c.muqawil_member_no
 ORDER BY c.muqawil_member_no,c.id""").fetchall()
 d.close()
 payload={'version':'MembershipAudit-1','type':'muqawil_membership_audit_queue','count':len(rows),
          'companies':[{'muqawilKey':r['muqawil_key'],'nameAr':r['name_ar'],'currentMembershipNo':r['muqawil_member_no'],
                        'profileUrl':r['muqawil_profile_url'] or r['source_url']} for r in rows]}
 return Response(json.dumps(payload,ensure_ascii=False,indent=2),mimetype='application/json',
                 headers={'Content-Disposition':'attachment; filename=muqawil_membership_audit_queue.json'})

@app.post('/admin/membership-audit/import')
@admin_required
def membership_audit_import():
 f=request.files.get('file')
 if not f: flash('اختر ملف تصحيح العضويات JSON.'); return redirect(url_for('membership_audit'))
 try: payload=json.load(f)
 except Exception: flash('ملف JSON غير صالح.'); return redirect(url_for('membership_audit'))
 if not isinstance(payload,dict) or payload.get('type')!='muqawil_membership_patch':
  flash('تم رفض الملف: النوع المطلوب muqawil_membership_patch.'); return redirect(url_for('membership_audit'))
 items=payload.get('companies')
 if not isinstance(items,list): flash('تم رفض الملف: companies غير صالح.'); return redirect(url_for('membership_audit'))
 seen=set(); clean=[]
 for i,x in enumerate(items,1):
  key=str(x.get('muqawilKey') or '').strip()
  member=str(x.get('membershipNo') or '').strip()
  if not key or key in seen or (member and not re.fullmatch(r'[0-9A-Za-z-]{3,40}',member)):
   flash(f'تم رفض الملف قبل التعديل عند السجل {i}.'); return redirect(url_for('membership_audit'))
  seen.add(key); clean.append((key,member))
 d=db()
 try:
  backup_dir=os.path.join(DATA_DIR,'backups'); os.makedirs(backup_dir,exist_ok=True)
  bp=os.path.join(backup_dir,'crm_before_membership_'+datetime.now().strftime('%Y%m%d_%H%M%S')+'.db')
  b=sqlite3.connect(bp); d.backup(b); b.close()
  d.execute('BEGIN IMMEDIATE'); updated=missing=0
  for key,member in clean:
   row=d.execute('SELECT id FROM companies WHERE muqawil_key=?',(key,)).fetchone()
   if not row: missing+=1; continue
   d.execute('UPDATE companies SET muqawil_member_no=?,updated_at=CURRENT_TIMESTAMP WHERE id=?',(member or None,row['id']))
   updated+=1
  d.commit(); flash(f'تم تصحيح {updated} عضوية. غير موجودة: {missing}. تم إنشاء نسخة احتياطية قبل التعديل.')
 except Exception as e:
  d.rollback(); flash('تم التراجع بالكامل: '+str(e))
 finally: d.close()
 return redirect(url_for('membership_audit'))


def _db_snapshot_info(path):
 try:
  st=os.stat(path)
  x=sqlite3.connect(path,timeout=10); x.row_factory=sqlite3.Row
  n=x.execute('SELECT COUNT(*) FROM companies').fetchone()[0]
  mobiles=x.execute("""SELECT COUNT(*) FROM companies WHERE COALESCE(TRIM(phone_mobile),'')<>''""").fetchone()[0]
  users=x.execute('SELECT COUNT(*) FROM users').fetchone()[0]
  interactions=x.execute('SELECT COUNT(*) FROM interactions').fetchone()[0]
  x.close()
  return {'path':path,'size':st.st_size,'modified':datetime.fromtimestamp(st.st_mtime).isoformat(timespec='seconds'),
          'companies':n,'mobiles':mobiles,'users':users,'interactions':interactions}
 except Exception as e: return {'path':path,'error':str(e)}



def _recovery_contact_columns(conn):
    cols=[r['name'] for r in conn.execute("PRAGMA table_info(companies)").fetchall()]
    wanted=[]
    for c in ('phone_mobile','whatsapp','muqawil_phone_mobile'):
        if c in cols and c not in wanted: wanted.append(c)
    for c in cols:
        lc=c.lower()
        if c not in wanted and ('mobile' in lc or 'whatsapp' in lc):
            wanted.append(c)
    return wanted

def _norm_sa_mobile(v):
    d=re.sub(r'\D','',str(v or ''))
    if d.startswith('00966'): d=d[5:]
    elif d.startswith('966'): d=d[3:]
    if re.fullmatch(r'5\d{8}',d): d='0'+d
    return d if re.fullmatch(r'05\d{8}',d) else ''

@app.get('/admin/recovery-center')
@admin_required
def recovery_center():
    d=db()
    try:
        companies=d.execute('SELECT COUNT(*) FROM companies').fetchone()[0]
        interactions=d.execute('SELECT COUNT(*) FROM interactions').fetchone()[0]
        cols=_recovery_contact_columns(d)
        field_counts={}
        for c in cols:
            q='"'+c.replace('"','""')+'"'
            field_counts[c]=d.execute("SELECT COUNT(*) FROM companies WHERE COALESCE(TRIM("+q+"),'')<>''").fetchone()[0]
        # Same definition used by company sorting: WhatsApp OR phone_mobile.
        protected_any=d.execute("""SELECT COUNT(*) FROM companies
          WHERE COALESCE(TRIM(phone_mobile),'')<>'' OR COALESCE(TRIM(whatsapp),'')<>''""").fetchone()[0]
    finally:d.close()
    pending=os.path.join(DATA_DIR,'pending_map_phone_13288.json')
    return render_template('recovery_center.html',is_admin=True,
      live={'companies':companies,'mobiles':protected_any,'interactions':interactions},
      pending_ready=os.path.isfile(pending),actual_contact_cols=cols,field_counts=field_counts)

@app.post('/admin/recovery-map-upload')
@admin_required
def recovery_map_upload():
    f=request.files.get('file')
    if not f:
        flash('اختر ملف JSON أولًا.'); return redirect(url_for('recovery_center'))
    try: payload=json.load(f)
    except Exception as e:
        flash('ملف JSON غير صالح: '+str(e)); return redirect(url_for('recovery_center'))
    items=payload.get('companies')
    if payload.get('type')!='muqawil_map_phone_patch' or payload.get('rule')!='fill-empty-only' or not isinstance(items,list) or len(items)!=13288:
        flash('تم رفض الملف: يجب أن يكون ملف جوالات خريطة مقاول النهائي 13,288 وبقاعدة fill-empty-only.')
        return redirect(url_for('recovery_center'))
    clean=[]; seen=set()
    for i,x in enumerate(items,1):
        key=str(x.get('muqawilKey') or x.get('muqawil_key') or '').strip()
        mob=_norm_sa_mobile(x.get('phoneMobile') or x.get('phone_mobile'))
        if not key or key in seen or not mob:
            flash(f'تم رفض الملف عند السجل {i}: مفتاح مكرر/ناقص أو جوال غير صالح.')
            return redirect(url_for('recovery_center'))
        seen.add(key); clean.append((key,mob))

    d=db()
    try:
        cols=_recovery_contact_columns(d)
        select_cols=','.join(['muqawil_key']+['"'+c.replace('"','""')+'"' for c in cols])
        rows=d.execute("SELECT "+select_cols+" FROM companies WHERE COALESCE(TRIM(muqawil_key),'')<>''").fetchall()
        live={}
        for r in rows:
            live[str(r['muqawil_key']).strip()]={c:str(r[c] or '').strip() for c in cols}
    finally:d.close()

    matched=sum(k in live for k,m in clean)
    protected=sum(k in live and any(live[k].values()) for k,m in clean)
    eligible=sum(k in live and not any(live[k].values()) for k,m in clean)
    missing=len(clean)-matched
    per_field={c:{
      'nonempty':sum(k in live and bool(live[k].get(c)) for k,m in clean),
      'valid':sum(k in live and bool(_norm_sa_mobile(live[k].get(c))) for k,m in clean)
    } for c in cols}

    pending=os.path.join(DATA_DIR,'pending_map_phone_13288.json')
    tmp=pending+'.tmp'
    with open(tmp,'w',encoding='utf-8') as out: json.dump(payload,out,ensure_ascii=False)
    os.replace(tmp,pending)
    return render_template('recovery_preview.html',is_admin=True,total=len(clean),matched=matched,
      protected=protected,eligible=eligible,missing=missing,per_field=per_field,contact_cols=cols)


@app.get('/admin/recovery-map-audit')
@admin_required
def recovery_map_audit():
    pending=os.path.join(DATA_DIR,'pending_map_phone_13288.json')
    if not os.path.isfile(pending):
        flash('ارفع ملف 13,288 وافحصه أولًا.'); return redirect(url_for('recovery_center'))
    try:
        with open(pending,encoding='utf-8') as f: payload=json.load(f)
    except Exception as e:
        flash('تعذر قراءة ملف الفحص: '+str(e)); return redirect(url_for('recovery_center'))
    items=payload.get('companies') or []
    patch_keys={str(x.get('muqawilKey') or x.get('muqawil_key') or '').strip() for x in items}
    patch_keys.discard('')
    d=db()
    try:
        cols={r['name'] for r in d.execute("PRAGMA table_info(companies)").fetchall()}
        has_mpm='muqawil_phone_mobile' in cols
        extra=",muqawil_phone_mobile" if has_mpm else ""
        rows=d.execute("""SELECT id,name_ar,muqawil_key,phone_mobile,whatsapp"""+extra+"""
                          FROM companies""").fetchall()
    finally:d.close()

    def has_contact(r):
        vals=[r['phone_mobile'],r['whatsapp']]
        if has_mpm: vals.append(r['muqawil_phone_mobile'])
        return any(str(v or '').strip() for v in vals)

    with_contact=[r for r in rows if has_contact(r)]
    contact_with_key=[r for r in with_contact if str(r['muqawil_key'] or '').strip()]
    contact_no_key=[r for r in with_contact if not str(r['muqawil_key'] or '').strip()]
    overlap=[r for r in contact_with_key if str(r['muqawil_key']).strip() in patch_keys]
    contact_outside=[r for r in contact_with_key if str(r['muqawil_key']).strip() not in patch_keys]

    live_keys={str(r['muqawil_key'] or '').strip() for r in rows if str(r['muqawil_key'] or '').strip()}
    patch_missing_live=patch_keys-live_keys

    def sample(rs,n=20):
        return [{'id':r['id'],'name':r['name_ar'] or '',
                 'key':str(r['muqawil_key'] or '').strip(),
                 'phone':str(r['phone_mobile'] or '').strip(),
                 'whatsapp':str(r['whatsapp'] or '').strip(),
                 'muqawil_phone':str(r['muqawil_phone_mobile'] or '').strip() if has_mpm else ''} for r in rs[:n]]

    return render_template('recovery_audit.html',is_admin=True,
      total_companies=len(rows),patch_total=len(patch_keys),with_contact=len(with_contact),
      contact_with_key=len(contact_with_key),contact_no_key=len(contact_no_key),
      overlap=len(overlap),contact_outside=len(contact_outside),
      patch_missing_live=len(patch_missing_live),
      sample_overlap=sample(overlap),sample_outside=sample(contact_outside),
      sample_no_key=sample(contact_no_key))

@app.post('/admin/recovery-map-commit')
@admin_required
def recovery_map_commit():
    if request.form.get('confirm')!='RECOVER13288':
        flash('لم يتم التنفيذ. اكتب RECOVER13288 للتأكيد.'); return redirect(url_for('recovery_center'))
    pending=os.path.join(DATA_DIR,'pending_map_phone_13288.json')
    if not os.path.isfile(pending):
        flash('لا يوجد ملف متحقق منه بانتظار التنفيذ.'); return redirect(url_for('recovery_center'))
    try:
        with open(pending,encoding='utf-8') as f: payload=json.load(f)
    except Exception as e:
        flash('تعذر قراءة الملف المؤقت: '+str(e)); return redirect(url_for('recovery_center'))
    items=payload.get('companies')
    if payload.get('type')!='muqawil_map_phone_patch' or payload.get('rule')!='fill-empty-only' or not isinstance(items,list) or len(items)!=13288:
        flash('فشل التحقق النهائي؛ لم يتم تعديل قاعدة البيانات.'); return redirect(url_for('recovery_center'))
    clean=[]; seen=set()
    for i,x in enumerate(items,1):
        key=str(x.get('muqawilKey') or x.get('muqawil_key') or '').strip()
        mob=_norm_sa_mobile(x.get('phoneMobile') or x.get('phone_mobile'))
        if not key or key in seen or not mob:
            flash(f'فشل التحقق النهائي عند السجل {i}.'); return redirect(url_for('recovery_center'))
        seen.add(key); clean.append((key,mob))

    d=db()
    try:
        backup_dir=os.path.join(DATA_DIR,'backups'); os.makedirs(backup_dir,exist_ok=True)
        bp=os.path.join(backup_dir,'crm_before_phone_recovery_'+datetime.now().strftime('%Y%m%d_%H%M%S')+'.db')
        b=sqlite3.connect(bp); d.backup(b); b.close()
        cols=_recovery_contact_columns(d)
        select_cols=','.join(['id','muqawil_key']+['"'+c.replace('"','""')+'"' for c in cols])
        rows=d.execute("SELECT "+select_cols+" FROM companies WHERE COALESCE(TRIM(muqawil_key),'')<>''").fetchall()
        live={}
        for r in rows:
            live[str(r['muqawil_key']).strip()]=(r['id'],{c:str(r[c] or '').strip() for c in cols})
        d.execute('BEGIN IMMEDIATE'); updated=skipped=missing=0
        for key,mob in clean:
            row=live.get(key)
            if not row: missing+=1; continue
            cid,existing=row
            # Critical lock: ANY existing mobile/WhatsApp-like value wins, even legacy formatting.
            if any(existing.values()): skipped+=1; continue
            d.execute("""UPDATE companies SET phone_mobile=?,
              whatsapp=CASE WHEN COALESCE(TRIM(whatsapp),'')='' THEN ? ELSE whatsapp END,
              phone_source='Muqawil Company Map',phone_verified_at=CURRENT_TIMESTAMP,updated_at=CURRENT_TIMESTAMP
              WHERE id=? AND COALESCE(TRIM(phone_mobile),'')='' AND COALESCE(TRIM(whatsapp),'')=''""",(mob,mob,cid))
            updated+=d.execute('SELECT changes()').fetchone()[0]
        d.commit()
        try: os.remove(pending)
        except OSError: pass
        flash(f'تمت الاستعادة بأمان: أضيف {updated}، حُمي {skipped} سجل موجود، وغير مطابق {missing}. Backup محفوظ قبل التنفيذ.')
    except Exception as e:
        try:d.rollback()
        except:pass
        flash('تم التراجع بالكامل: '+str(e))
    finally:d.close()
    return redirect(url_for('recovery_center'))

@app.get('/admin/data-lock')
@admin_required
def data_lock_status():
 live=_db_snapshot_info(DB)
 backup_dir=os.path.join(DATA_DIR,'backups'); os.makedirs(backup_dir,exist_ok=True)
 backups=[]
 for fn in sorted(os.listdir(backup_dir),reverse=True):
  if fn.endswith('.db'):
   info=_db_snapshot_info(os.path.join(backup_dir,fn)); info['name']=fn; backups.append(info)
 return render_template('data_lock.html',is_admin=True,live=live,backups=backups[:50],
                        data_dir=DATA_DIR,db_path=DB,is_render=IS_RENDER,lock_exists=os.path.exists(DATA_LOCK))

@app.get('/admin/data-lock/backup')
@admin_required
def data_lock_backup():
 backup_dir=os.path.join(DATA_DIR,'backups'); os.makedirs(backup_dir,exist_ok=True)
 name='crm_manual_'+datetime.now().strftime('%Y%m%d_%H%M%S')+'.db'; path=os.path.join(backup_dir,name)
 src=sqlite3.connect(DB,timeout=30); dst=sqlite3.connect(path)
 try: src.backup(dst)
 finally: dst.close(); src.close()
 return send_file(path,as_attachment=True,download_name=name)

@app.get('/admin/data-lock/backup/<path:name>')
@admin_required
def data_lock_download(name):
 if '/' in name or '\\\\' in name or not name.endswith('.db'): abort(400)
 path=os.path.join(DATA_DIR,'backups',name)
 if not os.path.isfile(path): abort(404)
 return send_file(path,as_attachment=True,download_name=name)

@app.post('/admin/data-lock/restore')
@admin_required
def data_lock_restore():
 name=request.form.get('name','')
 confirm=request.form.get('confirm','')
 if '/' in name or '\\\\' in name or not name.endswith('.db') or confirm!='RESTORE':
  flash('لم تتم الاستعادة: يلزم اختيار نسخة وكتابة RESTORE.'); return redirect(url_for('data_lock_status'))
 src_path=os.path.join(DATA_DIR,'backups',name)
 if not os.path.isfile(src_path): flash('النسخة غير موجودة.'); return redirect(url_for('data_lock_status'))
 # Validate backup before touching live DB.
 info=_db_snapshot_info(src_path)
 if info.get('error'): flash('النسخة غير صالحة: '+info['error']); return redirect(url_for('data_lock_status'))
 safety=os.path.join(DATA_DIR,'backups','crm_before_restore_'+datetime.now().strftime('%Y%m%d_%H%M%S')+'.db')
 live=sqlite3.connect(DB,timeout=30); saf=sqlite3.connect(safety)
 try: live.backup(saf)
 finally: saf.close(); live.close()
 source=sqlite3.connect(src_path,timeout=30); target=sqlite3.connect(DB,timeout=30)
 try: source.backup(target)
 finally: target.close(); source.close()
 flash(f'تمت استعادة {name}. تم حفظ نسخة من الحالة السابقة تلقائيًا.')
 return redirect(url_for('data_lock_status'))


@app.post('/admin/import-full-reaudit')
@admin_required
def import_full_reaudit():
 f=request.files.get('file')
 if not f: flash('اختر ملف muqawil_PLATFORM_FINAL أو muqawil_reaudit_FINAL.'); return redirect(url_for('recovery_center'))
 try: payload=json.load(f)
 except Exception: flash('ملف JSON غير صالح.'); return redirect(url_for('recovery_center'))
 if payload.get('type') not in ('muqawil_platform_full_crawl','muqawil_full_reaudit'):
  flash('تم رفض الملف: ليس ناتج تدقيق مقاول الشامل.'); return redirect(url_for('recovery_center'))
 items=payload.get('companies')
 if not isinstance(items,list) or not items: flash('لا توجد شركات موثقة في الملف.'); return redirect(url_for('recovery_center'))
 d=db()
 try:
  backup_dir=os.path.join(DATA_DIR,'backups'); os.makedirs(backup_dir,exist_ok=True)
  bp=os.path.join(backup_dir,'crm_before_full_reaudit_'+datetime.now().strftime('%Y%m%d_%H%M%S')+'.db')
  b=sqlite3.connect(bp); d.backup(b); b.close()
  d.execute('BEGIN IMMEDIATE'); updated=missing=kept_mobile=0; seen=set()
  for i,x in enumerate(items,1):
   key=str(x.get('muqawilKey') or '').strip()
   if not key or key in seen: raise ValueError(f'muqawilKey غير صالح/مكرر عند {i}')
   seen.add(key)
   row=d.execute('SELECT id,phone_mobile FROM companies WHERE muqawil_key=? LIMIT 1',(key,)).fetchone()
   if not row: missing+=1; continue
   # Verified descriptive fields may be refreshed, but live phone always wins.
   fields={'muqawil_member_no':x.get('membershipNo'),'email':x.get('email'),'phone_landline':x.get('phoneLandline'),
           'classification_status':x.get('classificationStatus'),'classification_grade':x.get('classificationGrade'),
           'region':x.get('region'),'city':x.get('city'),'sector':x.get('sector'),'address':x.get('address')}
   sets=[]; vals=[]
   for col,val in fields.items():
    if val is not None and str(val).strip()!='': sets.append(col+'=?'); vals.append(str(val).strip())
   mob=re.sub(r'\D','',str(x.get('phoneMobile') or ''))
   if not (row['phone_mobile'] or '').strip() and re.fullmatch(r'05\d{8}',mob):
    sets+=['phone_mobile=?','muqawil_phone_mobile=?','phone_source=?']; vals += [mob,mob,'Muqawil Contractor Profile']
   elif (row['phone_mobile'] or '').strip(): kept_mobile+=1
   if sets:
    vals.append(row['id']); d.execute('UPDATE companies SET '+','.join(sets)+',updated_at=CURRENT_TIMESTAMP WHERE id=?',vals); updated+=1
  d.commit()
  flash(f'تم دمج التدقيق الشامل بأمان: حدّث {updated} شركة، حافظ على {kept_mobile} جوال حي، ولم يجد {missing} مفتاحًا في الدليل الحالي. لم يتم حذف أي شركة.')
 except Exception as e:
  d.rollback(); flash('تم التراجع بالكامل: '+str(e))
 finally: d.close()
 return redirect(url_for('recovery_center'))

@app.route('/admin/system-update', methods=['GET','POST'])
@admin_required
def system_update():
 if request.method=='GET':
  return render_template('system_update.html', is_admin=True)
 f=request.files.get('file')
 if not f or not (f.filename or '').lower().endswith('.zip'):
  flash('اختر ملف تحديث ZIP صالح.'); return redirect(url_for('system_update'))
 import zipfile, shutil, py_compile, threading, signal
 update_dir=os.path.join(DATA_DIR,'app_updates'); os.makedirs(update_dir,exist_ok=True)
 stage=tempfile.mkdtemp(prefix='crm_update_')
 uploaded=os.path.join(stage,'update.zip'); f.save(uploaded)
 try:
  with zipfile.ZipFile(uploaded) as z:
   members=z.infolist()
   if not members: raise ValueError('ملف التحديث فارغ')
   for m in members:
    n=m.filename.replace('\\','/')
    if n.startswith('/') or '..' in n.split('/'):
     raise ValueError('مسار غير آمن داخل ملف التحديث')
   z.extractall(stage)
  candidates=[]
  for p in [os.path.join(stage,'Alrawasi_CRM_OneClick_Deploy'),stage]:
   if os.path.isfile(os.path.join(p,'app.py')): candidates.append(p)
  if not candidates:
   for base,dirs,files in os.walk(stage):
    if 'app.py' in files and 'templates' in dirs: candidates.append(base); break
  if not candidates: raise ValueError('الحزمة لا تحتوي مشروع CRM صالحًا')
  src=candidates[0]
  py_compile.compile(os.path.join(src,'app.py'),doraise=True)
  version_file=os.path.join(src,'VERSION')
  if not os.path.isfile(version_file): raise ValueError('ملف VERSION غير موجود')
  new_version=open(version_file,encoding='utf-8').read().strip()
  if not re.fullmatch(r'\d+\.\d+\.\d+',new_version): raise ValueError('رقم الإصدار غير صالح')
  # لا يسمح ملف التحديث باستبدال قاعدة البيانات الحية أو ملفات البيانات.
  blocked={'crm.db','.env','backups','app_updates','.alrawasi_data_initialized'}
  # نسخة DB تلقائية قبل أي تحديث كود.
  backup_dir=os.path.join(DATA_DIR,'backups'); os.makedirs(backup_dir,exist_ok=True)
  bp=os.path.join(backup_dir,'crm_before_code_update_'+datetime.now().strftime('%Y%m%d_%H%M%S')+'.db')
  _src=sqlite3.connect(DB,timeout=30); _dst=sqlite3.connect(bp)
  try: _src.backup(_dst)
  finally: _dst.close(); _src.close()
  current=os.path.join(update_dir,'current.zip'); previous=os.path.join(update_dir,'previous.zip')
  if os.path.exists(current): shutil.copy2(current,previous)
  shutil.copy2(uploaded,current)
  for name in os.listdir(src):
   if name in blocked or name.startswith('.git') or name=='__pycache__': continue
   a=os.path.join(src,name); b=os.path.join(BASE,name)
   if os.path.isdir(a):
    if name=='templates':
     if os.path.exists(b): shutil.rmtree(b)
     shutil.copytree(a,b)
   else: shutil.copy2(a,b)
  flash(f'تم تثبيت تحديث CRM V{new_version} وحفظه على القرص الدائم. سيتم إعادة تحميل النظام تلقائيًا. لم يتم استبدال قاعدة البيانات.')
  threading.Timer(1.0, lambda: os.kill(os.getppid(), signal.SIGHUP)).start()
  return redirect(url_for('system_update'))
 except Exception as e:
  flash('تم رفض التحديث ولم يتم تثبيته: '+str(e)); return redirect(url_for('system_update'))
 finally:
  shutil.rmtree(stage,ignore_errors=True)


@app.post('/admin/import-muqawil-map-phones')
@admin_required
def import_muqawil_map_phones():
 f=request.files.get('file')
 if not f: flash('اختر ملف muqawil_map_phone_FINAL JSON أولاً.'); return redirect(url_for('companies'))
 try: payload=json.load(f)
 except Exception: flash('ملف JSON غير صالح.'); return redirect(url_for('companies'))
 if not isinstance(payload,dict) or payload.get('type')!='muqawil_map_phone_patch' or payload.get('rule')!='fill-empty-only':
  flash('تم رفض الملف: يجب أن يكون ملف Map Phone Collector النهائي وبقاعدة fill-empty-only.'); return redirect(url_for('companies'))
 items=payload.get('companies')
 if not isinstance(items,list) or not items: flash('تم رفض الملف: لا توجد شركات صالحة للاستيراد.'); return redirect(url_for('companies'))
 mobile_re=re.compile(r'^05\d{8}$'); seen=set(); clean=[]
 for i,x in enumerate(items,1):
  if not isinstance(x,dict): flash(f'تم رفض الملف قبل أي تعديل: السجل رقم {i} غير صالح.'); return redirect(url_for('companies'))
  key=str(x.get('muqawilKey') or x.get('muqawil_key') or '').strip()
  mobile=re.sub(r'\D','',str(x.get('phoneMobile') or x.get('phone_mobile') or ''))
  if not key or key in seen or not mobile_re.fullmatch(mobile):
   flash(f'تم رفض الملف قبل أي تعديل: خطأ في muqawilKey أو الجوال عند السجل رقم {i}.'); return redirect(url_for('companies'))
  seen.add(key); clean.append((key,mobile))
 d=db()
 try:
  backup_dir=os.path.join(DATA_DIR,'backups'); os.makedirs(backup_dir,exist_ok=True)
  stamp=datetime.now().strftime('%Y%m%d_%H%M%S'); backup_path=os.path.join(backup_dir,f'crm_before_map_phone_{stamp}.db')
  b=sqlite3.connect(backup_path); d.backup(b); b.close()
  d.execute('BEGIN IMMEDIATE'); updated=already=missing=0
  for key,mobile in clean:
   row=d.execute('SELECT id,phone_mobile FROM companies WHERE muqawil_key=? LIMIT 1',(key,)).fetchone()
   if not row: missing+=1; continue
   if (row['phone_mobile'] or '').strip(): already+=1; continue
   d.execute("""UPDATE companies SET phone_mobile=?,
                muqawil_phone_mobile=CASE WHEN COALESCE(TRIM(muqawil_phone_mobile),'')='' THEN ? ELSE muqawil_phone_mobile END,
                phone_source='Muqawil Company Map',updated_at=CURRENT_TIMESTAMP
                WHERE id=? AND COALESCE(TRIM(phone_mobile),'')=''""",(mobile,mobile,row['id']))
   updated += d.execute('SELECT changes()').fetchone()[0]
  d.commit()
  flash(f'تم استيراد جوالات خريطة مقاول بأمان: أضيف {updated}، موجود مسبقًا ولم يُستبدل {already}، غير موجود في الدليل {missing}. تم إنشاء نسخة احتياطية تلقائية قبل التحديث.')
 except Exception as e:
  d.rollback(); flash('تم التراجع عن العملية بالكامل ولم تُحفظ أي تغييرات: '+str(e))
 finally: d.close()
 return redirect(url_for('companies'))


def _google_places_phone(company):
 key=(os.environ.get('GOOGLE_MAPS_API_KEY') or '').strip()
 if not key: return {'ok':False,'reason':'GOOGLE_MAPS_API_KEY غير مضبوط'}
 name=(company['name_ar'] or company['name_en'] or '').strip()
 city=(company['city'] or '').strip()
 address=(company['address'] or '').strip()
 query='، '.join(x for x in (name,address,city,'السعودية') if x)
 payload=json.dumps({'textQuery':query,'languageCode':'ar','regionCode':'SA','maxResultCount':3}).encode('utf-8')
 req=urllib.request.Request('https://places.googleapis.com/v1/places:searchText',data=payload,method='POST',
   headers={'Content-Type':'application/json','X-Goog-Api-Key':key,
            'X-Goog-FieldMask':'places.id,places.displayName,places.formattedAddress,places.nationalPhoneNumber,places.internationalPhoneNumber'})
 try:
  with urllib.request.urlopen(req,timeout=20) as r: data=json.load(r)
 except Exception as e:
  return {'ok':False,'reason':str(e)}
 places=data.get('places') or []
 if not places: return {'ok':False,'reason':'لا توجد نتيجة مطابقة'}
 # لا نختار نتيجة بعيدة: الاسم + المدينة/العنوان يجب أن يساندا المطابقة.
 def norm(s): return re.sub(r'[\W_]+','',str(s or '').lower(),flags=re.UNICODE)
 nn=norm(name); cc=norm(city)
 best=None
 for p in places:
  pn=norm((p.get('displayName') or {}).get('text',''))
  pa=norm(p.get('formattedAddress',''))
  name_ok=(nn and (nn in pn or pn in nn))
  city_ok=(not cc) or (cc in pa)
  phone=p.get('internationalPhoneNumber') or p.get('nationalPhoneNumber')
  if name_ok and city_ok and phone:
   best=p; break
 if not best: return {'ok':False,'reason':'المطابقة غير كافية أو لا يوجد هاتف'}
 raw_phone=best.get('internationalPhoneNumber') or best.get('nationalPhoneNumber') or ''
 digits=re.sub(r'\D','',raw_phone)
 if digits.startswith('00966'): digits=digits[2:]
 if digits.startswith('966'): digits=digits[3:]
 if re.fullmatch(r'5\d{8}',digits): normalized='0'+digits; phone_kind='mobile'
 elif re.fullmatch(r'05\d{8}',digits): normalized=digits; phone_kind='mobile'
 elif re.fullmatch(r'1[1-7]\d{7}',digits): normalized='0'+digits; phone_kind='landline'
 elif re.fullmatch(r'01[1-7]\d{7}',digits): normalized=digits; phone_kind='landline'
 else: return {'ok':False,'reason':'تم العثور على هاتف لكن صيغته غير مناسبة للاعتماد'}
 return {'ok':True,'phone':normalized,'phone_kind':phone_kind,
         'place_id':best.get('id'),'match_name':(best.get('displayName') or {}).get('text',''),
         'match_address':best.get('formattedAddress','')}

@app.post('/admin/google-phone/<int:cid>')
@admin_required
def google_phone_one(cid):
 return_to=(request.form.get('return_to') or '').strip()
 safe_return=return_to if return_to.startswith('/companies') and not return_to.startswith('//') else url_for('company',cid=cid)
 d=db(); c=d.execute('SELECT * FROM companies WHERE id=?',(cid,)).fetchone()
 if not c: d.close(); abort(404)
 existing=(c['phone_mobile'] or '').strip()
 if existing:
  d.close(); flash('لم يتم تغيير الرقم: الشركة لديها جوال مسجل مسبقاً.'); return redirect(safe_return)
 result=_google_places_phone(c)
 if result.get('ok'):
  if result.get('phone_kind')=='mobile':
   d.execute("""UPDATE companies SET phone_mobile=?,whatsapp=CASE WHEN COALESCE(TRIM(whatsapp),'')='' THEN ? ELSE whatsapp END,
                phone_source='Google Maps / Places',google_place_id=?,phone_verified_at=CURRENT_TIMESTAMP,
                phone_match_name=?,phone_match_address=? WHERE id=?""",
             (result['phone'],result['phone'],result['place_id'],result['match_name'],result['match_address'],cid))
   flash('تمت إضافة جوال من Google Maps وأصبح واتساب متاحاً، وسيعاد ترتيب الشركة تلقائياً.')
  else:
   d.execute("""UPDATE companies SET phone_landline=CASE WHEN COALESCE(TRIM(phone_landline),'')='' THEN ? ELSE phone_landline END,
                phone_source='Google Maps / Places',google_place_id=?,phone_verified_at=CURRENT_TIMESTAMP,
                phone_match_name=?,phone_match_address=? WHERE id=?""",
             (result['phone'],result['place_id'],result['match_name'],result['match_address'],cid))
   flash('تم العثور على هاتف ثابت من Google Maps وحُفظ كهاتف ثابت؛ لم يُعامل كواتساب.')
  d.commit()
 else:
  flash('لم تتم إضافة رقم: '+result.get('reason','تعذر التحقق'))
 d.close(); return redirect(safe_return)

@app.post('/admin/google-phone-batch')
@admin_required
def google_phone_batch():
 # دفعة صغيرة لحماية التكلفة والوقت؛ لا تستبدل أي رقم موجود.
 limit=min(max(int(request.form.get('limit') or 20),1),50)
 d=db()
 rows=d.execute("""SELECT * FROM companies
                   WHERE COALESCE(TRIM(phone_mobile),'')=''
                   ORDER BY id LIMIT ?""",(limit,)).fetchall()
 added=skipped=0
 for c in rows:
  r=_google_places_phone(c)
  if r.get('ok') and r.get('phone_kind')=='mobile':
   d.execute("""UPDATE companies SET phone_mobile=?,whatsapp=CASE WHEN COALESCE(TRIM(whatsapp),'')='' THEN ? ELSE whatsapp END,
                phone_source='Google Maps / Places',google_place_id=?,phone_verified_at=CURRENT_TIMESTAMP,
                phone_match_name=?,phone_match_address=? WHERE id=?""",
             (r['phone'],r['phone'],r['place_id'],r['match_name'],r['match_address'],c['id']))
   added+=1
  elif r.get('ok') and r.get('phone_kind')=='landline':
   d.execute("""UPDATE companies SET phone_landline=CASE WHEN COALESCE(TRIM(phone_landline),'')='' THEN ? ELSE phone_landline END,
                phone_source='Google Maps / Places',google_place_id=?,phone_verified_at=CURRENT_TIMESTAMP,
                phone_match_name=?,phone_match_address=? WHERE id=?""",
             (r['phone'],r['place_id'],r['match_name'],r['match_address'],c['id']))
   skipped+=1
  else: skipped+=1
 d.commit(); d.close()
 flash(f'مطابقة Google Maps: أضيف {added} رقم، ولم تعتمد {skipped} نتيجة لعدم كفاية المطابقة/عدم وجود هاتف.')
 return redirect(url_for('companies'))


def _report_month_bounds(value):
 value=(value or '').strip()
 if not re.fullmatch(r'\d{4}-\d{2}',value):
  value=datetime.now().strftime('%Y-%m')
 try:
  first=datetime.strptime(value+'-01','%Y-%m-%d')
 except Exception:
  first=datetime.now().replace(day=1,hour=0,minute=0,second=0,microsecond=0)
  value=first.strftime('%Y-%m')
 if first.month==12:
  nxt=first.replace(year=first.year+1,month=1)
 else:
  nxt=first.replace(month=first.month+1)
 return value,first.strftime('%Y-%m-%d 00:00:00'),nxt.strftime('%Y-%m-%d 00:00:00')

def _report_type(value):
 return value if value in ('contacted','unreachable','both') else 'both'

def _report_user_selection(d):
 """Non-admin users are always restricted to their own activity."""
 if session.get('role')!='admin':
  u=d.execute("SELECT id,name,email FROM users WHERE id=?",(session['uid'],)).fetchone()
  return str(session['uid']),u
 requested=(request.args.get('user_id') or 'all').strip()
 if requested=='all':
  return 'all',None
 if requested.isdigit():
  u=d.execute("SELECT id,name,email FROM users WHERE id=?",(int(requested),)).fetchone()
  if u: return str(u['id']),u
 return 'all',None

def _report_user_condition(selected_user,selected_row):
 if selected_user=='all':
  return '',[]
 uid=int(selected_user)
 uname=(selected_row['name'] or '') if selected_row else session.get('name','')
 return " AND (i.user_id=? OR (i.user_id IS NULL AND i.user_name_snapshot=?))",[uid,uname]

def _report_latest_rows(d,month_value,report_type,selected_user,selected_row):
 month_value,start,end=_report_month_bounds(month_value)
 user_sql,user_args=_report_user_condition(selected_user,selected_row)
 type_sql=''
 type_args=[]
 if report_type in ('contacted','unreachable'):
  type_sql=' AND outcome=?'
  type_args=[report_type]
 sql=f"""
 WITH ranked AS (
  SELECT
   i.id AS interaction_id,i.company_id,i.user_id,i.user_name_snapshot,
   i.channel,i.outcome,i.notes,i.contacted_at,i.next_followup_at,
   c.crm_number,c.name_ar,c.city,c.phone_mobile,c.whatsapp,c.email,
   c.muqawil_member_no,c.classification_status,c.classification_grade,
   COALESCE(u.name,i.user_name_snapshot,'مستخدم غير معروف') AS rep_name,
   ROW_NUMBER() OVER (
    PARTITION BY COALESCE(CAST(i.user_id AS TEXT),'name:'||COALESCE(i.user_name_snapshot,'')), i.company_id
    ORDER BY i.contacted_at DESC,i.id DESC
   ) AS rn
  FROM interactions i
  JOIN companies c ON c.id=i.company_id
  LEFT JOIN users u ON u.id=i.user_id
  WHERE i.outcome IN ('contacted','unreachable')
    AND i.contacted_at>=? AND i.contacted_at<?
    {user_sql}
 )
 SELECT * FROM ranked
 WHERE rn=1 {type_sql}
 ORDER BY contacted_at DESC,interaction_id DESC
 """
 return month_value,start,end,[dict(r) for r in d.execute(sql,[start,end]+user_args+type_args).fetchall()]

def _report_activity_counts(d,start,end,selected_user,selected_row):
 user_sql,user_args=_report_user_condition(selected_user,selected_row)
 rows=d.execute(f"""
  SELECT
   COALESCE(i.user_id,-1) AS user_id,
   COALESCE(u.name,i.user_name_snapshot,'مستخدم غير معروف') AS rep_name,
   COUNT(*) AS events
  FROM interactions i
  LEFT JOIN users u ON u.id=i.user_id
  WHERE i.outcome IN ('contacted','unreachable')
    AND i.contacted_at>=? AND i.contacted_at<?
    {user_sql}
  GROUP BY COALESCE(i.user_id,-1),COALESCE(u.name,i.user_name_snapshot,'مستخدم غير معروف')
 """,[start,end]+user_args).fetchall()
 return {(r['user_id'],r['rep_name']):r['events'] for r in rows}

def _report_summary(d,month_value,selected_user,selected_row):
 month_value,start,end,all_latest=_report_latest_rows(d,month_value,'both',selected_user,selected_row)
 activity=_report_activity_counts(d,start,end,selected_user,selected_row)

 # Active users are included even when their monthly count is zero.
 result={}
 if selected_user=='all' and session.get('role')=='admin':
  for u in d.execute("SELECT id,name FROM users WHERE active=1 ORDER BY name").fetchall():
   result[(u['id'],u['name'])]={
    'user_id':u['id'],'name':u['name'],'contacted':0,'unreachable':0,'total':0,'events':0
   }
 elif selected_row:
  result[(selected_row['id'],selected_row['name'])]={
   'user_id':selected_row['id'],'name':selected_row['name'],'contacted':0,'unreachable':0,'total':0,'events':0
  }

 for r in all_latest:
  key=(r['user_id'] if r['user_id'] is not None else -1,r['rep_name'])
  if key not in result:
   result[key]={'user_id':key[0],'name':key[1],'contacted':0,'unreachable':0,'total':0,'events':0}
  if r['outcome']=='contacted': result[key]['contacted']+=1
  elif r['outcome']=='unreachable': result[key]['unreachable']+=1
  result[key]['total']+=1

 for key,events in activity.items():
  if key not in result:
   result[key]={'user_id':key[0],'name':key[1],'contacted':0,'unreachable':0,'total':0,'events':0}
  result[key]['events']=events

 rows=sorted(result.values(),key=lambda x:(-x['total'],x['name']))
 totals={
  'contacted':sum(x['contacted'] for x in rows),
  'unreachable':sum(x['unreachable'] for x in rows),
  'total':sum(x['total'] for x in rows),
  'events':sum(x['events'] for x in rows)
 }
 return month_value,start,end,rows,totals

@app.get('/reports')
@login_required
def reports():
 d=db()
 month=(request.args.get('month') or datetime.now().strftime('%Y-%m')).strip()
 rtype=_report_type(request.args.get('type'))
 selected_user,selected_row=_report_user_selection(d)
 month,start,end,rows=_report_latest_rows(d,month,rtype,selected_user,selected_row)
 _,_,_,summary,totals=_report_summary(d,month,selected_user,selected_row)
 users=d.execute("SELECT id,name FROM users WHERE active=1 ORDER BY name").fetchall() if session.get('role')=='admin' else []
 d.close()
 return render_template(
  'reports.html',rows=rows,month=month,report_type=rtype,
  selected_user=selected_user,users=users,summary=summary,totals=totals,
  is_admin=session.get('role')=='admin'
 )

@app.get('/reports/export.csv')
@login_required
def reports_export_csv():
 d=db()
 month=(request.args.get('month') or datetime.now().strftime('%Y-%m')).strip()
 rtype=_report_type(request.args.get('type'))
 selected_user,selected_row=_report_user_selection(d)
 month,start,end,rows=_report_latest_rows(d,month,rtype,selected_user,selected_row)
 d.close()
 out=io.StringIO()
 w=csv.writer(out)
 w.writerow(['#','الشركة','رقم العضوية','المدينة','الجوال','الحالة','المستخدم','القناة','تاريخ التواصل','الملاحظات','المتابعة القادمة','رقم CRM'])
 for n,r in enumerate(rows,1):
  w.writerow([
   n,r['name_ar'],r['muqawil_member_no'] or '',r['city'] or '',
   r['phone_mobile'] or r['whatsapp'] or '',
   STATUS.get(r['outcome'],r['outcome']),r['rep_name'],r['channel'] or '',
   r['contacted_at'] or '',r['notes'] or '',r['next_followup_at'] or '',r['crm_number'] or ''
  ])
 filename=f"alrawasi_report_{month}_{rtype}.csv"
 return Response(
  '\ufeff'+out.getvalue(),mimetype='text/csv; charset=utf-8',
  headers={'Content-Disposition':f'attachment; filename="{filename}"'}
 )

@app.get('/reports/export.xlsx')
@login_required
def reports_export_xlsx():
 from openpyxl.styles import Font,PatternFill,Alignment,Border,Side
 from openpyxl.utils import get_column_letter

 d=db()
 month=(request.args.get('month') or datetime.now().strftime('%Y-%m')).strip()
 rtype=_report_type(request.args.get('type'))
 selected_user,selected_row=_report_user_selection(d)
 month,start,end,rows=_report_latest_rows(d,month,rtype,selected_user,selected_row)
 _,_,_,summary,totals=_report_summary(d,month,selected_user,selected_row)
 d.close()

 wb=Workbook()
 ws=wb.active
 ws.title='تفاصيل الشركات'
 ws.sheet_view.rightToLeft=True

 headers=['#','الشركة','رقم العضوية','المدينة','الجوال','الحالة','المستخدم','القناة','تاريخ التواصل','الملاحظات','المتابعة القادمة','رقم CRM']
 ws.append(headers)
 for n,r in enumerate(rows,1):
  ws.append([
   n,r['name_ar'],r['muqawil_member_no'] or '',r['city'] or '',
   r['phone_mobile'] or r['whatsapp'] or '',
   STATUS.get(r['outcome'],r['outcome']),r['rep_name'],r['channel'] or '',
   r['contacted_at'] or '',r['notes'] or '',r['next_followup_at'] or '',r['crm_number'] or ''
  ])

 navy='0B466F'; white='FFFFFF'; pale='EAF2F8'
 header_fill=PatternFill('solid',fgColor=navy)
 thin=Side(style='thin',color='D8E2EB')
 for cell in ws[1]:
  cell.fill=header_fill; cell.font=Font(color=white,bold=True)
  cell.alignment=Alignment(horizontal='center',vertical='center')
 for row in ws.iter_rows(min_row=2):
  for c in row:
   c.border=Border(bottom=thin)
   c.alignment=Alignment(horizontal='right',vertical='center',wrap_text=True)
 ws.freeze_panes='A2'
 ws.auto_filter.ref=ws.dimensions
 widths=[6,40,16,16,16,18,18,14,20,35,20,12]
 for i,wid in enumerate(widths,1):
  ws.column_dimensions[get_column_letter(i)].width=wid
 ws.row_dimensions[1].height=26

 ss=wb.create_sheet('ملخص الإنجاز')
 ss.sheet_view.rightToLeft=True
 ss['A1']='تقرير إنجاز المستخدمين'
 ss['A2']='الشهر'; ss['B2']=month
 ss['A3']='نوع التقرير'; ss['B3']={'contacted':'تم التواصل','unreachable':'لا يمكن التواصل','both':'الاثنان معاً'}[rtype]
 ss['A5']='المستخدم'; ss['B5']='تم التواصل'; ss['C5']='لا يمكن التواصل'; ss['D5']='إجمالي الشركات'; ss['E5']='سجلات النشاط'
 for c in ss[5]:
  c.fill=header_fill;c.font=Font(color=white,bold=True);c.alignment=Alignment(horizontal='center')
 row=6
 for item in summary:
  ss.append([item['name'],item['contacted'],item['unreachable'],item['total'],item['events']])
  row+=1
 total_row=6+len(summary)
 ss.cell(total_row,1,'الإجمالي')
 ss.cell(total_row,2,totals['contacted'])
 ss.cell(total_row,3,totals['unreachable'])
 ss.cell(total_row,4,totals['total'])
 ss.cell(total_row,5,totals['events'])
 for c in ss[total_row]:
  c.font=Font(bold=True);c.fill=PatternFill('solid',fgColor=pale)
 for col,wid in zip(range(1,6),[24,16,20,18,18]):
  ss.column_dimensions[get_column_letter(col)].width=wid
 ss['A1'].font=Font(size=16,bold=True,color=navy)
 ss.merge_cells('A1:E1')
 ss['A1'].alignment=Alignment(horizontal='center')
 ss.freeze_panes='A6'

 bio=io.BytesIO()
 wb.save(bio);bio.seek(0)
 filename=f"alrawasi_monthly_report_{month}_{rtype}.xlsx"
 return send_file(
  bio,as_attachment=True,download_name=filename,
  mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
 )


@app.get('/admin/backup-db')
@admin_required
def backup_db():
 """Download a transactionally consistent copy of the LIVE persistent SQLite database."""
 src=sqlite3.connect(DB,timeout=30)
 tmp=tempfile.NamedTemporaryFile(prefix='alrawasi_crm_',suffix='.db',delete=False)
 tmp.close()
 dst=sqlite3.connect(tmp.name)
 try:
  src.backup(dst)
 finally:
  dst.close(); src.close()
 return send_file(tmp.name,as_attachment=True,download_name='alrawasi_crm_live_backup.db',mimetype='application/octet-stream')

@app.get('/admin/export-full')
@admin_required
def export_full():
 """Export companies plus CRM history/users/followups to one Excel workbook."""
 d=db()
 wb=Workbook()
 # Remove default sheet and export every CRM table so history is not lost.
 wb.remove(wb.active)
 table_names=[r['name'] for r in d.execute("SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%' ORDER BY name").fetchall()]
 for table in table_names:
  ws=wb.create_sheet(title=table[:31])
  rows=d.execute(f'SELECT * FROM "{table}"').fetchall()
  cols=[r['name'] for r in d.execute(f'PRAGMA table_info("{table}")').fetchall()]
  ws.append(cols)
  for row in rows:
   ws.append([row[c] for c in cols])
  ws.freeze_panes='A2'
  ws.auto_filter.ref=ws.dimensions
 d.close()
 bio=io.BytesIO()
 wb.save(bio); bio.seek(0)
 return send_file(bio,as_attachment=True,download_name='alrawasi_crm_full_export.xlsx',
                  mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

@app.get('/export')
@login_required
def export():
 d=db(); me=current_user(d); cond,args=specialty_scope(me,'companies'); rows=d.execute('SELECT * FROM companies WHERE '+cond+' ORDER BY id',args).fetchall(); out=io.StringIO(); w=csv.writer(out); w.writerow(rows[0].keys() if rows else ['id']); [w.writerow(list(r)) for r in rows]; d.close(); return Response('\ufeff'+out.getvalue(),mimetype='text/csv',headers={'Content-Disposition':'attachment; filename=companies.csv'})

@app.get('/admin/export-json-snapshot')
@admin_required
def export_json_snapshot():
 """Download a complete logical JSON snapshot of the current live CRM database."""
 d=db()
 try:
  tables=[r['name'] for r in d.execute(
   "SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%' ORDER BY name"
  ).fetchall()]
  payload={
   'format':'alrawasi_crm_json_snapshot',
   'format_version':1,
   'crm_version':APP_VERSION,
   'exported_at':datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
   'data_dir':DATA_DIR,
   'storage_mode':('persistent_disk' if PERSISTENT_RENDER else ('render_free_ephemeral' if FREE_EPHEMERAL_RENDER else 'local')),
   'tables':{}
  }
  counts={}
  for table in tables:
   safe='"'+table.replace('"','""')+'"'
   rows=d.execute('SELECT * FROM '+safe).fetchall()
   payload['tables'][table]=[dict(r) for r in rows]
   counts[table]=len(rows)
  payload['counts']=counts
 finally:
  d.close()
 raw=json.dumps(payload,ensure_ascii=False,indent=2,default=str).encode('utf-8')
 bio=io.BytesIO(raw)
 stamp=datetime.now().strftime('%Y%m%d_%H%M%S')
 return send_file(bio,as_attachment=True,
  download_name=f'alrawasi_crm_snapshot_V{APP_VERSION}_{stamp}.json',
  mimetype='application/json; charset=utf-8')


@app.post('/admin/restore-json-snapshot')
@admin_required
def restore_json_snapshot():
 f=request.files.get('file')
 if (request.form.get('confirm') or '').strip()!='RESTORE':
  flash('للاستعادة اكتب RESTORE في خانة التأكيد.'); return redirect(url_for('recovery_center'))
 if not f:
  flash('اختر ملف نسخة JSON أولاً.'); return redirect(url_for('recovery_center'))
 try: payload=json.load(f)
 except Exception as e:
  flash('ملف JSON غير صالح: '+str(e)); return redirect(url_for('recovery_center'))
 incoming=payload.get('tables') if isinstance(payload,dict) else None
 required=('companies','interactions','templates','users')
 if payload.get('format')!='alrawasi_crm_json_snapshot' or not isinstance(incoming,dict) or any(not isinstance(incoming.get(t),list) for t in required):
  flash('تم رفض الملف: ليس Snapshot صحيحاً للنظام.'); return redirect(url_for('recovery_center'))
 if len(incoming['companies'])<15000 or len(incoming['users'])<1:
  flash('تم رفض الملف احترازياً: عدد الشركات أو المستخدمين غير منطقي.'); return redirect(url_for('recovery_center'))
 d=db()
 backup_dir=os.path.join(DATA_DIR,'backups'); os.makedirs(backup_dir,exist_ok=True)
 bp=os.path.join(backup_dir,'crm_before_json_restore_'+datetime.now().strftime('%Y%m%d_%H%M%S')+'.db')
 try:
  b=sqlite3.connect(bp); d.backup(b); b.close()
  d.execute('BEGIN IMMEDIATE')
  for table in ('interactions','companies','templates','users'):
   rows=incoming[table]
   dbcols=[r['name'] for r in d.execute('PRAGMA table_info("'+table+'")').fetchall()]
   cols=[x for x in dbcols if not rows or x in rows[0]]
   if rows and 'id' not in cols: raise ValueError('جدول '+table+' لا يحتوي id')
   d.execute('DELETE FROM "'+table+'"')
   if rows:
    qcols=','.join('"'+x+'"' for x in cols); qs=','.join('?' for _ in cols)
    d.executemany('INSERT INTO "'+table+'" ('+qcols+') VALUES ('+qs+')',[[r.get(x) for x in cols] for r in rows])
  try:
   for table in required:
    mx=d.execute('SELECT COALESCE(MAX(id),0) FROM "'+table+'"').fetchone()[0]
    d.execute('DELETE FROM sqlite_sequence WHERE name=?',(table,))
    d.execute('INSERT INTO sqlite_sequence(name,seq) VALUES(?,?)',(table,mx))
  except sqlite3.OperationalError: pass
  if d.execute('PRAGMA quick_check').fetchone()[0]!='ok': raise ValueError('فشل فحص سلامة قاعدة البيانات')
  d.commit()
  flash('تمت استعادة Snapshot بنجاح: '+str(len(incoming['companies']))+' شركة و '+str(len(incoming['interactions']))+' سجل تواصل.')
 except Exception as e:
  d.rollback(); flash('تم إلغاء الاستعادة بالكامل: '+str(e))
 finally: d.close()
 return redirect(url_for('recovery_center'))

@app.post('/company/<int:company_id>/delete')
@admin_required
def delete_company(company_id):
 return_to=(request.form.get('return_to') or session.get('companies_return_to') or url_for('companies')).strip()
 # Accept only an internal companies URL; never redirect to an external address.
 if not return_to.startswith('/companies') or return_to.startswith('//'):
  return_to=url_for('companies')
 d=db()
 c=d.execute('SELECT id,name_ar,name_en FROM companies WHERE id=?',(company_id,)).fetchone()
 if not c:
  d.close(); flash('الشركة غير موجودة.'); return redirect(return_to)
 try:
  d.execute('BEGIN IMMEDIATE')
  # Delete dependent rows from every table that has a company_id column.
  tables=[r['name'] for r in d.execute("SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%'").fetchall()]
  for table in tables:
   if table=='companies': continue
   cols={r['name'] for r in d.execute(f'PRAGMA table_info("{table}")').fetchall()}
   if 'company_id' in cols:
    d.execute(f'DELETE FROM "{table}" WHERE company_id=?',(company_id,))
  d.execute('DELETE FROM companies WHERE id=?',(company_id,))
  d.commit()
  flash('تم حذف الشركة نهائياً. تم الاحتفاظ بنفس الصفحة والفلاتر للمتابعة.')
 except Exception:
  d.rollback(); raise
 finally:
  d.close()
 return redirect(return_to)

if __name__=='__main__': app.run(host='0.0.0.0',port=int(os.environ.get('PORT',5000)))
