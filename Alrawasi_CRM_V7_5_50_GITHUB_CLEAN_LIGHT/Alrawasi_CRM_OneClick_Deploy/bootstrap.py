import os, zipfile, tempfile, shutil, py_compile, subprocess, sys
BASE=os.path.dirname(os.path.abspath(__file__))
DATA_DIR=os.environ.get('DATA_DIR',BASE)
PKG=os.path.join(DATA_DIR,'app_updates','current.zip')
BLOCKED={'crm.db','.env'}
def apply_persistent_update():
    if not os.path.isfile(PKG): return
    stage=tempfile.mkdtemp(prefix='crm_boot_')
    try:
        with zipfile.ZipFile(PKG) as z:
            for m in z.infolist():
                n=m.filename.replace('\\','/')
                if n.startswith('/') or '..' in n.split('/'): raise RuntimeError('unsafe update package')
            z.extractall(stage)
        src=None
        for p in (os.path.join(stage,'Alrawasi_CRM_OneClick_Deploy'),stage):
            if os.path.isfile(os.path.join(p,'app.py')): src=p; break
        if not src:
            for b,ds,fs in os.walk(stage):
                if 'app.py' in fs and 'templates' in ds: src=b; break
        if not src: raise RuntimeError('invalid update package')
        py_compile.compile(os.path.join(src,'app.py'),doraise=True)
        for name in os.listdir(src):
            if name in BLOCKED or name.startswith('.git') or name in {'__pycache__','bootstrap.py','Dockerfile','render.yaml'}: continue
            a=os.path.join(src,name); b=os.path.join(BASE,name)
            if os.path.isdir(a):
                if name=='templates':
                    if os.path.exists(b): shutil.rmtree(b)
                    shutil.copytree(a,b)
            else: shutil.copy2(a,b)
    finally: shutil.rmtree(stage,ignore_errors=True)
apply_persistent_update()
os.execvp('gunicorn',['gunicorn','--bind',f"0.0.0.0:{os.environ.get('PORT','5000')}",'--workers','2','--threads','4','--timeout','120','app:app'])
