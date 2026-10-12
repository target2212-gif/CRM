import os, requests, sys
base=os.environ["CRM_BASE_URL"].rstrip("/")
token=os.environ["MUQAWIL_SYNC_TOKEN"]
r=requests.post(base+"/internal/sync-muqawil",headers={"X-Sync-Token":token},timeout=900)
print(r.status_code, r.text)
r.raise_for_status()
