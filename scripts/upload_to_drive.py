import sys, os
missing = []
for pkg, imp in [("google-api-python-client","googleapiclient"),("google-auth","google.oauth2"),("pyyaml","yaml")]:
    try: __import__(imp)
    except ImportError: missing.append(pkg)
if missing:
    print("Missing packages: " + " ".join(missing))
    print("Fix: pip install " + " ".join(missing))
    sys.exit(1)
import yaml
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from googleapiclient.errors import HttpError
def main():
    if len(sys.argv)<2: print("Usage: upload_to_drive.py FILE"); sys.exit(1)
    fp=sys.argv[1]
    if not os.path.exists(fp): print("Error: "+fp+" not found"); sys.exit(1)
    cfg=yaml.safe_load(open("data/input/config.yaml"))["google_drive"]
    sa=cfg.get("service_account_path","credentials/service_account.json")
    fid=cfg.get("folder_id","")
    fn=cfg.get("filename",os.path.basename(fp))
    if not os.path.exists(sa): print("Error: "+sa+" not found"); sys.exit(1)
    if not fid: print("Error: folder_id not set in config.yaml"); sys.exit(1)
    creds=service_account.Credentials.from_service_account_file(sa,scopes=["https://www.googleapis.com/auth/drive"])
    svc=build("drive","v3",credentials=creds)
    q="name='%s' and '%s' in parents and trashed=false" % (fn,fid)
    ex=svc.files().list(q=q,fields="files(id)").execute().get("files",[])
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    media=MediaFileUpload(fp,mimetype=mime,resumable=True)
    try:
        if ex:
            svc.files().update(fileId=ex[0]["id"],media_body=media).execute()
            print("Updated on Drive: "+fn)
        else:
            r=svc.files().create(body={"name":fn,"parents":[fid]},media_body=media,fields="id,webViewLink").execute()
            print("Uploaded to Drive: "+fn)
            print("Link: "+r.get("webViewLink",""))
    except HttpError as e: print("Drive API error: "+str(e)); sys.exit(1)
if __name__=="__main__": main()
