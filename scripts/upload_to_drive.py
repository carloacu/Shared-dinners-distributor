import json
import os
import sys

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
from google.auth import exceptions as google_auth_exceptions
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from googleapiclient.errors import HttpError


def parse_http_error(err):
    reason = ""
    message = str(err)
    if not getattr(err, "content", None):
        return reason, message
    try:
        payload = json.loads(err.content.decode("utf-8"))
        root = payload.get("error", {})
        errors = root.get("errors", [])
        if errors:
            reason = errors[0].get("reason", "") or ""
        message = root.get("message", message) or message
    except Exception:
        pass
    return reason, message


def main():
    if len(sys.argv)<2: print("Usage: upload_to_drive.py FILE"); sys.exit(1)
    fp=sys.argv[1]
    if not os.path.exists(fp): print("Error: "+fp+" not found"); sys.exit(1)
    cfg=yaml.safe_load(open("data/input/config.yaml")).get("google_drive", {})
    sa=cfg.get("service_account_path","credentials/service_account.json")
    fid=cfg.get("folder_id","")
    fn=cfg.get("filename",os.path.basename(fp))
    impersonate_user = cfg.get("impersonate_user", "").strip()
    shared_drive_id = cfg.get("shared_drive_id", "").strip()
    if not os.path.exists(sa): print("Error: "+sa+" not found"); sys.exit(1)
    if not fid: print("Error: folder_id not set in config.yaml"); sys.exit(1)
    creds=service_account.Credentials.from_service_account_file(sa,scopes=["https://www.googleapis.com/auth/drive"])
    if impersonate_user:
        creds = creds.with_subject(impersonate_user)
    svc=build("drive","v3",credentials=creds)

    # Resolve the parent folder to detect whether it is in a Shared Drive.
    try:
        folder = svc.files().get(
            fileId=fid,
            fields="id,name,driveId",
            supportsAllDrives=True
        ).execute()
        if not shared_drive_id:
            shared_drive_id = folder.get("driveId", "")
    except HttpError as e:
        reason, message = parse_http_error(e)
        if reason in {"notFound", "insufficientFilePermissions"}:
            print("Drive API error: cannot access folder_id '%s' (%s)." % (fid, message))
            print("Make sure the service account email has Editor access to this folder.")
        else:
            print("Drive API error: " + str(e))
        sys.exit(1)
    except google_auth_exceptions.TransportError as e:
        print("Drive network error: " + str(e))
        print("Hint: check internet/DNS connectivity and retry.")
        sys.exit(1)

    safe_name = fn.replace("'", "\\'")
    q="name='%s' and '%s' in parents and trashed=false" % (safe_name,fid)
    list_args = {
        "q": q,
        "fields": "files(id)",
        "supportsAllDrives": True,
        "includeItemsFromAllDrives": True,
    }
    if shared_drive_id:
        list_args["corpora"] = "drive"
        list_args["driveId"] = shared_drive_id
    else:
        list_args["corpora"] = "allDrives"

    ex=svc.files().list(**list_args).execute().get("files",[])
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    media=MediaFileUpload(fp,mimetype=mime,resumable=True)
    try:
        if ex:
            svc.files().update(
                fileId=ex[0]["id"],
                media_body=media,
                supportsAllDrives=True
            ).execute()
            print("Updated on Drive: "+fn)
        else:
            r=svc.files().create(
                body={"name":fn,"parents":[fid]},
                media_body=media,
                fields="id,webViewLink",
                supportsAllDrives=True
            ).execute()
            print("Uploaded to Drive: "+fn)
            print("Link: "+r.get("webViewLink",""))
    except HttpError as e:
        reason, message = parse_http_error(e)
        print("Drive API error: " + str(e))
        if reason == "storageQuotaExceeded":
            print(
                "Hint: service accounts cannot upload to 'My Drive'. "
                "Use a Shared Drive folder (set google_drive.shared_drive_id optionally) "
                "or set google_drive.impersonate_user with domain-wide delegation."
            )
            print("Current folder_id: " + fid)
        elif reason in {"notFound", "insufficientFilePermissions"}:
            print("Hint: the service account may not have access to this folder_id: " + fid)
        elif message:
            print("Details: " + message)
        sys.exit(1)
    except google_auth_exceptions.TransportError as e:
        print("Drive network error: " + str(e))
        print("Hint: check internet/DNS connectivity and retry.")
        sys.exit(1)
if __name__=="__main__": main()
