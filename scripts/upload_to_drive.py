import json
import os
import sys

missing = []
for pkg, imp in [
    ("google-api-python-client", "googleapiclient"),
    ("google-auth", "google.oauth2"),
    ("google-auth-oauthlib", "google_auth_oauthlib"),
    ("pyyaml", "yaml"),
]:
    try: __import__(imp)
    except ImportError: missing.append(pkg)
if missing:
    print("Missing packages: " + " ".join(missing))
    print("Fix: pip install " + " ".join(missing))
    sys.exit(1)
import yaml
from google.oauth2 import credentials as user_credentials
from google.auth import exceptions as google_auth_exceptions
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from googleapiclient.errors import HttpError

SCOPES = ["https://www.googleapis.com/auth/drive"]


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


def load_oauth_credentials(client_secret_path, token_path):
    creds = None
    if os.path.exists(token_path):
        creds = user_credentials.Credentials.from_authorized_user_file(token_path, SCOPES)

    if creds and creds.valid:
        return creds

    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        if not os.path.exists(client_secret_path):
            print("Error: " + client_secret_path + " not found")
            sys.exit(1)
        flow = InstalledAppFlow.from_client_secrets_file(client_secret_path, SCOPES)
        try:
            creds = flow.run_local_server(port=0, open_browser=False)
        except Exception:
            creds = flow.run_console()

    token_dir = os.path.dirname(token_path)
    if token_dir:
        os.makedirs(token_dir, exist_ok=True)
    with open(token_path, "w") as token_file:
        token_file.write(creds.to_json())
    return creds


def main():
    if len(sys.argv)<2: print("Usage: upload_to_drive.py FILE"); sys.exit(1)
    fp=sys.argv[1]
    if not os.path.exists(fp): print("Error: "+fp+" not found"); sys.exit(1)
    cfg=yaml.safe_load(open("data/input/config.yaml")).get("google_drive", {})
    client_secret_path = cfg.get("client_secret_path", "credentials/client_secret.json")
    token_path = cfg.get("token_path", "credentials/token.json")
    fid=cfg.get("folder_id","")
    fn=cfg.get("filename",os.path.basename(fp))
    if not fid: print("Error: folder_id not set in config.yaml"); sys.exit(1)

    creds = load_oauth_credentials(client_secret_path, token_path)
    svc=build("drive","v3",credentials=creds)

    # Resolve the parent folder to detect whether it is in a Shared Drive.
    try:
        folder = svc.files().get(
            fileId=fid,
            fields="id,name,driveId",
            supportsAllDrives=True
        ).execute()
        shared_drive_id = folder.get("driveId", "")
    except HttpError as e:
        reason, message = parse_http_error(e)
        if reason in {"notFound", "insufficientFilePermissions"}:
            print("Drive API error: cannot access folder_id '%s' (%s)." % (fid, message))
            print("Make sure your authenticated Google account has Editor access to this folder.")
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
        if reason in {"notFound", "insufficientFilePermissions"}:
            print("Hint: your authenticated account may not have access to this folder_id: " + fid)
        elif message:
            print("Details: " + message)
        sys.exit(1)
    except google_auth_exceptions.TransportError as e:
        print("Drive network error: " + str(e))
        print("Hint: check internet/DNS connectivity and retry.")
        sys.exit(1)
if __name__=="__main__": main()
