#!/usr/bin/env python3
"""Upload the result file to Google Drive using a Service Account."""

import sys
import os
import yaml
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from googleapiclient.errors import HttpError

def main():
    if len(sys.argv) < 2:
        print("Usage: upload_to_drive.py <file_to_upload>")
        sys.exit(1)

    file_path = sys.argv[1]
    if not os.path.exists(file_path):
        print(f"Error: file not found: {file_path}")
        sys.exit(1)

    # Load config
    cfg = yaml.safe_load(open("data/input/config.yaml"))
    gd = cfg.get("google_drive", {})

    sa_path   = gd.get("service_account_path", "credentials/service_account.json")
    folder_id = gd.get("folder_id", "")
    filename  = gd.get("filename", os.path.basename(file_path))

    if not os.path.exists(sa_path):
        print(f"Error: service account file not found: {sa_path}")
        print("See README for setup instructions.")
        sys.exit(1)

    if not folder_id:
        print("Error: google_drive.folder_id is not set in config.yaml")
        sys.exit(1)

    # Authenticate
    creds = service_account.Credentials.from_service_account_file(
        sa_path,
        scopes=["https://www.googleapis.com/auth/drive"]
    )
    service = build("drive", "v3", credentials=creds)

    # Check if file already exists in the folder (to update instead of duplicate)
    existing = service.files().list(
        q=f"name='{filename}' and '{folder_id}' in parents and trashed=false",
        fields="files(id, name)"
    ).execute().get("files", [])

    mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    media = MediaFileUpload(file_path, mimetype=mime, resumable=True)

    try:
        if existing:
            # Update existing file
            file_id = existing[0]["id"]
            service.files().update(
                fileId=file_id,
                media_body=media
            ).execute()
            print(f"Updated existing file on Google Drive: {filename} (id: {file_id})")
        else:
            # Create new file
            meta = {"name": filename, "parents": [folder_id]}
            result = service.files().create(
                body=meta,
                media_body=media,
                fields="id, webViewLink"
            ).execute()
            print(f"Uploaded to Google Drive: {filename}")
            print(f"Link: {result.get('webViewLink', '')}")

    except HttpError as e:
        print(f"Google Drive API error: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
