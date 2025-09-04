#!/usr/bin/env python3
# Uploads output_en.pptx + supporting files to Google Drive:/translation (creates if missing).
import os, io, json, sys, mimetypes
from google.oauth2 import service_account, credentials as oauth_credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

FOLDER_NAME = os.environ.get("GDRIVE_FOLDER_NAME", "translation")

def build_service():
    """Build a Drive service using either OAuth user creds (if provided) or Service Account."""
    client_id = os.environ.get("GOOGLE_OAUTH_CLIENT_ID")
    client_secret = os.environ.get("GOOGLE_OAUTH_CLIENT_SECRET")
    refresh_token = os.environ.get("GOOGLE_OAUTH_REFRESH_TOKEN")

    if client_id and client_secret and refresh_token:
        # Use user's OAuth credentials (uploads count against the user quota)
        scopes = [
            "https://www.googleapis.com/auth/drive",
            "https://www.googleapis.com/auth/drive.file",
        ]
        creds = oauth_credentials.Credentials(
            token=None,
            refresh_token=refresh_token,
            token_uri="https://oauth2.googleapis.com/token",
            client_id=client_id,
            client_secret=client_secret,
            scopes=scopes,
        )
        return build("drive", "v3", credentials=creds)

    # Fallback to service account
    sa_json = os.environ["GDRIVE_SA_JSON"]
    try:
        creds = service_account.Credentials.from_service_account_info(json.loads(sa_json))
    except json.JSONDecodeError:
        creds = service_account.Credentials.from_service_account_file(sa_json)
    return build("drive", "v3", credentials=creds)

def ensure_folder(drive, name):
    q = (
        f"name = '{name}' and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
    )
    res = drive.files().list(
        q=q,
        fields="files(id,name)",
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
    ).execute()
    if res.get("files"):
        return res["files"][0]["id"]
    meta = {"name": name, "mimeType": "application/vnd.google-apps.folder"}
    return (
        drive.files()
        .create(body=meta, fields="id", supportsAllDrives=True)
        .execute()["id"]
    )

def validate_folder_id(drive, folder_id: str) -> str:
    """Return folder_id if it exists and is a folder; raise otherwise."""
    try:
        info = (
            drive.files()
            .get(fileId=folder_id, fields="id,name,mimeType,driveId", supportsAllDrives=True)
            .execute()
        )
    except Exception as e:
        print(f"ERROR: Unable to access folder id {folder_id}: {e}", file=sys.stderr)
        sys.exit(1)
    if info.get("mimeType") != "application/vnd.google-apps.folder":
        print(
            f"ERROR: Provided id {folder_id} is not a folder (mimeType={info.get('mimeType')}).",
            file=sys.stderr,
        )
        sys.exit(1)
    return folder_id

def upload(drive, path, folder_id):
    fname = os.path.basename(path)
    mime, _ = mimetypes.guess_type(path)
    meta = {"name": fname, "parents": [folder_id]}
    media = MediaFileUpload(path, mimetype=mime or "application/octet-stream", resumable=True)
    return (
        drive.files()
        .create(
            body=meta,
            media_body=media,
            fields="id,webViewLink",
            supportsAllDrives=True,
        )
        .execute()
    )

if __name__ == "__main__":
    files = [
        "output_en.pptx",
        "bilingual.csv",
        "translation_cache.json",
        "audit.json",
        "glossary.json",
        "translate_pptx_inplace.py",
        "audit_pptx_jp_count.py",
        "upload_to_drive.py",
    ]
    files = [f for f in files if os.path.exists(f)]
    if not files:
        print("Nothing to upload.", file=sys.stderr); sys.exit(1)

    drive = build_service()
    # Prefer explicit folder id via env (UPLOAD_FOLDER_ID or FOLDER_ID), else fallback by name
    explicit_id = os.environ.get("UPLOAD_FOLDER_ID") or os.environ.get("FOLDER_ID")
    if explicit_id:
        folder_id = validate_folder_id(drive, explicit_id)
    else:
        folder_id = ensure_folder(drive, FOLDER_NAME)
    for f in files:
        info = upload(drive, f, folder_id)
        print(f"Uploaded {f} -> {info['webViewLink']}")
