#!/usr/bin/env python3
# Uploads output_en.pptx + supporting files to Google Drive:/translation (creates if missing).
import os, io, json, sys, mimetypes
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

FOLDER_NAME = os.environ.get("GDRIVE_FOLDER_NAME", "translation")

def build_service():
    sa_json = os.environ["GDRIVE_SA_JSON"]
    creds = (service_account.Credentials.from_service_account_info(json.loads(sa_json))
             if sa_json.strip().startswith("{")
             else service_account.Credentials.from_service_account_file(sa_json))
    return build("drive", "v3", credentials=creds)

def ensure_folder(drive, name):
    q = f"name = '{name}' and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
    res = drive.files().list(q=q, fields="files(id,name)").execute()
    if res.get("files"):
        return res["files"][0]["id"]
    meta = {"name": name, "mimeType": "application/vnd.google-apps.folder"}
    return drive.files().create(body=meta, fields="id").execute()["id"]

def upload(drive, path, folder_id):
    fname = os.path.basename(path)
    mime, _ = mimetypes.guess_type(path)
    meta = {"name": fname, "parents": [folder_id]}
    media = MediaFileUpload(path, mimetype=mime or "application/octet-stream", resumable=True)
    return drive.files().create(body=meta, media_body=media, fields="id,webViewLink").execute()

if __name__ == "__main__":
    files = ["output_en.pptx","bilingual.csv","translation_cache.json","audit.json",
             "glossary.json","translate_pptx_inplace.py","audit_pptx_jp_count.py","upload_to_drive.py"]
    files = [f for f in files if os.path.exists(f)]
    if not files:
        print("Nothing to upload.", file=sys.stderr); sys.exit(1)

    drive = build_service()
    folder_id = ensure_folder(drive, FOLDER_NAME)
    for f in files:
        info = upload(drive, f, folder_id)
        print(f"Uploaded {f} -> {info['webViewLink']}")
