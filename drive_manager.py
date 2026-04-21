"""
drive_manager.py
Upload file Excel lên Google Drive và trả về link chia sẻ.
"""
import os
import json
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from google.oauth2.service_account import Credentials

SCOPES = ["https://www.googleapis.com/auth/drive"]
FOLDER_ID = os.environ.get("GDRIVE_FOLDER_ID", "13zAuILXXNPePLOm5v4P6SSLCBnt97xX8")
CREDS_FILE = os.environ.get("GOOGLE_CREDS_FILE", "google_credentials.json")


def _get_service():
    # Hỗ trợ cả file JSON trực tiếp (local) lẫn biến môi trường JSON string (Render)
    creds_json = os.environ.get("GOOGLE_CREDS_JSON")
    if creds_json:
        info = json.loads(creds_json)
        creds = Credentials.from_service_account_info(info, scopes=SCOPES)
    else:
        creds = Credentials.from_service_account_file(CREDS_FILE, scopes=SCOPES)
    return build("drive", "v3", credentials=creds)


def upload_report(local_path: str, filename: str) -> str:
    """
    Upload file lên Google Drive folder.
    Trả về link chia sẻ (view).
    """
    service = _get_service()

    file_metadata = {
        "name": filename,
        "parents": [FOLDER_ID],
    }
    media = MediaFileUpload(
        local_path,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    uploaded = service.files().create(
        body=file_metadata,
        media_body=media,
        fields="id"
    ).execute()

    file_id = uploaded.get("id")

    # Cấp quyền xem công khai (chỉ ai có link)
    service.permissions().create(
        fileId=file_id,
        body={"type": "anyone", "role": "reader"}
    ).execute()

    return f"https://drive.google.com/file/d/{file_id}/view"
