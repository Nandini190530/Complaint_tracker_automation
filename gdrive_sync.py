# gdrive_sync.py
import os, io, json
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload

GDRIVE_FILE_ID = "1aQb2LD8FgOXyc0oRjwx-sMGJ85xTYGTt"
LOCAL_FILE     = "FINAL_STRUCTURED_COMPLAINTS.xlsx"  # matches main_pipeline.py


def get_service():
    creds_json = os.environ.get("GDRIVE_CREDENTIALS", "")
    if not creds_json:
        # Colab fallback — look for JSON file in current directory
        json_files = [f for f in os.listdir(".")
                      if f.endswith(".json")
                      and "last_run" not in f]
        if not json_files:
            raise ValueError(
                "GDRIVE_CREDENTIALS not set and no JSON file found"
            )
        with open(json_files[0]) as f:
            creds_info = json.load(f)
        print(f"Using local credentials: {json_files[0]}")
    else:
        creds_info = json.loads(creds_json)

    creds = service_account.Credentials.from_service_account_info(
        creds_info,
        scopes=["https://www.googleapis.com/auth/drive"]
    )
    print(f"Service account: {creds_info.get('client_email','unknown')}")
    return build("drive", "v3", credentials=creds, cache_discovery=False)


def download_from_onedrive():
    """Downloads Excel from Google Drive. Named for compatibility."""
    print("Downloading from Google Drive...")
    try:
        service    = get_service()
        request    = service.files().get_media(fileId=GDRIVE_FILE_ID)
        buffer     = io.BytesIO()
        downloader = MediaIoBaseDownload(buffer, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()
        with open(LOCAL_FILE, "wb") as f:
            f.write(buffer.getvalue())
        size = os.path.getsize(LOCAL_FILE)
        print(f"Downloaded: {LOCAL_FILE} ({size:,} bytes)")
        return True
    except Exception as e:
        print(f"Download error: {e}")
        return False


def upload_to_onedrive():
    """Uploads Excel to Google Drive. Named for compatibility."""
    print("Uploading to Google Drive...")
    try:
        service = get_service()
        media   = MediaFileUpload(
            LOCAL_FILE,
            mimetype=(
                "application/vnd.openxmlformats-"
                "officedocument.spreadsheetml.sheet"
            ),
            resumable=True
        )
        service.files().update(
            fileId=GDRIVE_FILE_ID,
            media_body=media
        ).execute()
        print(f"Upload complete — file updated on Google Drive")
        return True
    except Exception as e:
        print(f"Upload error: {e}")
        return False
