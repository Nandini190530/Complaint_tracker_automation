import os
import io
import json
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload

# 🔥 USE YOUR FILE ID (VERY IMPORTANT)

GDRIVE_FILE_ID = "1aQb2LD8FgOXyc0oRjwx-sMGJ85xTYGTt"

LOCAL_FILE = "data.xlsx"

def get_service():
creds_json = os.environ.get("GDRIVE_CREDENTIALS")

```
if not creds_json:
    raise ValueError("GDRIVE_CREDENTIALS not set")

creds_info = json.loads(creds_json)

creds = service_account.Credentials.from_service_account_info(
    creds_info,
    scopes=["https://www.googleapis.com/auth/drive"]
)

return build("drive", "v3", credentials=creds)
```

def download_from_onedrive():   # keep same name (no change in main.py)
print("📥 Downloading from Google Drive...")

```
try:
    service = get_service()

    request = service.files().get_media(fileId=GDRIVE_FILE_ID)

    buffer = io.BytesIO()
    downloader = MediaIoBaseDownload(buffer, request)

    done = False
    while not done:
        _, done = downloader.next_chunk()

    with open(LOCAL_FILE, "wb") as f:
        f.write(buffer.getvalue())

    print("✅ Download complete")
    return True

except Exception as e:
    print("Download error:", e)
    return False
```

def upload_to_onedrive():
print("📤 Uploading to Google Drive...")

```
try:
    service = get_service()

    media = MediaFileUpload(
        LOCAL_FILE,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    service.files().update(
        fileId=GDRIVE_FILE_ID,
        media_body=media
    ).execute()

    print("✅ Upload complete")
    return True

except Exception as e:
    print("Upload error:", e)
    return False
```
