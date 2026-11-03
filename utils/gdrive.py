import io
import os
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload


SCOPES = ['https://www.googleapis.com/auth/drive']
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CREDENTIALS_FILE = os.path.join(BASE_DIR, 'credentials.json')
TOKEN_FILE = os.path.join(BASE_DIR, 'token_gdrive.json')

# Target folder in your Google Drive
GDRIVE_FOLDER_ID = "161vqbdDartUVPGlogkJ0hD-zDiw3_Oq4"


def get_drive_service():
    """Authenticates with user OAuth2 credentials so uploads use user storage quota."""
    creds = None
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists(CREDENTIALS_FILE):
                raise FileNotFoundError(
                    f"Missing {CREDENTIALS_FILE}. Place your downloaded OAuth client JSON file there."
                )
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)

        with open(TOKEN_FILE, 'w') as token:
            token.write(creds.to_json())

    return build('drive', 'v3', credentials=creds)


def upload_image_to_drive(local_path, file_name):
    """Uploads a local image to Google Drive and returns the file ID."""
    try:
        service = get_drive_service()
        file_metadata = {
            'name': file_name,
            'parents': [GDRIVE_FOLDER_ID]
        }
        media = MediaFileUpload(local_path, mimetype='image/jpeg', resumable=True)
        uploaded = service.files().create(
            body=file_metadata,
            media_body=media,
            fields='id'
        ).execute()
        drive_id = uploaded.get('id')
        print(f"Uploaded to Drive: {drive_id}")
        return drive_id
    except Exception as e:
        print(f"FAILED to upload {file_name} to Drive: {e}")
        return None


def download_image_as_stream(file_id):
    """Downloads an image from Drive into an in-memory BytesIO stream."""
    if not file_id or os.path.sep in file_id or len(file_id) < 15:
        return None
    try:
        service = get_drive_service()
        request = service.files().get_media(fileId=file_id)
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()
        fh.seek(0)
        return fh
    except Exception as e:
        print(f"Failed to stream Drive file {file_id}: {e}")
        return None


def delete_drive_file(file_id):
    """Deletes a file from Google Drive."""
    if not file_id or os.path.sep in file_id:
        return False
    try:
        service = get_drive_service()
        service.files().delete(fileId=file_id).execute()
        return True
    except Exception as e:
        print(f"Failed to delete Drive file {file_id}: {e}")
        return False