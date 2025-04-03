import os
import pickle
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# === CONFIGURATION ===
SCOPES = ['https://www.googleapis.com/auth/drive.file']
CREDENTIALS_FILE = 'credentials.json'  # Your OAuth Client ID JSON
TOKEN_FILE = 'token_drive.pickle'      # Will be created automatically
DRIVE_FOLDER_ID = '1rXqywoBKWUcdExU4jpIWjZ4mwbOlmHZF'  # Target folder inside shared drive
SHARED_DRIVE_ID = '0ADuXLLjfuoALUk9PVA'   # The shared drive ID (starts with '0A...')


def get_drive_service():
    """Authenticate and return the Drive service."""
    creds = None
    if os.path.exists(TOKEN_FILE):
        with open(TOKEN_FILE, 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_FILE, 'wb') as token:
            pickle.dump(creds, token)
    return build('drive', 'v3', credentials=creds)


def upload_to_shared_drive(file_path, folder_id, shared_drive_id):
    """Upload file to a Shared Drive folder."""
    service = get_drive_service()
    file_metadata = {
        'name': os.path.basename(file_path),
        'parents': [folder_id]
    }
    media = MediaFileUpload(file_path, resumable=True)
    file = service.files().create(
        body=file_metadata,
        media_body=media,
        fields='id',
        supportsAllDrives=True
    ).execute()
    print(f'✅ Uploaded to Shared Drive: {file_path} (File ID: {file.get("id")})')


if __name__ == "__main__":
    # === TEST ===
    test_file = r'C:\Users\Ryan\Downloads\12_dogbones_1.gcode'  # Change to any file you want to test
    if not os.path.exists(test_file):
        with open(test_file, 'w') as f:
            f.write('This is a test upload file.')

    upload_to_shared_drive(test_file, DRIVE_FOLDER_ID, SHARED_DRIVE_ID)
