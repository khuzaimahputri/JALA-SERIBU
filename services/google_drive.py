import io
import streamlit as st

from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

SCOPES = [
    "https://www.googleapis.com/auth/drive"
]

def get_drive_service():
    drive_config = st.secrets["google_drive_oauth"]

    creds = Credentials(
        token=drive_config.get("token"),
        refresh_token=drive_config["refresh_token"],
        token_uri=drive_config["token_uri"],
        client_id=drive_config["client_id"],
        client_secret=drive_config["client_secret"],
        scopes=SCOPES
    )

    return build(
        "drive",
        "v3",
        credentials=creds
    )

def buat_folder_bukti(nama_folder, parent_folder_id):
    service = get_drive_service()

    metadata = {
        "name": nama_folder,
        "mimeType": "application/vnd.google-apps.folder",
        "parents": [parent_folder_id],
    }

    folder = service.files().create(
        body=metadata,
        fields="id, webViewLink"
    ).execute()

    return folder["id"], folder["webViewLink"]

def upload_file_ke_folder(uploaded_file, folder_id):
    service = get_drive_service()

    uploaded_file.seek(0)

    media = MediaIoBaseUpload(
        io.BytesIO(uploaded_file.getvalue()),
        mimetype=uploaded_file.type,
        resumable=False
    )

    metadata = {
        "name": uploaded_file.name,
        "parents": [folder_id],
    }

    file_drive = service.files().create(
        body=metadata,
        media_body=media,
        fields="id, webViewLink"
    ).execute()

    return file_drive

def upload_screenshot(files, nama_folder, parent_folder_id):
    folder_id, folder_link = buat_folder_bukti(
        nama_folder,
        parent_folder_id
    )

    uploaded = []

    for file in files:
        hasil = upload_file_ke_folder(
            file,
            folder_id
        )
        uploaded.append(hasil)

    return {
        "folder_id": folder_id,
        "folder_link": folder_link,
        "files": uploaded,
    }