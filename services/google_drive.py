import io
import os

import streamlit as st

from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

SCOPES = [
    "https://www.googleapis.com/auth/drive"
]

def get_drive_service():
    token_path = "credentials/token.json"
    client_secret_path = "credentials/oauth_client.json"

    creds = None

    # Kalau token pernah dibuat, pakai lagi
    if os.path.exists(token_path):
        creds = Credentials.from_authorized_user_file(
            token_path,
            SCOPES
        )

    # Kalau token tidak ada / expired
    if not creds or not creds.valid:

        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())

        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                client_secret_path,
                SCOPES
            )

            creds = flow.run_local_server(
                port=0,
                access_type="offline",
                prompt="consent"
            )

        # Simpan token supaya tidak login ulang
        with open(token_path, "w") as token:
            token.write(creds.to_json())

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