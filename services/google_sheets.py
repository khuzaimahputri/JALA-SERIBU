import streamlit as st
import gspread
import pandas as pd
from google.oauth2.service_account import Credentials


def get_google_sheet_data():
    config = st.secrets["google_sheets"]

    credentials_info = {
        "type": config["type"],
        "project_id": config["project_id"],
        "private_key_id": config["private_key_id"],
        "private_key": config["private_key"],
        "client_email": config["client_email"],
        "client_id": config["client_id"],
        "auth_uri": config["auth_uri"],
        "token_uri": config["token_uri"],
        "auth_provider_x509_cert_url":
            config["auth_provider_x509_cert_url"],
        "client_x509_cert_url":
            config["client_x509_cert_url"],
    }

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets.readonly"
    ]

    credentials = Credentials.from_service_account_info(
        credentials_info,
        scopes=scopes,
    )

    client = gspread.authorize(credentials)

    spreadsheet = client.open_by_key(
        config["spreadsheet_id"]
    )

    worksheet = spreadsheet.worksheet(
        config["worksheet_name"]
    )

    data = worksheet.get_all_records()

    return pd.DataFrame(data)

def append_pengaduan_row(row_data):
    config = st.secrets["google_sheets"]
    db_config = st.secrets["jala_seribu_db"]

    creds = Credentials.from_service_account_info(
        dict(config),
        scopes=[
            "https://www.googleapis.com/auth/spreadsheets"
        ]
    )

    client = gspread.authorize(creds)

    spreadsheet = client.open_by_key(
        db_config["spreadsheet_id"]
    )

    worksheet = spreadsheet.worksheet(
        "Saran dan Pengaduan"
    )

    worksheet.append_row(
        row_data,
        value_input_option="USER_ENTERED"
    )

def get_pengaduan_data():
    config = st.secrets["google_sheets"]
    db_config = st.secrets["jala_seribu_db"]

    creds = Credentials.from_service_account_info(
        dict(config),
        scopes=[
            "https://www.googleapis.com/auth/spreadsheets"
        ]
    )

    client = gspread.authorize(creds)

    spreadsheet = client.open_by_key(
        db_config["spreadsheet_id"]
    )

    worksheet = spreadsheet.worksheet(
        "Saran dan Pengaduan"
    )

    data = worksheet.get_all_records()

    return pd.DataFrame(data)

def append_pemutakhiran_row(row_data):
    config = st.secrets["google_sheets"]
    db_config = st.secrets["jala_seribu_db"]

    creds = Credentials.from_service_account_info(
        dict(config),
        scopes=[
            "https://www.googleapis.com/auth/spreadsheets"
        ]
    )

    client = gspread.authorize(creds)

    spreadsheet = client.open_by_key(
        db_config["spreadsheet_id"]
    )

    worksheet = spreadsheet.worksheet(
        "Pemutakhiran Kanal Digital"
    )

    worksheet.append_row(
        row_data,
        value_input_option="USER_ENTERED"
    )

def get_pemutakhiran_data():
    config = st.secrets["google_sheets"]
    db_config = st.secrets["jala_seribu_db"]

    creds = Credentials.from_service_account_info(
        dict(config),
        scopes=[
            "https://www.googleapis.com/auth/spreadsheets"
        ]
    )

    client = gspread.authorize(creds)

    spreadsheet = client.open_by_key(
        db_config["spreadsheet_id"]
    )

    worksheet = spreadsheet.worksheet(
        "Pemutakhiran Kanal Digital"
    )

    data = worksheet.get_all_records()

    return pd.DataFrame(data)

def append_faq_row(row_data):
    config = st.secrets["google_sheets"]
    db_config = st.secrets["jala_seribu_db"]

    creds = Credentials.from_service_account_info(
        dict(config),
        scopes=[
            "https://www.googleapis.com/auth/spreadsheets"
        ]
    )

    client = gspread.authorize(creds)

    spreadsheet = client.open_by_key(
        db_config["spreadsheet_id"]
    )

    worksheet = spreadsheet.worksheet(
        "Pertanyaan dan FAQ"
    )

    worksheet.append_row(
        row_data,
        value_input_option="USER_ENTERED"
    )

def get_faq_data():
    config = st.secrets["google_sheets"]
    db_config = st.secrets["jala_seribu_db"]

    creds = Credentials.from_service_account_info(
        dict(config),
        scopes=[
            "https://www.googleapis.com/auth/spreadsheets"
        ]
    )

    client = gspread.authorize(creds)

    spreadsheet = client.open_by_key(
        db_config["spreadsheet_id"]
    )

    worksheet = spreadsheet.worksheet(
        "Pertanyaan dan FAQ"
    )

    data = worksheet.get_all_records()

    return pd.DataFrame(data)

def get_skd_data():
    config = st.secrets["google_sheets"]
    db_config = st.secrets["jala_seribu_db"]

    creds = Credentials.from_service_account_info(
        dict(config),
        scopes=[
            "https://www.googleapis.com/auth/spreadsheets"
        ]
    )

    client = gspread.authorize(creds)

    spreadsheet = client.open_by_key(
        db_config["spreadsheet_id"]
    )

    worksheet = spreadsheet.worksheet(
        "Progres SKD"
    )

    data = worksheet.get_all_records()

    return pd.DataFrame(data)

def upsert_skd_row(tanggal_cacah, nama, status_kuesioner):
    config = st.secrets["google_sheets"]
    db_config = st.secrets["jala_seribu_db"]

    creds = Credentials.from_service_account_info(
        dict(config),
        scopes=[
            "https://www.googleapis.com/auth/spreadsheets"
        ]
    )

    client = gspread.authorize(creds)

    spreadsheet = client.open_by_key(
        db_config["spreadsheet_id"]
    )

    worksheet = spreadsheet.worksheet(
        "Progres SKD"
    )

    data = worksheet.get_all_records()

    # Cari berdasarkan nama
    nama_baru = str(nama).strip().lower()

    for index, row in enumerate(data):
        nama_lama = str(row.get("Nama", "")).strip().lower()

        if nama_lama == nama_baru:
            # +2 karena:
            # index Python mulai dari 0
            # baris 1 Google Sheets adalah header
            sheet_row = index + 2

            worksheet.update(
                range_name=f"A{sheet_row}:C{sheet_row}",
                values=[[
                    tanggal_cacah,
                    nama,
                    status_kuesioner
                ]]
            )

            return "updated"

    # Kalau nama belum ditemukan → tambah baris baru
    worksheet.append_row(
        [
            tanggal_cacah,
            nama,
            status_kuesioner
        ],
        value_input_option="USER_ENTERED"
    )

    return "inserted"