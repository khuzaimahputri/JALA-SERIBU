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