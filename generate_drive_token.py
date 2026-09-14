from google_auth_oauthlib.flow import InstalledAppFlow

SCOPES = [
    "https://www.googleapis.com/auth/drive"
]

flow = InstalledAppFlow.from_client_secrets_file(
    "credentials/oauth_client.json",
    SCOPES
)

creds = flow.run_local_server(
    port=0,
    access_type="offline",
    prompt="consent"
)

print("\n=== REFRESH TOKEN BARU ===")
print(creds.refresh_token)

print("\n=== ACCESS TOKEN ===")
print(creds.token)