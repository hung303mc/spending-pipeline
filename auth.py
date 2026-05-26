"""One-time OAuth flow with combined scopes for spending pipeline."""
import sys
from pathlib import Path

from google_auth_oauthlib.flow import InstalledAppFlow

HERE = Path(__file__).parent
CREDS_PATH = HERE.parent / "gsheets" / "credentials.json"  # reuse OAuth desktop client
TOKEN_PATH = HERE / "token.json"

SCOPES = [
    "https://www.googleapis.com/auth/gmail.readonly",
    "https://www.googleapis.com/auth/gmail.send",
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


def main():
    flow = InstalledAppFlow.from_client_secrets_file(str(CREDS_PATH), SCOPES)
    creds = flow.run_local_server(port=0, open_browser=True)
    TOKEN_PATH.write_text(creds.to_json(), encoding="utf-8")
    print(f"Saved combined token: {TOKEN_PATH}")
    print(f"Scopes: {SCOPES}")


if __name__ == "__main__":
    main()
