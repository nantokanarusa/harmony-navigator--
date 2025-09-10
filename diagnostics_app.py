# diagnostics_app.py
import streamlit as st
import re
import gspread
from google.oauth2.service_account import Credentials
from google.auth.transport.requests import Request
import pandas as pd

st.set_page_config(layout="wide")
st.title("🕵️‍♂️ GSheets Authentication Diagnostics Tool")
st.caption("This tool diagnoses the connection between Streamlit Cloud and your private Google Sheet.")

st.header("Step 1: Checking Streamlit Secrets")
st.info("This first step checks if the secret keys stored in Streamlit are correctly formatted.")

try:
    # 1) secrets の存在確認
    if "gcp_service_account" not in st.secrets:
        st.error("❌ **Error:** `gcp_service_account` not found in `st.secrets`. Please check your secrets.toml file's location and the table name `[gcp_service_account]`.")
        st.stop()
    
    info = st.secrets["gcp_service_account"]
    
    # 簡易的なキーの存在チェック
    required_keys = ["type", "project_id", "private_key_id", "private_key", "client_email", "client_id"]
    missing_keys = [key for key in required_keys if key not in info]
    
    if missing_keys:
        st.error(f"❌ **Error:** The following keys are missing from your `[gcp_service_account]` secret: `{', '.join(missing_keys)}`")
    else:
        st.success("✅ `gcp_service_account` table and all required keys are present in secrets.")

    # client_email の形式チェック
    client_email = info.get("client_email", "")
    if "@" in client_email and ".iam.gserviceaccount.com" in client_email:
        st.success(f"✅ `client_email` seems correctly formatted: `{client_email}`")
    else:
        st.warning(f"⚠️ **Warning:** `client_email` format might be incorrect: `{client_email}`")

    # private_key の簡易チェック（中身は表示しない）
    pk = info.get("private_key", "")
    st.write("---")
    st.subheader("Private Key Format Check:")
    if pk.strip().startswith("-----BEGIN PRIVATE KEY-----"):
        st.success("✅ `private_key` starts with the correct header.")
    else:
        st.error("❌ **Error:** `private_key` does not start with `-----BEGIN PRIVATE KEY-----`. This is a common copy-paste error.")
    
    if pk.strip().endswith("-----END PRIVATE KEY-----"):
        st.success("✅ `private_key` ends with the correct footer.")
    else:
        st.error("❌ **Error:** `private_key` does not end with `-----END PRIVATE KEY-----`. This is a common copy-paste error.")

    if "\\n" in pk or "\n" in pk:
        st.success("✅ `private_key` contains newline characters, which is correct.")
    else:
        st.warning("⚠️ **Warning:** `private_key` does not seem to contain newline characters (`\\n`). This might cause issues if it was altered.")

except Exception as e:
    st.error("An unexpected error occurred while checking the secrets.")
    st.exception(e)

st.header("Step 2: Attempting to Authenticate with Google")
st.info("This step uses the secret key to request an authentication token from Google. If this fails, the key itself is likely invalid or corrupted.")

try:
    # 2) スコープ付きで Credentials を作り、トークンを取得できるかチェック
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    creds = Credentials.from_service_account_info(info, scopes=scopes)
    creds.refresh(Request())
    st.success("✅ **Authentication successful!** Google has accepted your credentials and issued a temporary access token.")
    st.write(f"Token is valid: `{'Yes' if creds.valid else 'No'}`")
    st.write(f"Token expires at: `{creds.expiry}`")

except Exception as e:
    st.error("❌ **Authentication Failed.** Google rejected your credentials. This strongly suggests an issue with the content of your `private_key` or other service account details in your secrets.")
    st.exception(e)
    st.stop()


st.header("Step 3: Attempting to Access the Google Sheet")
st.info("This final step uses the valid authentication token to open your specified Google Sheet. If this fails, the issue is likely with the Sheet's URL or its sharing settings.")

try:
    # 3) gspread で接続してシートにアクセスできるか
    if "connections" not in st.secrets or "gsheets" not in st.secrets["connections"] or "spreadsheet" not in st.secrets["connections"]["gsheets"]:
         st.error("❌ **Error:** `[connections.gsheets]` table or `spreadsheet` key not found in secrets.")
         st.stop()

    spreadsheet_url_or_id = st.secrets["connections"]["gsheets"]["spreadsheet"]
    st.write(f"Attempting to open spreadsheet: `{spreadsheet_url_or_id}`")
    
    # URLからIDを抽出する、より堅牢な方法
    match = re.search(r"/d/([a-zA-Z0-9-_]+)", spreadsheet_url_or_id)
    if match:
        sheet_id = match.group(1)
        st.write(f"Extracted Sheet ID: `{sheet_id}`")
    else:
        sheet_id = spreadsheet_url_or_id # IDが直接指定されていると仮定
        st.write("Assuming the provided value is the Sheet ID itself.")

    gc = gspread.authorize(creds)
    spreadsheet = gc.open_by_key(sheet_id)
    
    st.success(f"✅ **Success!** Successfully opened spreadsheet titled: **`{spreadsheet.title}`**")
    
    worksheet = spreadsheet.worksheet("Sheet1")
    df = pd.DataFrame(worksheet.get_all_records())
    st.write("Successfully read the first 5 rows from 'Sheet1':")
    st.dataframe(df.head())

except gspread.exceptions.SpreadsheetNotFound:
    st.error("❌ **Access Denied (SpreadsheetNotFound):** The spreadsheet with the given ID was not found. This could mean the ID is wrong, or the sheet has not been shared with your service account's `client_email`.")
    st.write("**Action:** Please double-check that the Sheet ID is correct and that you have shared the sheet with the email address listed in Step 1.")
except gspread.exceptions.APIError as e:
    st.error("❌ **Access Denied (APIError):** Google's API returned an error, likely related to permissions.")
    st.exception(e)
    st.write("**Action:** This often happens if the Google Drive API or Google Sheets API is not enabled in your GCP project, or if your organization's policies restrict sharing to service accounts.")
except Exception as e:
    st.error("❌ An unexpected error occurred while trying to open the sheet.")
    st.exception(e)
