"""
Uploads dashboard-data.json to SharePoint after Smartsheet sync.
Uses Microsoft Graph API with app-only authentication.
"""

import os, json, requests

TENANT_ID     = os.environ.get("SHAREPOINT_TENANT_ID", "YOUR_TENANT_ID")
CLIENT_ID     = os.environ.get("SHAREPOINT_CLIENT_ID", "YOUR_CLIENT_ID")
CLIENT_SECRET = os.environ.get("SHAREPOINT_CLIENT_SECRET", "YOUR_CLIENT_SECRET")
SITE_URL      = os.environ.get("SHAREPOINT_SITE_URL", "https://yourcompany.sharepoint.com/sites/yourteam")

# Path within SharePoint where dashboard-data.json will be stored
# Change this to match your SharePoint document library and folder
UPLOAD_PATH = "/Shared Documents/PMT Dashboard/dashboard-data.json"

def get_access_token():
    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    data = {
        "grant_type": "client_credentials",
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default",
    }
    r = requests.post(url, data=data)
    r.raise_for_status()
    return r.json()["access_token"]

def get_site_id(token):
    host = SITE_URL.split("/sites/")[0].replace("https://", "")
    site = SITE_URL.split("/sites/")[1]
    url = f"https://graph.microsoft.com/v1.0/sites/{host}:/sites/{site}"
    r = requests.get(url, headers={"Authorization": f"Bearer {token}"})
    r.raise_for_status()
    return r.json()["id"]

def upload_file(token, site_id, content):
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:{UPLOAD_PATH}:/content"
    r = requests.put(
        url,
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        data=content,
    )
    r.raise_for_status()
    print(f"  ✓ dashboard-data.json uploaded to SharePoint: {UPLOAD_PATH}")

def main():
    print("[SharePoint] Uploading dashboard-data.json...")
    with open("dashboard-data.json", "rb") as f:
        content = f.read()
    token   = get_access_token()
    site_id = get_site_id(token)
    upload_file(token, site_id, content)

if __name__ == "__main__":
    main()
