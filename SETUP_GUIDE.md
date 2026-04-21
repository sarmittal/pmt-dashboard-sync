# PMT Dashboard — Setup Guide
## One-time setup (~30 minutes)

---

## What you're setting up

```
Smartsheet ──► GitHub Actions ──► Anthropic Shared Storage ──► Claude Project (your team)
                              └──► dashboard-data.json ──────► MS Teams / SharePoint (wider team)
```

The Refresh button in the dashboard triggers GitHub Actions, which pulls fresh data from Smartsheet and pushes it to both surfaces.

---

## Step 1 — GitHub Repository (5 minutes)

1. Go to https://github.com and sign in (or create a free account)
2. Click **New repository**
3. Name it `pmt-dashboard-sync`
4. Set it to **Private** ← important
5. Click **Create repository**
6. Upload these 3 files to the repo:
   - `sync_smartsheet.py`
   - `upload_to_sharepoint.py`
   - `.github/workflows/sync.yml` (create the folder structure)

---

## Step 2 — GitHub Secrets (10 minutes)

In your GitHub repo: **Settings → Secrets and variables → Actions → New repository secret**

Add these 6 secrets:

| Secret Name | Value | Where to get it |
|---|---|---|
| `SMARTSHEET_TOKEN` | Your Smartsheet API token | Smartsheet → Account → Apps & Integrations → API Access → Generate token |
| `ANTHROPIC_API_KEY` | Your Anthropic API key | https://console.anthropic.com/api-keys |
| `PROJECT_ID` | `019d7781-f53c-74f9-8021-688603830be0` | Already known |
| `SHAREPOINT_TENANT_ID` | Your Microsoft tenant ID | Azure Portal → Azure Active Directory → Overview |
| `SHAREPOINT_CLIENT_ID` | App registration client ID | See Step 3 |
| `SHAREPOINT_CLIENT_SECRET` | App registration secret | See Step 3 |

---

## Step 3 — SharePoint App Registration (10 minutes, needs IT or Azure access)

This lets GitHub Actions write the JSON file to SharePoint automatically.

1. Go to https://portal.azure.com
2. Search for **App registrations** → **New registration**
3. Name: `PMT Dashboard Sync`
4. Click **Register**
5. Copy the **Application (client) ID** → this is your `SHAREPOINT_CLIENT_ID`
6. Go to **Certificates & secrets** → **New client secret** → copy the value → `SHAREPOINT_CLIENT_SECRET`
7. Go to **API permissions** → **Add permission** → **Microsoft Graph** → **Application permissions** → add `Sites.ReadWrite.All`
8. Click **Grant admin consent**

Also add this secret to GitHub:

| Secret Name | Value |
|---|---|
| `SHAREPOINT_SITE_URL` | `https://yourcompany.sharepoint.com/sites/yourteam` |

---

## Step 4 — GitHub Personal Access Token (2 minutes)

This is what the **Refresh button** in the dashboard uses to trigger GitHub Actions.

1. Go to https://github.com/settings/tokens
2. Click **Generate new token (classic)**
3. Give it a name: `PMT Dashboard Refresh`
4. Check the `repo` scope
5. Click **Generate token** → copy it

Then update the dashboard JSX — find these two lines near the top:
```js
const GITHUB_WEBHOOK_URL = "YOUR_GITHUB_WEBHOOK_URL";
const GITHUB_TOKEN       = "YOUR_GITHUB_PAT_TOKEN";
```

Replace with:
```js
const GITHUB_WEBHOOK_URL = "https://api.github.com/repos/YOUR_GITHUB_USERNAME/pmt-dashboard-sync/dispatches";
const GITHUB_TOKEN       = "ghp_xxxxxxxxxxxxxxxxxxxx"; // your token from above
```

---

## Step 5 — First Sync (1 minute)

1. In your GitHub repo, go to **Actions** tab
2. Click **PMT Dashboard Sync** → **Run workflow** → **Run workflow**
3. Watch it run — should take ~30 seconds
4. Check that shared storage is populated by opening the Claude Project dashboard

---

## Step 6 — Upload index.html to MS Teams (2 minutes)

1. In your MS Teams team, go to the channel you want
2. Click **+** to add a tab → **Website** or **SharePoint**
3. Upload `index.html` to your SharePoint document library
4. Share the link with your team

---

## Ongoing — How the refresh works

### Automatic
The sync runs every weekday at 7am UTC automatically. No one needs to do anything.

### Manual (anyone on the team)
1. Open the Claude Project dashboard
2. Click **🔄 Refresh from Smartsheet**
3. Wait ~35 seconds — page reloads with fresh data

The same button is in the MS Teams `index.html` version.

---

## Schedule — Changing the refresh time

In `sync.yml`, find this line:
```yaml
- cron: '0 7 * * 1-5'   # 7am UTC Monday-Friday
```

Change to your preferred time. Cron format: `minute hour day month weekday`
- Every day at 8am EST (1pm UTC): `0 13 * * *`
- Every 4 hours: `0 */4 * * *`
- Every Monday at 9am EST: `0 14 * * 1`

---

## Troubleshooting

**Dashboard shows "No data loaded"**
→ The first sync hasn't run yet. Go to GitHub Actions → run manually.

**Refresh button says "webhook not configured"**
→ You haven't updated `GITHUB_WEBHOOK_URL` in the dashboard JSX yet. See Step 4.

**Sync fails with 401 error**
→ Your `SMARTSHEET_TOKEN` is invalid or expired. Regenerate in Smartsheet.

**SharePoint upload fails**
→ Check that admin consent was granted in Step 3.

**Sheet not found warning**
→ The sheet name in `sync_smartsheet.py` doesn't exactly match your Smartsheet sheet name. Check for extra spaces.
