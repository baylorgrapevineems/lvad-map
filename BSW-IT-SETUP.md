# BSW IT Setup Guide — LVAD Patient Monitor

This document is for BSW IT staff. It explains how to connect the
LVAD Patient Monitor web app to your Microsoft 365 / SharePoint tenant.
The whole process takes about 15–20 minutes.

---

## What you need

- Global Administrator or Application Administrator role in Azure AD
- Access to the SharePoint site where patient data will live
- PowerShell with PnP.PowerShell module (`Install-Module PnP.PowerShell`)

---

## Step 1 — Create an Azure AD App Registration

1. Go to https://portal.azure.com → **Azure Active Directory** → **App registrations**
2. Click **New registration**
3. Fill in:
   - **Name:** `LVAD Map App`
   - **Supported account types:** *Accounts in this organizational directory only*
   - **Redirect URI:** leave blank
4. Click **Register**
5. Copy the following values — you will need them later:
   - **Application (client) ID** — looks like `xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx`
   - **Directory (tenant) ID** — same format

---

## Step 2 — Grant SharePoint permissions

1. In the App Registration page, go to **API permissions**
2. Click **Add a permission** → **Microsoft Graph** → **Application permissions**
3. Search for and add: **`Sites.ReadWrite.All`**
4. Click **Add permissions**
5. Click **Grant admin consent for [your org]** — confirm when prompted

> **Why Sites.ReadWrite.All?**
> The app reads and writes patient records in a SharePoint list. This
> permission is app-level (no user sign-in required), so only IT controls access.

---

## Step 3 — Create a client secret

1. In the App Registration page, go to **Certificates & secrets**
2. Click **New client secret**
3. Set a description (e.g., `LVAD Map`) and expiry (recommend 24 months)
4. Click **Add**
5. **Copy the secret Value immediately** — it is only shown once

---

## Step 4 — Create the SharePoint list

Run this PowerShell script against the SharePoint site you want to use:

```powershell
Install-Module PnP.PowerShell -Scope CurrentUser   # first time only

.\Create-LVADLists.ps1 -SiteUrl "https://bswh.sharepoint.com/sites/EMS"
```

The script creates the `LVAD_Patients` list with all required columns.
It is idempotent — safe to re-run.

---

## Step 5 — Configure the app environment

Create a file named `.env` in the app's root folder (same folder as `server.js`):

```
AZURE_TENANT_ID=<Directory (tenant) ID from Step 1>
AZURE_CLIENT_ID=<Application (client) ID from Step 1>
AZURE_CLIENT_SECRET=<Secret Value from Step 3>
SHAREPOINT_SITE_URL=https://bswh.sharepoint.com/sites/EMS
SP_PATIENTS_LIST=LVAD_Patients
```

For Vercel deployment, add these same key/value pairs in:
**Vercel project → Settings → Environment Variables**

---

## Step 6 — Verify

Start the app (`npm start`) and open http://localhost:3001/admin.html.
Add a test patient and confirm the record appears in the SharePoint list at:
`https://bswh.sharepoint.com/sites/EMS/Lists/LVAD_Patients`

---

## Security notes

- The client secret grants write access to that SharePoint site. Store it
  in Vercel environment variables, not in source code or email.
- Restrict the SharePoint site's member list to staff who need visibility.
- The app currently has no end-user login. All visitors on the network can
  view the map. If you need login enforcement, contact the dev team to
  enable MSAL browser-side authentication.
- This app handles PHI. Confirm your M365 tenant has an active Microsoft
  HIPAA Business Associate Agreement (BAA) before storing patient data.
  Check: https://www.microsoft.com/en-us/trust-center/compliance/hipaa

---

## Troubleshooting

| Error | Cause | Fix |
|-------|-------|-----|
| `Azure auth failed: AADSTS700016` | Wrong client ID or tenant ID | Re-check Step 1 values |
| `Graph 403` | Admin consent not granted | Re-do Step 2, grant consent |
| `Graph 404` on list | List name mismatch | Verify `SP_PATIENTS_LIST` matches the list title exactly |
| `invalid_client` | Wrong or expired secret | Create a new secret in Step 3 |
