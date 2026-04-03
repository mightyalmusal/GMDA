# EasyPC Marketing Hub — Deployment Guide

## What's in this package

```
easypc-dashboard/
├── netlify/
│   └── functions/
│       └── meta-insights.js   ← Serverless proxy to Meta API
├── src/
│   ├── main.jsx               ← React entry point
│   └── App.jsx                ← Full dashboard app
├── index.html
├── package.json
├── vite.config.js
└── netlify.toml               ← Netlify build config
```

---

## Step 1 — Install Node.js

If you don't have Node.js installed, download it from https://nodejs.org (LTS version).

---

## Step 2 — Install dependencies

Open a terminal in this folder and run:

```bash
npm install
```

---

## Step 3 — Deploy to Netlify

### Option A: Drag-and-drop (easiest)

1. Build the app:
   ```bash
   npm run build
   ```
2. Go to https://app.netlify.com
3. Drag the `dist/` folder onto the Netlify dashboard
4. Your site is live — but you still need Step 4 to set your token

### Option B: Git-based (recommended for ongoing use)

1. Push this folder to a GitHub repository
2. Go to https://app.netlify.com → "Add new site" → "Import from Git"
3. Connect your repo — Netlify will auto-detect the build settings from `netlify.toml`

---

## Step 4 — Require Microsoft Organization Login (SharePoint / M365)

This app now requires Microsoft Entra sign-in before any dashboard/API access.

1. Create an Entra app registration (single tenant)
2. Add redirect URIs:
   - `http://localhost:5173` (development)
   - your Netlify URL (production), e.g. `https://your-site.netlify.app`
3. In Netlify dashboard → **Site configuration** → **Environment variables**, add:
   - `AAD_TENANT_ID=973ec11f-980d-4bd7-9443-fe528f0a752b`
   - `AAD_CLIENT_ID=e7c8038f-4c5a-4be8-bce1-a3d42e0e38f5`
   - `AUTH_POLICY=emails`
   - `ALLOWED_EMAILS=user1@yourorg.com,user2@yourorg.com`
4. Redeploy the site.

Notes:
- `AUTH_POLICY=emails` restricts access to the exact email allowlist.
- If you want any tenant user to access, set `AUTH_POLICY=tenant`.

---

## Step 5 — Set your Meta Access Token (REQUIRED)

Your access token is stored as a server-side environment variable. It is NEVER exposed to the browser.

1. In Netlify dashboard → your site → **Site configuration** → **Environment variables**
2. Click **Add a variable**
3. Key: `META_ACCESS_TOKEN`
4. Value: your long-lived Meta System User token
5. Click **Save**
6. **Redeploy** the site (Deploys tab → "Trigger deploy")

---

## Step 6 — Configure the app

1. Open your live site
2. Go to **Settings** in the left sidebar
3. Switch Data Source to **"Live Meta API"**
4. Add your Ad Account IDs (just the numbers — e.g. `123456789`, not `act_123456789`)
5. Set your monthly budgets and inquiry targets
6. Add your Ad Set Identifier mappings
7. Click **Save All Settings**

---

## Step 7 — Persist data in SharePoint (recommended)

Without persistent storage, serverless local files can reset after redeploys. Enable SharePoint-backed storage so mappings, budget/targets, settings, and cache survive updates.

1. In Microsoft Entra, create an app registration for backend Graph access
2. Add **Application** permissions for Microsoft Graph:
   - `Sites.ReadWrite.All`
3. Grant admin consent
4. Create a client secret and copy its value
5. In Netlify environment variables, add:
   - `SP_STORAGE_ENABLED=true`
   - `SP_TENANT_ID=<your-tenant-id>`
   - `SP_CLIENT_ID=<backend-app-client-id>`
   - `SP_CLIENT_SECRET=<backend-app-client-secret>`
   - `SP_SITE_HOSTNAME=<yourtenant>.sharepoint.com`
   - `SP_SITE_PATH=/sites/<your-site-name>`
   - `SP_DOC_LIBRARY=Documents` (or your library name)
   - `SP_FOLDER=MarketingHubData`
6. Redeploy site

Once enabled, app data is stored in SharePoint files under your chosen folder and remains available after redeploys.

---

## Step 8 — Persist data in Supabase (free alternative)

If SharePoint admin consent takes time, use Supabase as a durable store.

1. Create a Supabase project
2. Open SQL Editor and run:

```sql
create table if not exists public.app_state (
   key text primary key,
   value text not null,
   updated_at timestamptz not null default now()
);
```

3. In Netlify environment variables, add:
    - `SUPABASE_URL=https://<project-ref>.supabase.co`
    - `SUPABASE_SERVICE_ROLE_KEY=<service_role_key>`
    - `SUPABASE_TABLE=app_state`
4. Redeploy

When `SUPABASE_URL` and `SUPABASE_SERVICE_ROLE_KEY` are set, backend persistence uses Supabase first (settings/mappings/budget-target/cache/status), so data survives redeploys.

---

## Running locally (for development)

Install the Netlify CLI to test your serverless function locally:

```bash
npm install -g netlify-cli
netlify dev
```

This starts both the React app and the Netlify function at http://localhost:8888

Set your token locally:
```bash
export META_ACCESS_TOKEN="your_token_here"
netlify dev
```

Or create a `.env` file (do NOT commit this):
```
META_ACCESS_TOKEN=your_token_here
```

---

## Token permissions required

Your Meta System User token needs these permissions:
- `ads_read`
- `ads_management`
- `read_insights`

To create a long-lived System User token:
1. Go to business.facebook.com → Settings → Users → System Users
2. Create a System User with "Employee" role
3. Click "Generate New Token" → select your app → check the permissions above
4. Copy the token — it doesn't expire

---

## Troubleshooting

| Issue | Fix |
|-------|-----|
| "META_ACCESS_TOKEN not set" | Check environment variable in Netlify dashboard, then redeploy |
| "Meta API error" | Verify token has correct permissions and hasn't expired |
| No data showing | Check that your Ad Account IDs are correct (numbers only) |
| LSA dashboard empty | Make sure Ad Set Identifiers are mapped with LSA as LOB/Division |
| CORS error in browser | You're calling the API directly — always use the Netlify function |
