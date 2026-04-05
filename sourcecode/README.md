# EasyPC Marketing Hub

## Project structure

```
marketing_hub/
|-- server/
|   `-- meta-insights.js      <- Local API proxy to Meta Graph API
|-- src/
|   |-- main.jsx              <- React entry point
|   `-- App.jsx               <- Dashboard app
|-- data/                     <- Local cache/settings data files
|-- settings.ini
|-- index.html
|-- vite.config.js
`-- package.json
```

## Run locally

1. Install Node.js (LTS): https://nodejs.org
2. Install dependencies:

```bash
npm install
```

3. Start the app:

```bash
npm run dev
```

4. Open in browser:

```text
http://localhost:5173
```

The `/api/meta-insights` endpoint is served locally by Vite middleware and handled by `server/meta-insights.js`.

## Local environment variables

Create a `.env` file for local development as needed:

```env
VITE_AAD_TENANT_ID=973ec11f-980d-4bd7-9443-fe528f0a752b
VITE_AAD_CLIENT_ID=e7c8038f-4c5a-4be8-bce1-a3d42e0e38f5
VITE_ALLOWED_EMAILS=

AAD_TENANT_ID=973ec11f-980d-4bd7-9443-fe528f0a752b
AAD_CLIENT_ID=e7c8038f-4c5a-4be8-bce1-a3d42e0e38f5
AUTH_POLICY=tenant
ALLOWED_EMAILS=

META_ACCESS_TOKEN=your_meta_token_here
```

## Build preview

```bash
npm run build
npm run preview
```
