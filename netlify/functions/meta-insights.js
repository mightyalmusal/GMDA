// netlify/functions/meta-insights.js
// Proxies requests to Meta Graph API to avoid CORS issues.

const fs = require("node:fs/promises");
const path = require("node:path");

const META_API_VERSION = "v19.0";
const META_BASE = `https://graph.facebook.com/${META_API_VERSION}`;
// Durable cache location (kept as a regular project file)
const CACHE_DIR = path.join(process.cwd(), "data");
const CACHE_FILE = path.join(CACHE_DIR, "meta-insights-cache.json");
const MAPPINGS_FILE = path.join(CACHE_DIR, "meta-mappings.json");
const BUDGET_TARGETS_FILE = path.join(CACHE_DIR, "meta-budget-targets.json");
// Legacy cache location used by earlier builds; retained for one-time migration.
const LEGACY_CACHE_FILE = path.join(process.cwd(), ".cache", "meta-insights-cache.json");
const STATUS_FILE = path.join(CACHE_DIR, "meta-insights-status.json");
const SETTINGS_FILE = path.join(process.cwd(), "settings.ini");
const AAD_TENANT_ID = (process.env.AAD_TENANT_ID || "973ec11f-980d-4bd7-9443-fe528f0a752b").trim();
const AAD_CLIENT_ID = (process.env.AAD_CLIENT_ID || "e7c8038f-4c5a-4be8-bce1-a3d42e0e38f5").trim();
const AUTH_POLICY = (process.env.AUTH_POLICY || "tenant").trim().toLowerCase();
const ALLOWED_EMAILS = new Set(
  String(process.env.ALLOWED_EMAILS || "")
    .split(",")
    .map(v => v.trim().toLowerCase())
    .filter(Boolean)
);
const SUPABASE_URL = String(process.env.SUPABASE_URL || "").trim().replace(/\/$/, "");
const SUPABASE_SERVICE_ROLE_KEY = String(process.env.SUPABASE_SERVICE_ROLE_KEY || "").trim();
const SUPABASE_TABLE = String(process.env.SUPABASE_TABLE || "app_state").trim();
const META_ACCESS_TOKEN_ENV = String(process.env.META_ACCESS_TOKEN || "").trim();
const META_APP_ID_ENV = String(process.env.META_APP_ID || "").trim();
const META_APP_SECRET_ENV = String(process.env.META_APP_SECRET || "").trim();
  const SP_STORAGE_ENABLED = String(process.env.SP_STORAGE_ENABLED || "").trim().toLowerCase() === "true";
  const SP_TENANT_ID = String(process.env.SP_TENANT_ID || AAD_TENANT_ID || "").trim();
  const SP_CLIENT_ID = String(process.env.SP_CLIENT_ID || "").trim();
  const SP_CLIENT_SECRET = String(process.env.SP_CLIENT_SECRET || "").trim();
  const SP_SITE_HOSTNAME = String(process.env.SP_SITE_HOSTNAME || "").trim();
  const SP_SITE_PATH = String(process.env.SP_SITE_PATH || "").trim();
  const SP_DOC_LIBRARY = String(process.env.SP_DOC_LIBRARY || "Documents").trim();
  const SP_FOLDER = String(process.env.SP_FOLDER || "MarketingHubData").trim();
const DEFAULT_MAPPING_OPTIONS = {
  divisions: ["Retail", "Ecomm", "HR", "LSA", "Desktop", "LTS"],
  lobs: ["Desktop", "LTS", "LSA", "Brand Value", "Housebrand", "Lazada"],
  segments: [
    "Productivity",
    "Gaming",
    "High-end Gamer",
    "Alabang",
    "Angeles",
    "Bacoor",
    "Baliwag",
    "Fairview",
    "Manila",
    "Marikina",
    "Muntinlupa",
    "Novaliches",
    "Pasig",
    "Quezon City",
    "San Jose del Monte",
    "Taytay",
    "DDS",
  ],
  objectives: ["Inquiry", "Engagement", "Awareness", "Sales"],
};

const INSIGHTS_FIELDS = [
  "campaign_name",
  "adset_name",
  "account_name",
  "account_currency",
  "spend",
  "actions",
  "action_values",
  "reach",
  "impressions",
  "clicks",
  "inline_link_clicks",
  "ctr",
  "inline_link_click_ctr",
  "frequency",
  "date_start",
  "date_stop",
].join(",");

let joseModPromise;
let microsoftJwks;
let graphAppTokenCache = { token: null, expiresAt: 0 };
let sharePointSiteIdCache = null;
let sharePointDriveIdCache = null;

function isSharePointConfigured() {
  return SP_STORAGE_ENABLED && !!SP_TENANT_ID && !!SP_CLIENT_ID && !!SP_CLIENT_SECRET && !!SP_SITE_HOSTNAME && !!SP_SITE_PATH;
}

function isSupabaseConfigured() {
  return !!SUPABASE_URL && !!SUPABASE_SERVICE_ROLE_KEY;
}

function supabaseHeaders(extra = {}) {
  return {
    apikey: SUPABASE_SERVICE_ROLE_KEY,
    Authorization: `Bearer ${SUPABASE_SERVICE_ROLE_KEY}`,
    "Content-Type": "application/json",
    ...extra,
  };
}

async function readSupabaseText(key) {
  const url = `${SUPABASE_URL}/rest/v1/${encodeURIComponent(SUPABASE_TABLE)}?key=eq.${encodeURIComponent(key)}&select=value&limit=1`;
  const res = await fetch(url, { headers: supabaseHeaders() });
  if (!res.ok) {
    const text = await res.text().catch(() => "");
    throw new Error(`Failed to read Supabase key '${key}': ${text || res.statusText}`);
  }
  const rows = await res.json();
  if (!Array.isArray(rows) || !rows.length) {
    const err = new Error("Not found");
    err.code = "ENOENT";
    throw err;
  }
  return typeof rows[0].value === "string" ? rows[0].value : JSON.stringify(rows[0].value ?? null);
}

async function writeSupabaseText(key, content) {
  const payload = [{ key, value: String(content || "") }];
  const url = `${SUPABASE_URL}/rest/v1/${encodeURIComponent(SUPABASE_TABLE)}`;
  const res = await fetch(url, {
    method: "POST",
    headers: supabaseHeaders({ Prefer: "resolution=merge-duplicates,return=minimal" }),
    body: JSON.stringify(payload),
  });
  if (!res.ok) {
    const text = await res.text().catch(() => "");
    throw new Error(`Failed to write Supabase key '${key}': ${text || res.statusText}`);
  }
}

async function deleteSupabaseKey(key) {
  const url = `${SUPABASE_URL}/rest/v1/${encodeURIComponent(SUPABASE_TABLE)}?key=eq.${encodeURIComponent(key)}`;
  const res = await fetch(url, {
    method: "DELETE",
    headers: supabaseHeaders({ Prefer: "return=minimal" }),
  });
  if (!res.ok) {
    const text = await res.text().catch(() => "");
    throw new Error(`Failed to delete Supabase key '${key}': ${text || res.statusText}`);
  }
}

function encodeGraphPath(pathValue) {
  return String(pathValue || "")
    .split("/")
    .filter(Boolean)
    .map(seg => encodeURIComponent(seg))
    .join("/");
}

async function graphFetch(url, options = {}) {
  const res = await fetch(url, options);
  if (!res.ok) {
    const text = await res.text().catch(() => "");
    throw new Error(`Graph request failed (${res.status}): ${text || res.statusText}`);
  }
  return res;
}

async function getGraphAppToken() {
  if (graphAppTokenCache.token && Date.now() < graphAppTokenCache.expiresAt - 120000) {
    return graphAppTokenCache.token;
  }
  const tokenUrl = `https://login.microsoftonline.com/${SP_TENANT_ID}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    client_id: SP_CLIENT_ID,
    client_secret: SP_CLIENT_SECRET,
    scope: "https://graph.microsoft.com/.default",
    grant_type: "client_credentials",
  });
  const res = await fetch(tokenUrl, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body,
  });
  if (!res.ok) {
    const text = await res.text().catch(() => "");
    throw new Error(`Failed to get SharePoint app token: ${text || res.statusText}`);
  }
  const json = await res.json();
  graphAppTokenCache = {
    token: json.access_token,
    expiresAt: Date.now() + (Number(json.expires_in || 3600) * 1000),
  };
  return graphAppTokenCache.token;
}

async function getSharePointSiteId() {
  if (sharePointSiteIdCache) return sharePointSiteIdCache;
  const token = await getGraphAppToken();
  const normalizedPath = SP_SITE_PATH.startsWith("/") ? SP_SITE_PATH : `/${SP_SITE_PATH}`;
  const url = `https://graph.microsoft.com/v1.0/sites/${SP_SITE_HOSTNAME}:${normalizedPath}`;
  const res = await graphFetch(url, { headers: { Authorization: `Bearer ${token}` } });
  const json = await res.json();
  sharePointSiteIdCache = json.id;
  if (!sharePointSiteIdCache) throw new Error("Unable to resolve SharePoint site id");
  return sharePointSiteIdCache;
}

async function getSharePointDriveId() {
  if (sharePointDriveIdCache) return sharePointDriveIdCache;
  const token = await getGraphAppToken();
  const siteId = await getSharePointSiteId();
  const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/drives`;
  const res = await graphFetch(url, { headers: { Authorization: `Bearer ${token}` } });
  const json = await res.json();
  const drives = Array.isArray(json.value) ? json.value : [];
  const drive = drives.find(d => String(d.name || "").toLowerCase() === SP_DOC_LIBRARY.toLowerCase()) || drives[0];
  if (!drive?.id) throw new Error(`Unable to find SharePoint document library '${SP_DOC_LIBRARY}'`);
  sharePointDriveIdCache = drive.id;
  return sharePointDriveIdCache;
}

async function readStorageText(localPath, sharePointName) {
  if (isSupabaseConfigured()) {
    return readSupabaseText(sharePointName);
  }
  if (!isSharePointConfigured()) {
    return fs.readFile(localPath, "utf8");
  }
  const token = await getGraphAppToken();
  const driveId = await getSharePointDriveId();
  const relPath = encodeGraphPath(`${SP_FOLDER}/${sharePointName}`);
  const url = `https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${relPath}:/content`;
  const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  if (res.status === 404) {
    const err = new Error("Not found");
    err.code = "ENOENT";
    throw err;
  }
  if (!res.ok) {
    const text = await res.text().catch(() => "");
    throw new Error(`Failed to read SharePoint file '${sharePointName}': ${text || res.statusText}`);
  }
  return res.text();
}

async function writeStorageText(localPath, sharePointName, content) {
  if (isSupabaseConfigured()) {
    await writeSupabaseText(sharePointName, content);
    return;
  }
  if (!isSharePointConfigured()) {
    await fs.mkdir(path.dirname(localPath), { recursive: true });
    await fs.writeFile(localPath, content, "utf8");
    return;
  }
  const token = await getGraphAppToken();
  const driveId = await getSharePointDriveId();
  const relPath = encodeGraphPath(`${SP_FOLDER}/${sharePointName}`);
  const url = `https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${relPath}:/content`;
  const res = await fetch(url, {
    method: "PUT",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    body: String(content || ""),
  });
  if (!res.ok) {
    const text = await res.text().catch(() => "");
    throw new Error(`Failed to write SharePoint file '${sharePointName}': ${text || res.statusText}`);
  }
}

async function deleteStorageFile(localPath, sharePointName) {
  if (isSupabaseConfigured()) {
    await deleteSupabaseKey(sharePointName);
    return;
  }
  if (!isSharePointConfigured()) {
    try {
      await fs.unlink(localPath);
    } catch (err) {
      if (err.code !== "ENOENT") throw err;
    }
    return;
  }
  const token = await getGraphAppToken();
  const driveId = await getSharePointDriveId();
  const relPath = encodeGraphPath(`${SP_FOLDER}/${sharePointName}`);
  const url = `https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${relPath}`;
  const res = await fetch(url, { method: "DELETE", headers: { Authorization: `Bearer ${token}` } });
  if (![204, 404].includes(res.status)) {
    const text = await res.text().catch(() => "");
    throw new Error(`Failed to delete SharePoint file '${sharePointName}': ${text || res.statusText}`);
  }
}

function getBearerToken(headers = {}) {
  const value = headers.authorization || headers.Authorization || "";
  const m = String(value).match(/^Bearer\s+(.+)$/i);
  return m ? m[1].trim() : "";
}

async function getJose() {
  if (!joseModPromise) joseModPromise = import("jose");
  return joseModPromise;
}

async function verifyOrgToken(token) {
  if (!AAD_TENANT_ID || !AAD_CLIENT_ID) {
    throw new Error("Server auth configuration is incomplete");
  }

  const { createRemoteJWKSet, jwtVerify } = await getJose();
  if (!microsoftJwks) {
    microsoftJwks = createRemoteJWKSet(new URL(`https://login.microsoftonline.com/${AAD_TENANT_ID}/discovery/v2.0/keys`));
  }

  const { payload } = await jwtVerify(token, microsoftJwks, {
    audience: AAD_CLIENT_ID,
    issuer: [
      `https://login.microsoftonline.com/${AAD_TENANT_ID}/v2.0`,
      `https://sts.windows.net/${AAD_TENANT_ID}/`,
    ],
  });

  const email = String(payload.preferred_username || payload.upn || payload.email || "").trim().toLowerCase();
  const tid = String(payload.tid || "").trim().toLowerCase();
  if (!email) throw new Error("Token is missing an email claim");
  if (tid !== AAD_TENANT_ID.toLowerCase()) throw new Error("Token tenant mismatch");

  if (AUTH_POLICY === "emails") {
    if (!ALLOWED_EMAILS.size) throw new Error("ALLOWED_EMAILS is empty while AUTH_POLICY=emails");
    if (!ALLOWED_EMAILS.has(email)) throw new Error("This account is not allowed for this app");
  }

  return {
    email,
    oid: String(payload.oid || ""),
    name: String(payload.name || email),
  };
}

function toISODate(date) {
  return date.toISOString().slice(0, 10);
}

function yesterdayISO() {
  const parts = new Intl.DateTimeFormat("en-CA", {
    timeZone: "Asia/Manila",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  }).formatToParts(new Date(Date.now() - 24 * 60 * 60 * 1000));
  const year = parts.find(p => p.type === "year")?.value;
  const month = parts.find(p => p.type === "month")?.value;
  const day = parts.find(p => p.type === "day")?.value;
  return `${year}-${month}-${day}`;
}

function addDaysISO(isoDate, days) {
  const d = new Date(`${isoDate}T00:00:00Z`);
  d.setUTCDate(d.getUTCDate() + days);
  return toISODate(d);
}

function compareISO(a, b) {
  return a.localeCompare(b);
}

function daysBetweenISO(since, until) {
  const s = new Date(`${since}T00:00:00Z`);
  const u = new Date(`${until}T00:00:00Z`);
  return Math.floor((u - s) / 86400000);
}

function midDateISO(since, until) {
  const span = daysBetweenISO(since, until);
  return addDaysISO(since, Math.floor(span / 2));
}

function latestDate(rows) {
  return rows.reduce((max, r) => {
    const day = r?.day;
    if (!day) return max;
    if (!max || compareISO(day, max) > 0) return day;
    return max;
  }, null);
}

function mergeRows(existingRows, newRows) {
  const keyOf = r => `${r.account_id || ""}|${r.adset_name || ""}|${r.day || ""}`;
  const map = new Map(existingRows.map(r => [keyOf(r), r]));
  newRows.forEach(r => map.set(keyOf(r), r));
  return [...map.values()].sort((a, b) => (a.day || "").localeCompare(b.day || ""));
}

function readActionMetric(items, types) {
  const list = Array.isArray(types) ? types : [types];
  for (const t of list) {
    const val = items?.find(a => a.action_type === t)?.value;
    if (val != null) return val;
  }
  return "0";
}

function parseBusinessIds(input) {
  if (!input) return [];
  if (Array.isArray(input)) {
    return [...new Set(input.map(v => String(v).trim()).filter(Boolean))];
  }
  return [...new Set(String(input).split(",").map(v => v.trim()).filter(Boolean))];
}

function parseIni(content) {
  const result = {};
  const lines = String(content || "").split(/\r?\n/);
  for (const rawLine of lines) {
    const line = rawLine.trim();
    if (!line || line.startsWith("#") || line.startsWith(";")) continue;
    const eq = line.indexOf("=");
    if (eq < 0) continue;
    const key = line.slice(0, eq).trim();
    const value = line.slice(eq + 1).trim();
    if (key) result[key] = value;
  }
  return result;
}

function resolveMetaAccessToken(requestToken = "", settingsToken = "") {
  return String(META_ACCESS_TOKEN_ENV || requestToken || settingsToken || "").trim();
}

function sanitizeOptionList(value, fallback = []) {
  if (!Array.isArray(value)) return [...fallback];
  const cleaned = [...new Set(value.map(v => String(v || "").trim()).filter(Boolean))];
  return cleaned.length ? cleaned : [...fallback];
}

function parseOptionList(raw, fallback = []) {
  if (!raw) return [...fallback];
  try {
    const parsed = JSON.parse(raw);
    return sanitizeOptionList(parsed, fallback);
  } catch {
    return sanitizeOptionList(String(raw).split(","), fallback);
  }
}

function serializeIni(values) {
  const businessAccountId = String(values.businessAccountId || "").replace(/\r?\n/g, "").trim();
  const mappingOptions = values.mappingOptions || {};
  const divisions = JSON.stringify(sanitizeOptionList(mappingOptions.divisions, DEFAULT_MAPPING_OPTIONS.divisions));
  const lobs = JSON.stringify(sanitizeOptionList(mappingOptions.lobs, DEFAULT_MAPPING_OPTIONS.lobs));
  const segments = JSON.stringify(sanitizeOptionList(mappingOptions.segments, DEFAULT_MAPPING_OPTIONS.segments));
  const objectives = JSON.stringify(sanitizeOptionList(mappingOptions.objectives, DEFAULT_MAPPING_OPTIONS.objectives));
  return [
    "[meta]",
    `business_account_ids=${businessAccountId}`,
    `mapping_divisions=${divisions}`,
    `mapping_lobs=${lobs}`,
    `mapping_segments=${segments}`,
    `mapping_objectives=${objectives}`,
    "",
  ].join("\n");
}

async function loadSettingsIni() {
  try {
    const raw = await readStorageText(SETTINGS_FILE, "settings.ini");
    const parsed = parseIni(raw);
    return {
      accessToken: "",
      businessAccountId: parsed.business_account_ids || "",
      appId: META_APP_ID_ENV || "",
      appSecret: "",
      metaTokenConfigured: Boolean(META_ACCESS_TOKEN_ENV),
      metaAppConfigured: Boolean(META_APP_ID_ENV),
      metaAppSecretConfigured: Boolean(META_APP_SECRET_ENV),
      mappingOptions: {
        divisions: parseOptionList(parsed.mapping_divisions, DEFAULT_MAPPING_OPTIONS.divisions),
        lobs: parseOptionList(parsed.mapping_lobs, DEFAULT_MAPPING_OPTIONS.lobs),
        segments: parseOptionList(parsed.mapping_segments, DEFAULT_MAPPING_OPTIONS.segments),
        objectives: parseOptionList(parsed.mapping_objectives, DEFAULT_MAPPING_OPTIONS.objectives),
      },
    };
  } catch (err) {
    if (err.code === "ENOENT") {
      return {
        accessToken: "",
        businessAccountId: "",
        appId: META_APP_ID_ENV || "",
        appSecret: "",
        metaTokenConfigured: Boolean(META_ACCESS_TOKEN_ENV),
        metaAppConfigured: Boolean(META_APP_ID_ENV),
        metaAppSecretConfigured: Boolean(META_APP_SECRET_ENV),
        mappingOptions: { ...DEFAULT_MAPPING_OPTIONS },
      };
    }
    throw err;
  }
}

async function saveSettingsIni(values) {
  const content = serializeIni(values);
  await writeStorageText(SETTINGS_FILE, "settings.ini", content);
}

function normalizeIdentifierRow(row) {
  return {
    adset: String(row?.adset || "").trim(),
    adsetId: String(row?.adsetId || "").trim(),
    division: String(row?.division || "").trim(),
    lob: String(row?.lob || "").trim(),
    segment: String(row?.segment || "").trim(),
    objective: String(row?.objective || "").trim(),
  };
}

async function loadMappingsFile() {
  try {
    const raw = await readStorageText(MAPPINGS_FILE, "meta-mappings.json");
    const parsed = JSON.parse(raw);
    const list = Array.isArray(parsed?.identifiers) ? parsed.identifiers : [];
    return list.map(normalizeIdentifierRow).filter(r => r.adset || r.adsetId);
  } catch (err) {
    if (err.code === "ENOENT") return [];
    throw err;
  }
}

async function saveMappingsFile(identifiers) {
  const list = Array.isArray(identifiers) ? identifiers.map(normalizeIdentifierRow).filter(r => r.adset || r.adsetId) : [];
  await writeStorageText(MAPPINGS_FILE, "meta-mappings.json", JSON.stringify({ identifiers: list, updatedAt: new Date().toISOString() }));
  return list;
}

function normalizeNestedMap(input) {
  const src = input && typeof input === "object" ? input : {};
  const out = {};
  for (const month of Object.keys(src)) {
    const row = src[month] && typeof src[month] === "object" ? src[month] : {};
    const cleanRow = {};
    for (const key of Object.keys(row)) {
      const n = Number(row[key]);
      if (Number.isFinite(n)) cleanRow[key] = n;
    }
    if (Object.keys(cleanRow).length) out[month] = cleanRow;
  }
  return out;
}

function normalizeFlatMap(input) {
  const src = input && typeof input === "object" ? input : {};
  const out = {};
  for (const key of Object.keys(src)) {
    const n = Number(src[key]);
    if (Number.isFinite(n)) out[key] = n;
  }
  return out;
}

function normalizeBudgetTargets(payload) {
  const src = payload && typeof payload === "object" ? payload : {};
  return {
    budgets: normalizeFlatMap(src.budgets),
    targets: normalizeFlatMap(src.targets),
    desktopBudgets: normalizeNestedMap(src.desktopBudgets),
    desktopTargets: normalizeNestedMap(src.desktopTargets),
    ltsBudgets: normalizeNestedMap(src.ltsBudgets),
    ltsTargets: normalizeNestedMap(src.ltsTargets),
    lsaBudgets: normalizeNestedMap(src.lsaBudgets),
    lsaTargets: normalizeNestedMap(src.lsaTargets),
  };
}

async function loadBudgetTargetsFile() {
  try {
    const raw = await readStorageText(BUDGET_TARGETS_FILE, "meta-budget-targets.json");
    const parsed = JSON.parse(raw);
    return normalizeBudgetTargets(parsed?.budgetTargets || parsed || {});
  } catch (err) {
    if (err.code === "ENOENT") {
      return normalizeBudgetTargets({});
    }
    throw err;
  }
}

async function saveBudgetTargetsFile(payload) {
  const budgetTargets = normalizeBudgetTargets(payload);
  await writeStorageText(BUDGET_TARGETS_FILE, "meta-budget-targets.json", JSON.stringify({ budgetTargets, updatedAt: new Date().toISOString() }));
  return budgetTargets;
}

async function loadCache() {
  try {
    const raw = await readStorageText(CACHE_FILE, "meta-insights-cache.json");
    const parsed = JSON.parse(raw);
    return {
      rows: Array.isArray(parsed.rows) ? parsed.rows : [],
      fetchedAt: parsed.fetchedAt || null,
      since: parsed.since || null,
      until: parsed.until || null,
      accountsQueried: parsed.accountsQueried || 0,
      accountsSucceeded: parsed.accountsSucceeded || 0,
      errors: Array.isArray(parsed.errors) ? parsed.errors : [],
      businessNames: Array.isArray(parsed.businessNames) ? parsed.businessNames : [],
      discoveredAccounts: Array.isArray(parsed.discoveredAccounts) ? parsed.discoveredAccounts : [],
      cacheLastDate: parsed.cacheLastDate || null,
    };
  } catch (err) {
    if (err.code === "ENOENT") {
      // Backward compatibility: if old .cache file exists, load and migrate it.
      try {
        const legacyRaw = await fs.readFile(LEGACY_CACHE_FILE, "utf8");
        const legacy = JSON.parse(legacyRaw);
        const migrated = {
          rows: Array.isArray(legacy.rows) ? legacy.rows : [],
          fetchedAt: legacy.fetchedAt || null,
          since: legacy.since || null,
          until: legacy.until || null,
          accountsQueried: legacy.accountsQueried || 0,
          accountsSucceeded: legacy.accountsSucceeded || 0,
          errors: Array.isArray(legacy.errors) ? legacy.errors : [],
          businessNames: Array.isArray(legacy.businessNames) ? legacy.businessNames : [],
          discoveredAccounts: Array.isArray(legacy.discoveredAccounts) ? legacy.discoveredAccounts : [],
          cacheLastDate: legacy.cacheLastDate || null,
        };
        await saveCache(migrated.rows, {
          fetchedAt: migrated.fetchedAt,
          since: migrated.since,
          until: migrated.until,
          accountsQueried: migrated.accountsQueried,
          accountsSucceeded: migrated.accountsSucceeded,
          errors: migrated.errors,
          businessNames: migrated.businessNames,
          discoveredAccounts: migrated.discoveredAccounts,
          cacheLastDate: migrated.cacheLastDate,
        });
        return migrated;
      } catch (legacyErr) {
        if (legacyErr.code !== "ENOENT") throw legacyErr;
      }
      return { rows: [], fetchedAt: null, since: null, until: null, accountsQueried: 0, accountsSucceeded: 0, errors: [], businessNames: [], discoveredAccounts: [], cacheLastDate: null };
    }
    throw err;
  }
}

async function saveCache(rows, meta = {}) {
  await writeStorageText(CACHE_FILE, "meta-insights-cache.json", JSON.stringify({ rows, ...meta }));
}

async function clearCache() {
  await deleteStorageFile(CACHE_FILE, "meta-insights-cache.json");
  if (!isSharePointConfigured()) await deleteStorageFile(LEGACY_CACHE_FILE, "meta-insights-cache-legacy.json");
}

async function loadStatus() {
  try {
    const raw = await readStorageText(STATUS_FILE, "meta-insights-status.json");
    return JSON.parse(raw);
  } catch (err) {
    if (err.code === "ENOENT") {
      return { inProgress: false, updatedAt: null };
    }
    throw err;
  }
}

async function saveStatus(status) {
  await writeStorageText(STATUS_FILE, "meta-insights-status.json", JSON.stringify({ ...status, updatedAt: new Date().toISOString() }));
}

async function clearStatus() {
  await saveStatus({ inProgress: false, updatedAt: new Date().toISOString() });
}

async function fetchWithTimeout(url, options = {}, timeoutMs = 45000) {
  const ctrl = new AbortController();
  const timer = setTimeout(() => ctrl.abort(), timeoutMs);
  try {
    return await fetch(url, { ...options, signal: ctrl.signal });
  } catch (err) {
    if (err?.name === "AbortError") {
      throw new Error(`Request timeout after ${timeoutMs}ms`);
    }
    throw err;
  } finally {
    clearTimeout(timer);
  }
}

function shortText(text, max = 220) {
  const s = String(text || "").replace(/\s+/g, " ").trim();
  return s.length > max ? `${s.slice(0, max)}...` : s;
}

async function fetchJsonOrText(url, options = {}, timeoutMs = 45000) {
  const res = await fetchWithTimeout(url, options, timeoutMs);
  const text = await res.text().catch(() => "");
  let json = null;
  try {
    json = text ? JSON.parse(text) : null;
  } catch {
    json = null;
  }
  return { res, json, text };
}

async function fetchBusinessName(businessId, accessToken) {
  try {
    const params = new URLSearchParams({ fields: "name", access_token: accessToken });
    const { json } = await fetchJsonOrText(`${META_BASE}/${businessId}?${params}`);
    if (json.error) return null;
    return json.name || null;
  } catch {
    return null;
  }
}

async function fetchOwnedAdAccountsForBusiness(accessToken, businessId) {
  const ids = [];
  const params = new URLSearchParams({ fields: "account_id,name", access_token: accessToken, limit: 100 });
  let url = `${META_BASE}/${businessId}/owned_ad_accounts?${params}`;
  while (url) {
    const { json } = await fetchJsonOrText(url);
    if (json.error) throw new Error(`Business ${businessId} lookup failed: ${json.error.message}`);
    (json.data || []).forEach(a => ids.push(a.account_id));
    url = json.paging?.next || null;
  }
  return ids;
}

async function fetchAdAccountIds(accessToken, businessIds = []) {
  if (businessIds.length) {
    const merged = [];
    for (const bid of businessIds) {
      const ids = await fetchOwnedAdAccountsForBusiness(accessToken, bid);
      merged.push(...ids);
    }
    const unique = [...new Set(merged)];
    if (!unique.length) throw new Error("No owned ad accounts found for the specified Business Account IDs");
    return unique;
  }

  const ids = [];
  const params = new URLSearchParams({ fields: "account_id,name", access_token: accessToken, limit: 100 });
  let url = `${META_BASE}/me/adaccounts?${params}`;
  while (url) {
    const { json } = await fetchJsonOrText(url);
    if (json.error) throw new Error(json.error.message);
    (json.data || []).forEach(a => ids.push(a.account_id));
    url = json.paging?.next || null;
  }
  return ids;
}

async function fetchInsightsForAccount(adAccountId, datePreset, since, until, accessToken) {
  const params = new URLSearchParams({
    level: "adset",
    fields: INSIGHTS_FIELDS,
    access_token: accessToken,
    limit: 500,
  });

  if (since && until) {
    params.set("time_range", JSON.stringify({ since, until }));
  } else if (datePreset) {
    params.set("date_preset", datePreset);
  } else {
    params.set("date_preset", "last_month");
  }

  params.set("time_increment", "1");

  let url = `${META_BASE}/act_${adAccountId.replace(/^act_/, "")}/insights?${params}`;
  const rows = [];

  while (url) {
    const { res, json, text } = await fetchJsonOrText(url, {}, 60000);
    if (!res.ok) {
      if (json?.error) throw new Error(formatMetaError(json.error, adAccountId));
      throw new Error(`Meta API error for account ${adAccountId}: HTTP ${res.status} ${shortText(text)}`);
    }
    if (!json || typeof json !== "object") {
      throw new Error(`Meta API error for account ${adAccountId}: Non-JSON response ${shortText(text)}`);
    }
    if (json.error) throw new Error(formatMetaError(json.error, adAccountId));

    const data = json.data || [];
    data.forEach(row => {
      const actions = row.actions || [];
      const actionValues = row.action_values || [];
      // Use only Messaging conversations started for inquiry totals.
      const inquiries = readActionMetric(actions, "onsite_conversion.messaging_conversation_started_7d");
      const postEngagement = readActionMetric(actions, "post_engagement");
      const addToCart = readActionMetric(actions, ["add_to_cart", "omni_add_to_cart"]);
      const addToCartValue = readActionMetric(actionValues, ["add_to_cart", "omni_add_to_cart"]);
      const purchaseValue = readActionMetric(actionValues, ["purchase", "omni_purchase"]);

      rows.push({
        campaign_name: row.campaign_name,
        adset_name: row.adset_name,
        day: row.date_start,
        account_name: row.account_name,
        account_id: adAccountId,
        currency: row.account_currency,
        spend: parseFloat(row.spend || 0),
        inquiries: parseInt(inquiries, 10),
        post_engagement: parseInt(postEngagement, 10),
        add_to_cart: parseInt(addToCart, 10),
        add_to_cart_conversion_value: parseFloat(addToCartValue || 0),
        purchases_conversion_value: parseFloat(purchaseValue || 0),
        reach: parseInt(row.reach || 0, 10),
        impressions: parseInt(row.impressions || 0, 10),
        clicks: parseInt(row.clicks || 0, 10),
        link_clicks: parseInt(row.inline_link_clicks || 0, 10),
        ctr_link: parseFloat(row.inline_link_click_ctr || 0),
        ctr_all: parseFloat(row.ctr || 0),
        frequency: parseFloat(row.frequency || 0),
        reporting_starts: row.date_start,
        reporting_ends: row.date_stop,
      });
    });

    url = json.paging?.next || null;
  }

  return rows;
}

function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

function isRateLimitError(message = "") {
  const msg = String(message).toLowerCase();
  return msg.includes("application request limit reached") || msg.includes("rate limit") || msg.includes("too many calls") || msg.includes("service temporarily unavailable");
}

function isUnknownMetaError(message = "") {
  return String(message).toLowerCase().includes("unknown error occurred");
}

function isInactivityTimeoutError(message = "") {
  const msg = String(message).toLowerCase();
  return msg.includes("inactivity timeout") || msg.includes("too much time has passed without sending any data") || msg.includes("request timeout") || msg.includes("timed out") || msg.includes("non-json response");
}

function isRangeTooLargeError(message = "") {
  const msg = String(message).toLowerCase();
  return msg.includes("please reduce the amount of data") || msg.includes("reduce the amount of data") || msg.includes("request code=1");
}

function formatMetaError(errorObj, adAccountId) {
  if (!errorObj) return `Meta API error for account ${adAccountId}: Unknown error`;
  const base = errorObj.message || "Unknown error";
  const code = errorObj.code != null ? ` code=${errorObj.code}` : "";
  const subcode = errorObj.error_subcode != null ? ` subcode=${errorObj.error_subcode}` : "";
  const type = errorObj.type ? ` type=${errorObj.type}` : "";
  return `Meta API error for account ${adAccountId}: ${base}${code}${subcode}${type}`;
}

async function fetchInsightsWithRetry(adAccountId, datePreset, since, until, accessToken, maxRetries = 3) {
  for (let attempt = 0; attempt <= maxRetries; attempt++) {
    try {
      return await fetchInsightsForAccount(adAccountId, datePreset, since, until, accessToken);
    } catch (err) {
      const message = err?.message || "Unknown Meta API error";
      const retryable = isRateLimitError(message) || isUnknownMetaError(message) || isRangeTooLargeError(message) || isInactivityTimeoutError(message);

      if (isRangeTooLargeError(message) && since && until && since !== until) {
        const mid = midDateISO(since, until);
        const next = addDaysISO(mid, 1);
        const left = await fetchInsightsWithRetry(adAccountId, datePreset, since, mid, accessToken, 1);
        const right = compareISO(next, until) <= 0 ? await fetchInsightsWithRetry(adAccountId, datePreset, next, until, accessToken, 1) : [];
        return [...left, ...right];
      }

      if (isUnknownMetaError(message) && since && until && since !== until) {
        const rows = [];
        let day = since;
        while (compareISO(day, until) <= 0) {
          const dayRows = await fetchInsightsWithRetry(adAccountId, datePreset, day, day, accessToken, 1);
          rows.push(...dayRows);
          day = addDaysISO(day, 1);
        }
        return rows;
      }

      if (isInactivityTimeoutError(message) && since && until && since !== until) {
        const mid = midDateISO(since, until);
        const next = addDaysISO(mid, 1);
        const left = await fetchInsightsWithRetry(adAccountId, datePreset, since, mid, accessToken, 1);
        const right = compareISO(next, until) <= 0 ? await fetchInsightsWithRetry(adAccountId, datePreset, next, until, accessToken, 1) : [];
        return [...left, ...right];
      }

      if (!retryable || attempt === maxRetries) throw new Error(message);
      const delayMs = (1000 * Math.pow(2, attempt)) + Math.floor(Math.random() * 250);
      await sleep(delayMs);
    }
  }
}

exports.handler = async function (event) {
  const headers = {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Headers": "Content-Type, Authorization",
    "Access-Control-Allow-Methods": "POST, OPTIONS",
    "Content-Type": "application/json",
  };

  if (event.httpMethod === "OPTIONS") {
    return { statusCode: 204, headers, body: "" };
  }

  if (event.httpMethod !== "POST") {
    return { statusCode: 405, headers, body: JSON.stringify({ error: "Method not allowed" }) };
  }

  const bearerToken = getBearerToken(event.headers || {});
  if (!bearerToken) {
    return { statusCode: 401, headers, body: JSON.stringify({ error: "Unauthorized: missing bearer token" }) };
  }

  try {
    await verifyOrgToken(bearerToken);
  } catch (err) {
    return { statusCode: 401, headers, body: JSON.stringify({ error: `Unauthorized: ${err.message}` }) };
  }

  let body;
  try {
    body = JSON.parse(event.body || "{}");
  } catch {
    return { statusCode: 400, headers, body: JSON.stringify({ error: "Invalid JSON body" }) };
  }

  const { accessToken, datePreset, since, until, action } = body;
  const businessIds = parseBusinessIds(body.businessAccountId);

  if (action === "load") {
    try {
      const cache = await loadCache();
      return {
        statusCode: 200,
        headers,
        body: JSON.stringify({
          data: cache.rows,
          meta: {
            totalRows: cache.rows.length,
            fetchedAt: cache.fetchedAt,
            since: cache.since,
            until: cache.until,
            accountsQueried: cache.accountsQueried,
            accountsSucceeded: cache.accountsSucceeded,
            discoveredAccounts: cache.discoveredAccounts,
            errors: cache.errors,
            businessNames: cache.businessNames,
            cacheLastDate: cache.cacheLastDate || latestDate(cache.rows),
            fromCache: true,
          },
        }),
      };
    } catch (err) {
      return { statusCode: 500, headers, body: JSON.stringify({ error: err.message }) };
    }
  }

  if (action === "status") {
    try {
      const status = await loadStatus();
      return { statusCode: 200, headers, body: JSON.stringify({ status }) };
    } catch (err) {
      return { statusCode: 500, headers, body: JSON.stringify({ error: err.message }) };
    }
  }

  if (action === "clear") {
    try {
      await clearCache();
      await clearStatus();
      return {
        statusCode: 200,
        headers,
        body: JSON.stringify({
          data: [],
          meta: {
            totalRows: 0,
            fetchedAt: null,
            since: null,
            until: null,
            accountsQueried: 0,
            accountsSucceeded: 0,
            discoveredAccounts: [],
            errors: [],
            businessNames: [],
            cacheLastDate: null,
            fromCache: true,
            cacheCleared: true,
          },
        }),
      };
    } catch (err) {
      return { statusCode: 500, headers, body: JSON.stringify({ error: err.message }) };
    }
  }

  if (action === "load_settings") {
    try {
      const settings = await loadSettingsIni();
      return {
        statusCode: 200,
        headers,
        body: JSON.stringify({ settings }),
      };
    } catch (err) {
      return { statusCode: 500, headers, body: JSON.stringify({ error: err.message }) };
    }
  }

  if (action === "save_settings") {
    try {
      await saveSettingsIni({
        businessAccountId: body.businessAccountId || "",
        mappingOptions: body.mappingOptions || DEFAULT_MAPPING_OPTIONS,
      });
      const settings = await loadSettingsIni();
      return {
        statusCode: 200,
        headers,
        body: JSON.stringify({ ok: true, settings }),
      };
    } catch (err) {
      return { statusCode: 500, headers, body: JSON.stringify({ error: err.message }) };
    }
  }

  if (action === "load_mappings") {
    try {
      const identifiers = await loadMappingsFile();
      return {
        statusCode: 200,
        headers,
        body: JSON.stringify({ identifiers }),
      };
    } catch (err) {
      return { statusCode: 500, headers, body: JSON.stringify({ error: err.message }) };
    }
  }

  if (action === "save_mappings") {
    try {
      const identifiers = await saveMappingsFile(body.identifiers || []);
      return {
        statusCode: 200,
        headers,
        body: JSON.stringify({ ok: true, identifiers }),
      };
    } catch (err) {
      return { statusCode: 500, headers, body: JSON.stringify({ error: err.message }) };
    }
  }

  if (action === "load_budget_targets") {
    try {
      const budgetTargets = await loadBudgetTargetsFile();
      return {
        statusCode: 200,
        headers,
        body: JSON.stringify({ budgetTargets }),
      };
    } catch (err) {
      return { statusCode: 500, headers, body: JSON.stringify({ error: err.message }) };
    }
  }

  if (action === "save_budget_targets") {
    try {
      const budgetTargets = await saveBudgetTargetsFile(body.budgetTargets || {});
      return {
        statusCode: 200,
        headers,
        body: JSON.stringify({ ok: true, budgetTargets }),
      };
    } catch (err) {
      return { statusCode: 500, headers, body: JSON.stringify({ error: err.message }) };
    }
  }

  const savedSettings = await loadSettingsIni();
  const effectiveAccessToken = resolveMetaAccessToken(accessToken, savedSettings.accessToken || "");

  if (!effectiveAccessToken) {
    return { statusCode: 400, headers, body: JSON.stringify({ error: "accessToken is required" }) };
  }

  try {
    const cache = await loadCache();
    const yday = yesterdayISO();
    const initialSince = since || "2025-01-01";
    const cacheLastDate = cache.cacheLastDate || latestDate(cache.rows);
    const syncSince = cacheLastDate ? addDaysISO(cacheLastDate, 1) : initialSince;
    const syncUntil = until || yday;

    // Nothing new to fetch; return cache only.
    if (compareISO(syncSince, syncUntil) > 0) {
      const meta = {
        totalRows: cache.rows.length,
        discoveredAccounts: cache.discoveredAccounts,
        fetchedAt: cache.fetchedAt,
        since: cache.since,
        until: cache.until,
        accountsQueried: cache.accountsQueried,
        accountsSucceeded: cache.accountsSucceeded,
        errors: cache.errors,
        businessNames: cache.businessNames,
        cacheLastDate,
        syncSince,
        syncUntil,
        syncedNow: false,
        fromCache: true,
      };
      await clearStatus();
      return { statusCode: 200, headers, body: JSON.stringify({ data: cache.rows, meta }) };
    }

    const businessNames = [];
    for (const bid of businessIds) {
      const name = await fetchBusinessName(bid, effectiveAccessToken);
      if (name) businessNames.push(name);
    }

    let adAccountIds;
    try {
      adAccountIds = await fetchAdAccountIds(effectiveAccessToken, businessIds);
    } catch (e) {
      return { statusCode: 400, headers, body: JSON.stringify({ error: `Failed to discover ad accounts: ${e.message}` }) };
    }

    if (!adAccountIds.length) {
      return { statusCode: 400, headers, body: JSON.stringify({ error: "No ad accounts found for this access token" }) };
    }

    const startedAt = new Date().toISOString();
    const accountStatuses = adAccountIds.map(id => ({ accountId: id, status: "pending", rows: 0, error: null }));
    await saveStatus({
      inProgress: true,
      startedAt,
      totalAccounts: adAccountIds.length,
      completedAccounts: 0,
      currentAccountId: adAccountIds[0],
      currentAccountIndex: 1,
      syncSince,
      syncUntil,
      accounts: accountStatuses,
      message: "Starting sync",
    });

    const fetchedRows = [];
    const errors = [];
    for (let i = 0; i < adAccountIds.length; i++) {
      const id = adAccountIds[i];
      await saveStatus({
        inProgress: true,
        startedAt,
        totalAccounts: adAccountIds.length,
        completedAccounts: i,
        currentAccountId: id,
        currentAccountIndex: i + 1,
        syncSince,
        syncUntil,
        accounts: accountStatuses,
        message: `Fetching account ${i + 1} of ${adAccountIds.length}`,
      });

      try {
        const rows = await fetchInsightsWithRetry(id, datePreset, syncSince, syncUntil, effectiveAccessToken);
        fetchedRows.push(...rows);
        accountStatuses[i] = { accountId: id, status: "success", rows: rows.length, error: null };
      } catch (err) {
        const msg = err?.message || "Unknown Meta API error";
        errors.push({ accountId: id, error: msg });
        accountStatuses[i] = { accountId: id, status: "error", rows: 0, error: msg };
      }

      await saveStatus({
        inProgress: true,
        startedAt,
        totalAccounts: adAccountIds.length,
        completedAccounts: i + 1,
        currentAccountId: id,
        currentAccountIndex: i + 1,
        syncSince,
        syncUntil,
        accounts: accountStatuses,
        message: `Completed ${i + 1} of ${adAccountIds.length}`,
      });
    }

    const merged = mergeRows(cache.rows, fetchedRows);
    const fetchedAt = new Date().toISOString();
    const nextCacheLastDate = latestDate(merged);

    const meta = {
      totalRows: merged.length,
      discoveredAccounts: adAccountIds,
      fetchedAt,
      since: cache.since || initialSince,
      until: syncUntil,
      accountsQueried: adAccountIds.length,
      accountsSucceeded: adAccountIds.length - errors.length,
      errors,
      businessNames,
      cacheLastDate: nextCacheLastDate,
      syncSince,
      syncUntil,
      syncedNow: true,
      fromCache: false,
    };

    await saveCache(merged, {
      fetchedAt,
      since: cache.since || initialSince,
      until: syncUntil,
      accountsQueried: adAccountIds.length,
      accountsSucceeded: adAccountIds.length - errors.length,
      discoveredAccounts: adAccountIds,
      errors,
      businessNames,
      cacheLastDate: nextCacheLastDate,
    });

    await saveStatus({
      inProgress: false,
      startedAt,
      finishedAt: new Date().toISOString(),
      totalAccounts: adAccountIds.length,
      completedAccounts: adAccountIds.length,
      currentAccountId: null,
      currentAccountIndex: adAccountIds.length,
      syncSince,
      syncUntil,
      accounts: accountStatuses,
      message: "Sync complete",
    });

    return { statusCode: 200, headers, body: JSON.stringify({ data: merged, meta }) };
  } catch (err) {
    await clearStatus();
    return { statusCode: 500, headers, body: JSON.stringify({ error: err.message }) };
  }
};
