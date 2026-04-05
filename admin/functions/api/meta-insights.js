// Cloudflare Pages Function: /api/meta-insights
// Replaces the local Node.js server/meta-insights.js for cloud deployment.
// Reads/writes data files via R2 binding instead of local filesystem.
//
// Required Cloudflare bindings & env vars:
//   BUCKET          — R2 bucket binding (name: BUCKET)
//   META_ACCESS_TOKEN — Meta Graph API access token (secret)
//   META_APP_ID     — Meta App ID (optional)
//   META_APP_SECRET — Meta App Secret (optional)
//   BUSINESS_ACCOUNT_IDS — comma-separated business account IDs (optional)

const META_API_VERSION = "v19.0";
const META_BASE = `https://graph.facebook.com/${META_API_VERSION}`;

const INSIGHTS_FIELDS = [
  "campaign_name","adset_name","adset_id","account_name","account_currency",
  "spend","actions","action_values","reach","impressions","clicks",
  "inline_link_clicks","ctr","inline_link_click_ctr","frequency",
  "date_start","date_stop",
].join(",");

function json(data, status = 200) {
  return new Response(JSON.stringify(data), {
    status,
    headers: { "Content-Type": "application/json" },
  });
}

async function r2Get(bucket, key, fallback = null) {
  try {
    const obj = await bucket.get(key);
    if (!obj) return fallback;
    return await obj.json();
  } catch {
    return fallback;
  }
}

async function r2Put(bucket, key, data) {
  await bucket.put(key, JSON.stringify(data, null, 2), {
    httpMetadata: { contentType: "application/json" },
  });
}

// ── Settings ─────────────────────────────────────────────────────────────────

async function handleLoadSettings(bucket, env) {
  const selections = await r2Get(bucket, "meta-selection-lists.json", {});
  return json({
    settings: {
      // Credentials come from Cloudflare env vars (not stored in R2 for security)
      accessToken: env.META_ACCESS_TOKEN || "",
      appSecret: env.META_APP_SECRET || "",
      appId: env.META_APP_ID || "",
      businessAccountId: env.BUSINESS_ACCOUNT_IDS || "",
      mappingOptions: selections.mappingOptions || null,
    },
  });
}

async function handleSaveSettings(bucket, body) {
  // Only save non-secret fields to R2 (credentials are managed via Cloudflare env vars)
  const existing = await r2Get(bucket, "meta-selection-lists.json", {});
  const updated = {
    ...existing,
    ...(body.mappingOptions ? { mappingOptions: body.mappingOptions } : {}),
  };
  await r2Put(bucket, "meta-selection-lists.json", updated);
  return json({ ok: true });
}

// ── Mappings ──────────────────────────────────────────────────────────────────

async function handleLoadMappings(bucket) {
  const data = await r2Get(bucket, "meta-mappings.json", { identifiers: [] });
  const identifiers = Array.isArray(data?.identifiers) ? data.identifiers
    : Array.isArray(data) ? data
    : [];
  return json({ identifiers });
}

async function handleSaveMappings(bucket, body) {
  const identifiers = Array.isArray(body.identifiers) ? body.identifiers : [];
  await r2Put(bucket, "meta-mappings.json", { identifiers });
  return json({ ok: true });
}

// ── Budget Targets ────────────────────────────────────────────────────────────

async function handleLoadBudgetTargets(bucket) {
  const data = await r2Get(bucket, "meta-budget-targets.json", {});
  const budgetTargets = data?.budgetTargets || data || {};
  return json({ budgetTargets });
}

async function handleSaveBudgetTargets(bucket, body) {
  const budgetTargets = body.budgetTargets || {};
  await r2Put(bucket, "meta-budget-targets.json", { budgetTargets });
  return json({ ok: true });
}

// ── Cache Load ────────────────────────────────────────────────────────────────

async function handleLoadCache(bucket) {
  const data = await r2Get(bucket, "meta-insights-cache.json", null);
  if (!data) return json({ rows: [], meta: { totalRows: 0, fromCache: true } });
  const rows = Array.isArray(data?.rows) ? data.rows
    : Array.isArray(data?.data) ? data.data
    : Array.isArray(data) ? data
    : [];
  const meta = data?.meta || {
    totalRows: rows.length,
    fetchedAt: null,
    fromCache: true,
  };
  return json({ rows, meta });
}

// ── Status ────────────────────────────────────────────────────────────────────

async function handleStatus(bucket) {
  const data = await r2Get(bucket, "meta-insights-cache.json", null);
  const rows = Array.isArray(data?.rows) ? data.rows : [];
  return json({
    status: "cloud",
    totalRows: rows.length,
    fetchedAt: data?.meta?.fetchedAt || null,
    fromCache: true,
  });
}

// ── Meta API Sync ─────────────────────────────────────────────────────────────

function readAction(actions, types) {
  const list = Array.isArray(types) ? types : [types];
  for (const t of list) {
    const val = actions?.find(a => a.action_type === t)?.value;
    if (val != null) return val;
  }
  return "0";
}

function normalizeRow(item, accountId) {
  const actions = item.actions || [];
  return {
    account_id: String(accountId || item.account_id || ""),
    account_name: item.account_name || "",
    campaign_name: item.campaign_name || "",
    adset_name: item.adset_name || "",
    adset_id: item.adset_id || "",
    day: item.date_start || "",
    spend: parseFloat(item.spend || 0),
    inquiries: parseInt(readAction(actions, ["lead", "onsite_web_lead", "offsite_conversion.fb_pixel_lead"]), 10),
    post_engagement: parseInt(readAction(actions, ["post_engagement"]), 10),
    reach: parseInt(item.reach || 0, 10),
    impressions: parseInt(item.impressions || 0, 10),
    clicks: parseInt(item.inline_link_clicks || item.clicks || 0, 10),
    ctr: parseFloat(item.inline_link_click_ctr || item.ctr || 0),
  };
}

function mergeRows(existing, incoming) {
  const key = r => `${r.account_id}|${r.adset_id || r.adset_name}|${r.day}`;
  const map = new Map(existing.map(r => [key(r), r]));
  incoming.forEach(r => map.set(key(r), r));
  return [...map.values()].sort((a, b) => (a.day || "").localeCompare(b.day || ""));
}

async function metaFetch(url, token) {
  const res = await fetch(url, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!res.ok) throw new Error(`Meta API error ${res.status}: ${await res.text()}`);
  return res.json();
}

async function fetchAccountIds(token, businessIds) {
  const ids = [];
  for (const biz of businessIds) {
    try {
      const data = await metaFetch(
        `${META_BASE}/${biz}/owned_ad_accounts?fields=account_id,name&limit=200`,
        token
      );
      (data.data || []).forEach(a => ids.push(a.account_id || a.id));
    } catch {
      // skip failed business accounts
    }
  }
  return [...new Set(ids)];
}

async function fetchInsightsForAccount(accountId, token, since, until) {
  const params = new URLSearchParams({
    level: "adset",
    fields: INSIGHTS_FIELDS,
    time_range: JSON.stringify({ since, until }),
    time_increment: "1",
    limit: "500",
  });
  let url = `${META_BASE}/act_${accountId}/insights?${params}`;
  const rows = [];
  while (url) {
    const page = await metaFetch(url, token);
    (page.data || []).forEach(item => rows.push(normalizeRow(item, accountId)));
    url = page.paging?.next || null;
  }
  return rows;
}

async function handleSync(bucket, body, env) {
  const token = env.META_ACCESS_TOKEN || body.accessToken || "";
  if (!token) return json({ error: "META_ACCESS_TOKEN not configured" }, 400);

  const rawBizIds = env.BUSINESS_ACCOUNT_IDS || body.businessAccountId || "";
  const businessIds = String(rawBizIds).split(",").map(v => v.trim()).filter(Boolean);

  // Default since: day after the latest cached date, or 2025-01-01
  const cached = await r2Get(bucket, "meta-insights-cache.json", { rows: [] });
  const existingRows = Array.isArray(cached?.rows) ? cached.rows : [];
  const cacheLastDate = existingRows.reduce((max, r) => {
    const d = r?.day || "";
    return (!max || d.localeCompare(max) > 0) ? d : max;
  }, null);

  // Yesterday in PHT (UTC+8)
  const phtYesterday = (() => {
    const now = new Date(Date.now() + 8 * 60 * 60 * 1000); // shift to PHT
    now.setUTCDate(now.getUTCDate() - 1); // go back one day
    return now.toISOString().slice(0, 10);
  })();

  // Re-fetch from last cached date (not +1) so the last day always gets refreshed with complete data
  const since = body.since || cacheLastDate || "2025-01-01";
  const until = body.until || phtYesterday;

  // Nothing to fetch — cache is already up to date
  if (since > until) {
    return json({
      rows: existingRows,
      meta: {
        totalRows: existingRows.length,
        fetchedAt: cached?.meta?.fetchedAt || null,
        since,
        until,
        accountsQueried: 0,
        accountsSucceeded: 0,
        discoveredAccounts: [],
        businessNames: [],
        errors: [],
        syncedNow: false,
        fromCache: true,
      },
    });
  }

  try {
    const accountIds = businessIds.length
      ? await fetchAccountIds(token, businessIds)
      : (body.accountIds || []);

    if (!accountIds.length) return json({ error: "No ad account IDs found" }, 400);

    // existingRows already loaded above for date calculation

    const newRows = [];
    const errors = [];
    const businessNames = [];
    const discoveredAccounts = [];

    for (const accountId of accountIds) {
      try {
        const rows = await fetchInsightsForAccount(accountId, token, since, until);
        newRows.push(...rows);
        discoveredAccounts.push(accountId);
      } catch (err) {
        errors.push({ accountId, error: err.message });
      }
    }

    const merged = mergeRows(existingRows, newRows);
    const fetchedAt = new Date().toISOString();

    await r2Put(bucket, "meta-insights-cache.json", {
      rows: merged,
      meta: {
        totalRows: merged.length,
        fetchedAt,
        since,
        until,
        accountsQueried: accountIds.length,
        accountsSucceeded: discoveredAccounts.length,
        discoveredAccounts,
        businessNames,
        errors,
        fromCache: false,
      },
    });

    return json({
      rows: merged,
      meta: {
        totalRows: merged.length,
        fetchedAt,
        since,
        until,
        accountsQueried: accountIds.length,
        accountsSucceeded: discoveredAccounts.length,
        discoveredAccounts,
        businessNames,
        errors,
      },
    });
  } catch (err) {
    return json({ error: err?.message || "Sync failed" }, 500);
  }
}

// ── Main handler ──────────────────────────────────────────────────────────────

export async function onRequest(context) {
  const { request, env } = context;

  if (!env.BUCKET) {
    return json({ error: "R2 bucket binding (BUCKET) is not configured in Cloudflare dashboard" }, 500);
  }

  let body = {};
  try { body = await request.json(); } catch {}

  const { action } = body;

  try {
    switch (action) {
      case "load_settings":     return handleLoadSettings(env.BUCKET, env);
      case "save_settings":     return handleSaveSettings(env.BUCKET, body);
      case "load_mappings":     return handleLoadMappings(env.BUCKET);
      case "save_mappings":     return handleSaveMappings(env.BUCKET, body);
      case "load_budget_targets": return handleLoadBudgetTargets(env.BUCKET);
      case "save_budget_targets": return handleSaveBudgetTargets(env.BUCKET, body);
      case "load":              return handleLoadCache(env.BUCKET);
      case "status":            return handleStatus(env.BUCKET);
      case "clear":             return json({ ok: true });
      default:                  return handleSync(env.BUCKET, body, env);
    }
  } catch (err) {
    return json({ error: err?.message || "Internal error" }, 500);
  }
}
