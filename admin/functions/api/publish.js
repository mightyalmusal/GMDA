// Cloudflare Pages Function: POST /api/publish
// Increments the manifest version so viewers know to re-download data from R2.
// Data files are already in R2 (written by meta-insights function on save/sync).
// R2 binding name: BUCKET (configure in Cloudflare dashboard)

const DATA_FILES = [
  "meta-insights-cache.json",
  "meta-budget-targets.json",
  "meta-mappings.json",
  "meta-selection-lists.json",
];

function json(data, status = 200) {
  return new Response(JSON.stringify(data), {
    status,
    headers: { "Content-Type": "application/json" },
  });
}

function getBearerToken(headers) {
  const value = headers.get("authorization") || "";
  const m = value.match(/^Bearer\s+(.+)$/i);
  return m ? m[1].trim() : "";
}

function decodeJwtPayload(token) {
  try {
    const parts = token.split(".");
    if (parts.length < 2) return {};
    const padded = parts[1].replace(/-/g, "+").replace(/_/g, "/");
    return JSON.parse(atob(padded));
  } catch {
    return {};
  }
}

export async function onRequestPost(context) {
  const { request, env } = context;

  if (!env.BUCKET) {
    return json({ error: "R2 bucket binding (BUCKET) is not configured" }, 500);
  }

  try {
    // Identify publisher from JWT
    const token = getBearerToken(request.headers);
    let publishedBy = "admin";
    if (token) {
      const payload = decodeJwtPayload(token);
      publishedBy = payload.preferred_username || payload.upn || payload.email || publishedBy;
    }

    // Read current manifest version
    let currentVersion = 0;
    try {
      const obj = await env.BUCKET.get("manifest.json");
      if (obj) {
        const manifest = await obj.json();
        currentVersion = Number(manifest.version) || 0;
      }
    } catch {
      // No manifest yet — start at version 0
    }

    const newVersion = currentVersion + 1;
    const publishedAt = new Date().toISOString();

    await env.BUCKET.put(
      "manifest.json",
      JSON.stringify({ version: newVersion, publishedAt, publishedBy, files: DATA_FILES }, null, 2),
      { httpMetadata: { contentType: "application/json" } }
    );

    return json({ success: true, version: newVersion, publishedAt });
  } catch (err) {
    return json({ error: err?.message || "Internal Server Error" }, 500);
  }
}
