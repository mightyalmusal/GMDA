import fs from "node:fs/promises";
import path from "node:path";

function resolveInputPath() {
  const argPath = process.argv[2];
  if (argPath) return path.resolve(process.cwd(), argPath);
  return path.resolve(process.cwd(), ".cache", "meta-insights-cache.json");
}

function getEnv(name) {
  const value = String(process.env[name] || "").trim();
  if (!value) throw new Error(`Missing required environment variable: ${name}`);
  return value;
}

async function main() {
  const inputPath = resolveInputPath();
  const raw = await fs.readFile(inputPath, "utf8");
  const parsed = JSON.parse(raw);
  const rows = Array.isArray(parsed?.rows) ? parsed.rows : (Array.isArray(parsed?.data) ? parsed.data : []);

  const supabaseUrl = getEnv("SUPABASE_URL").replace(/\/$/, "");
  const supabaseKey = getEnv("SUPABASE_SERVICE_ROLE_KEY");
  const table = String(process.env.SUPABASE_TABLE || "app_state").trim();

  const payload = [{ key: "meta-insights-cache.json", value: raw }];
  const endpoint = `${supabaseUrl}/rest/v1/${encodeURIComponent(table)}`;

  const res = await fetch(endpoint, {
    method: "POST",
    headers: {
      apikey: supabaseKey,
      Authorization: `Bearer ${supabaseKey}`,
      "Content-Type": "application/json",
      Prefer: "resolution=merge-duplicates,return=minimal",
    },
    body: JSON.stringify(payload),
  });

  if (!res.ok) {
    const text = await res.text().catch(() => "");
    throw new Error(`Supabase upsert failed (${res.status}): ${text || res.statusText}`);
  }

  console.log(`Imported cache to Supabase from ${inputPath}`);
  console.log(`Rows detected: ${rows.length.toLocaleString()}`);
  console.log(`Target key: meta-insights-cache.json`);
}

main().catch((err) => {
  console.error(err.message || err);
  process.exitCode = 1;
});
