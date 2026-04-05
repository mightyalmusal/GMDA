// Sync logic: checks R2 manifest version, downloads updated files, stores in IndexedDB
import { getVersion, saveVersion, saveRows, saveBudgets, saveMappings, saveSelections } from './db.js';

const R2_PUBLIC_URL = (import.meta.env.VITE_R2_PUBLIC_URL || '').replace(/\/$/, '');

export async function fetchManifest() {
  const res = await fetch(`${R2_PUBLIC_URL}/manifest.json`, { cache: 'no-store' });
  if (!res.ok) throw new Error(`Failed to fetch manifest: ${res.status}`);
  return res.json();
}

async function fetchJsonFile(filename) {
  const res = await fetch(`${R2_PUBLIC_URL}/${filename}`, { cache: 'no-store' });
  if (!res.ok) throw new Error(`Failed to fetch ${filename}: ${res.status}`);
  return res.json();
}

export async function syncIfNeeded() {
  const manifest = await fetchManifest();
  const remoteVersion = Number(manifest.version) || 0;
  const localVersion = await getVersion();

  if (remoteVersion <= localVersion) {
    return { updated: false, version: localVersion };
  }

  // Download all files in parallel
  const [cacheData, budgets, mappings, selections] = await Promise.all([
    fetchJsonFile('meta-insights-cache.json'),
    fetchJsonFile('meta-budget-targets.json'),
    fetchJsonFile('meta-mappings.json'),
    fetchJsonFile('meta-selection-lists.json'),
  ]);

  // Store in IndexedDB
  // meta-insights-cache.json has a "rows" array (or "data" array) at root
  const rows = Array.isArray(cacheData?.rows) ? cacheData.rows
    : Array.isArray(cacheData?.data) ? cacheData.data
    : Array.isArray(cacheData) ? cacheData
    : [];

  // meta-mappings.json may be {"identifiers":[...]} or a raw array
  const identifierList = Array.isArray(mappings) ? mappings
    : Array.isArray(mappings?.identifiers) ? mappings.identifiers
    : [];

  await Promise.all([
    saveRows(rows),
    saveBudgets(budgets || {}),
    saveMappings(identifierList),
    saveSelections(selections || {}),
    saveVersion(remoteVersion),
  ]);

  return { updated: true, version: remoteVersion };
}
