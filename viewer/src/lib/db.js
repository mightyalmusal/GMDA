// IndexedDB wrapper for marketing hub viewer
const DB_NAME = 'marketing-hub';
const DB_VERSION = 1;
const STORES = {
  meta: 'meta',           // manifest version info
  rows: 'rows',           // meta-insights-cache rows (array stored as single record)
  budgets: 'budgets',     // meta-budget-targets.json
  mappings: 'mappings',   // meta-mappings.json
  selections: 'selections' // meta-selection-lists.json
};

function openDB() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, DB_VERSION);
    req.onupgradeneeded = (e) => {
      const db = e.target.result;
      Object.values(STORES).forEach(name => {
        if (!db.objectStoreNames.contains(name)) {
          db.createObjectStore(name);
        }
      });
    };
    req.onsuccess = (e) => resolve(e.target.result);
    req.onerror = () => reject(req.error);
  });
}

async function dbGet(store, key) {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(store, 'readonly');
    const req = tx.objectStore(store).get(key);
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

async function dbPut(store, key, value) {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(store, 'readwrite');
    const req = tx.objectStore(store).put(value, key);
    req.onsuccess = () => resolve();
    req.onerror = () => reject(req.error);
  });
}

export async function getVersion() {
  try {
    const v = await dbGet(STORES.meta, 'version');
    return typeof v === 'number' ? v : -1;
  } catch {
    return -1;
  }
}

export async function saveVersion(v) {
  await dbPut(STORES.meta, 'version', v);
}

export async function saveRows(rows) {
  await dbPut(STORES.rows, 'data', rows);
}

export async function loadRows() {
  try {
    const rows = await dbGet(STORES.rows, 'data');
    return Array.isArray(rows) ? rows : [];
  } catch {
    return [];
  }
}

export async function saveBudgets(data) {
  await dbPut(STORES.budgets, 'data', data);
}

export async function loadBudgets() {
  try {
    return (await dbGet(STORES.budgets, 'data')) || {};
  } catch {
    return {};
  }
}

export async function saveMappings(data) {
  await dbPut(STORES.mappings, 'data', data);
}

export async function loadMappings() {
  try {
    const m = await dbGet(STORES.mappings, 'data');
    return Array.isArray(m) ? m : [];
  } catch {
    return [];
  }
}

export async function saveSelections(data) {
  await dbPut(STORES.selections, 'data', data);
}

export async function loadSelections() {
  try {
    return (await dbGet(STORES.selections, 'data')) || {};
  } catch {
    return {};
  }
}

export async function loadAllData() {
  const [rows, budgets, mappings, selections] = await Promise.all([
    loadRows(),
    loadBudgets(),
    loadMappings(),
    loadSelections(),
  ]);
  return { rows, budgets, mappings, selections };
}
