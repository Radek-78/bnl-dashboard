/**
 * Repository vrstva nad Google Sheets.
 *
 * Tabulka = list, první řádek = hlavičky podle DB_SCHEMA. Veškerá čtení
 * a zápisy běží dávkově přes getValues/setValues, zápisy pod zámkem.
 *
 * Nové tabulky projektu se přidávají rozšířením DB_SCHEMA — funkce
 * dbEnsureSchema_() chybějící listy a sloupce doplní (funguje i jako
 * migrace při rozšíření šablony).
 */
const DB_SCHEMA = {
  '_users': ['id', 'email', 'firstName', 'lastName', 'role', 'active', 'created_at', 'created_by', 'updated_at', 'last_visit_at', 'permission', 'location', 'allowed_apps'],
  '_settings': ['key', 'value', 'updated_at', 'updated_by'],
  '_audit_log': ['timestamp', 'user', 'action', 'detail'],
  'stores': [
    'id', 'code', 'name', 'lc_code', 'phone', 'area_manager', 'regional_manager', 'rm_phone',
    'mon_open', 'mon_close', 'tue_open', 'tue_close', 'wed_open', 'wed_close',
    'thu_open', 'thu_close', 'fri_open', 'fri_close', 'sat_open', 'sat_close',
    'sun_open', 'sun_close',
    'temporarily_closed', 'active', 'manually_inactive', 'synced_at', 'created_at', 'created_by', 'updated_at', 'temp_closed_ranges',
    'metropolitni',
  ],
  'logistics': [
    'id', 'code', 'name', 'abbreviation',
    'active', 'created_at', 'created_by', 'updated_at',
  ],
  'apps': [
    'id', 'name', 'description', 'icon', 'color', 'status', 'order',
    'active', 'created_at', 'created_by', 'updated_at', 'slug',
  ],
  '_role_permissions': [
    'id', 'role', 'stores_read', 'stores_write', 'logistics_read', 'logistics_write', 'users_manage', 'settings_manage', 'allowed_apps', 'updated_at', 'updated_by'
  ],
};

let dbHandle_ = null;     // spreadsheet pro aktuální běh skriptu
let dbLockHeld_ = false;  // reentrantní zámek v rámci jednoho běhu

/* ── Cache vrstva ─────────────────────────────────────────────────
 * Výsledky dbGetAll_ se drží v CacheService (sdílené mezi běhy skriptu),
 * takže opakovaná čtení nemusí otevírat spreadsheet. Každý zápis do
 * tabulky svůj záznam v cache zneplatní. */
const DB_CACHE_TTL_ = 300; // sekund

function dbCacheKey_(table) {
  return 'tbl:' + table;
}

function dbCacheInvalidate_(table) {
  try {
    CacheService.getScriptCache().remove(dbCacheKey_(table));
  } catch (e) { /* cache je jen optimalizace */ }
}

/**
 * Smaže veškerou DB cache (data všech tabulek + kontrolu schématu). Použij
 * po přímé úpravě dat mimo appku i při kompletním resetu — jinak může po
 * přepnutí na novou databázi ještě chvíli "prosvítat" obsah té staré.
 */
function clearAllDbCache_() {
  try {
    CacheService.getScriptCache().removeAll(
      Object.values(SHEETS).map((t) => dbCacheKey_(t)).concat([SCHEMA_CHECK_CACHE_KEY_])
    );
  } catch (e) { /* cache je jen optimalizace */ }
}

function dbSpreadsheet_() {
  if (dbHandle_) return dbHandle_;
  const id = PropertiesService.getScriptProperties().getProperty(PROPS.DB_ID);
  if (!id) throw new Error('Databáze není inicializována. Spusťte úvodního průvodce.');
  dbHandle_ = SpreadsheetApp.openById(id);
  return dbHandle_;
}

function dbSheet_(table) {
  const sheet = dbSpreadsheet_().getSheetByName(table);
  if (!sheet) throw new Error('Tabulka "' + table + '" v databázi neexistuje.');
  return sheet;
}

/** Doplní chybějící listy a hlavičky podle DB_SCHEMA. Nic nemaže. */
function dbEnsureSchema_(ss) {
  Object.keys(DB_SCHEMA).forEach((name) => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);
    const headers = DB_SCHEMA[name];
    const current = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
    if (headers.some((header, i) => current[i] !== header)) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.setFrozenRows(1);
    }
  });
}

/** Spustí fn pod script zámkem; vnořená volání zámek nepřebírají znovu. */
function withLock_(fn) {
  if (dbLockHeld_) return fn();
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  dbLockHeld_ = true;
  try {
    return fn();
  } finally {
    dbLockHeld_ = false;
    lock.releaseLock();
  }
}

/** Vrátí všechny záznamy tabulky jako pole objektů podle hlaviček (s cache). */
function dbGetAll_(table) {
  let cache = null;
  try {
    cache = CacheService.getScriptCache();
    const hit = cache.get(dbCacheKey_(table));
    if (hit) return JSON.parse(hit);
  } catch (e) { /* cache je jen optimalizace */ }
  const rows = dbReadAll_(table);
  try {
    if (cache) cache.put(dbCacheKey_(table), JSON.stringify(rows), DB_CACHE_TTL_);
  } catch (e) { /* příliš velká data se prostě necachují */ }
  return rows;
}

function dbRowsToRecords_(headers, values) {
  return values.map((row) => {
    const record = {};
    headers.forEach((header, i) => { record[header] = dbDeserialize_(row[i]); });
    return record;
  });
}

/** Skutečné čtení tabulky ze spreadsheetu (bez cache). */
function dbReadAll_(table) {
  const sheet = dbSheet_(table);
  const headers = DB_SCHEMA[table];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  return dbRowsToRecords_(headers, values);
}

/**
 * Vrátí posledních n záznamů tabulky přímým čtením konce listu (bez cache).
 * Pro rostoucí logy (audit) — vyhne se čtení a cachování celé tabulky jen kvůli
 * pár posledním řádkům.
 */
function dbReadTail_(table, n) {
  const sheet = dbSheet_(table);
  const headers = DB_SCHEMA[table];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const startRow = Math.max(2, lastRow - n + 1);
  const values = sheet.getRange(startRow, 1, lastRow - startRow + 1, headers.length).getValues();
  return dbRowsToRecords_(headers, values);
}

function dbDeserialize_(val) {
  if (val instanceof Date) {
    if (val.getFullYear() === 1899 && val.getMonth() === 11 && val.getDate() === 30) {
      const h = val.getHours();
      const m = val.getMinutes();
      return h + ':' + (m < 10 ? '0' + m : m);
    }
    return val.toISOString();
  }
  return val;
}

function dbGetById_(table, id) {
  return dbGetAll_(table).find((record) => record.id === id) || null;
}

/** Surový zápis řádku podle hlaviček (bez generovaných polí). */
function dbAppendRow_(table, record) {
  const headers = DB_SCHEMA[table];
  withLock_(() => {
    dbSheet_(table).appendRow(headers.map((header) => (record[header] !== undefined ? record[header] : '')));
  });
  dbCacheInvalidate_(table);
}

/** Vloží záznam a doplní id, created_at, created_by, updated_at. */
function dbInsert_(table, data) {
  const record = Object.assign({}, data, {
    id: uuid_(),
    created_at: nowIso_(),
    created_by: currentEmail_(),
    updated_at: nowIso_(),
  });
  dbAppendRow_(table, record);
  return record;
}

/** Aktualizuje záznam podle id, vrací sloučený záznam. */
function dbUpdate_(table, id, patch) {
  const result = withLock_(() => {
    const sheet = dbSheet_(table);
    const headers = DB_SCHEMA[table];
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) throw new Error('Záznam nenalezen.');
    const values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
    const idCol = headers.indexOf('id');
    const rowIndex = values.findIndex((row) => row[idCol] === id);
    if (rowIndex === -1) throw new Error('Záznam nenalezen.');
    const merged = {};
    headers.forEach((header, i) => { merged[header] = values[rowIndex][i]; });
    Object.assign(merged, patch, { updated_at: nowIso_() });
    sheet.getRange(rowIndex + 2, 1, 1, headers.length).setValues([headers.map((header) => merged[header])]);
    return merged;
  });
  dbCacheInvalidate_(table);
  return result;
}

/**
 * Nahradí celý obsah tabulky novými záznamy v jedné dávce.
 * Používá se při synchronizaci — výrazně rychlejší než jednotlivé UPDATE/INSERT.
 */
function dbBatchReplace_(table, records) {
  withLock_(() => {
    const sheet = dbSheet_(table);
    const headers = DB_SCHEMA[table];
    const oldLastRow = sheet.getLastRow();
    if (oldLastRow > 1) {
      sheet.getRange(2, 1, oldLastRow - 1, headers.length).clearContent();
    }
    if (records.length > 0) {
      const values = records.map((record) =>
        headers.map((h) => (record[h] !== undefined && record[h] !== null) ? record[h] : '')
      );
      sheet.getRange(2, 1, values.length, headers.length).setValues(values);
    }
  });
  dbCacheInvalidate_(table);
}

function dbDelete_(table, id) {
  withLock_(() => {
    const sheet = dbSheet_(table);
    const headers = DB_SCHEMA[table];
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) throw new Error('Záznam nenalezen.');
    const values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
    const idCol = headers.indexOf('id');
    const rowIndex = values.findIndex((row) => row[idCol] === id);
    if (rowIndex === -1) throw new Error('Záznam nenalezen.');
    sheet.deleteRow(rowIndex + 2);
  });
  dbCacheInvalidate_(table);
}

/* ── Nastavení (klíč/hodnota v listu _settings) ─────────────────── */

function settingsAll_() {
  const settings = {};
  dbGetAll_(SHEETS.SETTINGS).forEach((row) => { settings[row.key] = row.value; });
  return settings;
}

function settingsSet_(key, value) {
  withLock_(() => {
    const sheet = dbSheet_(SHEETS.SETTINGS);
    const lastRow = sheet.getLastRow();
    const keys = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, 1).getValues() : [];
    const index = keys.findIndex((row) => row[0] === key);
    const record = [key, value, nowIso_(), currentEmail_()];
    if (index === -1) {
      sheet.appendRow(record);
    } else {
      sheet.getRange(index + 2, 1, 1, record.length).setValues([record]);
    }
  });
  dbCacheInvalidate_(SHEETS.SETTINGS);
}
