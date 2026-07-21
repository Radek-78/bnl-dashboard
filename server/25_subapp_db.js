/**
 * Repository vrstva pro subaplikace s vlastní DB (samostatný spreadsheet
 * mimo hlavní databázi — viz provisionSubAppDb_ a apiSaveApp v 50_api.js).
 *
 * Na rozdíl od dbGetAll_/dbInsert_/... v 20_db.js (napevno svázaných s jedinou
 * hlavní DB přes PROPS.DB_ID) je tahle vrstva parametrizovaná konkrétním
 * spreadsheetId a schématem listů — každá subaplikace si createDbRepo_()
 * zavolá se svým db_spreadsheet_id a vlastním schématem (stejný tvar jako
 * DB_SCHEMA). Cache klíče jsou navíc namespacované spreadsheetId, aby se
 * nepletly mezi subaplikacemi ani s hlavní DB.
 */
function createDbRepo_(spreadsheetId, schema) {
  let handle = null;
  const spreadsheet_ = () => {
    if (handle) return handle;
    handle = SpreadsheetApp.openById(spreadsheetId);
    return handle;
  };
  const sheet_ = (table) => {
    const sheet = spreadsheet_().getSheetByName(table);
    if (!sheet) throw new Error('Tabulka "' + table + '" v databázi subaplikace neexistuje.');
    return sheet;
  };
  const cacheKey_ = (table) => 'subtbl:' + spreadsheetId + ':' + table;
  const cacheInvalidate_ = (table) => {
    try { CacheService.getScriptCache().remove(cacheKey_(table)); } catch (e) { /* cache je jen optimalizace */ }
  };

  /**
   * Doplní chybějící listy a hlavičky podle schema. Nic nemaže.
   *
   * Volá se na začátku prakticky každého endpointu, ale schéma se v praxi
   * nemění - proto se úspěšné ověření zapamatuje v cache a další volání
   * (i v dalších exekucích) ho přeskočí. Bez toho to znamenalo čtení hlaviček
   * všech tabulek schématu při KAŽDÉM uložení jednoho políčka.
   */
  const ensureSchema = (force) => {
    // Otisk schématu v klíči - jakmile se schéma v kódu změní (nový sloupec),
    // cache se sama zneplatní a hlavičky se doplní, ne až po expiraci.
    const fingerprint = Object.keys(schema).map((n) => n + ':' + schema[n].join(',')).join('|');
    let hash = 0;
    for (let i = 0; i < fingerprint.length; i++) {
      hash = ((hash << 5) - hash + fingerprint.charCodeAt(i)) | 0;
    }
    const doneKey = 'subschema:' + spreadsheetId + ':' + hash;
    if (!force) {
      try {
        if (CacheService.getScriptCache().get(doneKey)) return;
      } catch (e) { /* cache je jen optimalizace */ }
    }
    const ss = spreadsheet_();
    Object.keys(schema).forEach((name) => {
      let sheet = ss.getSheetByName(name);
      if (!sheet) sheet = ss.insertSheet(name);
      const headers = schema[name];
      const current = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
      if (headers.some((header, i) => current[i] !== header)) {
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
        sheet.setFrozenRows(1);
      }
    });
    try {
      CacheService.getScriptCache().put(doneKey, '1', 21600);
    } catch (e) { /* cache je jen optimalizace */ }
  };

  const readAll = (table) => {
    const sheet = sheet_(table);
    const headers = schema[table];
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    const values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
    return dbRowsToRecords_(headers, values);
  };

  const getAll = (table) => {
    let cache = null;
    try {
      cache = CacheService.getScriptCache();
      const hit = cache.get(cacheKey_(table));
      if (hit) return JSON.parse(hit);
    } catch (e) { /* cache je jen optimalizace */ }
    const rows = readAll(table);
    try {
      if (cache) cache.put(cacheKey_(table), JSON.stringify(rows), DB_CACHE_TTL_);
    } catch (e) { /* příliš velká data se prostě necachují */ }
    return rows;
  };

  const getById = (table, id) => getAll(table).find((record) => record.id === id) || null;

  const insert = (table, data) => {
    const record = Object.assign({}, data, {
      id: uuid_(),
      created_at: nowIso_(),
      created_by: currentEmail_(),
      updated_at: nowIso_(),
    });
    const headers = schema[table];
    withLock_(() => {
      sheet_(table).appendRow(headers.map((header) => (record[header] !== undefined ? record[header] : '')));
    });
    cacheInvalidate_(table);
    return record;
  };

  const update = (table, id, patch) => {
    const result = withLock_(() => {
      const sheet = sheet_(table);
      const headers = schema[table];
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
    cacheInvalidate_(table);
    return result;
  };

  const remove = (table, id) => {
    withLock_(() => {
      const sheet = sheet_(table);
      const headers = schema[table];
      const lastRow = sheet.getLastRow();
      if (lastRow < 2) throw new Error('Záznam nenalezen.');
      const values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
      const idCol = headers.indexOf('id');
      const rowIndex = values.findIndex((row) => row[idCol] === id);
      if (rowIndex === -1) throw new Error('Záznam nenalezen.');
      sheet.deleteRow(rowIndex + 2);
    });
    cacheInvalidate_(table);
  };

  /** Smaže všechny řádky tabulky (hlavičku ponechá) - jedno rychlé mazání, ne cyklus přes delete() pro každý řádek zvlášť. */
  const clearTable = (table) => {
    withLock_(() => {
      const sheet = sheet_(table);
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);
    });
    cacheInvalidate_(table);
  };

  return {
    spreadsheet: spreadsheet_,
    ensureSchema: ensureSchema,
    getAll: getAll,
    getById: getById,
    insert: insert,
    update: update,
    delete: remove,
    clearTable: clearTable,
    // Pro hromadné operace, které si čtení/zápis dělají samy napřímo přes
    // sheet (jeden zámek na celou dávku místo jednoho na řádek) - repo.insert/
    // update by se pro to muselo volat opakovaně, což je právě to pomalé.
    invalidateCache: (table) => cacheInvalidate_(table),
  };
}

/**
 * Založí vlastní podsložku a spreadsheet DB pro nově vytvořenou subaplikaci
 * (ve stejné složce, kde leží hlavní DB — viz scriptFolder_ v 40_setup.js).
 * Vrací ID nového spreadsheetu (apps.db_spreadsheet_id) i ID podsložky —
 * tu využívají specifické subaplikace, které si v ní zakládají další vlastní
 * podsložky (viz rzProvisionFolders_ v 70_rozdelovnik.js).
 *
 * Nový spreadsheet má vždy jeden výchozí list ("List1"/"Sheet1" podle jazyka
 * Disku) — smazat ho hned nejde, spreadsheet musí mít vždy aspoň jeden list,
 * a schéma konkrétní subaplikace tady ještě neznáme (ensureSchema se volá
 * až později, líně). Místo mazání se proto rovnou přejmenuje na _settings,
 * což je list, který má (podle stejné konvence jako hlavní DB) každá
 * subaplikace — ensureSchema ho pak jen doplní o hlavičku jako kterýkoli
 * jiný list schématu, žádný prázdný "List1" navíc už nezůstane.
 */
function provisionSubAppDb_(appName) {
  const parentFolder = scriptFolder_();
  const subFolder = parentFolder ? parentFolder.createFolder(appName) : DriveApp.createFolder(appName);
  const ss = SpreadsheetApp.create(appName + ' – databáze');
  ss.getSheets()[0].setName('_settings');
  DriveApp.getFileById(ss.getId()).moveTo(subFolder);
  return { dbSpreadsheetId: ss.getId(), folderId: subFolder.getId() };
}
