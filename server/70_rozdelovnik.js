/**
 * Subaplikace "Rozdělovník 20 artiklů" (slug rozdelovnik-20-artiklu).
 *
 * Vlastní DB (apps.db_spreadsheet_id, viz 25_subapp_db.js) odděleně od hlavní
 * databáze appky. Záložky Artikly/_settings mají pevné schéma (RZ_SCHEMA) přes
 * createDbRepo_. Záložky Odprodej/Teo_stavy/Vyskladňovací_listy/Přidělení_po_artiklech
 * mají schéma dynamické — sloupce odpovídají přesně hlavičce naimportovaného
 * zdrojového souboru, proto se ukládají mimo RZ_SCHEMA jako syrová mřížka
 * (rzReadGrid_/rzWriteGrid_).
 */
const RZ_SLUG = 'rozdelovnik-20-artiklu';
const RZ_ARTIKLY_ROWS = 20;

const RZ_SCHEMA = {
  '_settings': ['id', 'key', 'value', 'updated_at', 'updated_by'],
  'artikly': ['id', 'poradi', 'cislo_artiklu', 'nazev', 'obsah', 'k_rozdeleni', 'pocet_dni', 'metropol', 'created_at', 'created_by', 'updated_at'],
  // Řádek s prázdnou prodejna = úroveň artiklu (min/max/rw); řádek s vyplněnou prodejna = úroveň konkrétní prodejny (uprava).
  // Pozor: nové sloupce se přidávají VŽDY na konec pole - řádky uložené podle
  // staršího schématu mají data pozičně svázaná se starým pořadím sloupců,
  // vložení doprostřed by je při čtení posunulo a rozbilo (viz updated_at/rw).
  'rozdeleni': ['id', 'cislo_artiklu', 'prodejna', 'min', 'max', 'uprava', 'created_at', 'created_by', 'updated_at', 'rw'],
};

const RZ_IMPORT_TABLES = ['odprodej', 'teo_stavy', 'vyskladnovaci_listy', 'prideleni_po_artiklech'];

const RZ_PATTERN_SETTING_KEY = {
  odprodej: 'patternOdprodej',
  teo_stavy: 'patternTeoStavy',
  vyskladnovaci_listy: 'patternVyskladnovaciListy',
  prideleni_po_artiklech: 'patternPrideleniPoArtiklech',
};

const RZ_IMPORT_LABELS = {
  odprodej: 'Odprodej',
  teo_stavy: 'Teoretické stavy',
  vyskladnovaci_listy: 'Vyskladňovací listy',
  prideleni_po_artiklech: 'Přidělení po artiklech',
};

// Informace o artiklech nemá (zatím) vlastní záložku/import - jen uložená
// složka a název souboru, dokud nebude domluveno, jak přesně appka data z něj
// využije. Má vlastní složku (jinou než ostatní 4 soubory) a hledá se v ní
// i podle čísla LC, ne jen podle výrazu v názvu.
// folderExport - cílová složka, kam se ukládá vyexportované rozdělení
// (.xlsx + Google Sheet). Samotný export se dodělá, až bude domluvený přesný
// formát listu; složka jde nastavit už teď.
const RZ_EXTRA_SETTING_KEYS = ['folderInformaceOArtiklech', 'patternInformaceOArtiklech', 'folderExport', 'prideliColorScale'];

function rzApp_() {
  const app = dbGetAll_(SHEETS.APPS).find((a) => a.slug === RZ_SLUG);
  if (!app) throw new Error('Subaplikace Rozdělovník 20 artiklů nenalezena (založte ji v sekci Aplikace).');
  if (!app.db_spreadsheet_id) throw new Error('Subaplikace ještě nemá vlastní databázi — zkuste ji znovu uložit v sekci Aplikace.');
  return app;
}

let rzRepoCache_ = null;
function rzRepo_() {
  if (rzRepoCache_) return rzRepoCache_;
  rzRepoCache_ = createDbRepo_(rzApp_().db_spreadsheet_id, RZ_SCHEMA);
  return rzRepoCache_;
}

/** Obal endpointů subaplikace: základní přístup (ROLES.USER) + kontrola allowed_apps pro tuto appku. */
function rzGuard_(fn) {
  return guard_(ROLES.USER, (user) => {
    const app = rzApp_();
    if (!isAppAllowedForUser_(user, app)) throw new Error('Nemáte přístup do této subaplikace.');
    return fn(user, app);
  });
}

function rzCanWrite_(user) {
  return user.role === 'SUPERADMIN' || user.permission === 'EDITOR';
}

/* ── Nastavení ────────────────────────────────────────────────── */

function rzSettingsAll_() {
  const repo = rzRepo_();
  repo.ensureSchema();
  const settings = {};
  repo.getAll('_settings').forEach((row) => { settings[row.key] = row.value; });
  return settings;
}

function rzSettingsSet_(key, value) {
  const repo = rzRepo_();
  const rows = repo.getAll('_settings');
  const existing = rows.find((r) => r.key === key);
  if (existing) {
    repo.update('_settings', existing.id, { value: value, updated_by: currentEmail_() });
  } else {
    repo.insert('_settings', { key: key, value: value, updated_by: currentEmail_() });
  }
}

function apiRzGetSettings() {
  return rzGuard_(() => rzSettingsAll_());
}

function apiRzSaveSettings(payload) {
  return rzGuard_((user) => {
    if ((ROLE_LEVEL[user.role] || 0) < ROLE_LEVEL[ROLES.ADMIN]) throw new Error('Nemáte oprávnění měnit nastavení.');
    const keys = ['syncFolderUrl'].concat(Object.values(RZ_PATTERN_SETTING_KEY), RZ_EXTRA_SETTING_KEYS);
    keys.forEach((key) => rzSettingsSet_(key, String((payload && payload[key]) || '').trim()));
    audit_('rz_settings_update', 'Aktualizace nastavení Rozdělovníku 20 artiklů.');
    return rzSettingsAll_();
  });
}

/** Vrátí seznam souborů (.xlsx/.csv/Google Sheets) v zadané složce. */
function rzListFolderFiles_(folderId) {
  const folder = DriveApp.getFolderById(folderId);
  const files = [];
  [MimeType.MICROSOFT_EXCEL, MimeType.GOOGLE_SHEETS, MimeType.CSV].forEach((mimeType) => {
    const it = folder.getFilesByType(mimeType);
    while (it.hasNext()) {
      const f = it.next();
      files.push({ name: f.getName(), updatedAt: f.getLastUpdated().toISOString() });
    }
  });
  files.sort((a, b) => b.updatedAt.localeCompare(a.updatedAt));
  return { folderName: folder.getName(), files: files };
}

/** Počet naimportovaných datových řádků (bez hlavičky) pro daný list — 0, pokud je prázdný. */
function rzImportedRowCount_(fileKey) {
  const lastRow = rzImportSheet_(fileKey).getLastRow();
  return lastRow > 1 ? lastRow - 1 : 0;
}

/**
 * Přehled pro záložku Artikly: co je ve složce se zdrojovými soubory, který
 * ze 4 očekávaných souborů se podle nastaveného výrazu právě najde, a jestli
 * ho už máme naimportovaný (dva nezávislé stavy — soubor ve složce může být
 * novější než poslední import, nebo naopak zmizet a import přitom zůstává).
 * Nahrazuje dřívější živý náhled v Nastavení.
 */
function apiRzGetImportOverview() {
  return rzGuard_(() => {
    const settings = rzSettingsAll_();
    const folderId = rzExtractFolderId_(settings.syncFolderUrl);
    if (!folderId) {
      return {
        folderName: '',
        items: RZ_IMPORT_TABLES.map((key) => ({
          key: key, label: RZ_IMPORT_LABELS[key], found: false, fileName: '', updatedAt: '',
          importedRowCount: rzImportedRowCount_(key),
        })),
      };
    }

    let listing;
    try {
      listing = rzListFolderFiles_(folderId);
    } catch (e) {
      throw new Error('Složku se nepodařilo otevřít: ' + e.message);
    }

    const items = RZ_IMPORT_TABLES.map((key) => {
      const pattern = settings[RZ_PATTERN_SETTING_KEY[key]] || '';
      const needle = pattern.toLowerCase();
      const match = needle ? listing.files.find((f) => f.name.toLowerCase().indexOf(needle) !== -1) : null;
      return {
        key: key,
        label: RZ_IMPORT_LABELS[key],
        found: !!match,
        fileName: match ? match.name : '',
        updatedAt: match ? match.updatedAt : '',
        importedRowCount: rzImportedRowCount_(key),
      };
    });

    return { folderName: listing.folderName, items: items };
  });
}

/** Vymaže obsah všech 4 listů s naimportovanými daty (hlavička i řádky) — nesahá na zdrojové soubory ani Artikly. */
function apiRzClearImports() {
  return rzGuard_((user) => {
    if (!rzCanWrite_(user)) throw new Error('Nemáte oprávnění k mazání.');
    RZ_IMPORT_TABLES.forEach((key) => rzWriteGrid_(key, [], []));
    audit_('rz_import_clear', 'Vymazán obsah naimportovaných dat (' + RZ_IMPORT_TABLES.length + ' listů).');
    return { ok: true };
  });
}

/** Naimportuje najednou všechny nalezené soubory (jedno tlačítko na záložce Artikly). */
function apiRzImportAll() {
  return rzGuard_((user) => {
    if (!rzCanWrite_(user)) throw new Error('Nemáte oprávnění k importu.');
    const settings = rzSettingsAll_();
    const folderId = rzExtractFolderId_(settings.syncFolderUrl);
    if (!folderId) throw new Error('Není nastavena složka se zdrojovými soubory (Nastavení).');

    const results = RZ_IMPORT_TABLES.map((key) => {
      const pattern = settings[RZ_PATTERN_SETTING_KEY[key]] || '';
      if (!pattern) return { key: key, label: RZ_IMPORT_LABELS[key], ok: false, message: 'Není nastaven výraz v názvu.' };
      try {
        const file = rzFindFileInFolderByName_(folderId, pattern);
        if (!file) return { key: key, label: RZ_IMPORT_LABELS[key], ok: false, message: 'Soubor nenalezen.' };
        const { headers, rows } = rzReadSourceFile_(file);
        rzWriteGrid_(key, headers, rows);
        return { key: key, label: RZ_IMPORT_LABELS[key], ok: true, fileName: file.getName(), rowCount: rows.length };
      } catch (e) {
        return { key: key, label: RZ_IMPORT_LABELS[key], ok: false, message: e.message };
      }
    });

    const okCount = results.filter((r) => r.ok).length;
    audit_('rz_import_all', okCount + '/' + RZ_IMPORT_TABLES.length + ' souborů naimportováno');
    return results;
  });
}

/** Číslo a zkratka výchozího LC appky (Informace o artiklech se podle čísla LC dohledává). */
function rzDefaultLcInfo_() {
  const abbreviation = String(settingsAll_().defaultLcCode || '').trim().toUpperCase();
  if (!abbreviation) return { abbreviation: '', code: '' };
  const lc = dbGetAll_(SHEETS.LOGISTICS).find((l) => String(l.abbreviation).toUpperCase() === abbreviation);
  return { abbreviation: abbreviation, code: lc ? String(lc.code) : '' };
}

/** Podporuje URL i přímo vložené ID složky. */
function rzExtractFolderId_(input) {
  const raw = String(input || '').trim();
  if (!raw) return null;
  const match = raw.match(/\/folders\/([a-zA-Z0-9_-]+)/);
  if (match) return match[1];
  if (/^[a-zA-Z0-9_-]{10,}$/.test(raw)) return raw;
  return null;
}

/* ── Artikly ──────────────────────────────────────────────────── */

function apiRzListArtikly() {
  return rzGuard_(() => {
    const repo = rzRepo_();
    repo.ensureSchema();
    return repo.getAll('artikly').sort((a, b) => (Number(a.poradi) || 0) - (Number(b.poradi) || 0));
  });
}

/** Uloží najednou všech 20 řádků (přepíše existující podle pořadí, chybějící založí). */
/**
 * Ukládá všech 20 řádků najednou - jeden zámek, jedno čtení, jeden (až dva)
 * zápis. Původně se pro každý řádek volalo repo.update/insert zvlášť (každé
 * se svým zámkem i vlastním čtením celého listu) - při psaní do polí (volá se
 * na blur) to dělalo až 20x tolik zámků/API volání, než bylo nutné, a citelně
 * to appku zpomalovalo.
 */
function apiRzSaveArtikly(rows) {
  return rzGuard_((user) => {
    if (!rzCanWrite_(user)) throw new Error('Nemáte oprávnění k zápisu.');
    if (!Array.isArray(rows)) throw new Error('Neplatná data.');
    const repo = rzRepo_();
    repo.ensureSchema();
    const headers = RZ_SCHEMA.artikly;
    const poradiCol = headers.indexOf('poradi');

    const saved = withLock_(() => {
      const sheet = repo.spreadsheet().getSheetByName('artikly');
      const lastRow = sheet.getLastRow();
      const existingValues = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, headers.length).getValues() : [];
      const now = nowIso_();
      const email = currentEmail_();
      const toAppend = [];
      const results = [];

      rows.slice(0, RZ_ARTIKLY_ROWS).forEach((row, i) => {
        const poradi = i + 1;
        const data = {
          poradi: poradi,
          cislo_artiklu: String((row && row.cislo_artiklu) || '').trim(),
          nazev: String((row && row.nazev) || '').trim(),
          obsah: String((row && row.obsah) || '').trim(),
          k_rozdeleni: (row && row.k_rozdeleni !== '' && row.k_rozdeleni != null) ? Number(row.k_rozdeleni) || 0 : '',
          pocet_dni: (row && row.pocet_dni !== '' && row.pocet_dni != null) ? Number(row.pocet_dni) || 0 : '',
          metropol: !!(row && row.metropol),
        };
        const rowIdx = existingValues.findIndex((r) => Number(r[poradiCol]) === poradi);
        if (rowIdx !== -1) {
          const merged = {};
          headers.forEach((h, ci) => { merged[h] = existingValues[rowIdx][ci]; });
          Object.assign(merged, data, { updated_at: now });
          existingValues[rowIdx] = headers.map((h) => (merged[h] !== undefined ? merged[h] : ''));
          results.push(merged);
        } else {
          const record = Object.assign({ id: uuid_(), created_at: now, created_by: email, updated_at: now }, data);
          toAppend.push(headers.map((h) => (record[h] !== undefined ? record[h] : '')));
          results.push(record);
        }
      });

      if (existingValues.length) {
        sheet.getRange(2, 1, existingValues.length, headers.length).setValues(existingValues);
      }
      if (toAppend.length) {
        sheet.getRange(lastRow + 1, 1, toAppend.length, headers.length).setValues(toAppend);
      }
      return results;
    });

    repo.invalidateCache('artikly');
    // Bez audit_ - volá se automaticky při opuštění každého políčka, takže by
    // to byl zápis do hlavní DB (druhý spreadsheet + vlastní zámek) při každém
    // stisku Tab, a v auditu jen šum. Audit zůstává u zásadních akcí (import,
    // reset, mazání dat).
    return saved.sort((a, b) => (Number(a.poradi) || 0) - (Number(b.poradi) || 0));
  });
}

/* ── Rozdělení ────────────────────────────────────────────────── */

function apiRzGetRozdeleni() {
  return rzGuard_(() => {
    const repo = rzRepo_();
    repo.ensureSchema();
    return repo.getAll('rozdeleni');
  });
}

/** Volá se při resetu tabulky Artiklů - bez toho by min/max/Úprava zůstaly viset na starých číslech artiklů a při opětovném zadání stejného čísla by se nečekaně vynořily. */
function apiRzResetRozdeleni() {
  return rzGuard_((user) => {
    if (!rzCanWrite_(user)) throw new Error('Nemáte oprávnění k mazání.');
    rzRepo_().clearTable('rozdeleni');
    audit_('rz_rozdeleni_reset', 'Vymazána data záložky Rozdělení (reset tabulky Artiklů).');
    return { ok: true };
  });
}

/** Porovná číslo artiklu i mezi Number/String - Sheets si číslo artiklu (vypadá jako číslo) tiše převádí na skutečné číslo, takže striktní === mezi "6717079" a 6717079 by nikdy neplatilo. */
function rzArtiklMatches_(cellVal, cislo) {
  const a = String(cellVal == null ? '' : cellVal).trim();
  const b = String(cislo == null ? '' : cislo).trim();
  if (a === b) return true;
  const na = Number(a), nb = Number(b);
  return !isNaN(na) && !isNaN(nb) && na === nb;
}

/**
 * Najde existující řádek (artikl + prodejna, prodejna '' = úroveň artiklu) a
 * přepíše ho, jinak založí nový - jeden zámek/čtení/zápis (dřív getAll +
 * repo.update/insert dělaly dohromady dvě čtení a zvlášť svůj zámek, zbytečně
 * pomalé pro tak časté volání jako je uložení jednoho pole na blur).
 */
function rzUpsertRozdeleni_(cisloArtiklu, prodejna, patch) {
  const repo = rzRepo_();
  repo.ensureSchema();
  const headers = RZ_SCHEMA.rozdeleni;
  const cisloCol = headers.indexOf('cislo_artiklu');
  const prodejnaCol = headers.indexOf('prodejna');
  const prodejnaNorm = String(prodejna || '');

  const result = withLock_(() => {
    const sheet = repo.spreadsheet().getSheetByName('rozdeleni');
    const lastRow = sheet.getLastRow();
    const existingValues = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, headers.length).getValues() : [];
    const rowIdx = existingValues.findIndex((r) =>
      rzArtiklMatches_(r[cisloCol], cisloArtiklu) && String(r[prodejnaCol] || '') === prodejnaNorm
    );
    const now = nowIso_();

    if (rowIdx !== -1) {
      const merged = {};
      headers.forEach((h, ci) => { merged[h] = existingValues[rowIdx][ci]; });
      Object.assign(merged, patch, { updated_at: now });
      sheet.getRange(rowIdx + 2, 1, 1, headers.length).setValues([headers.map((h) => (merged[h] !== undefined ? merged[h] : ''))]);
      return merged;
    }

    const record = Object.assign(
      { id: uuid_(), cislo_artiklu: cisloArtiklu, prodejna: prodejna || '', created_at: now, created_by: currentEmail_(), updated_at: now },
      patch
    );
    sheet.getRange(lastRow + 1, 1, 1, headers.length).setValues([headers.map((h) => (record[h] !== undefined ? record[h] : ''))]);
    return record;
  });

  repo.invalidateCache('rozdeleni');
  return result;
}

function apiRzSaveRozdeleniMinMax(payload) {
  return rzGuard_((user) => {
    if (!rzCanWrite_(user)) throw new Error('Nemáte oprávnění k zápisu.');
    const cislo = String((payload && payload.cislo_artiklu) || '').trim();
    if (!cislo) throw new Error('Chybí číslo artiklu.');
    const min = (payload && payload.min !== '' && payload.min != null) ? Number(payload.min) || 0 : '';
    const max = (payload && payload.max !== '' && payload.max != null) ? Number(payload.max) || 0 : '';
    return rzUpsertRozdeleni_(cislo, '', { min: min, max: max });
  });
}

/** RW (denní násobitel) - vodorovné doplňování (viz _computePrideleni/_solveRw ve frontendu): příděl prodejny = RW × MEDIAN − Stav, ohraničeno min/max. */
function apiRzSaveRozdeleniRw(payload) {
  return rzGuard_((user) => {
    if (!rzCanWrite_(user)) throw new Error('Nemáte oprávnění k zápisu.');
    const cislo = String((payload && payload.cislo_artiklu) || '').trim();
    if (!cislo) throw new Error('Chybí číslo artiklu.');
    const rw = (payload && payload.rw !== '' && payload.rw != null) ? Number(payload.rw) || 0 : '';
    return rzUpsertRozdeleni_(cislo, '', { rw: rw });
  });
}

/**
 * Uloží RW pro víc artiklů v jednom zámku/čtení/zápisu - používá se při
 * prvním zobrazení záložky Rozdělení, kde se RW dopočítá a uloží pro každý
 * vyplněný artikl zvlášť; bez dávky by to bylo tolik samostatných
 * round-tripů (a zámků), kolik je artiklů. items: [{cislo_artiklu, rw}].
 */
function apiRzSaveRozdeleniRwBatch(items) {
  return rzGuard_((user) => {
    if (!rzCanWrite_(user)) throw new Error('Nemáte oprávnění k zápisu.');
    if (!Array.isArray(items) || !items.length) return { ok: true };
    const repo = rzRepo_();
    repo.ensureSchema();
    const headers = RZ_SCHEMA.rozdeleni;
    const cisloCol = headers.indexOf('cislo_artiklu');
    const prodejnaCol = headers.indexOf('prodejna');

    withLock_(() => {
      const sheet = repo.spreadsheet().getSheetByName('rozdeleni');
      const lastRow = sheet.getLastRow();
      const existingValues = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, headers.length).getValues() : [];
      const now = nowIso_();
      const email = currentEmail_();
      const toAppend = [];

      items.forEach((item) => {
        const cislo = String((item && item.cislo_artiklu) || '').trim();
        if (!cislo) return;
        const rw = (item && item.rw !== '' && item.rw != null) ? Number(item.rw) || 0 : '';
        const rowIdx = existingValues.findIndex((r) => rzArtiklMatches_(r[cisloCol], cislo) && !String(r[prodejnaCol] || ''));
        if (rowIdx !== -1) {
          const merged = {};
          headers.forEach((h, ci) => { merged[h] = existingValues[rowIdx][ci]; });
          merged.rw = rw;
          merged.updated_at = now;
          existingValues[rowIdx] = headers.map((h) => (merged[h] !== undefined ? merged[h] : ''));
        } else {
          const record = { id: uuid_(), cislo_artiklu: cislo, prodejna: '', rw: rw, created_at: now, created_by: email, updated_at: now };
          toAppend.push(headers.map((h) => (record[h] !== undefined ? record[h] : '')));
        }
      });

      if (existingValues.length) {
        sheet.getRange(2, 1, existingValues.length, headers.length).setValues(existingValues);
      }
      if (toAppend.length) {
        sheet.getRange(lastRow + 1, 1, toAppend.length, headers.length).setValues(toAppend);
      }
    });

    repo.invalidateCache('rozdeleni');
    return { ok: true };
  });
}

function apiRzSaveRozdeleniUprava(payload) {
  return rzGuard_((user) => {
    if (!rzCanWrite_(user)) throw new Error('Nemáte oprávnění k zápisu.');
    const cislo = String((payload && payload.cislo_artiklu) || '').trim();
    const prodejna = String((payload && payload.prodejna) || '').trim();
    if (!cislo || !prodejna) throw new Error('Chybí číslo artiklu nebo prodejna.');
    const uprava = (payload && payload.uprava !== '' && payload.uprava != null) ? Number(payload.uprava) || 0 : '';
    return rzUpsertRozdeleni_(cislo, prodejna, { uprava: uprava });
  });
}

/**
 * Informace o artiklech má běžně desítky až stovky tisíc řádků — appka ho
 * proto nikdy celý nečte ani nikam neukládá, jen v něm cíleně vyhledá
 * konkrétní číslo artiklu (TextFinder nad sloupcem ARTIKL, ne getValues přes
 * celou tabulku). U .xlsx je nutná jedna konverze na Google Sheet; ta se
 * krátkodobě (~30 min) znovupoužívá (ID uložené v _settings), aby se při
 * postupném vyplňování 20 řádků nekonvertovalo opakovaně.
 */
const RZ_LOOKUP_TEMP_TTL_MS = 30 * 60 * 1000;

/** Otevře Informace o artiklech jako Sheet — bez kopírování, pokud je zdroj už nativní Google Sheet. */
function rzOpenInfoArtiklechSheet_(file) {
  if (file.getMimeType() === MimeType.GOOGLE_SHEETS) {
    return SpreadsheetApp.openById(file.getId()).getSheets()[0];
  }

  const settings = rzSettingsAll_();
  const cachedId = settings.lookupTempSheetId || '';
  const cachedAt = settings.lookupTempSheetAt ? new Date(settings.lookupTempSheetAt).getTime() : 0;
  const sameSource = settings.lookupTempSourceId === file.getId();
  const fresh = cachedId && sameSource && (Date.now() - cachedAt) < RZ_LOOKUP_TEMP_TTL_MS;
  if (fresh) {
    try {
      return SpreadsheetApp.openById(cachedId).getSheets()[0];
    } catch (e) { /* kopie zmizela/je neplatná — vytvoří se nová níže */ }
  }

  // Stará dočasná kopie (pokud existuje) se smaže, ať se nehromadí v Disku.
  if (cachedId) { try { Drive.Files.remove(cachedId); } catch (e) { /* není kritické */ } }

  const scriptFolder = scriptFolder_();
  const copyMeta = { name: '__rz_lookup_tmp__', mimeType: 'application/vnd.google-apps.spreadsheet' };
  if (scriptFolder) copyMeta.parents = [scriptFolder.getId()];
  const copy = Drive.Files.copy(copyMeta, file.getId(), { supportsAllDrives: true });

  rzSettingsSet_('lookupTempSheetId', copy.id);
  rzSettingsSet_('lookupTempSheetAt', nowIso_());
  rzSettingsSet_('lookupTempSourceId', file.getId());

  return SpreadsheetApp.openById(copy.id).getSheets()[0];
}

/** V listu najde řádek s daným číslem artiklu ve sloupci ARTIKL (TextFinder) a vrátí Název/Obsah. */
function rzFindArtiklInSheet_(sheet, cisloArtiklu) {
  const lastCol = sheet.getLastColumn();
  const lastRow = sheet.getLastRow();
  if (lastCol < 1 || lastRow < 2) return null;
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  const norm = (h) => String(h || '').trim().toUpperCase();
  const idxArtikl = headers.findIndex((h) => norm(h) === 'ARTIKL');
  const idxNazev = headers.findIndex((h) => norm(h) === 'NAZEV');
  const idxObsah = headers.findIndex((h) => norm(h) === 'OBSAH');
  if (idxArtikl === -1) throw new Error('Soubor Informace o artiklech nemá sloupec s hlavičkou ARTIKL.');

  const artiklCol = sheet.getRange(2, idxArtikl + 1, lastRow - 1, 1);
  const match = artiklCol.createTextFinder(String(cisloArtiklu)).matchEntireCell(true).findNext();
  if (!match) return null;

  const rowVals = sheet.getRange(match.getRow(), 1, 1, lastCol).getValues()[0];
  return {
    nazev: idxNazev !== -1 ? String(rowVals[idxNazev] || '').trim() : '',
    obsah: idxObsah !== -1 ? String(rowVals[idxObsah] || '').trim() : '',
  };
}

/** Vyhledá Název a Obsah pro jedno číslo artiklu — vrací null, pokud nenalezeno. Volá se na pozadí po opuštění pole Artikl. */
function apiRzLookupArtikl(cisloArtiklu) {
  return rzGuard_(() => {
    const cislo = String(cisloArtiklu || '').trim();
    if (!cislo) return null;

    const settings = rzSettingsAll_();
    const folderId = rzExtractFolderId_(settings.folderInformaceOArtiklech);
    if (!folderId) throw new Error('Není nastavena složka pro Informace o artiklech (Nastavení).');
    const pattern = settings.patternInformaceOArtiklech || '';
    if (!pattern) throw new Error('Není nastaven výraz pro soubor Informace o artiklech (Nastavení).');
    const lcCode = rzDefaultLcInfo_().code;
    if (!lcCode) throw new Error('V hlavním dashboardu není nastaveno výchozí logistické centrum.');

    const file = rzFindFileInFolderByNameAndLc_(folderId, pattern, lcCode);
    if (!file) throw new Error('Soubor Informace o artiklech (výraz „' + pattern + '" + LC ' + lcCode + ') nebyl ve složce nalezen.');

    if (file.getMimeType() === MimeType.CSV) {
      const text = rzReadCsvText_(file);
      const data = Utilities.parseCsv(text, rzDetectCsvDelimiter_(text));
      if (!data.length) return null;
      const headers = data[0];
      const norm = (h) => String(h || '').trim().toUpperCase();
      const idxArtikl = headers.findIndex((h) => norm(h) === 'ARTIKL');
      const idxNazev = headers.findIndex((h) => norm(h) === 'NAZEV');
      const idxObsah = headers.findIndex((h) => norm(h) === 'OBSAH');
      if (idxArtikl === -1) throw new Error('Soubor Informace o artiklech nemá sloupec s hlavičkou ARTIKL.');
      const row = data.slice(1).find((r) => String(r[idxArtikl] || '').trim() === cislo);
      if (!row) return null;
      return {
        nazev: idxNazev !== -1 ? String(row[idxNazev] || '').trim() : '',
        obsah: idxObsah !== -1 ? String(row[idxObsah] || '').trim() : '',
      };
    }

    return rzFindArtiklInSheet_(rzOpenInfoArtiklechSheet_(file), cislo);
  });
}

/**
 * Dohledá Název/Obsah pro víc čísel artiklů najednou - jedno vyhledání
 * souboru, jedno otevření (u .xlsx případně jedna konverze) pro celou dávku
 * místo samostatného volání (a kopie souboru) na každý artikl.
 * Vrací { '<cislo>': { nazev, obsah }, ... }; nenalezené artikly ve výsledku chybí.
 */
function apiRzLookupArtiklBatch(cisla) {
  return rzGuard_(() => {
    if (!Array.isArray(cisla) || !cisla.length) return {};
    const wanted = cisla.map((c) => String(c || '').trim()).filter((c) => c);
    if (!wanted.length) return {};

    const settings = rzSettingsAll_();
    const folderId = rzExtractFolderId_(settings.folderInformaceOArtiklech);
    if (!folderId) throw new Error('Není nastavena složka pro Informace o artiklech (Nastavení).');
    const pattern = settings.patternInformaceOArtiklech || '';
    if (!pattern) throw new Error('Není nastaven výraz pro soubor Informace o artiklech (Nastavení).');
    const lcCode = rzDefaultLcInfo_().code;
    if (!lcCode) throw new Error('V hlavním dashboardu není nastaveno výchozí logistické centrum.');

    const file = rzFindFileInFolderByNameAndLc_(folderId, pattern, lcCode);
    if (!file) throw new Error('Soubor Informace o artiklech (výraz „' + pattern + '" + LC ' + lcCode + ') nebyl ve složce nalezen.');

    const norm = (h) => String(h || '').trim().toUpperCase();
    const result = {};
    // Číselné porovnání i pro text - zdroj může mít artikl jako číslo (viz rzArtiklMatches_).
    const remaining = {};
    wanted.forEach((c) => { remaining[c] = true; });

    if (file.getMimeType() === MimeType.CSV) {
      const text = rzReadCsvText_(file);
      const data = Utilities.parseCsv(text, rzDetectCsvDelimiter_(text));
      if (!data.length) return {};
      const headers = data[0];
      const idxArtikl = headers.findIndex((h) => norm(h) === 'ARTIKL');
      const idxNazev = headers.findIndex((h) => norm(h) === 'NAZEV');
      const idxObsah = headers.findIndex((h) => norm(h) === 'OBSAH');
      if (idxArtikl === -1) throw new Error('Soubor Informace o artiklech nemá sloupec s hlavičkou ARTIKL.');
      // Jeden průchod souborem pro všechna hledaná čísla, ne jeden průchod na číslo.
      for (let i = 1; i < data.length; i++) {
        const cell = data[i][idxArtikl];
        const hit = wanted.find((c) => remaining[c] && rzArtiklMatches_(cell, c));
        if (!hit) continue;
        result[hit] = {
          nazev: idxNazev !== -1 ? String(data[i][idxNazev] || '').trim() : '',
          obsah: idxObsah !== -1 ? String(data[i][idxObsah] || '').trim() : '',
        };
        delete remaining[hit];
        if (!Object.keys(remaining).length) break;
      }
      return result;
    }

    const sheet = rzOpenInfoArtiklechSheet_(file);
    wanted.forEach((cislo) => {
      const found = rzFindArtiklInSheet_(sheet, cislo);
      if (found) result[cislo] = found;
    });
    return result;
  });
}

/**
 * Referenční přehled filiálek (Metropol se nastavuje v hlavním dashboardu — čteme napřímo hlavní DB).
 * Appka je postavená pro jedno konkrétní LC — vracíme jen filiálky spadající pod
 * výchozí LC nastavené v hlavním dashboardu (Log. centra), ne celou síť.
 */
function apiRzListStores() {
  return rzGuard_(() => {
    const defaultLcCode = String(settingsAll_().defaultLcCode || '').trim().toUpperCase();
    if (!defaultLcCode) throw new Error('V hlavním dashboardu není nastaveno výchozí logistické centrum (sekce Log. centra).');
    return dbGetAll_(SHEETS.STORES)
      .filter((s) => s.active === true && String(s.lc_code).trim().toUpperCase() === defaultLcCode)
      .map((s) => ({ code: s.code, name: s.name, metropolitni: !!s.metropolitni }))
      .sort((a, b) => String(a.code).localeCompare(String(b.code)));
  });
}

/* ── Import zdrojových souborů (dynamické schéma) ────────────────
 * Odprodej/Teo_stavy/Vyskladňovací_listy/Přidělení_po_artiklech nemají
 * pevně daný seznam sloupců v kódu — ukládají se jako syrová mřížka
 * (hlavička + řádky) přesně podle toho, co obsahuje zdrojový soubor. */

function rzImportSheet_(fileKey) {
  const ss = rzRepo_().spreadsheet();
  let sheet = ss.getSheetByName(fileKey);
  if (!sheet) sheet = ss.insertSheet(fileKey);
  return sheet;
}

/**
 * Buňku vždy vrátí jako čitelný text. Sheets občas přetypuje zapsaný text na
 * Date (i když jsme níže při zápisu vynutili textový formát — u dat zapsaných
 * před touto opravou to tak zůstalo) — takový Date objekt se přes
 * google.script.run přenáší nespolehlivě, což se projevovalo jako obecná
 * chyba serveru na klientovi.
 */
function rzCellToText_(val) {
  if (val instanceof Date) {
    // GAS vrací čistě časové buňky (bez data) jako Date s pevným datem
    // 30.12.1899 — stejná kuriozita, kterou už appka řeší u synchronizace
    // filiálek (formatCellValue_ v 60_sync.js). Bez téhle výjimky by se
    // takové buňky (např. "čas posledního prodeje") zobrazovaly jako
    // nesmyslné "30.12.1899 19:02:00" místo prostého času.
    if (val.getFullYear() === 1899 && val.getMonth() === 11 && val.getDate() === 30) {
      const h = val.getHours();
      const m = val.getMinutes();
      return h + ':' + (m < 10 ? '0' + m : m);
    }
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'dd.MM.yyyy HH:mm:ss').replace(/ 00:00:00$/, '');
  }
  return (val === undefined || val === null) ? '' : String(val);
}

function rzReadGrid_(fileKey) {
  const sheet = rzImportSheet_(fileKey);
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) return { headers: [], rows: [] };
  const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  return {
    headers: values[0].map(rzCellToText_),
    rows: values.slice(1).map((row) => row.map(rzCellToText_)),
  };
}

/**
 * Zapíše mřížku do listu jako čistý text — bez tohoto by Sheets textové
 * hodnoty vypadající jako datum nebo číslo (např. "06.07.2026" nebo kód
 * s vedoucí nulou) samo přetypoval, a import by tak nevratně změnil obsah
 * zdrojových dat.
 */
function rzWriteGrid_(fileKey, headers, rows) {
  const sheet = rzImportSheet_(fileKey);
  sheet.clearContents();
  if (!headers || !headers.length) return;
  const totalRows = 1 + (rows ? rows.length : 0);
  sheet.getRange(1, 1, totalRows, headers.length).setNumberFormat('@');
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (rows && rows.length) sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
}

/** Najde ve složce nejnovější soubor (.xlsx/.csv/Google Sheets), jehož název obsahuje daný výraz. */
function rzFindFileInFolderByName_(folderId, namePattern) {
  const folder = DriveApp.getFolderById(folderId);
  const needle = String(namePattern).toLowerCase();
  let newest = null;
  let newestDate = null;
  [MimeType.MICROSOFT_EXCEL, MimeType.GOOGLE_SHEETS, MimeType.CSV].forEach((mimeType) => {
    const files = folder.getFilesByType(mimeType);
    while (files.hasNext()) {
      const file = files.next();
      if (file.getName().toLowerCase().indexOf(needle) === -1) continue;
      const date = file.getLastUpdated();
      if (!newestDate || date > newestDate) { newest = file; newestDate = date; }
    }
  });
  return newest;
}

/**
 * Odhadne oddělovač CSV podle prvního řádku — české exporty (Excel v CZ
 * lokalizaci) běžně používají středník, protože čárka je desetinný oddělovač.
 * Utilities.parseCsv bez druhého parametru počítá vždy jen s čárkou, což by
 * takový soubor rozparsovalo jako jeden sloupec místo mnoha.
 */
function rzDetectCsvDelimiter_(text) {
  const firstLine = String(text).split(/\r?\n/, 1)[0] || '';
  const semicolons = (firstLine.match(/;/g) || []).length;
  const commas = (firstLine.match(/,/g) || []).length;
  return semicolons > commas ? ';' : ',';
}

/**
 * Přečte CSV soubor jako text a ošetří kódování — starší/české exporty z Excelu
 * často nejsou v UTF-8, ale ve Windows-1250 (čeština bez UTF-8 BOM). Chybné
 * dekódování se pozná podle náhradního znaku U+FFFD (ten se objeví jen tehdy,
 * když bajty nejsou platné UTF-8), a v tom případě se soubor přečte znovu
 * s kódováním Windows-1250.
 */
function rzReadCsvText_(file) {
  const blob = file.getBlob();
  const utf8 = blob.getDataAsString('UTF-8');
  if (utf8.indexOf('�') === -1) return utf8;
  return blob.getDataAsString('Windows-1250');
}

/** Jako rzFindFileInFolderByName_, ale název musí obsahovat i zadané číslo LC (Informace o artiklech). */
function rzFindFileInFolderByNameAndLc_(folderId, namePattern, lcCode) {
  const folder = DriveApp.getFolderById(folderId);
  const needle = String(namePattern).toLowerCase();
  const lcNeedle = String(lcCode || '').toLowerCase();
  let newest = null;
  let newestDate = null;
  [MimeType.MICROSOFT_EXCEL, MimeType.GOOGLE_SHEETS, MimeType.CSV].forEach((mimeType) => {
    const files = folder.getFilesByType(mimeType);
    while (files.hasNext()) {
      const file = files.next();
      const name = file.getName().toLowerCase();
      if (name.indexOf(needle) === -1) continue;
      if (lcNeedle && name.indexOf(lcNeedle) === -1) continue;
      const date = file.getLastUpdated();
      if (!newestDate || date > newestDate) { newest = file; newestDate = date; }
    }
  });
  return newest;
}

/**
 * Přečte zdrojový soubor (.xlsx/.csv/Google Sheets) a vrátí { headers, rows }.
 * U .xlsx/Sheets použije stejný trik jako hlavní synchronizace (60_sync.js) —
 * dočasná kopie jako Google Sheet, přečtení, smazání kopie.
 */
function rzReadSourceFile_(file) {
  if (file.getMimeType() === MimeType.CSV) {
    const text = rzReadCsvText_(file);
    const data = Utilities.parseCsv(text, rzDetectCsvDelimiter_(text));
    if (!data.length) return { headers: [], rows: [] };
    return { headers: data[0].map(String), rows: data.slice(1) };
  }

  let tempSheetId = null;
  try {
    const scriptFolder = scriptFolder_();
    const copyMeta = { name: '__rz_import_tmp__', mimeType: 'application/vnd.google-apps.spreadsheet' };
    if (scriptFolder) copyMeta.parents = [scriptFolder.getId()];
    // supportsAllDrives: soubor může ležet ve sdíleném disku (Shared Drive) - bez
    // tohoto parametru Drive.Files.copy hlásí "File not found", i když DriveApp
    // stejný soubor bez problémů najde a přečte.
    const copy = Drive.Files.copy(copyMeta, file.getId(), { supportsAllDrives: true });
    tempSheetId = copy.id;
    const ss = SpreadsheetApp.openById(tempSheetId);
    const sheet = ss.getSheets()[0];
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 1 || lastCol < 1) return { headers: [], rows: [] };
    const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    return { headers: values[0].map(String), rows: values.slice(1) };
  } catch (e) {
    throw new Error('Nepodařilo se načíst soubor "' + file.getName() + '": ' + e.message);
  } finally {
    if (tempSheetId) { try { Drive.Files.remove(tempSheetId); } catch (_) { /* dočasný soubor, chyba mazání není kritická */ } }
  }
}

function apiRzGetImportTable(fileKey) {
  return rzGuard_(() => {
    if (RZ_IMPORT_TABLES.indexOf(fileKey) === -1) throw new Error('Neplatný typ importu.');
    return rzReadGrid_(fileKey);
  });
}

function apiRzSaveImportTable(payload) {
  return rzGuard_((user) => {
    if (!rzCanWrite_(user)) throw new Error('Nemáte oprávnění k zápisu.');
    const fileKey = payload && payload.fileKey;
    if (RZ_IMPORT_TABLES.indexOf(fileKey) === -1) throw new Error('Neplatný typ importu.');
    rzWriteGrid_(fileKey, (payload && payload.headers) || [], (payload && payload.rows) || []);
    return { ok: true };
  });
}
