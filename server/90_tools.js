/**
 * Jednorázové nástroje pro vlastníka — spouštějí se ručně z editoru
 * Apps Script (Spustit ▸ vybraná funkce). Nejsou volány z aplikace,
 * ale jako funkce bez koncového "_" jsou technicky volatelné i přes
 * google.script.run — proto každá ověřuje, že ji spouští vlastník skriptu.
 */
function assertToolsOwner_() {
  const activeEmail = currentEmail_();
  if (!activeEmail || activeEmail !== Session.getEffectiveUser().getEmail()) {
    throw new Error('Tento nástroj smí spustit pouze vlastník skriptu.');
  }
}

/** Vypíše, kde v Drive skript leží a kam by se vytvořila databáze. */
function TOOLS_kdeJeSkript() {
  assertToolsOwner_();
  const folder = scriptFolder_();
  console.log(folder
    ? 'Skript leží ve složce: ' + folder.getName() + ' (' + folder.getUrl() + ')'
    : 'Skript leží v kořeni Můj disk (nebo složku nelze přečíst).');
  const dbId = PropertiesService.getScriptProperties().getProperty(PROPS.DB_ID);
  console.log(dbId ? 'Databáze: https://docs.google.com/spreadsheets/d/' + dbId : 'Databáze zatím nebyla vytvořena.');
}

/**
 * Přesune tento skript do zadané složky v Drive. Použij, když přesun
 * v Drive UI hlásí chybu. Před spuštěním vlož ID složky (z její URL).
 */
function TOOLS_presunSkriptDoSlozky() {
  assertToolsOwner_();
  const FOLDER_ID = 'SEM_VLOZ_ID_SLOZKY';
  if (FOLDER_ID === 'SEM_VLOZ_ID_SLOZKY') {
    throw new Error('Nejdřív do konstanty FOLDER_ID vlož ID cílové složky.');
  }
  const folder = DriveApp.getFolderById(FOLDER_ID);
  DriveApp.getFileById(ScriptApp.getScriptId()).moveTo(folder);
  console.log('Skript přesunut do složky: ' + folder.getName());
}

/**
 * POZOR: odpojí aplikaci od databáze (smaže Script Properties) a vymaže
 * server cache, takže při dalším otevření znovu naběhne úvodní průvodce
 * bez rizika, že se ještě chvíli zobrazují data ze staré databáze.
 * Samotný spreadsheet v Drive zůstává — případné smazání proveď ručně.
 */
function TOOLS_resetInicializace() {
  assertToolsOwner_();
  PropertiesService.getScriptProperties().deleteAllProperties();
  clearAllDbCache_();
  console.log('Inicializace zrušena a cache vymazána. Další otevření aplikace spustí průvodce.');
}

/** Smaže celou server-side cache (CacheService). Použij po přímých úpravách v DB sheetu. */
function TOOLS_clearCache() {
  assertToolsOwner_();
  clearAllDbCache_();
  console.log('Cache smazána.');
}

/** Rychlý test DB vrstvy: vlož, načti, uprav a smaž testovací záznam v _settings. */
function TOOLS_testDb() {
  assertToolsOwner_();
  const key = '_test_' + uuid_();
  settingsSet_(key, 'hodnota1');
  let value = settingsAll_()[key];
  if (value !== 'hodnota1') throw new Error('Zápis/čtení selhalo: ' + value);
  settingsSet_(key, 'hodnota2');
  value = settingsAll_()[key];
  if (value !== 'hodnota2') throw new Error('Aktualizace selhala: ' + value);
  const sheet = dbSheet_(SHEETS.SETTINGS);
  const keys = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  sheet.deleteRow(keys.findIndex((row) => row[0] === key) + 2);
  console.log('DB vrstva funguje (insert, read, update, delete).');
}

/**
 * Vynutí žádost o autorizaci OAuth scope script.scriptapp (přidán ve v3.1.8
 * pro automatickou synchronizaci). Nemění žádná data — jen čte seznam
 * triggerů, čímž GAS donutí zobrazit dialog s novým oprávněním. Spusť
 * jednou ručně z editoru; teprve pak lze v UI zapnout přepínač
 * "Automatická synchronizace" (bez toho hlásí chybějící autorizaci).
 */
function TOOLS_authorizeAutoSync() {
  assertToolsOwner_();
  const count = ScriptApp.getProjectTriggers().length;
  console.log('Autorizace scriptapp OK. Aktuální počet triggerů v projektu: ' + count);
}

/**
 * Zkontroluje stav triggeru automatické synchronizace a srovná ho podle
 * uloženého nastavení (zapnuto/vypnuto + hodina). Spusť z editoru, pokud
 * zapnutí v aplikaci trigger nevytvořilo — editor je pro zakládání
 * triggerů vždy spolehlivý kontext.
 */
function TOOLS_zkontrolujAutoSync() {
  assertToolsOwner_();
  const s = settingsAll_();
  const enabled = s.autoSyncEnabled === true || s.autoSyncEnabled === 'true';
  const hour = Math.min(23, Math.max(0, parseInt(s.autoSyncHour, 10) || 0));

  const before = ScriptApp.getProjectTriggers().filter((t) => t.getHandlerFunction() === 'autoSyncCheck_').length;
  console.log('Nastaveni: autoSyncEnabled=' + enabled + ', hodina=' + hour + ':00 | triggeru pred srovnanim: ' + before);

  applyAutoSyncTrigger_(enabled, hour);

  const after = ScriptApp.getProjectTriggers().filter((t) => t.getHandlerFunction() === 'autoSyncCheck_').length;
  console.log('Triggeru po srovnani: ' + after + (enabled
    ? (after > 0 ? ' — OK, denni kontrola pobezi cca v ' + hour + ':00.' : ' — CHYBA: trigger se nepodarilo vytvorit!')
    : ' — auto sync je vypnuta, zadny trigger nema existovat.'));
}
