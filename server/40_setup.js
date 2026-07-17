/**
 * Inicializace aplikace — backend úvodního průvodce.
 *
 * První spuštění: doGet() zjistí, že chybí DB_SPREADSHEET_ID, a místo
 * aplikace vyrenderuje wizard. Ten smí dokončit pouze vlastník skriptu;
 * stane se superadminem a v jeho složce (tam, kde leží skript) vznikne
 * DB spreadsheet.
 */
function isSetupDone_() {
  const id = PropertiesService.getScriptProperties().getProperty(PROPS.DB_ID);
  if (!id) return false;
  try {
    SpreadsheetApp.openById(id);
    return true;
  } catch (e) {
    // Spreadsheet byl smazán — vynulujeme property a spustíme wizard znovu
    PropertiesService.getScriptProperties().deleteProperty(PROPS.DB_ID);
    return false;
  }
}

/** Složka v Drive, ve které leží tento skript, nebo null (root / bez přístupu). */
function scriptFolder_() {
  try {
    const parents = DriveApp.getFileById(ScriptApp.getScriptId()).getParents();
    return parents.hasNext() ? parents.next() : null;
  } catch (e) {
    console.error('Složku skriptu se nepodařilo zjistit: ' + e);
    return null;
  }
}

/** Data pro úvodní obrazovku wizardu (volá doGet, jen před inicializací). */
function wizardInfo_() {
  // getEffectiveUser() vrací vlastníka skriptu spolehlivě v USER_DEPLOYING módu
  const email = Session.getEffectiveUser().getEmail() || Session.getActiveUser().getEmail() || '';
  const folder = scriptFolder_();
  return {
    email: email,
    isOwner: !!email && email === Session.getEffectiveUser().getEmail(),
    folderName: folder ? folder.getName() : null,
    defaultAppName: CONFIG.defaultAppName,
    defaultAppSubtitle: CONFIG.defaultAppSubtitle,
  };
}

/** Ověří, zda má aktuální uživatel přístup ke složce zadané URL. Jen před inicializací. */
function wizardCheckFolderAccess(url) {
  try {
    if (isSetupDone_()) return ok_({ ok: false, message: 'Aplikace už je inicializována.' });
    const folderId = extractFolderIdFromUrl_(url || '');
    if (!folderId) return ok_({ ok: false, message: 'Nepodařilo se rozpoznat ID složky z URL.' });
    const folder = DriveApp.getFolderById(folderId);
    return ok_({ ok: true, message: 'Přístup potvrzen: „' + folder.getName() + '"' });
  } catch (e) {
    return ok_({ ok: false, message: 'Přístup odepřen nebo složka neexistuje.' });
  }
}

/** Vrátí email a složku pro wizard — volá se z klienta přes google.script.run. */
function wizardGetOwnerEmail() {
  try {
    const email = Session.getEffectiveUser().getEmail() || Session.getActiveUser().getEmail() || '';
    const folder = scriptFolder_();
    return ok_({ email: email, folderName: folder ? folder.getName() : null });
  } catch (e) {
    console.error('wizardGetOwnerEmail chyba: ' + e);
    return fail_(String(e));
  }
}

/**
 * Dokončení wizardu: vytvoří DB spreadsheet, založí schéma, zapíše
 * superadmina a nastavení. Chráněno zámkem proti dvojí inicializaci.
 */
function setupInitialize(payload) {
  try {
    if (isSetupDone_()) return fail_('Aplikace už je inicializována.');

    const email = currentEmail_();
    if (!email || email !== Session.getEffectiveUser().getEmail()) {
      return fail_('Inicializaci může provést pouze vlastník skriptu.');
    }

    const appName = String((payload && payload.appName) || '').trim();
    if (!appName) return fail_('Vyplňte název aplikace.');
    const appSubtitle = String((payload && payload.appSubtitle) || '').trim();
    const firstName = String((payload && payload.firstName) || '').trim();
    const lastName = String((payload && payload.lastName) || '').trim();
    const syncFolderUrl = String((payload && payload.syncFolderUrl) || '').trim();

    return withLock_(() => {
      if (isSetupDone_()) return fail_('Aplikace už je inicializována.');

      const ss = SpreadsheetApp.create(appName + ' – databáze');
      const defaultSheet = ss.getSheets()[0];
      const folder = scriptFolder_();
      if (folder) DriveApp.getFileById(ss.getId()).moveTo(folder);

      dbEnsureSchema_(ss);
      ss.deleteSheet(defaultSheet);

      const properties = {};
      properties[PROPS.DB_ID] = ss.getId();
      properties[PROPS.SETUP_AT] = nowIso_();
      PropertiesService.getScriptProperties().setProperties(properties);
      dbHandle_ = ss;

      dbInsert_(SHEETS.USERS, {
        email: email,
        firstName: firstName,
        lastName: lastName,
        role: ROLES.SUPERADMIN,
        active: true,
      });
      settingsSet_('appName', appName);
      settingsSet_('appSubtitle', appSubtitle);
      if (syncFolderUrl) settingsSet_('syncFolderUrl', syncFolderUrl);
      audit_('setup', 'Inicializace aplikace. DB: ' + ss.getId()
        + (folder ? ', složka: ' + folder.getName() : ', složka: kořen Disku'));

      return ok_({
        spreadsheetUrl: ss.getUrl(),
        appUrl: ScriptApp.getService().getUrl(),
        folderName: folder ? folder.getName() : null,
      });
    });
  } catch (e) {
    console.error(e);
    return fail_(e);
  }
}
