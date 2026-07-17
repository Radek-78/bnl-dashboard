/**
 * Veřejné API volané z frontendu přes google.script.run.
 * Každý endpoint je obalený guard_() — vrací { ok, data } / { ok, error }.
 */

function isTempClosedNow_(store) {
  var raw = store.temp_closed_ranges;
  if (!raw) return false;
  var ranges;
  try { ranges = JSON.parse(String(raw)); } catch(e) { return false; }
  if (!Array.isArray(ranges) || ranges.length === 0) return false;
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return ranges.some(function(r) { return r.from <= today && today <= r.to; });
}

function apiGetCurrentUser() {
  return guard_(ROLES.USER, (user) => {
    let allowed = String(user.allowed_apps || '').trim();
    if (!allowed || allowed.toLowerCase() === 'default') {
      const permSet = getRolePermissions_(user.role);
      allowed = permSet.allowed_apps || '';
    }
    return Object.assign({}, user, {
      allowed_apps: allowed,
      stores_write: isAllowed_(user, 'stores_write'),
    });
  });
}

/**
 * Hromadný bootstrap: vrátí data všech sekcí jedním voláním.
 * Frontend si je uloží do paměti a přepínání záložek je pak okamžité.
 */
function apiGetBootstrap() {
  return guard_(ROLES.USER, (user) => {
    dbEnsureApps_();
    const isAdmin = (ROLE_LEVEL[user.role] || 0) >= ROLE_LEVEL[ROLES.ADMIN];
    const data = {
      apps: dbGetAll_(SHEETS.APPS)
        .filter((a) => a.active === true && (isAdmin || isAppAllowedForUser_(user, a)))
        .sort((a, b) => (Number(a.order) || 0) - (Number(b.order) || 0)),
    };
    if (isAdmin) {
      const s = settingsAll_();
      data.users = dbGetAll_(SHEETS.USERS);
      data.stores = dbGetAll_(SHEETS.STORES);
      data.logistics = dbGetAll_(SHEETS.LOGISTICS);
      data.audit = dbReadTail_(SHEETS.AUDIT, 100).reverse();
      data.settings = s;
      data.syncSettings = {
        syncFolderUrl: s.syncFolderUrl || '',
        syncStoresSheet: s.syncStoresSheet || 'VTBZL_export',
        syncTempClosedPrefix: s.syncTempClosedPrefix || 'Dočasné zavření',
        autoSyncEnabled: s.autoSyncEnabled === true || s.autoSyncEnabled === 'true',
        autoSyncHour: s.autoSyncHour !== undefined && s.autoSyncHour !== '' ? Number(s.autoSyncHour) : 3,
        autoSyncLastCheckAt: s.autoSyncLastCheckAt || null,
        autoSyncLastCheckResult: s.autoSyncLastCheckResult || '',
        lastSyncAt: s.lastSyncAt || null,
        lastSyncResult: s.lastSyncResult ? JSON.parse(s.lastSyncResult) : null,
        syncHistory: s.syncHistory ? JSON.parse(s.syncHistory) : [],
      };
      dbEnsureRolePermissions_();
      data.rolePermissions = dbGetAll_(SHEETS.ROLE_PERMISSIONS);
    }
    return data;
  });
}

/**
 * Ověří přístupnost Drive složky s 5min cache — DriveApp.getFolderById
 * je pomalé (externí volání) a výsledek se mění jen výjimečně.
 */
function driveFolderAccessible_(folderUrl) {
  const folderId = extractFolderIdFromUrl_(folderUrl);
  if (!folderId) return false;
  let cache = null;
  try {
    cache = CacheService.getScriptCache();
    const hit = cache.get('drivecheck:' + folderId);
    if (hit) return hit === 'ok';
  } catch (e) { /* cache je jen optimalizace */ }
  let accessible = true;
  try {
    DriveApp.getFolderById(folderId);
  } catch (e) {
    accessible = false;
  }
  try {
    if (cache) cache.put('drivecheck:' + folderId, accessible ? 'ok' : 'fail', 300);
  } catch (e) { /* cache je jen optimalizace */ }
  return accessible;
}

function apiGetHome() {
  return guard_(ROLES.USER, (user) => {
    const isAdmin = (ROLE_LEVEL[user.role] || 0) >= ROLE_LEVEL[ROLES.ADMIN];
    const s = isAdmin ? settingsAll_() : {};
    const warnings = [];
    if (isAdmin) {
      if (!s.syncFolderUrl) {
        warnings.push({ key: 'no_sync_folder', message: 'Není nastavena složka pro synchronizaci filiálek.', action: 'Nastavit', section: 'sync' });
      } else if (!driveFolderAccessible_(s.syncFolderUrl)) {
        warnings.push({ key: 'sync_folder_inaccessible', message: 'Složka pro synchronizaci filiálek není přístupná.', action: 'Zkontrolovat', section: 'sync' });
      }
      if (!s.lastSyncAt) {
        warnings.push({ key: 'never_synced', message: 'Synchronizace filiálek nebyla ještě nikdy spuštěna.', action: 'Synchronizovat', section: 'sync' });
      }
    }
    const stores = dbGetAll_(SHEETS.STORES);
    const logistics = dbGetAll_(SHEETS.LOGISTICS);
    const logisticsActive = logistics.filter((lc) => lc.active === true).length;
    const lastSyncAt = isAdmin ? s.lastSyncAt : settingsAll_().lastSyncAt;
    return {
      warnings: warnings,
      lastSyncAt: isAdmin ? (lastSyncAt || null) : null,
      storesActive: stores.filter((st) => st.active === true).length,
      storesTempClosed: stores.filter((st) => st.active === true && isTempClosedNow_(st)).length,
      logisticsActive: logisticsActive,
      lastChange: isAdmin ? (dbReadTail_(SHEETS.AUDIT, 1)[0] || null) : null,
      // Dokud neproběhla první synchronizace filiálek a neexistuje aktivní LC,
      // frontend touto hodnotou zamyká administrační sekce menu.
      onboardingComplete: !!lastSyncAt && logisticsActive > 0,
    };
  });
}

/* ── Uživatelé ──────────────────────────────────────────────────── */

function apiListUsers() {
  return guard_(ROLES.ADMIN, () => dbGetAll_(SHEETS.USERS));
}

function apiSaveUser(payload) {
  return guard_(ROLES.ADMIN, (actor) => {
    const email = String((payload && payload.email) || '').trim().toLowerCase();
    if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) throw new Error('Zadejte platný e-mail.');
    const firstName = String((payload && payload.firstName) || '').trim();
    const lastName = String((payload && payload.lastName) || '').trim();

    const role = String((payload && payload.role) || ROLES.USER).toUpperCase();
    if (!ROLES[role]) throw new Error('Neplatná role.');

    const actorIsSuper = actor.role === ROLES.SUPERADMIN;
    if (role === ROLES.SUPERADMIN && !actorIsSuper) {
      throw new Error('Roli superadmin může přidělit pouze superadmin.');
    }

    const permission = String((payload && payload.permission) || 'READER').toUpperCase();
    if (permission !== 'READER' && permission !== 'EDITOR') throw new Error('Neplatné oprávnění (musí být Čtenář nebo Editor).');

    const location = String((payload && payload.location) || 'HQ').trim().toUpperCase();
    if (location !== 'HQ' && location !== 'CENTRÁLA' && location !== 'CENTRAL') {
      const logistics = dbGetAll_(SHEETS.LOGISTICS);
      const exists = logistics.some(lc => String(lc.abbreviation).toUpperCase() === location && lc.active === true);
      if (!exists) throw new Error('Neplatné logistické centrum: ' + location);
    }

    const users = dbGetAll_(SHEETS.USERS);
    const existing = payload.id ? users.find((user) => user.id === payload.id) : null;
    if (payload.id && !existing) throw new Error('Uživatel nenalezen.');
    if (existing && existing.role === ROLES.SUPERADMIN && !actorIsSuper) {
      throw new Error('Superadmina může upravit pouze superadmin.');
    }

    const duplicate = users.find(
      (user) => String(user.email).toLowerCase() === email && (!existing || user.id !== existing.id)
    );
    if (duplicate) throw new Error('Uživatel s tímto e-mailem už existuje.');

    const active = !payload || payload.active !== false;

    // Poslední aktivní superadmin nesmí přijít o roli ani být zablokován.
    if (existing && existing.role === ROLES.SUPERADMIN && (role !== ROLES.SUPERADMIN || !active)) {
      const superadmins = users.filter(
        (user) => user.role === ROLES.SUPERADMIN && user.active === true
      );
      if (superadmins.length <= 1) {
        throw new Error('Aplikace musí mít alespoň jednoho aktivního superadmina.');
      }
    }

    const allowed_apps = String((payload && payload.allowed_apps) || 'default').trim();

    const data = {
      email: email,
      firstName: firstName,
      lastName: lastName,
      role: role,
      active: active,
      permission: permission,
      location: location,
      allowed_apps: allowed_apps,
    };

    const saved = existing
      ? dbUpdate_(SHEETS.USERS, existing.id, data)
      : dbInsert_(SHEETS.USERS, data);
    audit_(existing ? 'user_update' : 'user_create', email + ' (' + role + ', ' + permission + ', ' + location + (active ? '' : ', blokován') + ')');
    return saved;
  });
}

function apiDeleteUser(id) {
  return guard_(ROLES.ADMIN, (actor) => {
    const users = dbGetAll_(SHEETS.USERS);
    const user = users.find((u) => u.id === id);
    if (!user) throw new Error('Uživatel nenalezen.');

    if (String(user.email).toLowerCase() === String(actor.email).toLowerCase()) {
      throw new Error('Nemůžete smazat sami sebe.');
    }

    if (user.role === ROLES.SUPERADMIN) {
      if (actor.role !== ROLES.SUPERADMIN) {
        throw new Error('Superadmina může smazat pouze superadmin.');
      }
      const activeSuperadmins = users.filter((u) => u.role === ROLES.SUPERADMIN && u.active === true);
      if (user.active === true && activeSuperadmins.length <= 1) {
        throw new Error('Aplikace musí mít alespoň jednoho aktivního superadmina.');
      }
    }

    dbDelete_(SHEETS.USERS, id);
    audit_('user_delete', user.email);
    return null;
  });
}

/* ── Nastavení ──────────────────────────────────────────────────── */

function apiGetSettings() {
  return guard_(ROLES.ADMIN, () => settingsAll_());
}

function apiSaveSettings(payload) {
  return guard_(ROLES.ADMIN, () => {
    const appName = String((payload && payload.appName) || '').trim();
    if (!appName) throw new Error('Vyplňte název aplikace.');
    settingsSet_('appName', appName);
    settingsSet_('appSubtitle', String((payload && payload.appSubtitle) || '').trim());
    audit_('settings_update', appName);
    return settingsAll_();
  });
}

function apiClearCache() {
  return guard_(ROLES.ADMIN, () => {
    clearAllDbCache_();
    audit_('cache_clear', 'Manuální vymazání cache');
    return null;
  });
}

/**
 * Kompletní reset aplikace — odpojí od databáze a vymaže server cache.
 * Appka při dalším načtení znovu spustí úvodní průvodce. Samotný databázový
 * spreadsheet v Drive zůstává nedotčený (lze ho v Disku smazat ručně).
 */
function apiResetApplication() {
  return guard_('SUPERADMIN', () => {
    audit_('app_reset', 'Kompletní reset aplikace — inicializace zrušena, spustí se úvodní průvodce.');
    PropertiesService.getScriptProperties().deleteAllProperties();
    clearAllDbCache_();
    dbHandle_ = null;
    return null;
  });
}

/* ── Stores ─────────────────────────────────────────────────────── */

function apiListStores() {
  return guard_(ROLES.USER, (user) => {
    const allStores = dbGetAll_(SHEETS.STORES);
    const userLoc = String(user.location || 'HQ').toUpperCase();
    
    if (userLoc === 'HQ' || userLoc === 'CENTRÁLA' || userLoc === 'CENTRAL' || user.role === 'SUPERADMIN') {
      return allStores;
    }
    return allStores.filter(store => String(store.lc_code).toUpperCase() === userLoc);
  });
}

function apiSaveStore(payload) {
  return guard_(ROLES.USER, (actor) => {
    if (!isAllowed_(actor, 'stores_write')) {
      throw new Error('Nemáte oprávnění k zápisu/úpravám dat.');
    }

    const lc_code = String((payload && payload.lc_code) || '').trim();

    if (payload && payload.id) {
      const existing = dbGetById_(SHEETS.STORES, payload.id);
      if (existing && !isAllowed_(actor, 'stores_write', existing.lc_code)) {
        throw new Error('Nemáte oprávnění upravovat prodejnu v lokaci ' + existing.lc_code);
      }
    }

    if (!isAllowed_(actor, 'stores_write', lc_code)) {
      throw new Error('Nemáte oprávnění přiřadit prodejnu pod lokaci ' + lc_code);
    }

    const code = String((payload && payload.code) || '').trim();
    if (!/^\d{1,3}$/.test(code)) throw new Error('Číslo prodejny musí být 1–3 cifry.');
    const paddedCode = code.padStart(3, '0');

    const name = String((payload && payload.name) || '').trim();
    if (!name) throw new Error('Název je povinný.');

    const HOUR_FIELDS = [
      'mon_open','mon_close','tue_open','tue_close','wed_open','wed_close',
      'thu_open','thu_close','fri_open','fri_close','sat_open','sat_close',
      'sun_open','sun_close',
    ];
    const data = {
      code: paddedCode,
      name,
      lc_code: lc_code,
      phone: String((payload && payload.phone) || '').trim(),
      area_manager: String((payload && payload.area_manager) || '').trim(),
      regional_manager: String((payload && payload.regional_manager) || '').trim(),
      rm_phone: String((payload && payload.rm_phone) || '').trim(),
      temporarily_closed: payload && payload.temporarily_closed === true,
      active: !payload || payload.active !== false,
      metropolitni: !!(payload && payload.metropolitni),
    };
    // Ruční deaktivace přes editaci se chová stejně jako tlačítko vypnutí —
    // sync ji nesmí automaticky vrátit zpět (viz manually_inactive v syncStores_).
    data.manually_inactive = !data.active;
    HOUR_FIELDS.forEach((f) => { data[f] = String((payload && payload[f]) || '').trim(); });
    if (payload && Array.isArray(payload.temp_closed_ranges)) {
      data.temp_closed_ranges = JSON.stringify(payload.temp_closed_ranges);
    }

    const stores = dbGetAll_(SHEETS.STORES);
    const existing = payload && payload.id ? stores.find((s) => s.id === payload.id) : null;
    if (payload && payload.id && !existing) throw new Error('Filiálka nenalezena.');

    const duplicate = stores.find(
      (s) => s.code === paddedCode && (!existing || s.id !== existing.id)
    );
    if (duplicate) throw new Error('Filiálka s tímto číslem už existuje.');

    const saved = existing
      ? dbUpdate_(SHEETS.STORES, existing.id, data)
      : dbInsert_(SHEETS.STORES, data);
    audit_('store_' + (existing ? 'update' : 'create'), paddedCode + ' ' + name);
    return saved;
  });
}

function apiDeleteStore(id) {
  return guard_(ROLES.USER, (actor) => {
    const store = dbGetById_(SHEETS.STORES, id);
    if (!store) throw new Error('Filiálka nenalezena.');
    if (!isAllowed_(actor, 'stores_write', store.lc_code)) {
      throw new Error('Nemáte oprávnění mazat filiálku v lokaci ' + store.lc_code);
    }
    dbUpdate_(SHEETS.STORES, id, { active: false });
    audit_('store_delete', store.code + ' ' + store.name);
    return null;
  });
}

function apiSetStoreMetropolitni(payload) {
  return guard_(ROLES.USER, (actor) => {
    const id = payload && payload.id;
    const store = dbGetById_(SHEETS.STORES, id);
    if (!store) throw new Error('Filiálka nenalezena.');
    if (!isAllowed_(actor, 'stores_write', store.lc_code)) {
      throw new Error('Nemáte oprávnění upravovat filiálku v lokaci ' + store.lc_code);
    }
    const value = !!(payload && payload.metropolitni);
    dbUpdate_(SHEETS.STORES, id, { metropolitni: value });
    audit_('store_metropolitni', store.code + ' ' + store.name + ' → ' + (value ? 'metropolitní' : 'běžná'));
    return null;
  });
}

function apiToggleStoreActive(id) {
  return guard_(ROLES.USER, (actor) => {
    const store = dbGetById_(SHEETS.STORES, id);
    if (!store) throw new Error('Filiálka nenalezena.');
    if (!isAllowed_(actor, 'stores_write', store.lc_code)) {
      throw new Error('Nemáte oprávnění upravovat filiálku v lokaci ' + store.lc_code);
    }
    const newActive = store.active !== true;
    dbUpdate_(SHEETS.STORES, id, { active: newActive, manually_inactive: !newActive });
    audit_('store_toggle', store.code + ' ' + store.name + ' → ' + (newActive ? 'aktivní' : 'neaktivní (ručně)'));
    return null;
  });
}

function apiSearchWorkspaceUsers(query) {
  return guard_(ROLES.ADMIN, () => {
    if (!query || String(query).trim().length < 2) return [];
    var q = String(query).trim();
    try {
      var url = 'https://people.googleapis.com/v1/people:searchDirectoryPeople'
        + '?query=' + encodeURIComponent(q)
        + '&readMask=names%2CemailAddresses'
        + '&sources=DIRECTORY_SOURCE_TYPE_DOMAIN_PROFILE'
        + '&pageSize=20';
      var resp = UrlFetchApp.fetch(url, {
        headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
        muteHttpExceptions: true,
      });
      if (resp.getResponseCode() !== 200) {
        throw new Error(JSON.parse(resp.getContentText()).error.message);
      }
      var result = JSON.parse(resp.getContentText());
      return (result.people || []).map(function(p) {
        var name  = (p.names && p.names[0]) || {};
        var email = (p.emailAddresses && p.emailAddresses[0]) || {};
        return {
          email:     email.value     || '',
          firstName: name.givenName  || '',
          lastName:  name.familyName || '',
        };
      });
    } catch (e) {
      throw new Error('Nepodařilo se načíst uživatele z Google Workspace: ' + e.message);
    }
  });
}

function apiSaveStoreTempRanges(payload) {
  return guard_(ROLES.USER, (actor) => {
    if (!isAllowed_(actor, 'stores_write')) {
      throw new Error('Nemáte oprávnění k úpravám filiálek.');
    }
    const id = payload && payload.id;
    if (!id) throw new Error('Chybí ID filiálky.');
    const store = dbGetById_(SHEETS.STORES, id);
    if (!store) throw new Error('Filiálka nenalezena.');
    if (!isAllowed_(actor, 'stores_write', store.lc_code)) {
      throw new Error('Nemáte oprávnění upravovat filiálku v lokaci ' + store.lc_code);
    }
    const ranges = payload.temp_closed_ranges;
    if (!Array.isArray(ranges)) throw new Error('Neplatný formát dat.');
    dbUpdate_(SHEETS.STORES, id, { temp_closed_ranges: JSON.stringify(ranges) });
    audit_('store_temp_ranges', store.code + ' ' + store.name + ': uloženo ' + ranges.length + ' rozsahů dočasného uzavření');
    return null;
  });
}

/* ── Logistics ──────────────────────────────────────────────────── */

function apiListLogistics() {
  return guard_(ROLES.USER, (user) => {
    const allLogistics = dbGetAll_(SHEETS.LOGISTICS);
    const userLoc = String(user.location || 'HQ').toUpperCase();
    if (userLoc === 'HQ' || userLoc === 'CENTRÁLA' || userLoc === 'CENTRAL' || user.role === 'SUPERADMIN') {
      return allLogistics;
    }
    return allLogistics.filter((lc) => String(lc.abbreviation).toUpperCase() === userLoc);
  });
}

function apiSaveLogistic(payload) {
  return guard_(ROLES.USER, (user) => {
    if (!isAllowed_(user, 'logistics_write')) {
      throw new Error('Nemáte oprávnění k zápisu/úpravám dat.');
    }

    const abbreviation = String((payload && payload.abbreviation) || '').trim().toUpperCase();
    if (!/^[A-Z]{2,4}$/.test(abbreviation)) throw new Error('Zkratka musí mít 2–4 písmena.');

    if (payload && payload.id) {
      const existing = dbGetById_(SHEETS.LOGISTICS, payload.id);
      if (existing && !isAllowed_(user, 'logistics_write', existing.abbreviation)) {
        throw new Error('Nemáte oprávnění upravovat logistické centrum ' + existing.abbreviation);
      }
    }

    if (!isAllowed_(user, 'logistics_write', abbreviation)) {
      throw new Error('Nemáte oprávnění spravovat logistické centrum se zkratkou ' + abbreviation);
    }

    const code = String((payload && payload.code) || '').trim();
    if (!code) throw new Error('Číslo LC je povinné.');
    const name = String((payload && payload.name) || '').trim();
    if (!name) throw new Error('Název je povinný.');

    const data = {
      code,
      name,
      abbreviation,
      active: !payload || payload.active !== false,
    };

    const list = dbGetAll_(SHEETS.LOGISTICS);
    const existing = payload && payload.id ? list.find((l) => l.id === payload.id) : null;
    if (payload && payload.id && !existing) throw new Error('LC nenalezeno.');

    const dup = list.find(
      (l) => l.abbreviation === abbreviation && (!existing || l.id !== existing.id)
    );
    if (dup) throw new Error('LC s touto zkratkou už existuje.');

    const saved = existing
      ? dbUpdate_(SHEETS.LOGISTICS, existing.id, data)
      : dbInsert_(SHEETS.LOGISTICS, data);
    audit_('logistic_' + (existing ? 'update' : 'create'), abbreviation + ' ' + name);
    return saved;
  });
}

function apiDeleteLogistic(id) {
  return guard_(ROLES.USER, (user) => {
    const lc = dbGetById_(SHEETS.LOGISTICS, id);
    if (!lc) throw new Error('LC nenalezeno.');
    if (!isAllowed_(user, 'logistics_write', lc.abbreviation)) {
      throw new Error('Nemáte oprávnění mazat logistické centrum ' + lc.abbreviation);
    }
    dbUpdate_(SHEETS.LOGISTICS, id, { active: false });
    audit_('logistic_delete', lc.abbreviation + ' ' + lc.name);
    return null;
  });
}

/**
 * Nastaví dané LC jako výchozí (_settings.defaultLcCode) — nové záznamy
 * (např. formulář nového uživatele) se pak přednastaví na tuto lokaci.
 */
function apiSetDefaultLogistic(id) {
  return guard_(ROLES.ADMIN, () => {
    const lc = dbGetById_(SHEETS.LOGISTICS, id);
    if (!lc) throw new Error('LC nenalezeno.');
    if (lc.active !== true) throw new Error('Nelze nastavit neaktivní LC jako výchozí.');
    const abbreviation = String(lc.abbreviation).toUpperCase();
    settingsSet_('defaultLcCode', abbreviation);
    audit_('logistic_set_default', abbreviation + ' ' + lc.name);
    return { defaultLcCode: abbreviation };
  });
}

/**
 * Hromadné vytvoření logistických center — z modalu, kdy uživatel doplňuje
 * zkratky LC nalezené ve filiálkách, ale zatím bez vlastního záznamu.
 * Vrací počet vytvořených + pole chyb (jednotlivé řádky se navzájem neblokují).
 */
function apiBulkCreateLogistics(rows) {
  return guard_(ROLES.USER, (user) => {
    if (!Array.isArray(rows) || !rows.length) throw new Error('Nebyla vybrána žádná LC k vytvoření.');
    const existing = dbGetAll_(SHEETS.LOGISTICS);
    const created = [];
    const errors = [];
    rows.forEach((row) => {
      const abbreviation = String((row && row.abbreviation) || '').trim().toUpperCase();
      try {
        if (!/^[A-Z]{2,4}$/.test(abbreviation)) throw new Error('neplatná zkratka');
        if (!isAllowed_(user, 'logistics_write', abbreviation)) throw new Error('nemáte oprávnění');
        const code = String((row && row.code) || '').trim();
        if (!code) throw new Error('chybí číslo LC');
        const name = String((row && row.name) || '').trim();
        if (!name) throw new Error('chybí název');
        const isDup = existing.some((l) => l.abbreviation === abbreviation) || created.some((c) => c.abbreviation === abbreviation);
        if (isDup) throw new Error('LC už existuje');
        const active = !row || row.active !== false;
        created.push(dbInsert_(SHEETS.LOGISTICS, { code, name, abbreviation, active }));
      } catch (e) {
        errors.push(abbreviation + ': ' + e.message);
      }
    });
    if (created.length) audit_('logistic_bulk_create', created.map((c) => c.abbreviation).join(', '));
    return { created: created.length, errors: errors };
  });
}

function apiUpdateLastVisit() {
  return guard_(ROLES.USER, (actor) => {
    const users = dbGetAll_(SHEETS.USERS);
    const user = users.find((u) => String(u.email).toLowerCase() === actor.email.toLowerCase());
    if (!user) return { previous: null };
    const previous = user.last_visit_at || null;
    dbUpdate_(SHEETS.USERS, user.id, { last_visit_at: nowIso_() });
    return { previous: previous };
  });
}

/* ── Sync settings ──────────────────────────────────────────────── */

function apiGetSyncSettings() {
  return guard_(ROLES.ADMIN, () => {
    const s = settingsAll_();
    return {
      syncFolderUrl: s.syncFolderUrl || '',
      syncStoresSheet: s.syncStoresSheet || 'VTBZL_export',
      syncTempClosedPrefix: s.syncTempClosedPrefix || 'Dočasné zavření',
      autoSyncEnabled: s.autoSyncEnabled === true || s.autoSyncEnabled === 'true',
      autoSyncHour: s.autoSyncHour !== undefined && s.autoSyncHour !== '' ? Number(s.autoSyncHour) : 3,
      autoSyncLastCheckAt: s.autoSyncLastCheckAt || null,
      autoSyncLastCheckResult: s.autoSyncLastCheckResult || '',
      lastSyncAt: s.lastSyncAt || null,
      lastSyncResult: s.lastSyncResult ? JSON.parse(s.lastSyncResult) : null,
      syncHistory: s.syncHistory ? JSON.parse(s.syncHistory) : [],
    };
  });
}

/** Smaže případné existující triggery automatické synchronizace a založí nový, je-li zapnutá. */
function applyAutoSyncTrigger_(enabled, hour) {
  ScriptApp.getProjectTriggers().forEach((t) => {
    if (t.getHandlerFunction() === 'autoSyncCheck_') ScriptApp.deleteTrigger(t);
  });
  if (enabled) {
    ScriptApp.newTrigger('autoSyncCheck_').timeBased().everyDays(1).atHour(hour).create();
  }
}

function apiSaveSyncSettings(payload) {
  return guard_(ROLES.ADMIN, () => {
    settingsSet_('syncFolderUrl', String((payload && payload.syncFolderUrl) || '').trim());
    settingsSet_('syncStoresSheet', String((payload && payload.syncStoresSheet) || 'VTBZL_export').trim());
    settingsSet_('syncTempClosedPrefix', String((payload && payload.syncTempClosedPrefix) || 'Dočasné zavření').trim());

    const autoSyncEnabled = !!(payload && payload.autoSyncEnabled);
    const autoSyncHour = Math.min(23, Math.max(0, parseInt(payload && payload.autoSyncHour, 10) || 0));
    applyAutoSyncTrigger_(autoSyncEnabled, autoSyncHour);

    // Ověření, že trigger po vytvoření skutečně existuje — jinak by selhání
    // (např. chybějící autorizace scope script.scriptapp) zůstalo neviditelné.
    const triggerActive = ScriptApp.getProjectTriggers().some((t) => t.getHandlerFunction() === 'autoSyncCheck_');
    if (autoSyncEnabled && !triggerActive) {
      throw new Error('Trigger automatické synchronizace se nepodařilo vytvořit. Spusťte v editoru Apps Script funkci TOOLS_zkontrolujAutoSync.');
    }

    settingsSet_('autoSyncEnabled', autoSyncEnabled);
    settingsSet_('autoSyncHour', autoSyncHour);

    audit_('sync_settings_update', 'Aktualizace konfigurace synchronizace.'
      + (autoSyncEnabled ? ' Auto. sync zapnuta, kontrola cca v ' + autoSyncHour + ':00, trigger aktivní: ' + triggerActive + '.' : ' Auto. sync vypnuta.'));
    return { autoSyncTriggerActive: triggerActive };
  });
}

/* ── Apps ───────────────────────────────────────────────────────── */

/**
 * Doplní list apps včetně nově přidaných sloupců (levná kontrola posledního záhlaví).
 *
 * isAllowed_ volá tuto funkci (přes getRolePermissions_) při každé kontrole oprávnění,
 * takže jeden request může jinak spustit kontrolu schématu 2-3×. dbSchemaEnsuredThisRun_
 * to omezí na jednou za běh skriptu, CacheService pak i napříč requesty.
 */
let dbSchemaEnsuredThisRun_ = false;
const SCHEMA_CHECK_CACHE_KEY_ = 'schema:checked:2'; // změna klíče vynutí novou kontrolu po přidání sloupce
const SCHEMA_CHECK_TTL_ = 1800; // sekund

function dbEnsureApps_() {
  if (dbSchemaEnsuredThisRun_) return;

  try {
    if (CacheService.getScriptCache().get(SCHEMA_CHECK_CACHE_KEY_)) {
      dbSchemaEnsuredThisRun_ = true;
      return;
    }
  } catch (e) { /* cache je jen optimalizace */ }

  const ss = dbSpreadsheet_();
  const sheetsToCheck = [SHEETS.APPS, SHEETS.USERS, SHEETS.ROLE_PERMISSIONS, SHEETS.STORES];
  let needsMigration = false;

  for (const name of sheetsToCheck) {
    const sheet = ss.getSheetByName(name);
    const headers = DB_SCHEMA[name];
    if (!sheet || sheet.getLastColumn() < headers.length || sheet.getRange(1, headers.length).getValue() !== headers[headers.length - 1]) {
      needsMigration = true;
      break;
    }
  }

  if (needsMigration) {
    dbEnsureSchema_(ss);
    sheetsToCheck.forEach(name => dbCacheInvalidate_(name));
  }

  dbSchemaEnsuredThisRun_ = true;
  try { CacheService.getScriptCache().put(SCHEMA_CHECK_CACHE_KEY_, '1', SCHEMA_CHECK_TTL_); } catch (e) { /* cache je jen optimalizace */ }
}

/** Převede text na slug do URL: malá písmena bez diakritiky, pomlčky. */
function slugify_(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .normalize('NFD')
    .replace(/[̀-ͯ]/g, '')
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '');
}

function apiListApps() {
  return guard_(ROLES.USER, (user) => {
    dbEnsureApps_();
    const isAdmin = (ROLE_LEVEL[user.role] || 0) >= ROLE_LEVEL[ROLES.ADMIN];
    return dbGetAll_(SHEETS.APPS)
      .filter((a) => a.active === true && (isAdmin || isAppAllowedForUser_(user, a)))
      .sort((a, b) => (Number(a.order) || 0) - (Number(b.order) || 0));
  });
}

function apiSaveApp(payload) {
  return guard_(ROLES.ADMIN, () => {
    const name = String((payload && payload.name) || '').trim();
    if (!name) throw new Error('Název je povinný.');
    const STATUSES = ['available', 'coming', 'inactive'];
    const COLORS = ['', 'dark', 'light', 'yellow', 'red', 'green', 'teal', 'orange', 'purple', 'slate', 'pink', 'indigo'];
    const data = {
      name,
      description: String((payload && payload.description) || '').trim(),
      icon: String((payload && payload.icon) || 'info').trim(),
      color: COLORS.includes(payload && payload.color) ? (payload.color || '') : '',
      status: STATUSES.includes(payload && payload.status) ? payload.status : 'coming',
      order: parseInt(payload && payload.order) || 0,
      active: payload && payload.active !== false,
    };
    const apps = dbGetAll_(SHEETS.APPS);
    const existing = payload && payload.id ? apps.find((a) => a.id === payload.id) : null;
    if (payload && payload.id && !existing) throw new Error('Aplikace nenalezena.');

    const slug = slugify_((payload && payload.slug) || name);
    if (!slug) throw new Error('Odkaz (slug) nelze vytvořit — použijte písmena nebo číslice.');
    const slugTaken = apps.find((a) => a.slug === slug && (!existing || a.id !== existing.id));
    if (slugTaken) throw new Error('Aplikace s tímto odkazem už existuje.');
    data.slug = slug;

    const saved = existing
      ? dbUpdate_(SHEETS.APPS, existing.id, data)
      : dbInsert_(SHEETS.APPS, data);

    // Posun ostatních aplikací pokud bylo zadáno konkrétní pořadí — jedním dávkovým zápisem
    // místo N jednotlivých dbUpdate_ (každý čte a zamyká celý list zvlášť).
    const targetOrder = data.order;
    if (targetOrder > 0) {
      const shiftIds = new Set(
        apps
          .filter((a) => a.active !== false && (!existing || a.id !== existing.id))
          .filter((a) => (Number(a.order) || 0) >= targetOrder)
          .map((a) => a.id)
      );
      if (shiftIds.size > 0) {
        const now = nowIso_();
        const freshApps = dbGetAll_(SHEETS.APPS); // po dbInsert_/dbUpdate_ výše, včetně právě uloženého záznamu
        const updated = freshApps.map((a) =>
          shiftIds.has(a.id) ? Object.assign({}, a, { order: (Number(a.order) || 0) + 1, updated_at: now }) : a
        );
        dbBatchReplace_(SHEETS.APPS, updated);
      }
    }

    audit_('app_' + (existing ? 'update' : 'create'), name);
    return saved;
  });
}

function apiDeleteApp(id) {
  return guard_(ROLES.ADMIN, () => {
    const app = dbGetById_(SHEETS.APPS, id);
    if (!app) throw new Error('Aplikace nenalezena.');
    dbUpdate_(SHEETS.APPS, id, { active: false });
    audit_('app_delete', app.name);
    return null;
  });
}

function apiGetAudit() {
  return guard_(ROLES.ADMIN, () => dbReadTail_(SHEETS.AUDIT, 100).reverse());
}

/* ── Role Permissions ───────────────────────────────────────────── */

function apiGetRolePermissions() {
  return guard_('ADMIN', () => {
    dbEnsureRolePermissions_();
    return dbGetAll_(SHEETS.ROLE_PERMISSIONS);
  });
}

function apiSaveRolePermissions(payload) {
  return guard_('SUPERADMIN', (actor) => {
    if (!payload || !Array.isArray(payload)) {
      throw new Error('Neplatná data pro uložení oprávnění.');
    }
    
    // Načteme stávající role z DB
    const list = dbGetAll_(SHEETS.ROLE_PERMISSIONS);
    
    payload.forEach((row) => {
      const roleName = String(row.role || '').trim().toUpperCase();
      if (!roleName) return;
      
      // Superadmina nelze přes API měnit, je chráněn vestavěnou logikou
      if (roleName === 'SUPERADMIN') return;
      
      const existing = list.find((r) => String(r.role).trim().toUpperCase() === roleName);
      if (!existing) return;
      
      const data = {
        stores_read: row.stores_read === true || row.stores_read === 'true',
        stores_write: row.stores_write === true || row.stores_write === 'true',
        logistics_read: row.logistics_read === true || row.logistics_read === 'true',
        logistics_write: row.logistics_write === true || row.logistics_write === 'true',
        users_manage: row.users_manage === true || row.users_manage === 'true',
        settings_manage: row.settings_manage === true || row.settings_manage === 'true',
        allowed_apps: String(row.allowed_apps || '').trim(),
        updated_at: new Date().toISOString(),
        updated_by: actor.email
      };
      
      dbUpdate_(SHEETS.ROLE_PERMISSIONS, existing.id, data);
    });
    
    dbCacheInvalidate_(SHEETS.ROLE_PERMISSIONS);
    audit_('roles_update', 'Změna matice oprávnění');
    return null;
  });
}
