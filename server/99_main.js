/**
 * Vstupní bod webové aplikace.
 *
 * Routing podle stavu:
 *  - bez inicializace  → úvodní průvodce (wizard)
 *  - ?app=<id>         → samostatné okno subaplikace
 *  - uživatel v _users → aplikace
 *  - jinak             → obrazovka bez přístupu
 */
function doGet(e) {
  let page = 'wizard';
  let settings = {};
  let subApp = null;

  try {
    if (isSetupDone_()) {
      page = 'app';
      settings = settingsAll_();
      const appKey = e && e.parameter && e.parameter.app;
      if (appKey) {
        subApp = dbGetAll_(SHEETS.APPS).find(
          (a) => a.active === true && (a.slug === appKey || a.id === appKey)
        ) || null;
        if (subApp) page = 'subapp';
      }
    }
  } catch (err) {
    console.error(err);
  }

  // Identita uživatele se načítá klientsky přes apiGetCurrentUser (getActiveUser nefunguje v doGet)
  const data = {
    page: page,
    appName: settings.appName || CONFIG.defaultAppName,
    appSubtitle: settings.appSubtitle || CONFIG.defaultAppSubtitle,
    logoUrl: CONFIG.logoUrl,
    version: CONFIG.version,
    releaseDate: CONFIG.releaseDate,
    user: null,
    subApp: subApp,
    execUrl: ScriptApp.getService().getUrl(),
    setup: page === 'wizard' ? wizardInfo_() : null,
  };

  const template = HtmlService.createTemplateFromFile('index');
  template.app = data;
  template.bootstrapJson = JSON.stringify(data).replace(/</g, '\\u003c');
  return template.evaluate()
    .setTitle(page === 'subapp' ? subApp.name : data.appName)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(fileName) {
  return HtmlService.createHtmlOutputFromFile(fileName).getContent();
}
