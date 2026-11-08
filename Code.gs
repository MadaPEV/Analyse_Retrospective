/**
 * Atelier de résolution de problèmes — liaison avec Google Sheet
 *
 * DÉPLOIEMENT
 * 1. Ouvrir le Google Sheet > Extensions > Apps Script, coller ce code (remplacer Code.gs).
 * 2. Déployer > Nouveau déploiement > type « Application web »
 *      Exécuter en tant que : Moi
 *      Qui a accès : Tous (ou « Tous les utilisateurs de votre organisation »)
 * 3. Copier l'URL se terminant par /exec et la coller dans l'outil (⚙ Liaison Sheet).
 * 4. À chaque modification de ce code : Déployer > Gérer les déploiements > Modifier > Nouvelle version.
 */

const SHEET_ID = '1Kr7LsNXrBQMuYUElNAfOQ55OS4P67fg7ofREqVz9Ta8';
const TOKEN = '';              // facultatif : mettre la même valeur dans l'outil
const DATA_SHEET = '_DATA';    // sauvegarde complète (JSON découpé en blocs)
const CHUNK = 40000;           // limite d'une cellule : 50 000 caractères

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function ss_() {
  try {
    const a = SpreadsheetApp.getActiveSpreadsheet();
    if (a) return a;
  } catch (e) {}
  return SpreadsheetApp.openById(SHEET_ID);
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (TOKEN && p.token !== TOKEN) return out_({ ok: false, error: 'jeton invalide' });
  if (p.action === 'load') return out_({ ok: true, state: readState_() });
  return out_({ ok: true, message: 'API Atelier de résolution de problèmes' });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const body = JSON.parse(e.postData.contents);
    if (TOKEN && body.token !== TOKEN) return out_({ ok: false, error: 'jeton invalide' });
    if (body.action !== 'save') return out_({ ok: false, error: 'action inconnue' });
    writeState_(body.state);
    const tables = body.tables || {};
    Object.keys(tables).forEach(function (name) { putTable_(name, tables[name]); });
    return out_({ ok: true, saved: new Date().toISOString() });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

function readState_() {
  const sh = ss_().getSheetByName(DATA_SHEET);
  if (!sh || sh.getLastRow() < 1) return null;
  const vals = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
  const txt = vals.map(function (r) { return r[0]; }).join('');
  return txt ? JSON.parse(txt) : null;
}

function writeState_(state) {
  const ss = ss_();
  let sh = ss.getSheetByName(DATA_SHEET);
  if (!sh) sh = ss.insertSheet(DATA_SHEET);
  const txt = JSON.stringify(state);
  const rows = [];
  for (let i = 0; i < txt.length; i += CHUNK) rows.push([txt.substring(i, i + CHUNK)]);
  sh.clearContents();
  if (rows.length) sh.getRange(1, 1, rows.length, 1).setValues(rows);
  sh.getRange('B1').setValue('Sauvegarde technique de l\'outil — ne pas modifier. Mise à jour : ' + new Date().toLocaleString('fr-FR'));
}

function putTable_(name, rows) {
  if (!rows || !rows.length) return;
  const ss = ss_();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  const w = rows.reduce(function (m, r) { return Math.max(m, r.length); }, 1);
  const grid = rows.map(function (r) {
    const x = r.slice();
    while (x.length < w) x.push('');
    return x.map(function (c) { return c === null || c === undefined ? '' : c; });
  });
  sh.clear();
  sh.getRange(1, 1, grid.length, w).setValues(grid);
  sh.getRange(1, 1, 1, w).setFontWeight('bold').setBackground('#e2e8f0');
  sh.setFrozenRows(1);
  sh.getRange(1, 1, grid.length, w).setWrap(true).setVerticalAlignment('top');
  sh.autoResizeColumns(1, w);
}
