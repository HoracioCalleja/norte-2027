// Guarda los votos de alojamiento de la página en la pestaña "Votos" de la planilla de respuestas.
// Implementar como aplicación web: Ejecutar como "Yo", acceso "Cualquier persona".
const SHEET_ID = '1MoXd355xeRHr0ScwCCYQ3hnW83dSZNsEvGahPnI4hUo';

function hojaVotos_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName('Votos');
  if (!sh) { sh = ss.insertSheet('Votos'); sh.appendRow(['ts', 'voter', 'stop', 'opt']); }
  return sh;
}

function doPost(e) {
  const d = JSON.parse(e.postData.contents);
  const ok = s => typeof s === 'string' && s.length <= 80 && /^[\w+\-]*$/.test(s);
  if (!ok(d.voter) || !ok(d.stop) || !ok(d.opt) || !d.voter || !d.stop) {
    return ContentService.createTextOutput('bad');
  }
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try { hojaVotos_().appendRow([new Date(), d.voter, d.stop, d.opt]); } finally { lock.releaseLock(); }
  return ContentService.createTextOutput('ok');
}

// Correr una vez a mano para crear la pestaña y aceptar permisos.
function prepararVotos() { hojaVotos_(); }
