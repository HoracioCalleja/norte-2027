/**
 * NORTE 2027 – guard de identidad.
 * Va en el Apps Script del formulario (reemplaza a votos.gs).
 *
 * - Las respuestas del form (con mails) quedan en una planilla PRIVADA.
 * - Este script copia a una planilla PÚBLICA solo apodos y respuestas, sin mails.
 * - Los votos de alojamiento se aceptan solo con sesión de Google verificada,
 *   y se guardan con el apodo de esa persona: un voto por persona y parada.
 *
 * Pasos: 1) completar CLIENT_ID  2) correr configurarGuard()  3) volver a implementar la app web.
 */
const CLIENT_ID = 'PEGAR_ACA_EL_ID_DE_CLIENTE.apps.googleusercontent.com';
const FORM_ID = '1bhGv7hy_EebO_W3f660CenENwRugl80K7OyWq2fF3cY';
const PRIVADA_ID = '1MoXd355xeRHr0ScwCCYQ3hnW83dSZNsEvGahPnI4hUo';   // planilla de respuestas del form

const props_ = () => PropertiesService.getScriptProperties();
const slug_ = n => 'n_' + String(n || '').trim().toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '')
  .replace(/[^a-z0-9]+/g, '_').slice(0, 60);

function publica_() {
  const id = props_().getProperty('PUBLICA_ID');
  if (!id) throw new Error('Falta correr configurarGuard()');
  return SpreadsheetApp.openById(id);
}

/** Correr UNA vez a mano. */
function configurarGuard() {
  if (!props_().getProperty('PUBLICA_ID')) {
    const ss = SpreadsheetApp.create('NORTE 2027 – Público (sin mails)');
    ss.getSheets()[0].setName('Respuestas');
    ss.insertSheet('Votos').appendRow(['ts', 'voter', 'stop', 'opt']);
    DriveApp.getFileById(ss.getId()).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    props_().setProperty('PUBLICA_ID', ss.getId());
  }
  // la planilla con mails deja de ser pública
  DriveApp.getFileById(PRIVADA_ID).setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.EDIT);
  // copiar a la pública cada vez que alguien responde o edita su respuesta
  const f = FormApp.openById(FORM_ID);
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'alEnviar').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('alEnviar').forForm(f).onFormSubmit().create();
  sincronizar();
  Logger.log('ID de la planilla pública (pasáselo a Claude): ' + props_().getProperty('PUBLICA_ID'));
}

function alEnviar() { Utilities.sleep(3000); sincronizar(); }

/** Copia las respuestas a la planilla pública sin la columna de mail. Una fila por persona (la última). */
function sincronizar() {
  const data = SpreadsheetApp.openById(PRIVADA_ID).getSheets()[0].getDataRange().getValues();
  if (!data.length) return;
  const H = data[0].map(String);
  const iMail = H.findIndex(h => /correo|email|e-mail/i.test(h));
  const ultima = {};
  data.slice(1).forEach((r, i) => { ultima[iMail >= 0 && r[iMail] ? String(r[iMail]).toLowerCase() : '#' + i] = r; });
  const keep = H.map((_, i) => i).filter(i => i !== iMail);
  const out = [keep.map(i => H[i])].concat(Object.values(ultima).map(r => keep.map(i => r[i])));
  const sh = publica_().getSheetByName('Respuestas');
  sh.clearContents();
  sh.getRange(1, 1, out.length, out[0].length).setValues(out);
}

function mailVerificado_(token) {
  if (!token) return null;
  const res = UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(token), { muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) return null;
  const t = JSON.parse(res.getContentText());
  if (t.aud !== CLIENT_ID || String(t.email_verified) !== 'true' || Number(t.exp) * 1000 < Date.now()) return null;
  return String(t.email).toLowerCase();
}

function apodoDe_(mail) {
  const data = SpreadsheetApp.openById(PRIVADA_ID).getSheets()[0].getDataRange().getValues();
  const H = data[0].map(String);
  const iMail = H.findIndex(h => /correo|email|e-mail/i.test(h));
  const iApodo = H.findIndex(h => /nombre o apodo/i.test(h));
  if (iMail < 0 || iApodo < 0) return '';
  let apodo = '';
  data.slice(1).forEach(r => { if (String(r[iMail]).toLowerCase() === mail && r[iApodo]) apodo = String(r[iApodo]).trim(); });
  return apodo;
}

const json_ = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);

function doPost(e) {
  let d; try { d = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad_request' }); }
  const mail = mailVerificado_(d.token);
  if (!mail) return json_({ ok: false, error: 'login' });
  const apodo = apodoDe_(mail);
  if (!apodo) return json_({ ok: false, error: 'sin_encuesta' });
  if (d.action === 'me') return json_({ ok: true, apodo: apodo });
  if (d.action === 'vote') {
    const ok = s => typeof s === 'string' && s.length <= 80 && /^[\w+\-]*$/.test(s);
    if (!ok(d.stop) || !ok(d.opt) || !d.stop) return json_({ ok: false, error: 'bad_request' });
    const lock = LockService.getScriptLock(); lock.waitLock(10000);
    try { publica_().getSheetByName('Votos').appendRow([new Date(), slug_(apodo), d.stop, d.opt]); } finally { lock.releaseLock(); }
    return json_({ ok: true, apodo: apodo });
  }
  return json_({ ok: false, error: 'bad_request' });
}
