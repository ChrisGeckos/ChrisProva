/**
 * ChrisGeckos - esemplari disponibili
 * Legge il foglio Google "disponibili" e la cartella Drive "Foto disponibili",
 * abbina ogni foto al codice (COD) dell'esemplare dal nome del file
 * e restituisce tutto al sito in formato JSON.
 *
 * Nome foto -> codice:  21CF84.jpg  (le maiuscole e l'estensione non contano)
 * Più foto per lo stesso esemplare:  21CF84_2.jpg  oppure  21CF84 (2).jpg
 */

const SHEET_ID  = '13MEM5y0usM3giikW9E2GTF2bNzbhJ1eM_9EBFbgZ5Ps';
const FOLDER_ID = '1sQvIupby0TDnJECdCITYYiJGLb2Qb0-G';
const CACHE_SECONDS = 300; // i dati si aggiornano al massimo ogni 5 minuti

function doGet() {
  let body = null;
  try {
    body = CacheService.getScriptCache().get('animali');
  } catch (e) { /* la cache è facoltativa */ }

  if (!body) {
    body = JSON.stringify(buildData_());
    try {
      CacheService.getScriptCache().put('animali', body, CACHE_SECONDS);
    } catch (e) { /* dati troppo grandi per la cache: si ricalcolano ogni volta */ }
  }
  return ContentService.createTextOutput(body).setMimeType(ContentService.MimeType.JSON);
}

function buildData_() {
  // ---- Foglio (primo foglio del file) ----
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
  const values = sheet.getDataRange().getDisplayValues();
  const head = values[0].map(function (h) { return String(h).trim().toLowerCase().replace(/[:\s]+$/, ''); });
  const aliases = {
    cod: ['cod', 'codice'], sesso: ['sesso'], mf: ['m/f', 'mf'], nascita: ['nascita'],
    morph: ['morph'], note: ['note'], proven: ['riproduttore', 'proven'],
    specie: ['specie'], prezzo: ['prezzo'], stato: ['stato', 'prenotato', 'prenotata', 'prenotazione']   // facoltative (stato = casella di spunta o testo)
  };
  const col = {};
  Object.keys(aliases).forEach(function (k) {
    col[k] = head.findIndex(function (h) { return aliases[k].indexOf(h) !== -1; });
  });
  const get = function (row, k) { return col[k] >= 0 ? String(row[col[k]]).trim() : ''; };

  // ---- Foto della cartella Drive ----
  const photos = [];
  const files = DriveApp.getFolderById(FOLDER_ID).getFiles();
  while (files.hasNext()) {
    const f = files.next();
    if (String(f.getMimeType()).indexOf('image/') === 0) {
      photos.push({ name: f.getName(), key: normalize_(f.getName()), id: f.getId(), used: false });
    }
  }
  photos.sort(function (a, b) { return a.key < b.key ? -1 : a.key > b.key ? 1 : 0; });

  // ---- Abbinamento nome foto <-> codice ----
  const animali = [];
  const senzaFoto = [];
  values.slice(1).forEach(function (row) {
    const cod = get(row, 'cod');
    if (!cod) return;
    const key = normalize_(cod);
    const mine = photos.filter(function (p) {
      if (p.key === key) return true;
      return p.key.indexOf(key) === 0 && /^[_\-(]\d+\)?$/.test(p.key.slice(key.length));
    });
    mine.forEach(function (p) { p.used = true; });
    if (!mine.length) senzaFoto.push(cod);

    animali.push({
      cod: cod,
      sesso: get(row, 'sesso'),
      mf: get(row, 'mf'),
      nascita: get(row, 'nascita'),
      morph: get(row, 'morph'),
      note: get(row, 'note'),
      proven: get(row, 'proven'),
      specie: get(row, 'specie'),
      prezzo: get(row, 'prezzo'),
      stato: get(row, 'stato'),
      foto: mine.map(function (p) { return p.id; })
    });
  });

  return {
    ok: true,
    aggiornato: new Date().toISOString(),
    animali: animali,
    senzaFoto: senzaFoto,                                   // codici nel foglio senza foto
    fotoOrfane: photos.filter(function (p) { return !p.used; }).map(function (p) { return p.name; }) // foto senza riga nel foglio
  };
}

function normalize_(name) {
  return String(name).replace(/\.[^.]+$/, '').replace(/\s+/g, '').toUpperCase();
}

/** Prova manuale: eseguila dall'editor (menu Esegui) e guarda il risultato in Registro di esecuzione. */
function test() {
  const d = buildData_();
  Logger.log('Esemplari: ' + d.animali.length);
  Logger.log('Senza foto: ' + JSON.stringify(d.senzaFoto));
  Logger.log('Foto senza esemplare: ' + JSON.stringify(d.fotoOrfane));
}
