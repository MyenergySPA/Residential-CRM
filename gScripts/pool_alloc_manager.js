/***** CONFIG SOLO DI QUESTO MODULO (nessuna duplicazione con main.js) *****/

// Tabella pool
const POOL_SHEET_NAME = 'pool_alloc';

// Cartella radice dove creare i placeholder (ID reale che mi hai fornito)
const CLIENTI_ROOT_FOLDER_ID = '1kpBsmlPAaeCFWvgCEIw38tEk5Q-xQpH_';

// Scorta minima/limite e policy di pulizia
const TARGET_UNUSED = 100;        // minimo slot liberi da garantire
const MAX_UNUSED_CAP = 200;       // cap per evitare esplosioni
const CLEANUP_USED_AFTER_DAYS = 7;

// Nomi sottocartelle (coerenti con la tua struttura)
const SUBFOLDERS = {
  contratto: 'contratto',
  progetto: 'progetto',
  allegati: 'allegati',
  documenti: 'documenti'
};

/* =========================================================================
   HELPER: accesso a Sheets/Drive riusando oggetti del main.js
   ========================================================================= */

function _getCRM_() {
  // Riusa l'oggetto globale CRMdatabase se esiste (definito in main.js)
  if (typeof CRMdatabase !== 'undefined' && CRMdatabase.getId) return CRMdatabase;
  throw new Error('CRMdatabase non è definito nel progetto. Assicurati che main.js sia caricato.');
}

function _getPoolSheet_() {
  const ss = _getCRM_();
  const sh = ss.getSheetByName(POOL_SHEET_NAME);
  if (!sh) throw new Error('Sheet "' + POOL_SHEET_NAME + '" non trovato nel CRM.');
  return sh;
}

function _getCronoSheet_() {
  // Se sheetCronologia globale esiste, riusalo. Altrimenti apri dal CRMdatabase
  if (typeof sheetCronologia !== 'undefined' && sheetCronologia.getLastRow) return sheetCronologia;
  const sh = _getCRM_().getSheetByName('cronologia');
  if (!sh) throw new Error('Sheet "cronologia" non trovato nel CRM.');
  return sh;
}

function _headers_(sh) {
  const cols = sh.getLastColumn();
  if (cols < 1) throw new Error('Foglio senza colonne: ' + sh.getName());
  const head = sh.getRange(1,1,1,cols).getValues()[0].map(String);
  const idx  = Object.fromEntries(head.map((h,i)=>[h,i]));
  const col  = name => {
    const k = String(name).trim();
    if (!(k in idx)) throw new Error('Colonna "'+k+'" non trovata su ' + sh.getName());
    return idx[k] + 1;
  };
  return { head, col };
}

function _getAllRows_(sh) {
  const last = sh.getLastRow();
  if (last <= 1) return [];
  return sh.getRange(2,1,last-1,sh.getLastColumn()).getValues();
}

function _asBool_(v){ return v === true || String(v).toLowerCase() === 'true'; }
function _toDate_(v){ return v instanceof Date ? v : (v ? new Date(v) : null); }

/* =========================================================================
   GENERATORE ID 4 CIFRE (formato invariato) con controllo unicità
   ========================================================================= */

function _buildUsedIdSet_() {
  const used = new Set();

  // cronologia[id]
  const cr = _getCronoSheet_();
  const { head: hC, col: colC } = _headers_(cr);
  const lastC = cr.getLastRow();
  if (lastC > 1 && hC.includes('id')) {
    const vals = cr.getRange(2, colC('id'), lastC-1, 1).getDisplayValues().flat();
    vals.forEach(v => { if (v) used.add(String(v).trim()); });
  }

  // pool_alloc[id_univoco]
  const pl = _getPoolSheet_();
  const { head: hP, col: colP } = _headers_(pl);
  const lastP = pl.getLastRow();
  if (lastP > 1 && hP.includes('id_univoco')) {
    const vals = pl.getRange(2, colP('id_univoco'), lastP-1, 1).getDisplayValues().flat();
    vals.forEach(v => { if (v) used.add(String(v).trim()); });
  }

  return used;
}

function _pickUnique4Digit_(used) {
  if (used.size >= 10000) throw new Error('Spazio ID 4-cifre esaurito (10.000).');

  // tentativi random veloci
  for (let t=0; t<40; t++) {
    const n = Math.floor(Math.random() * 10000);
    const s = ('0000' + n).slice(-4);
    if (!used.has(s)) { used.add(s); return s; }
  }
  // fallback deterministico
  for (let n=0; n<10000; n++) {
    const s = ('0000' + n).slice(-4);
    if (!used.has(s)) { used.add(s); return s; }
  }
  throw new Error('Impossibile ottenere un ID a 4 cifre univoco.');
}

/* =========================================================================
   CREAZIONE PLACEHOLDER (Drive) + SCRITTURA RIGA POOL (Sheets)
   ========================================================================= */

function _createPlaceholderFolders_() {
  const root = DriveApp.getFolderById(CLIENTI_ROOT_FOLDER_ID);
  const key  = Utilities.getUuid(); // appIDpool (KEY in AppSheet)
  const main = root.createFolder('placeholder-' + key.slice(0,8));
  const urlMain = main.getUrl();

  const urlC = main.createFolder(SUBFOLDERS.contratto).getUrl();
  const urlP = main.createFolder(SUBFOLDERS.progetto ).getUrl();
  const urlA = main.createFolder(SUBFOLDERS.allegati ).getUrl();
  const urlD = main.createFolder(SUBFOLDERS.documenti).getUrl();

  return { key, urlMain, urlC, urlP, urlA, urlD };
}

function _appendPoolRow_(sh, head, obj) {
  const row = head.map(h => (h in obj ? obj[h] : ''));
  const r   = sh.getLastRow() + 1;
  // assicura che esista almeno l'header
  if (r === 1) throw new Error('Sheet "'+sh.getName()+'" senza intestazioni.');
  sh.getRange(r, 1, 1, head.length).setValues([row]);
  SpreadsheetApp.flush();
}

/* =========================================================================
   PUBBLICHE: TOP-UP, CLEANUP, RENAME
   ========================================================================= */

/**
 * Porta gli slot liberi in pool_alloc ad almeno TARGET_UNUSED.
 * Crea cartelle + sottocartelle e scrive riga per riga (robusto).
 */
function topUpPoolToTarget() {
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const shPool = _getPoolSheet_();
    const { head: hP, col: colP } = _headers_(shPool);

    // conta liberi
    const rows = _getAllRows_(shPool);
    const free = rows.filter(r => !_asBool_(r[colP('assegnato_check')-1]) && !r[colP('assegnato_ref')-1]).length;

    if (free >= TARGET_UNUSED) return;
    const need = Math.min(TARGET_UNUSED - free, Math.max(0, MAX_UNUSED_CAP - free));
    if (need <= 0) return;

    const used = _buildUsedIdSet_(); // garantisce unicità 4 cifre
    for (let i=0; i<need; i++) {
      const { key, urlMain, urlC, urlP, urlA, urlD } = _createPlaceholderFolders_();
      const id4 = _pickUnique4Digit_(used);

      const kv = {
        appIDpool: key,
        id_univoco: id4,
        cartella_pool: urlMain,
        sottocartella_contratto_pool: urlC,
        sottocartella_progetto_pool:  urlP,
        sottocartella_allegati_pool:  urlA,
        sottocartella_documenti_pool: urlD,
        assegnato_check: false,
        assegnato_ref: '',
        assegnato_data: ''
      };

      _appendPoolRow_(shPool, hP, kv);
      Utilities.sleep(120); // piccolo respiro per stabilità UI/Drive
    }
  } finally {
    lock.releaseLock();
  }
}

/**
 * Elimina da pool_alloc le righe "usate" più vecchie di CLEANUP_USED_AFTER_DAYS.
 * NON tocca le cartelle Drive.
 */
function cleanupUsedPool() {
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const shPool = _getPoolSheet_();
    const { col: colP } = _headers_(shPool);
    const last = shPool.getLastRow();
    if (last <= 1) return;

    const data = shPool.getRange(2,1,last-1,shPool.getLastColumn()).getValues();
    const toDelete = [];
    for (let i=0;i<data.length;i++) {
      const used = _asBool_(data[i][colP('assegnato_check')-1]);
      const dt   = _toDate_(data[i][colP('assegnato_data')-1]);
      if (!used || !dt) continue;
      const age = (new Date() - dt) / (1000*3600*24);
      if (age > CLEANUP_USED_AFTER_DAYS) toDelete.push(i+2);
    }
    toDelete.reverse().forEach(r => shPool.deleteRow(r));
  } finally {
    lock.releaseLock();
  }
}

/**
 * Rinomina la cartella del lead (richiamata dal Bot AppSheet con Call a script).
 * Formato: <tipo_opportunità>-<id>-<yy> <nome referente> <cognome referente>
 */
function renameAssignedFolderForLead(appID) {
  const shC = _getCronoSheet_();
  const shP = _getPoolSheet_();
  const { head: hC, col: colC } = _headers_(shC);
  const { head: hP, col: colP } = _headers_(shP);

  // trova riga lead
  const lastC = shC.getLastRow();
  if (lastC <= 1) throw new Error('cronologia vuota');
  const dataC = shC.getRange(2,1,lastC-1,shC.getLastColumn()).getValues();
  const iLead = dataC.findIndex(r => String(r[colC('appID')-1]).trim() === String(appID).trim());
  if (iLead < 0) throw new Error('Lead non trovato: ' + appID);
  const rowC = dataC[iLead];

  // campi per nome cartella (come da tuo formato attuale)
  const tipo = String(rowC[colC('tipo_opportunità')-1] || '').trim();
  const id   = String(rowC[colC('id')-1] || '').trim();
  const yy   = String(rowC[colC('yy')-1] || '').trim();
  const nome = String(rowC[colC('nome referente')-1] || '').trim();
  const cogn = String(rowC[colC('cognome referente')-1] || '').trim();

  // slot pool collegato (REF)
  const poolKey = String(rowC[colC('pool_ID')-1] || '').trim();
  if (!poolKey) throw new Error('pool_ID mancante nella riga lead ' + appID);

  // trova riga pool e folder
  const lastP = shP.getLastRow();
  if (lastP <= 1) throw new Error('pool_alloc vuoto');
  const dataP = shP.getRange(2,1,lastP-1,shP.getLastColumn()).getValues();
  const rowP  = dataP.find(r => String(r[colP('appIDpool')-1]).trim() === poolKey);
  if (!rowP) throw new Error('slot pool non trovato: ' + poolKey);

  const folderUrl = String(rowP[colP('cartella_pool')-1] || '').trim();
  const folderId  = _extractFolderId_(folderUrl);

  const newName = _safeName_(`${tipo}-${id}-${yy} ${nome} ${cogn}`);
  DriveApp.getFolderById(folderId).setName(newName);
}

/* =========================================================================
   PICCOLI HELPER GENERICI
   ========================================================================= */

function _extractFolderId_(urlOrId) {
  if (!urlOrId) throw new Error('URL/ID cartella mancante');
  const m = String(urlOrId).match(/[-\w]{25,}/);
  return m ? m[0] : String(urlOrId);
}

function _safeName_(s) {
  return String(s).replace(/[\\/:*?"<>|#\[\]]/g, ' ').replace(/\s+/g, ' ').trim();
}

