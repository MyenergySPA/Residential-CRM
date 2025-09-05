/**
 * PRE-GEN “centralizzata”: crea/aggiorna i file PRE nella cartella
 * …/contratto/placeholder/ e assicura “dati tecnici” in …/progetto/.
 *
 * Trigger: AppSheet Bot quando stato = "offerta da effettuare".
 *
 * Dipendenze riusate dal progetto:
 *  - CRMdatabase (da main.js)
 *  - nomeFileDatiTecnici, DATI_TECNICI_TEMPLATE_ID (da main.js)
 *  - newSubfolder(parentId, name) (da newSubfolders.js)
 *  - TEMPLATES (da docTemplates.js)
 */

function ensurePlaceholders(leadAppID) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    // === BEGIN: ottenimento ID (semplificato, solo cronologia.appID) ===
    const got = _loadCronologiaRow_(leadAppID);
    if (!got || !got.row) throw new Error('Lead non trovato per appID=' + leadAppID);
    const row  = got.row;
    const cols = got.cols; // funzione col(name) già fornita da _loadCronologiaRow_
    // === END: ottenimento ID ===

    // ====== TUO CODICE INVARIATO DA QUI IN POI ======
    const cartellaClienteId   = _extractFolderId_(row[cols('cartella')]);
    const folderContrattoId   = newSubfolder(cartellaClienteId, 'contratto');
    const folderProgettoId    = newSubfolder(cartellaClienteId, 'progetto');
    const folderPlaceholderId = newSubfolder(folderContrattoId, 'placeholder');

    _ensureDatiTecnici_(folderProgettoId);

    const tipoOp  = String(row[cols('tipo_opportunità')] || '').toUpperCase().trim();
    const tipoPay = String(row[cols('tipo pagamento scelto')] || '').trim();
    const id      = String(row[cols('id')] || '').trim();
    const yy      = String(row[cols('yy')] || '').trim();

    const needed = _selectDocsNoIncentive_(tipoOp, tipoPay);
    needed.forEach(d => {
      const preName   = _buildPreName_(d.docType, id, yy);
      const existsPre = _findFileInFolderByName_(folderPlaceholderId, preName);
      if (!existsPre) {
        const templateId = _templateIdForDocType_(d.docType);
        DriveApp.getFileById(templateId).makeCopy(preName, DriveApp.getFolderById(folderPlaceholderId));
      }
    });

  } finally {
    lock.releaseLock();
  }
}

/* ===================== Helpers dominio ===================== */

function _selectDocsNoIncentive_(tipo_opportunita, tipo_pagamento) {
  // Matrice richiesta dal cliente (niente incentivo):
  // - MAT → solo OFFERTA-MAT
  // - Finanziamento → PRESENTAZIONE-FIN + OFFERTA-FIN
  // - Diretto + COND → PRESENTAZIONE-COND + OFFERTA
  // - Diretto + (non COND & non MAT) → PRESENTAZIONE + OFFERTA

  if (tipo_opportunita === 'MAT') {
    return [{ docType: 'OFFERTA-MAT' }];
  }

  if (tipo_pagamento === 'Finanziamento') {
    return [{ docType: 'PRESENTAZIONE-FIN' }, { docType: 'OFFERTA-FIN' }];
  }

  if (tipo_opportunita === 'COND') {
    return [{ docType: 'PRESENTAZIONE-COND' }, { docType: 'OFFERTA' }];
  }

  // default diretto “normale”
  return [{ docType: 'PRESENTAZIONE' }, { docType: 'OFFERTA' }];
}

function _templateIdForDocType_(docType) {
  // Map ai tuoi ID template in docTemplates.js (già caricati nel progetto)
  switch (docType) {
    case 'OFFERTA-MAT':        return TEMPLATES.offerta_MAT;
    case 'OFFERTA-FIN':        return TEMPLATES.offerta_Finanz;
    case 'OFFERTA':            return TEMPLATES.offerta;
    case 'PRESENTAZIONE-FIN':  return TEMPLATES.presentazione_Finanz;
    case 'PRESENTAZIONE-COND': return TEMPLATES.presentazione_COND;
    case 'PRESENTAZIONE':      return TEMPLATES.presentazione;
    default:
      throw new Error('docType non supportato: ' + docType);
  }
}

function _buildPreName_(docType, id, yy) {
  // Nome PRE indipendente dalla data, così resta “riutilizzabile” nel tempo
  // Esempio: "OFFERTA-FIN - 0123-25 PRE"
  const cleanId = (id || '').trim();
  const cleanYy = (yy || '').trim();
  return `${docType} - ${cleanId}-${cleanYy} PRE`;
}

/* ===================== Helpers dati tecnici ===================== */

function _ensureDatiTecnici_(folderProgettoId) {
  // nomeFileDatiTecnici & DATI_TECNICI_TEMPLATE_ID sono nel main.js
  const name = String(nomeFileDatiTecnici || '').trim();
  if (!name) throw new Error('nomeFileDatiTecnici non definito in main.js');

  if (!_findFileInFolderByName_(folderProgettoId, name)) {
    DriveApp.getFileById(DATI_TECNICI_TEMPLATE_ID)
            .makeCopy(name, DriveApp.getFolderById(folderProgettoId));
  }
}

/* ===================== Helpers Sheets/Drive ===================== */

function _loadCronologiaRow_(appID) {
  if (typeof CRMdatabase === 'undefined' || !CRMdatabase.getSheetByName) {
    throw new Error('CRMdatabase non è disponibile (main.js).');
  }
  const sh = CRMdatabase.getSheetByName('cronologia');
  if (!sh) throw new Error('Sheet "cronologia" non trovato.');

  const head = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(h=>String(h).trim());
  const colIdx = Object.fromEntries(head.map((h,i)=>[h,i]));
  const col = name => {
    if (!(name in colIdx)) throw new Error('Colonna mancante: ' + name);
    return colIdx[name];
  };

  const last = sh.getLastRow();
  if (last <= 1) return { row: null, cols: col };
  const vals = sh.getRange(2,1,last-1,sh.getLastColumn()).getValues();

  const i = vals.findIndex(r => String(r[col('appID')]).trim() === String(appID).trim());
  return { row: i>=0 ? vals[i] : null, cols: col };
}

function _extractFolderId_(urlOrId) {
  const m = String(urlOrId||'').match(/[-\w]{25,}/);
  if (!m) throw new Error('URL cartella non valido: ' + urlOrId);
  return m[0];
}

function _findFileInFolderByName_(folderId, name) {
  const it = DriveApp.getFolderById(folderId).getFiles();
  while (it.hasNext()) {
    const f = it.next();
    if (f.getName() === name) return f;
  }
  return null;
}

