/*** vesrsione 1.0
 * PRECOMPUTE OFFERTA  (Fase 2) ***/
// Dipendenze riusate da main.js: CRMdatabase, sheetOfferte, sheetCronologia,
// nomeFileDatiTecnici, DATI_TECNICI_TEMPLATE_ID, newSubfolder, processDatiTecnici, newCharts, CHART_DEFINITIONS

function precomputeOffer(offerAppID) {
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const off = _openSheet_(sheetOfferte || CRMdatabase.getSheetByName('offerte'));
    const cro = _openSheet_(sheetCronologia || CRMdatabase.getSheetByName('cronologia'));
    const out = _openSheet_(CRMdatabase.getSheetByName('offerte_output'));

    // 1) Carica riga OFFERTA
    const o = _findRowByKey_(off, 'appID', offerAppID);
    if (!o) throw new Error('Offerta non trovata: ' + offerAppID);

    // campi offerta (usa header esatti)
    const oppRef            = _val(o, 'opportunitàREF');          // FK → cronologia.id
    const tilt              = _val(o, 'tilt');
    const azimuth           = _val(o, 'azimuth');
    const tipoModuli        = _val(o, 'tipo moduli');
    const numeroModuli      = _val(o, 'numero moduli');
    const potenzaKwp        = _val(o, 'Potenza [kWp]');
    const modelloBatteria   = _val(o, 'modello batteria');
    const numeroBatterie    = _val(o, 'numero batterie');
    const accumulo          = _val(o, 'Accumulo');
    const prezzoTotIncl     = _val(o, 'prezzo offerta appros. - tutto incluso');
    const anniDurIncentivo  = _val(o, 'anni durata incentivo');
    const rid               = _val(o, 'RID');
    const tipoPagamento     = _val(o, 'tipo_pagamento');
    const anniFinanziamento = _val(o, 'anni finanziamento');
    const accontoDiretto    = _val(o, 'Acconto diretto');
    const esposizione       = _val(o, 'esposizione'); // es. "30° est"

    // 2) Carica riga LEAD (cronologia) via id=opportunitàREF
    const lead = _findRowByKey_(cro, 'id', oppRef);
    if (!lead) throw new Error('Lead non trovato (cronologia.id=' + oppRef + ')');
    const leadAppID   = _val(lead, 'appID'); // per refIDlead in out
    const cartellaURL = _val(lead, 'cartella');
    const cartellaId  = _extractId_(cartellaURL);
    let   prjURL      = _val(lead, 'sottocartella_progetto'); // se presente
    let   prjId       = prjURL ? _extractId_(prjURL) : newSubfolder(cartellaId, 'progetto');
    prjURL            = 'https://drive.google.com/drive/folders/' + prjId;

    // Dati per processDatiTecnici (da cronologia)
    const consumiAnnui      = _val(lead, 'kwh annui');
    const profiloConsumo    = _val(lead, 'profilo di consumo');
    const provincia         = _val(lead, 'provincia');
    const prezzoEnergia     = _val(lead, 'prezzo energia');
    const indirizzo         = _val(lead, 'indirizzo');
    const coordinate        = _val(lead, 'coordinate'); // "lat,lon"

    // 3) Fingerprint (loss fissato a 14 per coerenza con uso attuale)
    const loss = 14;
    const fpObj = {
      tilt, azimuth, tipoModuli, numeroModuli, potenzaKwp, modelloBatteria, numeroBatterie,
      accumulo, prezzoTotIncl, anniDurIncentivo, rid, tipoPagamento, anniFinanziamento,
      accontoDiretto, coordinate, loss
    };
    const calc_fingerprint = _md5_(JSON.stringify(fpObj));

    // 4) Idempotenza su offerte_output
    const outRow = _findRowByKey_(out, 'refIDofferta', offerAppID);
    if (outRow && _val(outRow, 'calc_ready') === true && _val(outRow, 'calc_fingerprint') === calc_fingerprint) {
      // già pronto e identico → esci
      return;
    }

    // 5) Assicura DATI TECNICI in progetto/
    const dtId = _ensureDatiTecniciFile_(prjId, (typeof nomeFileDatiTecnici !== 'undefined' ? nomeFileDatiTecnici : 'dati tecnici'));
    const dtSS = SpreadsheetApp.openById(dtId);

    // 6) Scrivi log in dati tecnici e lancia calcoli PVGIS + analisi
    if (typeof newLogDatiTecnici === 'function') {
      newLogDatiTecnici(dtSS, offerAppID);
    }
    const res = processDatiTecnici(
      dtSS,
      consumiAnnui,           // kwh annui (cronologia)
      profiloConsumo,         // profilo di consumo (cronologia)
      provincia,              // provincia (cronologia)
      esposizione,            // esposizione (offerte)
      prezzoEnergia,          // prezzo energia (cronologia)
      indirizzo,              // indirizzo (cronologia)
      coordinate,             // coordinate (cronologia)
      offerAppID,             // appID offerta
      prjId                   // id cartella progetto
    );

    // 7) Leggi i risultati finali (named ranges) dal file Dati Tecnici
    const results = _readAnalisiEnergetica_(dtSS);

    // 8) Genera/aggiorna CHART e spostalo in progetto/chart/ con nome timestamp
    const chartFolderId = _ensureSubfolder_(prjId, 'chart');
    const ts = _fmtTs_(new Date());
    let chartFileId = null, chartFileUrl = null;

    if (typeof newCharts === 'function' && typeof CHART_DEFINITIONS !== 'undefined') {
      const mapping = newCharts(dtSS, prjId, CHART_DEFINITIONS);
      // prova a prendere il primo chart (o CHART_RITORNO_25_ANNI se presente)
      const key = mapping['CHART_RITORNO_25_ANNI'] ? 'CHART_RITORNO_25_ANNI' : Object.keys(mapping)[0];
      if (key) {
        chartFileId = mapping[key].fileId || mapping[key];
        const f = DriveApp.getFileById(chartFileId);
        // move & rename
        DriveApp.getFolderById(chartFolderId).addFile(f);
        f.setName(`chart_ritorno_25_anni_${offerAppID}_${ts}.png`);
        chartFileId = f.getId();
        chartFileUrl = 'https://drive.google.com/file/d/' + chartFileId + '/view';
      }
    }

    // 9) Scrivi/aggiorna OFFERTA_OUTPUT
    const now = new Date();
    if (outRow) {
      _writeRow_(out, outRow._rowNumber, {
        calc_fingerprint: calc_fingerprint,
        calc_ready: true,
        last_calc_ts: now,
        percentuale_autoconsumo: results.percentuale_autoconsumo,
        anni_ritorno_investimento: results.anni_ritorno_investimento,
        percentuale_risparmio_energetico: results.percentuale_risparmio_energetico,
        media_vendita: results.media_vendita,
        utile_25_anni: results.utile_25_anni,
        incentivo_effettivo: results.incentivo_effettivo,
        massimale: results.massimale,
        rata_mensile: results.rata_mensile,
        produzione_primo_anno: results.produzione_primo_anno,
        chart_file_id: chartFileId,
        chart_file_url: chartFileUrl
      });
    } else {
      const rowObj = {
        appIDoutput: Utilities.getUuid(),
        refIDofferta: offerAppID,
        refIDlead: leadAppID,
        calc_fingerprint: calc_fingerprint,
        calc_ready: true,
        last_calc_ts: now,
        percentuale_autoconsumo: results.percentuale_autoconsumo,
        anni_ritorno_investimento: results.anni_ritorno_investimento,
        percentuale_risparmio_energetico: results.percentuale_risparmio_energetico,
        media_vendita: results.media_vendita, 
        utile_25_anni: results.utile_25_anni,
        incentivo_effettivo: results.incentivo_effettivo,
        massimale: results.massimale,
        rata_mensile: results.rata_mensile,
        produzione_primo_anno: results.produzione_primo_anno,
        chart_file_id: chartFileId,
        chart_file_url: chartFileUrl,
        // campi doc e print_ts rimangono vuoti: li compilerà il print runner
        presentazione_doc_id: '',
        presentazione_doc_url: '',
        offerta_doc_id: '',
        offerta_doc_url: '',
        print_ts: ''
      };
      _appendRow_(out, rowObj);
    }
  } finally {
    lock.releaseLock();
  }
}

/*** Helpers ***/
function _openSheet_(sh) {
  if (!sh) throw new Error('Sheet non trovato');
  const head = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(h => String(h).trim());
  const idx  = Object.fromEntries(head.map((h,i)=>[h,i]));
  return { sh, head, idx };
}
function _val(rowObj, colName) {
  const { head, idx } = rowObj._table;
  const i = idx[colName]; if (i == null) throw new Error('Colonna mancante: ' + colName);
  return rowObj._values[i];
}
function _findRowByKey_(table, keyCol, keyVal) {
  const { sh, head, idx } = table;
  const iKey = idx[keyCol]; if (iKey == null) throw new Error('Colonna chiave mancante: ' + keyCol);
  const last = sh.getLastRow(); if (last <= 1) return null;
  const rng  = sh.getRange(2,1,last-1,sh.getLastColumn()).getValues();
  for (let r=0; r<rng.length; r++){
    if (String(rng[r][iKey]).trim() === String(keyVal).trim()){
      return { _table: table, _rowNumber: r+2, _values: rng[r] };
    }
  }
  return null;
}
function _writeRow_(table, rowNumber, obj) {
  const { sh, idx } = table;
  const row = sh.getRange(rowNumber, 1, 1, sh.getLastColumn()).getValues()[0];
  Object.keys(obj).forEach(k=>{
    if (idx[k] != null) row[idx[k]] = obj[k];
  });
  sh.getRange(rowNumber, 1, 1, sh.length || sh.getLastColumn()).setValues([row]);
}
function _appendRow_(table, obj) {
  const { sh, head, idx } = table;
  const row = head.map(h => idx[h] != null ? (obj[h] ?? '') : '');
  // sopra non riempie correttamente (head order). Meglio costruire array per indice:
  const arr = new Array(head.length).fill('');
  Object.keys(obj).forEach(k=>{
    if (idx[k] != null) arr[idx[k]] = obj[k];
  });
  sh.appendRow(arr);
}
function _extractId_(urlOrId) {
  const m = String(urlOrId||'').match(/[-\\w]{25,}/);
  return m ? m[0] : '';
}
function _ensureDatiTecniciFile_(projectFolderId, fileName) {
  const f = DriveApp.getFolderById(projectFolderId);
  const it = f.getFilesByName(fileName);
  if (it.hasNext()) return it.next().getId();
  const file = DriveApp.getFileById(DATI_TECNICI_TEMPLATE_ID).makeCopy(fileName, f);
  return file.getId();
}
function _ensureSubfolder_(parentId, name) {
  const parent = DriveApp.getFolderById(parentId);
  const it = parent.getFoldersByName(name);
  if (it.hasNext()) return it.next().getId();
  return parent.createFolder(name).getId();
}
function _md5_(s) {
  const raw = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, s);
  return raw.map(b => (b+256&255).toString(16).padStart(2,'0')).join('');
}
function _readAnalisiEnergetica_(ss) {
  // Named ranges nel file Dati Tecnici (allineati con main.js aggiornato)
  const get = n => ss.getRangeByName(n).getValue();
  return {
    percentuale_autoconsumo:       get('percentuale_autoconsumo'),
    media_vendita:                 get('media_vendita'),
    anni_ritorno_investimento:     get('anni_ritorno_investimento'),
    percentuale_risparmio_energetico:get('percentuale_risparmio_energetico'),
    utile_25_anni:                 get('utile_25_anni'),
    incentivo_effettivo:           get('incentivo_effettivo'),
    massimale:                     get('massimale'),
    rata_mensile:                  get('rata_mensile'),
    produzione_primo_anno:         get('produzione_primo_anno')
  };
}
function _fmtTs_(d) {
  const pad = n => String(n).padStart(2,'0');
  return d.getFullYear().toString() +
         pad(d.getMonth()+1) +
         pad(d.getDate()) + '_' +
         pad(d.getHours()) + pad(d.getMinutes()) + pad(d.getSeconds());
}

