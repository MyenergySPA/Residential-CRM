/*** v3 — precompute con 2 percorsi: FAST (clone) / STANDARD (calcolo)
 * Dipendenze esterne presenti nel progetto:
 *  - nomeFileDatiTecnici, DATI_TECNICI_TEMPLATE_ID
 *  - newSubfolder, newLogDatiTecnici, processDatiTecnici, newCharts, CHART_DEFINITIONS
 *  - CRMdatabase (Spreadsheet globale)
 ***/

/* -------------------- Util minimi -------------------- */
function _extractFolderId_(urlOrId) {
  const m = String(urlOrId || '').match(/[-\w]{25,}/);
  if (!m) throw new Error('URL/ID cartella non valido');
  return m[0];
}
function _norm_(v) {
  if (v === null || v === undefined) return '';
  if (v instanceof Date) return v.toISOString();
  if (typeof v === 'number') return String(Math.round(v * 1000) / 1000);
  return String(v).trim();
}
function _toNumber_(v) {
  if (typeof v === 'number') return v;
  const s = String(v || '').replace(/\./g, '').replace(',', '.').replace(/[^\d.-]/g, '');
  const n = Number(s);
  if (!isFinite(n)) throw new Error('Valore numerico non valido: ' + v);
  return n;
}
function _toEuroPerKWh_(v) {
  // Accetta "0,35", "0.35", "35" (cent/kWh) e normalizza a €/kWh
  const n = _toNumber_(v);
  // Se sembra in centesimi (tra 3 e 200), convertilo in euro
  const fixed = (n > 3 && n < 200) ? (n / 100) : n;
  return Math.round(fixed * 10000) / 10000; // 4 decimali
}
function _openByName_(ss, tabName) {
  const sh = ss.getSheetByName(tabName);
  if (!sh) throw new Error('Foglio mancante: ' + tabName);
  const head = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(x => String(x).trim());
  const idx  = Object.fromEntries(head.map((h,i)=>[h,i]));
  const col  = name => { if (!(name in idx)) throw new Error('Colonna mancante in ' + tabName + ': ' + name); return idx[name]+1; };
  return { sh, head, col };
}
function _idx_(sh) {
  const head = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(h => String(h).trim());
  const map = {}; head.forEach((h,i)=> map[h]=i);
  return { head, map, col: n => (n in map ? map[n]+1 : 0) };
}
function _uuid_() { return Utilities.getUuid(); }

/* Fingerprint tecnico dell’offerta (stessa logica della tua versione) */
function _buildFingerprint_({
  tilt, azimuth, tipoModuli, numeroModuli, potenzaKWp,
  modelloBatteria, numeroBatterie, accumulo, prezzoOffertaTuttoIncluso,
  anniDurataIncentivo, rid, tipoPagamento, anniFinanziamento, accontoDiretto, coordinate,
  prezzoEnergia
}) {
  const pairs = [
    ['tilt', tilt],
    ['azimuth', azimuth],
    ['tipo moduli', tipoModuli],
    ['numero moduli', numeroModuli],
    ['Potenza [kWp]', potenzaKWp],
    ['modello batteria', modelloBatteria],
    ['numero batterie', numeroBatterie],
    ['Accumulo', accumulo],
    ['prezzo energia', prezzoEnergia], // <<--- NEW (€/kWh)
    ['prezzo offerta appros. - tutto incluso', prezzoOffertaTuttoIncluso],
    ['anni durata incentivo', anniDurataIncentivo],
    ['RID', rid],
    ['tipo_pagamento', tipoPagamento],
    ['anni finanziamento', anniFinanziamento],
    ['Acconto diretto', accontoDiretto],
    ['coordinate', coordinate],
  ];
  return pairs.map(([k, v]) => k + ':' + _norm_(v)).join('|');
}


/* -------------------- offerte_output helpers -------------------- */

/** Trova la riga per un dato appID (refIDofferta) */
function _getOfferteOutputRowByOffer_(offerId) {
  const { sh, col } = _openByName_(CRMdatabase, 'offerte_output');
  const last = sh.getLastRow(); if (last <= 1) return null;
  const vals = sh.getRange(2,1,last-1,sh.getLastColumn()).getValues();
  const iRef = col('refIDofferta')-1;
  const idx  = vals.findIndex(r => String(r[iRef]).trim() === String(offerId).trim());
  if (idx < 0) return null;
  const row = vals[idx];
  return {
    rowNum: idx+2,
    appIDoutput:      String(row[col('appIDoutput')-1] || ''),
    calc_ready:       row[col('calc_ready')-1] === true || String(row[col('calc_ready')-1]).toLowerCase()==='true',
    calc_fingerprint: String(row[col('calc_fingerprint')-1] || ''),
    chart_file_url:   String(row[col('chart_file_url')-1] || '')
  };
}

/** Aggiorna timestamp + data_version su una riga esistente */
function _touchOfferteOutputRow_(rowNum) {
  const { sh, col } = _openByName_(CRMdatabase, 'offerte_output');
  sh.getRange(rowNum, col('last_calc_ts')).setValue(new Date());
  if (col('data_version')) sh.getRange(rowNum, col('data_version')).setValue('v3');
}

/** FAST-CLONE: clona (in NUOVA riga) l’ultima riga pronta con stesso fingerprint */
function _fastCloneFromFingerprint_({ refIDofferta, refIDlead, fingerprint }) {
  try {
    const sh = CRMdatabase.getSheetByName('offerte_output');
    if (!sh) return null;

    const head = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(h=>String(h).trim());
    const col  = (name) => head.indexOf(name) + 1;
    const need = [
      'appIDoutput','refIDofferta','refIDlead','calc_fingerprint','calc_ready','last_calc_ts',
      'percentuale_autoconsumo','anni_ritorno_investimento','percentuale_risparmio_energetico',
      'produzione_primo_anno','media_vendita','utile_25_anni','incentivo_effettivo','massimale','rata_mensile',
      'chart_file_url','presentazione_doc_url','offerta_doc_url','print_ts','data_version'
    ];
    if (!need.every(n => head.indexOf(n) >= 0)) return null;

    const last = sh.getLastRow();
    if (last <= 1) return null;

    const vals = sh.getRange(2,1,last-1,sh.getLastColumn()).getValues();
    const iFp    = col('calc_fingerprint') - 1;
    const iReady = col('calc_ready') - 1;

    // Cerca DALLA FINE la più recente con fingerprint identico e pronta
    let srcRowIdx = -1;
    for (let i = vals.length - 1; i >= 0; i--) {
      const row = vals[i];
      if (String(row[iFp]).trim() === String(fingerprint).trim()) {
        const ready = (row[iReady] === true) || (String(row[iReady]).toLowerCase() === 'true');
        if (ready) { srcRowIdx = i; break; }
      }
    }
    if (srcRowIdx < 0) return null;

    const getFrom = (n) => vals[srcRowIdx][col(n)-1];
    const cloned = {
      percentuale_autoconsumo:          getFrom('percentuale_autoconsumo'),
      anni_ritorno_investimento:        getFrom('anni_ritorno_investimento'),
      percentuale_risparmio_energetico: getFrom('percentuale_risparmio_energetico'),
      produzione_primo_anno:            getFrom('produzione_primo_anno'),
      media_vendita:                    getFrom('media_vendita'),
      utile_25_anni:                    getFrom('utile_25_anni'),
      incentivo_effettivo:              getFrom('incentivo_effettivo'),
      massimale:                        getFrom('massimale'),
      rata_mensile:                     getFrom('rata_mensile'),
      chart_file_url:                   String(getFrom('chart_file_url') || '')
    };

    const newRow = sh.getLastRow() + 1;
    const newKey = Utilities.getUuid();

    sh.getRange(newRow, col('appIDoutput')).setValue(newKey);
    sh.getRange(newRow, col('refIDofferta')).setValue(refIDofferta);
    sh.getRange(newRow, col('refIDlead')).setValue(refIDlead);
    sh.getRange(newRow, col('calc_fingerprint')).setValue(fingerprint);
    sh.getRange(newRow, col('calc_ready')).setValue(true);
    sh.getRange(newRow, col('last_calc_ts')).setValue(new Date());
    sh.getRange(newRow, col('percentuale_autoconsumo')).setValue(cloned.percentuale_autoconsumo);
    sh.getRange(newRow, col('anni_ritorno_investimento')).setValue(cloned.anni_ritorno_investimento);
    sh.getRange(newRow, col('percentuale_risparmio_energetico')).setValue(cloned.percentuale_risparmio_energetico);
    sh.getRange(newRow, col('produzione_primo_anno')).setValue(cloned.produzione_primo_anno);
    sh.getRange(newRow, col('media_vendita')).setValue(cloned.media_vendita);
    sh.getRange(newRow, col('utile_25_anni')).setValue(cloned.utile_25_anni);
    sh.getRange(newRow, col('incentivo_effettivo')).setValue(cloned.incentivo_effettivo);
    sh.getRange(newRow, col('massimale')).setValue(cloned.massimale);
    sh.getRange(newRow, col('rata_mensile')).setValue(cloned.rata_mensile);
    if (cloned.chart_file_url) sh.getRange(newRow, col('chart_file_url')).setValue(cloned.chart_file_url);
    sh.getRange(newRow, col('data_version')).setValue('v3');

    _writeOutputKeyToOfferte_(refIDofferta, newKey);
    Logger.log('[FAST-CLONE] Nuova riga clonata per appID=' + refIDofferta);
    return { ok:true, appID: refIDofferta, offerte_outputID: newKey, chart_file_url: cloned.chart_file_url, data_version:'v3', clonedFromFingerprint:true };
  } catch (e) {
    Logger.log('[FAST-CLONE] errore: ' + e);
    return null;
  }
}

/** Upsert v3: crea/aggiorna riga per un refIDofferta (un solo setValues) */
function _upsertOfferteOutputV3_(payload) {
  const sh = CRMdatabase.getSheetByName('offerte_output') || CRMdatabase.insertSheet('offerte_output');
  const HEADERS = [
    'appIDoutput','refIDofferta','refIDlead','calc_fingerprint','calc_ready','last_calc_ts',
    'percentuale_autoconsumo','anni_ritorno_investimento','percentuale_risparmio_energetico',
    'produzione_primo_anno','media_vendita','utile_25_anni','incentivo_effettivo','massimale','rata_mensile',
    'chart_file_url','presentazione_doc_url','offerta_doc_url','print_ts','data_version'
  ];
  // Assumiamo che l’intestazione sia già corretta; altrimenti qui potresti chiamare _ensureHeaders_

  const { head, col } = _idx_(sh);
  const last = sh.getLastRow();
  let rowNum = 0, appIDoutput = '';

  if (last > 1) {
    const vals = sh.getRange(2,1,last-1,sh.getLastColumn()).getValues();
    const iRef = col('refIDofferta')-1;
    const idx  = vals.findIndex(r => String(r[iRef]).trim() === String(payload.refIDofferta).trim());
    if (idx >= 0) {
      rowNum = idx + 2;
      appIDoutput = String(vals[idx][col('appIDoutput')-1] || '');
    }
  }
  if (!rowNum) { rowNum = last + 1; appIDoutput = _uuid_(); }

  const row = new Array(head.length).fill('');
  const put = (name, val) => { const i = col(name)-1; if (i>=0) row[i] = val; };

  put('appIDoutput', appIDoutput);
  put('refIDofferta', payload.refIDofferta);
  put('refIDlead', payload.refIDlead || '');
  put('calc_fingerprint', payload.calc_fingerprint || '');
  put('calc_ready', payload.calc_ready === true);
  put('last_calc_ts', payload.last_calc_ts || new Date());

  put('percentuale_autoconsumo',          payload.percentuale_autoconsumo ?? '');
  put('anni_ritorno_investimento',        payload.anni_ritorno_investimento ?? '');
  put('percentuale_risparmio_energetico', payload.percentuale_risparmio_energetico ?? '');
  put('produzione_primo_anno',            payload.produzione_primo_anno ?? '');
  put('media_vendita',                    payload.media_vendita ?? '');
  put('utile_25_anni',                    payload.utile_25_anni ?? '');
  put('incentivo_effettivo',              payload.incentivo_effettivo ?? '');
  put('massimale',                        payload.massimale ?? '');
  put('rata_mensile',                     payload.rata_mensile ?? '');
  put('chart_file_url',                   payload.chart_file_url || '');
  put('presentazione_doc_url',            payload.presentazione_doc_url || '');
  put('offerta_doc_url',                  payload.offerta_doc_url || '');
  put('data_version',                     'v3');

  sh.getRange(rowNum, 1, 1, row.length).setValues([row]);
  return { appIDoutput, rowNum };
}

/** Scrive la chiave di ponte su “offerte” */
function _writeOutputKeyToOfferte_(refIDofferta, appIDoutput) {
  const { sh, col } = _openByName_(CRMdatabase, 'offerte');
  const last = sh.getLastRow(); if (last <= 1) return;
  const vals = sh.getRange(2,1,last-1,sh.getLastColumn()).getValues();
  const iApp = col('appID') - 1;
  const iKey = col('offerte_outputID') - 1; // colonna presente nel foglio
  const idx = vals.findIndex(r => String(r[iApp]).trim() === String(refIDofferta).trim());
  if (idx >= 0) sh.getRange(idx+2, iKey+1).setValue(appIDoutput);
}

/* -------------------- Entry point: precomputeOffer -------------------- */
function precomputeOffer(
  appID,                 // string
  leadAppID,             // string
  cartella,              // URL/ID cartella cliente
  profiloDiConsumo,      // string
  kwhAnnui,              // number/string
  provincia,             // string
  esposizione,           // string
  prezzoEnergia,         // number/string
  indirizzo,             // string
  coordinate,            // string
  tilt,                  // number/string
  azimuth,               // number/string
  tipoModuli,            // string
  numeroModuli,          // number/string
  potenzaKWp,            // number/string
  modelloBatteria,       // string
  numeroBatterie,        // number/string
  accumulo,              // string
  prezzoOffertaTuttoIncluso, // number/string
  anniDurataIncentivo,   // number/string
  rid,                   // string
  tipoPagamento,         // string
  anniFinanziamento,     // number/string
  accontoDiretto         // number/string
) {
  'use strict';
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    if (!appID)     throw new Error('appID mancante');
    if (!leadAppID) throw new Error('leadAppID mancante');
    if (!cartella)  throw new Error('cartella mancante');

    // 🔹 NORMALIZZAZIONI PRIMA DEL FINGERPRINT (non toccano Drive/Sheets)
    const _kwhAnnui          = _toNumber_(kwhAnnui);
    const _prezzoEnergia     = _toEuroPerKWh_(prezzoEnergia);
    const _cartellaClienteId = _extractFolderId_(cartella);

    
    // 1) Calcolo fingerprint (prima di toccare Drive/Sheets)
    const fingerprint = _buildFingerprint_({
     tilt, azimuth, tipoModuli, numeroModuli, potenzaKWp,
     modelloBatteria, numeroBatterie, accumulo, prezzoOffertaTuttoIncluso,
     anniDurataIncentivo, rid, tipoPagamento, anniFinanziamento, accontoDiretto, coordinate,
     prezzoEnergia: _prezzoEnergia
    });


    // 1a) Guardia: stessa offerta rilanciata con stesso fingerprint e PNG presente → riuso (no nuova riga)
    const prev = _getOfferteOutputRowByOffer_(appID);
    if (prev && prev.calc_ready && prev.calc_fingerprint === fingerprint && prev.chart_file_url) {
      _writeOutputKeyToOfferte_(appID, prev.appIDoutput);
      _touchOfferteOutputRow_(prev.rowNum);
      Logger.log('[PRECOMPUTE] Reused previous row for appID=' + appID);
      return { ok:true, appID, offerte_outputID: prev.appIDoutput, chart_file_url: prev.chart_file_url, data_version:'v3', reused:true };
    }

    // 2) FAST PATH: clona ultima riga pronta con stesso fingerprint → nuova riga per questo appID
    const cloned = _fastCloneFromFingerprint_({ refIDofferta: appID, refIDlead: leadAppID, fingerprint });
    if (cloned) return cloned;

    // 3) STANDARD PATH: servono calcoli/grafico
    const folderProgettoId = newSubfolder(_cartellaClienteId, 'progetto');
    const chartFolderId    = newSubfolder(folderProgettoId, 'chart');

    // Apri/crea “dati tecnici” e logga la riga origine (usa la tua newLogDatiTecnici)
    let datiTecFile;
    const it = DriveApp.getFolderById(folderProgettoId).getFilesByName(nomeFileDatiTecnici);
    if (it.hasNext()) {
      datiTecFile = SpreadsheetApp.openById(it.next().getId());
    } else {
      const copy = DriveApp.getFileById(DATI_TECNICI_TEMPLATE_ID)
        .makeCopy(nomeFileDatiTecnici, DriveApp.getFolderById(folderProgettoId));
      datiTecFile = SpreadsheetApp.openById(copy.getId());
    }
    newLogDatiTecnici(datiTecFile, appID);

    // Calcoli (processDatiTecnici usa già la cache PVGIS via fetchWithCache_)
    const risultati = processDatiTecnici(
      datiTecFile,
      _kwhAnnui, String(profiloDiConsumo||'').trim(), String(provincia||'').trim(),
      String(esposizione||'').trim(), _prezzoEnergia, String(indirizzo||'').trim(),
      String(coordinate||'').trim(), appID /* (ok se passi anche folderProgettoId: l’extra arg non rompe) */
    );

    // Grafico: genera solo se non ne abbiamo già uno riusabile (prev eventuale)
    let chartUrl = prev?.chart_file_url || '';
    if (!chartUrl) {
      const mappings = newCharts(datiTecFile, folderProgettoId, CHART_DEFINITIONS);
      const chartId  = (function pick(m){ try{ if (m?.RITORNO_25_ANNI?.fileId) return m.RITORNO_25_ANNI.fileId; }catch(_){} const ks=Object.keys(m||{}); return ks.length ? (m[ks[0]]?.fileId||'') : ''; })(mappings);
      if (chartId) {
        const f  = DriveApp.getFileById(chartId);
        const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
        f.setName(`chart_ritorno_25_anni_${appID}_${ts}.png`);
        f.moveTo(DriveApp.getFolderById(chartFolderId));
        chartUrl = f.getUrl();
      }
    }

    // Persistenza risultati
    const up = _upsertOfferteOutputV3_({
      refIDofferta: appID,
      refIDlead:    leadAppID,
      calc_fingerprint: fingerprint,
      calc_ready: true,
      last_calc_ts: new Date(),
      percentuale_autoconsumo:          risultati.percentuale_autoconsumo,
      anni_ritorno_investimento:        risultati.anni_ritorno_investimento,
      percentuale_risparmio_energetico: risultati.percentuale_risparmio_energetico,
      produzione_primo_anno:            risultati.produzione_primo_anno,
      media_vendita:                    risultati.media_vendita,
      utile_25_anni:                    risultati.utile_25_anni,
      incentivo_effettivo:              risultati.incentivo_effettivo,
      massimale:                        risultati.massimale,
      rata_mensile:                     risultati.rata_mensile,
      chart_file_url:                   chartUrl
    });

    _writeOutputKeyToOfferte_(appID, up.appIDoutput);
    return { ok:true, appID, offerte_outputID: up.appIDoutput, chart_file_url: chartUrl, data_version:'v3' };

  } finally {
    lock.releaseLock();
  }
}
