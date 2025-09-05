// ===== Helpers offerte_output v3 =====

function _uuid_() {
  return Utilities.getUuid();
}

function _ensureHeadersExact_(sh, headers) {
  sh.getRange(1,1,1,headers.length).setValues([headers]);
  if (sh.getLastColumn() > headers.length) {
    const extra = sh.getLastColumn() - headers.length;
    sh.deleteColumns(headers.length+1, extra);
  }
}

function _idx_(sh) {
  const head = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(h => String(h).trim());
  const map = {}; head.forEach((h,i)=> map[h]=i);
  return { head, map, col: n => (n in map ? map[n]+1 : 0) };
}

/**
 * Upsert offerte_output v3. Ritorna { appIDoutput, rowNum }.
 * Se non esiste riga per refIDofferta, ne crea una nuova con appIDoutput=UUID.
 */
function _upsertOfferteOutputV3_(payload) {
  const sh = CRMdatabase.getSheetByName('offerte_output') || CRMdatabase.insertSheet('offerte_output');

  const HEADERS = [
    'appIDoutput','refIDofferta','refIDlead','calc_fingerprint','calc_ready','last_calc_ts',
    'percentuale_autoconsumo','anni_ritorno_investimento','percentuale_risparmio_energetico',
    'produzione_primo_anno','media_vendita','utile_25_anni','incentivo_effettivo','massimale','rata_mensile',
    'chart_file_url','presentazione_doc_url','offerta_doc_url','print_ts','data_version'
  ];
  _ensureHeadersExact_(sh, HEADERS);
  const { head, map, col } = _idx_(sh);

  // trova riga per refIDofferta
  const last = sh.getLastRow();
  let rowNum = 0, appIDoutput = '';
  if (last > 1) {
    const iRef = col('refIDofferta') - 1;
    const vals = sh.getRange(2,1,last-1,sh.getLastColumn()).getValues();
    const idx = vals.findIndex(r => String(r[iRef]).trim() === String(payload.refIDofferta).trim());
    if (idx >= 0) {
      rowNum = idx + 2;
      appIDoutput = String(vals[idx][col('appIDoutput')-1] || '');
    }
  }
  if (!rowNum) {
    rowNum = last + 1;
    appIDoutput = _uuid_();
    sh.getRange(rowNum, col('appIDoutput')).setValue(appIDoutput);
    sh.getRange(rowNum, col('refIDofferta')).setValue(payload.refIDofferta);
  }

  // scrittura campi (non sovrascrivo gli URL doc se payload non li porta)
  const toSet = {
    refIDlead: payload.refIDlead || '',
    calc_fingerprint: payload.calc_fingerprint || '',
    calc_ready: payload.calc_ready === true,
    last_calc_ts: payload.last_calc_ts || new Date(),
    percentuale_autoconsumo: payload.percentuale_autoconsumo ?? '',
    anni_ritorno_investimento: payload.anni_ritorno_investimento ?? '',
    percentuale_risparmio_energetico: payload.percentuale_risparmio_energetico ?? '',
    produzione_primo_anno: payload.produzione_primo_anno ?? '',
    media_vendita: payload.media_vendita ?? '',
    utile_25_anni: payload.utile_25_anni ?? '',
    incentivo_effettivo: payload.incentivo_effettivo ?? '',
    massimale: payload.massimale ?? '',
    rata_mensile: payload.rata_mensile ?? '',
    chart_file_url: payload.chart_file_url || '', // ok sovrascrivere
    data_version: 'v3'
  };

  Object.keys(toSet).forEach(k => {
    const c = col(k); if (c) sh.getRange(rowNum, c).setValue(toSet[k]);
  });

  return { appIDoutput, rowNum };
}

/** Scrive la chiave appIDoutput nella tab offerte (colonna "offerte_outputID") */
function _writeOutputKeyToOfferte_(refIDofferta, appIDoutput) {
  const { sh, col } = _openByName_(CRMdatabase, 'offerte');
  const last = sh.getLastRow(); if (last <= 1) return;
  const vals = sh.getRange(2,1,last-1,sh.getLastColumn()).getValues();
  const iApp = col('appID') - 1;
  const iKey = col('offerte_outputID') - 1;
  const idx = vals.findIndex(r => String(r[iApp]).trim() === String(refIDofferta).trim());
  if (idx >= 0) {
    sh.getRange(idx+2, iKey+1).setValue(appIDoutput);
  }
}
