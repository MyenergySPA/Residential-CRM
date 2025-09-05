// versione 2.2 — main.js (completo, ordinato)
// - Mantiene il flusso originale
// - Aggiunge FAST_MODE (usa offerte_output)
// - Inserisce grafico da PNG se disponibile
// - Nomina duplicati come "(1)", "(2)", ...
// - Registra sempre gli ID/URL documenti su offerte_output (recordDocsInOutput_)
// Dipendenze esterne (già nel progetto):
//   newSubfolder, determineDocumentTemplates, createPlaceholderMapping,
//   replacePlaceholders, addHyperlink, processDatiTecnici, newLogDatiTecnici,
//   newCharts, insertCharts, TEMPLATES, CRMdatabase

// ======= COSTANTI =======
// nome del file dati tecnici creato per il cliente
const nomeFileDatiTecnici = 'dati tecnici 6.7'; // MAIN
// const nomeFileDatiTecnici = 'dati tecnici 6.0 test'; // TEST

// ID del template di base del foglio "dati tecnici"
const DATI_TECNICI_TEMPLATE_ID = '1cmiPctIhVJu3xA8jqnOYO0eXSBBmWkTBTUkEko1GLFs'; // MAIN
// const DATI_TECNICI_TEMPLATE_ID = '1MQ4GH3HxhEpJzg9qNw5CGfJhaSS_dKcdJBCADaQUz_k'; // TEST

// ID del database CRM
const CRMdatabase = SpreadsheetApp.openById('1_QEo5ynx_29j3I3uJJff5g7ZzGZJnPcIarIXfr5O2gQ'); // MAIN
// const CRMdatabase = SpreadsheetApp.openById('1HRxyXQ1xwlGk25sRNfrcZ3DXEEXzPpjvTnAKBBliKyg'); // TEST

// Nome dei fogli chiave (se servono altrove)
const sheetOfferte = CRMdatabase.getSheetByName('offerte');
const sheetCronologia = CRMdatabase.getSheetByName('cronologia');

// punti tipografici (DocumentApp usa pt)
const CHART_TARGET_WIDTH_PT = 360; // <-- rimetti qui il valore del backup (quello che impaginava bene)

/**
 * Calcola il TAEG e l'Importo Totale Dovuto per un finanziamento a Tasso Zero.
 * Utilizza le funzioni di formattazione globali (es. formatPercentage).
 * Le spese fisse sono definite come costanti all'interno della funzione.
 *
 * @param {number} importoFinanziato - L'importo totale del finanziamento (es. il prezzo dell'offerta).
 * @param {number} anniFinanziamento - La durata del finanziamento in anni.
 * @returns {object} Un oggetto contenente 'taeg' (stringa formattata) e 'importoTotaleDovuto' (numero grezzo).
 */
function calcolaDatiFinanziamento(importoFinanziato, anniFinanziamento) {
  // --- Costi fissi del finanziamento ---
  const BOLLO_CONTRATTO = 16.00;
  const INCASSO_RATA = 3.00;
  const COMUNICAZIONE_ANNUALE = 1.20;
  const BOLLO_EC_ANNUALE = 2.00;

  // --- Validazione Input ---
  if (!importoFinanziato || importoFinanziato <= 0 || !anniFinanziamento || anniFinanziamento <= 0) {
    return { taeg: 'N/A', importoTotaleDovuto: 0 };
  }

  const numeroRate = anniFinanziamento * 12;

  // --- Calcolo Importo Totale Dovuto (restituisce il numero grezzo) ---
  const costoTotaleIncassoRate = INCASSO_RATA * numeroRate;
  const costoTotaleAnnuale = (COMUNICAZIONE_ANNUALE + BOLLO_EC_ANNUALE) * anniFinanziamento;
  const costoTotaleFinanziamento = BOLLO_CONTRATTO + costoTotaleIncassoRate + costoTotaleAnnuale;
  const importoTotaleDovuto = importoFinanziato + costoTotaleFinanziamento;

  // --- Calcolo TAEG ---
  const cashFlow = [];
  cashFlow.push(importoFinanziato - BOLLO_CONTRATTO);
  
  const rataCapitale = importoFinanziato / numeroRate;
  const esborsoMensileBase = -(rataCapitale + INCASSO_RATA);

  for (let i = 1; i <= numeroRate; i++) {
    let esborsoMeseCorrente = esborsoMensileBase;
    if (i % 12 === 0) {
      esborsoMeseCorrente -= (COMUNICAZIONE_ANNUALE + BOLLO_EC_ANNUALE);
    }
    cashFlow.push(esborsoMeseCorrente);
  }

  let mensileIRR = 0;
  try {
    let guess = 0.001;
    const MAX_ITERATIONS = 20;
    const PRECISION = 1e-7;

    for (let i = 0; i < MAX_ITERATIONS; i++) {
      let npv = 0;
      let derivative = 0;
      for (let j = 0; j < cashFlow.length; j++) {
        npv += cashFlow[j] / Math.pow(1 + guess, j);
        derivative += -j * cashFlow[j] / Math.pow(1 + guess, j + 1);
      }
      const newGuess = guess - npv / derivative;
      if (Math.abs(newGuess - guess) < PRECISION) {
        break;
      }
      guess = newGuess;
    }
    mensileIRR = guess;
  } catch (e) {
    mensileIRR = 0;
  }
  
  const taegAnnuale = mensileIRR * 12; // Questo è il valore decimale (es. 0.0057)

  return {
    // MODIFICATO: Utilizza la funzione formatPercentage esistente, passando il valore decimale.
    taeg: formatPercentage(taegAnnuale),
    importoTotaleDovuto: importoTotaleDovuto // Restituiamo il numero, la formattazione avverrà in createPlaceholderMapping
  };
}

/**
 * Stampa offerta (documenti) a partire dai parametri ricevuti da AppSheet/WebApp.
 * Con FAST_MODE: non ricalcola nulla, usa i dati precomputati in `offerte_output`
 * e inserisce il grafico (PNG) già pronto.
 *
 * Con fallback (FAST_MODE=false): comportamento identico alla versione precedente.
 */
function main(
  appID, tipo_opportunita, id, yy, nome_referente, cognome_referente, indirizzo, telefono, email,
  numero_moduli, numero_inverter, marca_moduli, marca_inverter, numero_batteria, capacita_batteria, totale_capacita_batterie, marca_batteria, tetto,
  potenza_impianto, alberi, testo_aggiuntivo, tipo_pagamento,
  condizione_pagamento_1, condizione_pagamento_2, condizione_pagamento_3, condizione_pagamento_4, imponibile_offerta,
  iva_offerta, iva_percentuale, prezzo_offerta, cartella, anni_finanziamento, esposizione, area_m2_impianto,
  numero_colonnina_74kw, numero_colonnina_22kw, numero_ottimizzatori, marca_ottimizzatori, numero_linea_vita,
  scheda_tecnica_moduli, scheda_tecnica_inverter, scheda_tecnica_batterie, scheda_tecnica_ottimizzatori, consumi_annui,
  profilo_di_consumo, provincia, prezzo_energia, tipo_incentivo, durata_incentivo, coordinate, ragione_sociale, garanzia_moduli, garanzia_inverter, garanzia_batterie,
  acconto_diretto, tilt, azimuth, nome_incentivo, descrizione_offerta
) {
  Logger.log('Avvio funzione stampaOffertaV2 per appID: ' + appID);

  // -------------------------------------------------------------------
  // 0) Data e cartelle base (come prima)
  // -------------------------------------------------------------------
  const oggi = new Date();
  const dataOggi = new Intl.DateTimeFormat('it-IT', { day: '2-digit', month: '2-digit', year: 'numeric' }).format(oggi);
  Logger.log('Data odierna: ' + dataOggi);

  // cartella principale del cliente (da URL → ID)
  const cartellaDestinazioneId = cartella.split('/folders/')[1];
  Logger.log('ID della cartella principale: ' + cartellaDestinazioneId);

  // sottocartelle: contratto/ e progetto/
  const cartellaContrattoId = newSubfolder(cartellaDestinazioneId, 'contratto');
  Logger.log('Cartella contratto ID: ' + cartellaContrattoId);

  const cartellaDataId = newSubfolder(cartellaContrattoId, dataOggi); // contratto/gg/mm/aaaa
  Logger.log('Cartella con la data ID: ' + cartellaDataId);

  const cartellaProgettoId = newSubfolder(cartellaDestinazioneId, 'progetto');
  Logger.log('ID cartella progetto: ' + cartellaProgettoId);

  //calcoli rate
  const anniFin = Number(anni_finanziamento || 0);
  const numero_rate_mensili = Math.round(anniFin * 12);

  // --- NUOVO: Calcolo TAEG e Importo Totale Dovuto ---
  let datiFinanziamento = { taeg: 'N/A', importoTotaleDovuto: 0 };
  // Esegui il calcolo solo se il pagamento è di tipo "Finanziamento" e i dati sono validi
  if (tipo_pagamento === 'Finanziamento' && anniFin > 0 && prezzo_offerta > 0) {
    datiFinanziamento = calcolaDatiFinanziamento(prezzo_offerta, anniFin);
    Logger.log('Calcolo finanziamento eseguito: ' + JSON.stringify(datiFinanziamento));
  }

  // -------------------------------------------------------------------
  // 1) Selezione template (immutata)
  // -------------------------------------------------------------------
  const datiDocumento = determineDocumentTemplates(
    tipo_opportunita, tipo_pagamento, tipo_incentivo, nome_referente, cognome_referente, dataOggi, id, yy
  );
  Logger.log('Dati documenti: ' + JSON.stringify(datiDocumento));

  // -------------------------------------------------------------------
  // 2) FAST-PATH: se `offerte_output` è pronto, saltiamo calcoli/grafici
  //    Altrimenti usiamo il flusso originale
  // -------------------------------------------------------------------
  const fastCtx = loadFastModeFromOfferteOutput_(appID,(typeof __FAST_OUTPUT_KEY !== 'undefined' ? __FAST_OUTPUT_KEY : '')); // {FAST_MODE, datiTecniciData, chartFileIdFast}
  const FAST_MODE = !!fastCtx.FAST_MODE;
  let datiTecniciData = fastCtx.datiTecniciData || null;   // oggetto usato per i placeholder
  let chartMappings   = null;                               // solo se fallback
  const chartFileIdFast = fastCtx.chartFileIdFast || '';

  if (!FAST_MODE) {
    // -----------------------------------------------------------------
    // 3) PREPARA DATI TECNICI & CALCOLI (comportamento originale)
    // -----------------------------------------------------------------

    // 3.1) Prepara/apri il file "dati tecnici" in progetto/
    let nuovoFileDatiTecnici;
    const fileDatiTecnici = DriveApp.getFolderById(cartellaProgettoId).getFilesByName(nomeFileDatiTecnici);
    if (fileDatiTecnici.hasNext()) {
      nuovoFileDatiTecnici = SpreadsheetApp.openById(fileDatiTecnici.next().getId());
      Logger.log('File dati tecnici esistente trovato e aperto: ' + nuovoFileDatiTecnici.getId());
    } else {
      const modelloDatiTecnici = DriveApp.getFileById(DATI_TECNICI_TEMPLATE_ID)
        .makeCopy(nomeFileDatiTecnici, DriveApp.getFolderById(cartellaProgettoId));
      nuovoFileDatiTecnici = SpreadsheetApp.openById(modelloDatiTecnici.getId());
      Logger.log('Nuovo file dati tecnici creato: ' + nuovoFileDatiTecnici.getId());
    }

    // 3.2) Scrivi riga in "log" di dati tecnici (come prima)
    newLogDatiTecnici(nuovoFileDatiTecnici, appID);

    // 3.3) Esegui calcoli + PVGIS (come prima)
    datiTecniciData = processDatiTecnici(
      nuovoFileDatiTecnici,
      consumi_annui, profilo_di_consumo, provincia, esposizione, prezzo_energia,
      indirizzo, coordinate, appID, cartellaProgettoId
    );

    // 3.4) Genera e carica grafici (come prima)
    chartMappings = newCharts(nuovoFileDatiTecnici, cartellaProgettoId, CHART_DEFINITIONS);

    // 3.5) Persisti anche nello sheet offerte_output (così da ora FAST_MODE sarà disponibile)
    try {
  const chartFileIdForTrace = _pickChartFileIdFromMappings_(chartMappings);
  const chartFileUrlForTrace = chartFileIdForTrace ? DriveApp.getFileById(chartFileIdForTrace).getUrl() : '';
  // upsert v3: solo colonne nuove
  const sh = CRMdatabase.getSheetByName('offerte_output');
  if (sh) {
    const HEADERS = [
      'appIDoutput','refIDofferta','refIDlead','calc_fingerprint','calc_ready','last_calc_ts',
      'percentuale_autoconsumo','anni_ritorno_investimento','percentuale_risparmio_energetico',
      'produzione_primo_anno','media_vendita','utile_25_anni','incentivo_effettivo','massimale','rata_mensile',
      'chart_file_url','presentazione_doc_url','offerta_doc_url','print_ts','data_version'
    ];
    _ensureHeadersExact_(sh, HEADERS);

    const head = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(h=>String(h).trim());
    const col  = name => head.indexOf(name)+1;
    const last = sh.getLastRow();
    let rowNum = 0;
    if (last > 1) {
      const vals = sh.getRange(2,1,last-1,sh.getLastColumn()).getValues();
      const iRef = col('refIDofferta') - 1;
      const idx  = vals.findIndex(r => String(r[iRef]).trim() === String(appID).trim());
      if (idx >= 0) rowNum = idx + 2;
    }
    if (!rowNum) {
      rowNum = last + 1;
      sh.getRange(rowNum, col('appIDoutput')).setValue(Utilities.getUuid());
      sh.getRange(rowNum, col('refIDofferta')).setValue(appID);
    }
    sh.getRange(rowNum, col('calc_ready')).setValue(true);
    sh.getRange(rowNum, col('last_calc_ts')).setValue(new Date());
    sh.getRange(rowNum, col('percentuale_autoconsumo')).setValue(datiTecniciData.percentuale_autoconsumo);
    sh.getRange(rowNum, col('anni_ritorno_investimento')).setValue(datiTecniciData.anni_ritorno_investimento);
    sh.getRange(rowNum, col('percentuale_risparmio_energetico')).setValue(datiTecniciData.percentuale_risparmio_energetico);
    sh.getRange(rowNum, col('produzione_primo_anno')).setValue(datiTecniciData.produzione_primo_anno);
    sh.getRange(rowNum, col('media_vendita')).setValue(datiTecniciData.media_vendita);
    sh.getRange(rowNum, col('utile_25_anni')).setValue(datiTecniciData.utile_25_anni);
    sh.getRange(rowNum, col('incentivo_effettivo')).setValue(datiTecniciData.incentivo_effettivo);
    sh.getRange(rowNum, col('massimale')).setValue(datiTecniciData.massimale);
    sh.getRange(rowNum, col('rata_mensile')).setValue(datiTecniciData.rata_mensile);
    sh.getRange(rowNum, col('chart_file_url')).setValue(chartFileUrlForTrace);
    if (head.indexOf('data_version') >= 0) sh.getRange(rowNum, col('data_version')).setValue('v3');
  }
} catch (e) {
  Logger.log('offerte_output upsert (fallback) error: ' + e);
}
  }

  // -------------------------------------------------------------------
  // 4) CREAZIONE/COMPILAZIONE DOCUMENTI (comune ai due rami)
  // -------------------------------------------------------------------
  const createdDocIds = [];
  let presentazioneDocId = '';
  let offertaDocId = '';

  datiDocumento.forEach(dato => {
    // 4.1) Apri il doc: prova a prendere un PRE coerente, altrimenti copia il template
    const doc = openDocForWrite_FromPlaceholderOrTemplate(
      cartellaContrattoId,       // esiste
      cartellaDataId,            // contratto/gg-mm-aaaa
      dato.nomeFile,             // baseName
      dato.templateId,           // template scelto
      id, yy
    );
    const corpo = doc.getBody();

    // 4.2) Mapping dei segnaposto (usa i datiTecniciData dal FAST_MODE o dal calcolo)
    const mappaturaSegnapostov2 = createPlaceholderMapping({
      // anagrafica & offerta
      tipo_opportunita, id, yy, nome_referente, cognome_referente, indirizzo, telefono, email, dataOggi,
      numero_moduli, marca_moduli, numero_inverter, marca_inverter, numero_batteria, capacita_batteria,
      totale_capacita_batterie, marca_batteria, tetto, potenza_impianto, alberi, testo_aggiuntivo, tipo_pagamento,
      condizione_pagamento_1, condizione_pagamento_2, condizione_pagamento_3, condizione_pagamento_4,
      imponibile_offerta, iva_offerta, iva_percentuale, prezzo_offerta,
      anni_finanziamento, esposizione, area_m2_impianto,
      scheda_tecnica_moduli, scheda_tecnica_inverter, scheda_tecnica_batterie, scheda_tecnica_ottimizzatori,
      numero_colonnina_74kw, numero_colonnina_22kw, numero_ottimizzatori, marca_ottimizzatori, numero_linea_vita,
      prezzo_energia, numero_rate_mensili, 
      anni_finanziamento_dup: anniFin, 
      tipo_incentivo, durata_incentivo,
      ragione_sociale, garanzia_moduli, garanzia_inverter, garanzia_batterie,
      acconto_diretto, tilt, azimuth, nome_incentivo, descrizione_offerta,

      // risultati calcolati (fast o calcolati ora)
      percentuale_risparmio_energetico: datiTecniciData.percentuale_risparmio_energetico,
      anni_ritorno_investimento:        datiTecniciData.anni_ritorno_investimento,
      utile_25_anni:                    datiTecniciData.utile_25_anni,
      percentuale_autoconsumo:          datiTecniciData.percentuale_autoconsumo,
      media_vendita:                    datiTecniciData.media_vendita,
      massimale:                        datiTecniciData.massimale,
      rata_mensile:                     datiTecniciData.rata_mensile,
      produzione_primo_anno:            datiTecniciData.produzione_primo_anno,
      incentivo_effettivo:              datiTecniciData.incentivo_effettivo,
      detrazione:                       datiTecniciData.incentivo_effettivo, // alias legacy
      taeg:                             datiFinanziamento.taeg,
      importo_totale_dovuto:            datiFinanziamento.importoTotaleDovuto
    });

    // 4.3) Sostituzione placeholder testuali
    replacePlaceholders(corpo, mappaturaSegnapostov2);

    // 4.4) Aggiungi hyperlink alle schede tecniche
    addHyperlink(corpo, 'Link scheda tecnica moduli',        scheda_tecnica_moduli);
    addHyperlink(corpo, 'Link scheda tecnica inverter',      scheda_tecnica_inverter);
    addHyperlink(corpo, 'Link scheda tecnica batterie',      scheda_tecnica_batterie);
    addHyperlink(corpo, 'Link scheda tecnica ottimizzatori', scheda_tecnica_ottimizzatori);

    // 4.5) Grafici
    if (FAST_MODE && chartFileIdFast && typeof insertImageByFileId === 'function') {
      // Inserisce PNG già pronto al posto del placeholder (es. {{CHART_RITORNO_25_ANNI}})
      insertImageByFileId(doc, '{{CHART_RITORNO_25_ANNI}}', chartFileIdFast, CHART_TARGET_WIDTH_PT);
    } else {
      // Fallback: genera e inserisce grafici come prima
      insertCharts(doc, chartMappings);
    }

    // 4.6) Salva, traccia ID e tipo documento per offerte_output
    doc.saveAndClose();
    createdDocIds.push(doc.getId());

    const docType = _docTypeFromTemplateId_(dato.templateId);
    if (docType.indexOf('PRESENTAZIONE') === 0) presentazioneDocId = doc.getId();
    if (docType.indexOf('OFFERTA') === 0)       offertaDocId       = doc.getId();

    Logger.log('Documento generato: ' + dato.nomeFile);
  });

  // 4.7) Registra gli ID/URL dei documenti stampati su offerte_output
  try {
    if (typeof recordDocsInOutput_ === 'function') {
      recordDocsInOutput_(appID, presentazioneDocId, offertaDocId);
    }
  } catch (e) {
    Logger.log('recordDocsInOutput_ error: ' + e);
  }

  Logger.log('Fine esecuzione funzione stampaOffertaV2');
}

// ======= HELPERS LOCALI =======

/**
 * Carica il contesto FAST_MODE da `offerte_output`.
 * Ritorna: { FAST_MODE, datiTecniciData: {...}, chartFileIdFast }
 */

function _getFastKeyFromCache_(offerId) {
  try { return CacheService.getScriptCache().get('fastkey_' + offerId) || ''; } catch(e){ return ''; }
}

function _toBoolFast_(v) {
  if (v === true) return true;
  const s = String(v||'').trim().toLowerCase();
  return s === 'true' || s === 'vero' || s === '1' || s === 'yes' || s === 'y';
}
function _tsFast_(v) {
  if (v instanceof Date) return v.getTime();
  const s = String(v||'').trim(); if (!s) return 0;
  const d = new Date(s); return isFinite(d) ? d.getTime() : 0;
}
function _extractIdFromUrl_(url) {
  const m = String(url||'').match(/[-\w]{25,}/);
  return m ? m[0] : '';
}

function loadFastModeFromOfferteOutput_(offerId, outputKeyOpt) {
  // se non passa esplicitamente, prova da cache impostata dal router
  if (!outputKeyOpt) outputKeyOpt = _getFastKeyFromCache_(offerId);
  try {
    const sh = CRMdatabase.getSheetByName('offerte_output');
    if (!sh) return { FAST_MODE:false, reason:'no_sheet' };

    const head = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(h => String(h).trim());
    const idx = {}; head.forEach((h,i)=>{ if (idx[h] == null) idx[h] = i; }); // prima occorrenza

    const iOffer = idx['refIDofferta'];
    const iKey   = idx['appIDoutput'];
    const iReady = idx['calc_ready'];
    const iTs    = idx['last_calc_ts'];

    if (iOffer == null) return { FAST_MODE:false, reason:'no_refIDofferta_col' };

    const last = sh.getLastRow();
    if (last <= 1) return { FAST_MODE:false, reason:'empty_sheet' };

    const vals = sh.getRange(2,1,last-1,sh.getLastColumn()).getValues();

    // canditati per offerId o per chiave
    const cand = [];
    for (let r=0;r<vals.length;r++) {
      const row = vals[r];
      const okByKey   = outputKeyOpt && iKey!=null && String(row[iKey]).trim() === String(outputKeyOpt).trim();
      const okByOffer = String(row[iOffer]).trim() === String(offerId).trim();
      if (okByKey || okByOffer) cand.push({ r, row });
    }
    if (!cand.length) return { FAST_MODE:false, reason:'no_matching_row' };

    // scegli la più recente, preferendo quelle ready=true
    cand.sort((a,b)=> _tsFast_(b.row[iTs]) - _tsFast_(a.row[iTs]));
    const pick = cand.find(x => _toBoolFast_(x.row[iReady])) || cand[0];

    const get = name => {
      const i = idx[name];
      return i == null ? '' : pick.row[i];
    };

    const ready = _toBoolFast_(get('calc_ready'));
    Logger.log('[FAST] candidates=%s | pickedRow=%s | ready=%s | ts=%s',
      cand.length, (pick.r+2), ready, get('last_calc_ts'));

    if (!ready) return { FAST_MODE:false, reason:'latest_not_ready' };

    const data = {
      percentuale_autoconsumo:          get('percentuale_autoconsumo'),
      anni_ritorno_investimento:        get('anni_ritorno_investimento'),
      percentuale_risparmio_energetico: get('percentuale_risparmio_energetico'),
      produzione_primo_anno:            get('produzione_primo_anno'),
      utile_25_anni:                    get('utile_25_anni'),
      incentivo_effettivo:              get('incentivo_effettivo'),
      massimale:                        get('massimale'),
      rata_mensile:                     get('rata_mensile'),
      media_vendita:                    get('media_vendita')
    };

    const chartUrl = get('chart_file_url') || '';
    const chartId  = _extractIdFromUrl_(chartUrl);

    return { FAST_MODE:true, datiTecniciData: data, chartFileIdFast: chartId };
  } catch (e) {
    Logger.log('loadFastModeFromOfferteOutput_ error: ' + e);
    return { FAST_MODE:false, reason:'exception' };
  }
}

/**
 * Estrae un fileId da chartMappings (per il tracciamento su offerte_output).
 */
function _pickChartFileIdFromMappings_(m) {
  try {
    if (!m) return '';
    // Se usi una chiave fissa per il grafico principale, mettila qui:
    if (m.RITORNO_25_ANNI && m.RITORNO_25_ANNI.fileId) return m.RITORNO_25_ANNI.fileId;
    const keys = Object.keys(m);
    if (!keys.length) return '';
    const k = keys[0];
    return (m[k] && m[k].fileId) ? m[k].fileId : '';
  } catch (e) { return ''; }
}

/**
 * Ritorna un Document “scrivibile” per la stampa del singolo file:
 * - garantisce la cartella data (contratto/gg/mm/aaaa),
 * - se esiste un PRE coerente in contratto/placeholder → lo sposta e rinomina,
 * - altrimenti copia dal template,
 * - assicura nome finale univoco: baseName, baseName (1), baseName (2)...
 */
function openDocForWrite_FromPlaceholderOrTemplate(contrattoFolderId, dateFolderId, baseName, templateId, id, yy) {
  const folderContratto   = DriveApp.getFolderById(contrattoFolderId);
  const folderData        = DriveApp.getFolderById(dateFolderId);
  const folderPlaceholder = DriveApp.getFolderById(newSubfolder(contrattoFolderId, 'placeholder'));

  // Nome finale UNICO nella cartella data (parti da baseName, poi (1), (2)...)
  const finalName = _uniqueNameInFolder_(folderData, baseName);

  // Individua docType dal templateId (stessa mappa usata in pre-gen)
  const docType = _docTypeFromTemplateId_(templateId);
  const preName = `${docType} - ${String(id||'').trim()}-${String(yy||'').trim()} PRE`;

  // 1) se c'è un PRE coerente -> move + rename
  const preFile = _findFileInFolderByName_(folderPlaceholder.getId(), preName);
  if (preFile) {
    preFile.moveTo(folderData);   // sposta (rimuove dagli altri parent)
    preFile.setName(finalName);   // rinomina PRE -> finale
    return DocumentApp.openById(preFile.getId());
  }

  // 2) altrimenti copia dal template
  const copy = DriveApp.getFileById(templateId).makeCopy(finalName, folderData);
  return DocumentApp.openById(copy.getId());
}

function _docTypeFromTemplateId_(templateId) {
  // allinea ai tuoi TEMPLATES in docTemplates.js
  if (templateId === TEMPLATES.offerta_MAT)           return 'OFFERTA-MAT';
  if (templateId === TEMPLATES.offerta_Finanz)        return 'OFFERTA-FIN';
  if (templateId === TEMPLATES.offerta)               return 'OFFERTA';
  if (templateId === TEMPLATES.presentazione_Finanz)  return 'PRESENTAZIONE-FIN';
  if (templateId === TEMPLATES.presentazione_COND)    return 'PRESENTAZIONE-COND';
  if (templateId === TEMPLATES.presentazione)         return 'PRESENTAZIONE';
  // fallback
  return 'OFFERTA';
}

function _uniqueNameInFolder_(folder, baseName) {
  // Se baseName non esiste -> lo uso così.
  let candidate = baseName;
  if (!_findFileInFolderByName_(folder.getId(), candidate)) return candidate;
  // Prima variante: baseName (1), poi (2), ...
  let n = 1;
  while (_findFileInFolderByName_(folder.getId(), `${baseName} (${n})`)) {
    n++;
  }
  return `${baseName} (${n})`;
}

function _findFileInFolderByName_(folderId, name) {
  const it = DriveApp.getFolderById(folderId).getFilesByName(name);
  return it.hasNext() ? it.next() : null;
}

function recordDocsInOutput_(refIDofferta, presentazioneDocId, offertaDocId) {
  const sh = CRMdatabase.getSheetByName('offerte_output');
  if (!sh) throw new Error('Foglio "offerte_output" mancante');

  const HEADERS = [
    'appIDoutput','refIDofferta','refIDlead','calc_fingerprint','calc_ready','last_calc_ts',
    'percentuale_autoconsumo','anni_ritorno_investimento','percentuale_risparmio_energetico',
    'produzione_primo_anno','media_vendita','utile_25_anni','incentivo_effettivo','massimale','rata_mensile',
    'chart_file_url','presentazione_doc_url','offerta_doc_url','print_ts','data_version'
  ];
  _ensureHeadersExact_(sh, HEADERS);

  const head = sh.getRange(1,1,1,HEADERS.length).getValues()[0].map(h => String(h).trim());
  const col  = name => {
    const i = head.indexOf(name);
    if (i < 0) throw new Error('Colonna mancante in offerte_output: ' + name);
    return i + 1;
  };

  const last = sh.getLastRow();
  let rowNum = 0, appIDoutput = '';
  if (last > 1) {
    const vals = sh.getRange(2,1,last-1,HEADERS.length).getValues();
    const iRef = col('refIDofferta') - 1;
    const idx  = vals.findIndex(r => String(r[iRef]).trim() === String(refIDofferta).trim());
    if (idx >= 0) {
      rowNum = idx + 2;
      const iKey = col('appIDoutput') - 1;
      appIDoutput = String(vals[idx][iKey] || '');
    }
  }
  if (!rowNum) {
    rowNum = last + 1;
    appIDoutput = _uuid_();
    sh.getRange(rowNum, col('appIDoutput')).setValue(appIDoutput);
    sh.getRange(rowNum, col('refIDofferta')).setValue(refIDofferta);
  }

  const presUrl    = presentazioneDocId ? DriveApp.getFileById(presentazioneDocId).getUrl() : '';
  const offertaUrl = offertaDocId       ? DriveApp.getFileById(offertaDocId).getUrl()       : '';

  if (presUrl)    sh.getRange(rowNum, col('presentazione_doc_url')).setValue(presUrl);
  if (offertaUrl) sh.getRange(rowNum, col('offerta_doc_url')).setValue(offertaUrl);
  sh.getRange(rowNum, col('print_ts')).setValue(new Date());
  sh.getRange(rowNum, col('data_version')).setValue('v3');
}

/** Restituisce lo Spreadsheet CRM.
 *  Preferisce la costante globale CRMdatabase (definita in main.js).
 *  Fallback: SpreadsheetApp.getActive() (container-bound).
 */
function getCRM_() {
  if (typeof CRMdatabase !== 'undefined' && CRMdatabase.getSheetByName) {
    return CRMdatabase;
  }
  const ss = SpreadsheetApp.getActive();
  if (ss) return ss;
  throw new Error('CRMdatabase non definito e nessun Spreadsheet attivo disponibile.');
}
