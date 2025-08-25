// versione 2.1 (ordinata + FAST_MODE)

//nome del file dati tecnici creato per il cliente
const nomeFileDatiTecnici = 'dati tecnici 6.7'; // MAIN
// const nomeFileDatiTecnici = 'dati tecnici 6.0 test'; // TEST

// ID del template di base del foglio "dati tecnici"
const DATI_TECNICI_TEMPLATE_ID = '1cmiPctIhVJu3xA8jqnOYO0eXSBBmWkTBTUkEko1GLFs'; // MAIN
// const DATI_TECNICI_TEMPLATE_ID = '1MQ4GH3HxhEpJzg9qNw5CGfJhaSS_dKcdJBCADaQUz_k'; // TEST

// ID del database CRM
const CRMdatabase = SpreadsheetApp.openById('1_QEo5ynx_29j3I3uJJff5g7ZzGZJnPcIarIXfr5O2gQ'); // MAIN
// const CRMdatabase = SpreadsheetApp.openById('1HRxyXQ1xwlGk25sRNfrcZ3DXEEXzPpjvTnAKBBliKyg'); // TEST

// Nome del foglio offerte nel database CRM
const sheetOfferte = CRMdatabase.getSheetByName('offerte');

// Nome del foglio cronologia nel database CRM
const sheetCronologia = CRMdatabase.getSheetByName('cronologia');

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
  profilo_di_consumo, provincia, prezzo_energia, numero_rate_mensili, anni_finanziamento_dup, // (dup tenuto per compatibilità)
  tipo_incentivo, durata_incentivo, coordinate, ragione_sociale, garanzia_moduli, garanzia_inverter, garanzia_batterie,
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
  const fastCtx = loadFastModeFromOfferteOutput_(appID); // {FAST_MODE, datiTecniciData, chartFileIdFast}
  const FAST_MODE = !!fastCtx.FAST_MODE;
  let datiTecniciData = fastCtx.datiTecniciData || null;   // oggetto usato per i placeholder
  let chartMappings   = null;                               // solo se fallback
  const chartFileIdFast = fastCtx.chartFileIdFast || '';    // PNG già pronto (se FAST)

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
  } else {
    Logger.log('FAST_MODE attivo: uso valori precomputati da offerte_output e PNG del grafico.');
  }

  // -------------------------------------------------------------------
  // 4) CREAZIONE/COMPILAZIONE DOCUMENTI (comune ai due rami)
  // -------------------------------------------------------------------
  const createdDocIds = [];

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
      prezzo_energia, numero_rate_mensili, anni_finanziamento_dup, tipo_incentivo, durata_incentivo,
      indirizzo, ragione_sociale, garanzia_moduli, garanzia_inverter, garanzia_batterie,
      acconto_diretto, tilt, azimuth, nome_incentivo, descrizione_offerta,

      // risultati calcolati (fast o calcolati ora)
      percentuale_risparmio_energetico: datiTecniciData.percentuale_risparmio_energetico,
      anni_ritorno_investimento:       datiTecniciData.anni_ritorno_investimento,
      utile_25_anni:                   datiTecniciData.utile_25_anni,
      percentuale_autoconsumo:         datiTecniciData.percentuale_autoconsumo,
      media_vendita:                   datiTecniciData.media_vendita,
      massimale:                       datiTecniciData.massimale,
      rata_mensile:                    datiTecniciData.rata_mensile,
      produzione_primo_anno:           datiTecniciData.produzione_primo_anno,
      incentivo_effettivo:             datiTecniciData.incentivo_effettivo,
      detrazione:                      datiTecniciData.incentivo_effettivo // alias legacy
    });

    Logger.log('Placeholder Mapping pronto');

    // 4.3) Sostituzione placeholder testuali
    replacePlaceholders(corpo, mappaturaSegnapostov2);

    // 4.4) Aggiungi hyperlink alle schede tecniche (come prima)
    addHyperlink(corpo, 'Link scheda tecnica moduli',       scheda_tecnica_moduli);
    addHyperlink(corpo, 'Link scheda tecnica inverter',     scheda_tecnica_inverter);
    addHyperlink(corpo, 'Link scheda tecnica batterie',     scheda_tecnica_batterie);
    addHyperlink(corpo, 'Link scheda tecnica ottimizzatori',scheda_tecnica_ottimizzatori);

    // 4.5) Grafici:
    if (FAST_MODE && fastCtx.chartFileIdFast) {
      // Inserisce PNG già pronto al posto del placeholder (es. {{CHART_RITORNO_25_ANNI}})
      insertImageByFileId(doc, '{{CHART_RITORNO_25_ANNI}}', fastCtx.chartFileIdFast, 430);
    } else {
      // Fallback: genera e inserisce grafici come prima
      insertCharts(doc, chartMappings);
    }

    // 4.6) Salva e traccia ID
    doc.saveAndClose();
    createdDocIds.push(doc.getId());
    Logger.log('Documento generato: ' + dato.nomeFile);
  });

  Logger.log('Fine esecuzione funzione stampaOffertaV2');
}

/**
 * Carica il contesto FAST_MODE da `offerte_output`.
 * Ritorna:
 *  {
 *    FAST_MODE: boolean,
 *    datiTecniciData: {...} // valori già pronti per il mapping
 *    chartFileIdFast: string // fileId PNG grafico
 *  }
 */
function loadFastModeFromOfferteOutput_(offerId) {
  try {
    const sh = CRMdatabase.getSheetByName('offerte_output');
    if (!sh) return { FAST_MODE: false };

    const head = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(h => String(h).trim());
    const idx = {}; head.forEach((h,i) => idx[h] = i);

    const iKey = idx['refIDofferta'];
    if (iKey == null) return { FAST_MODE: false };

    const last = sh.getLastRow();
    if (last <= 1) return { FAST_MODE: false };

    const vals = sh.getRange(2,1,last-1,sh.getLastColumn()).getValues();
    for (let r=0;r<vals.length;r++) {
      if (String(vals[r][iKey]).trim() === String(offerId).trim()) {
        const get = n => {
          const i = idx[n];
          return i == null ? '' : vals[r][i];
        };
        const ready = String(get('calc_ready')).toLowerCase() === 'true';
        if (!ready) return { FAST_MODE: false };

        // Valori precomputati
        const data = {
          percentuale_autoconsumo:         get('percentuale_autoconsumo'),
          anni_ritorno_investimento:       get('anni_ritorno_investimento'),
          percentuale_risparmio_energetico:get('percentuale_risparmio_energetico'),
          utile_25_anni:                   get('utile_25_anni'),
          incentivo_effettivo:             get('incentivo_effettivo'),
          massimale:                       get('massimale'),
          rata_mensile:                    get('rata_mensile'),
          produzione_primo_anno:           get('produzione_primo_anno'),
          media_vendita:                   get('media_vendita')
        };

        return {
          FAST_MODE: true,
          datiTecniciData: data,
          chartFileIdFast: get('chart_file_id') || ''
        };
      }
    }
    return { FAST_MODE: false };
  } catch (e) {
    Logger.log('loadFastModeFromOfferteOutput_ error: ' + e);
    return { FAST_MODE: false };
  }
}
