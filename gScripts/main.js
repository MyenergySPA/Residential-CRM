//nome del file dati tecnici creato per il cliente
const nomeFileDatiTecnici = 'dati tecnici 6.5'; // MAIN
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
 * Questo script automatizza la generazione di documenti di offerta basandosi su template predefiniti e dati forniti dall'utente.
 * Gestisce la creazione e l'organizzazione di cartelle su Google Drive, l'elaborazione di dati tecnici ed energetici, e la sostituzione
 * di segnaposto nei documenti finali. 
 */

function main(appID, tipo_opportunita, id, yy, nome_referente, cognome_referente, indirizzo, telefono, email, numero_moduli, numero_inverter, marca_moduli, 
                        marca_inverter, numero_batteria, capacita_batteria, totale_capacita_batterie, marca_batteria, tetto, 
                        potenza_impianto, alberi, testo_aggiuntivo, tipo_pagamento, 
                        condizione_pagamento_1, condizione_pagamento_2, condizione_pagamento_3, condizione_pagamento_4, imponibile_offerta,
                        iva_offerta, iva_percentuale, prezzo_offerta, cartella, anni_finanziamento, esposizione, area_m2_impianto, 
                        numero_colonnina_74kw, numero_colonnina_22kw, numero_ottimizzatori, marca_ottimizzatori, numero_linea_vita, 
                        scheda_tecnica_moduli, scheda_tecnica_inverter, scheda_tecnica_batterie, scheda_tecnica_ottimizzatori, consumi_annui, 
                        profilo_di_consumo, provincia, prezzo_energia, numero_rate_mensili, anni_finanziamento, tipo_incentivo,durata_incentivo, coordinate, ragione_sociale, garanzia_moduli, garanzia_inverter, garanzia_batterie, acconto_diretto, tilt, azimuth) {

  Logger.log('Avvio funzione stampaOffertaV2 per appID: ' + appID);


  // Creazione stringa di data nel formato 'dd/mm/yyyy'
  const oggi = new Date();
  const dataOggi = new Intl.DateTimeFormat('it-IT', { day: '2-digit', month: '2-digit', year: 'numeric' }).format(oggi);
  Logger.log('Data odierna: ' + dataOggi);


  // ottieni l'ID della cartella di destinazione del cliente
  const cartellaDestinazioneId = cartella.split('/folders/')[1];
  Logger.log('ID della cartella principale: ' + cartellaDestinazioneId);

  // Recupero o creazione della cartella "contratto"
  const cartellaContrattoId = newSubfolder  (cartellaDestinazioneId, 'contratto');
  Logger.log('Cartella contratto ID: ' + cartellaContrattoId);

  // Creazione della cartella con la data odierna
  const cartellaDataId = newSubfolder(cartellaContrattoId, dataOggi);
  Logger.log('Cartella con la data ID: ' + cartellaDataId);

  //creazione della cartella "progetto"
  const cartellaProgettoId = newSubfolder(cartellaDestinazioneId, 'progetto');
  Logger.log('ID cartella progetto: ' + cartellaProgettoId); 


  // Determina i template dei documenti da usare in base al tipo di opportunità e pagamento
  const datiDocumento = determineDocumentTemplates(tipo_opportunita, tipo_pagamento, tipo_incentivo, nome_referente, cognome_referente, dataOggi, id, yy);
  Logger.log('Dati documenti: ' + JSON.stringify(datiDocumento));


  // Prepara il file dati tecnici
  const fileDatiTecnici = DriveApp.getFolderById(cartellaProgettoId).getFilesByName(nomeFileDatiTecnici); 
  if (fileDatiTecnici.hasNext()) {
         nuovoFileDatiTecnici = SpreadsheetApp.openById(fileDatiTecnici.next().getId());
        Logger.log('File dati tecnici esistente trovato e aperto: ' + nuovoFileDatiTecnici.getId());
  } else {  
        var modelloDatiTecnici = DriveApp.getFileById(DATI_TECNICI_TEMPLATE_ID).makeCopy(nomeFileDatiTecnici, DriveApp.getFolderById(cartellaProgettoId));
         nuovoFileDatiTecnici = SpreadsheetApp.openById(modelloDatiTecnici.getId());
        Logger.log('Nuovo file dati tecnici creato: ' + nuovoFileDatiTecnici.getId());
  }


  //creazione nuova riga in sheet "log" in "dati tecnici", contenente info offerta
  newLogDatiTecnici(nuovoFileDatiTecnici, appID);


  // elaborazione dati e grafici su "dati tecnici" sheet
  const datiTecniciData = processDatiTecnici(nuovoFileDatiTecnici, consumi_annui, profilo_di_consumo, provincia, esposizione, prezzo_energia, indirizzo, coordinate, appID, cartellaProgettoId);


  // Upload charts using CHART_DEFINITIONS from newCharts.gs
  const chartMappings = newCharts(nuovoFileDatiTecnici, cartellaProgettoId, CHART_DEFINITIONS);


  // crea un array con tutti i doc d'offerta:
  const createdDocIds = [];
  datiDocumento.forEach(dato => {
    const doc = createDocumentFromTemplate(dato.templateId, cartellaDataId, dato.nomeFile);
    const corpo = doc.getBody();
    
  let mappaturaSegnapostov2 = createPlaceholderMapping({
      tipo_opportunita, id, yy, nome_referente, cognome_referente, indirizzo, telefono, email, dataOggi, numero_moduli, marca_moduli, numero_inverter, marca_inverter, numero_batteria, capacita_batteria, totale_capacita_batterie, marca_batteria, tetto, potenza_impianto, alberi, testo_aggiuntivo, tipo_pagamento, condizione_pagamento_1, condizione_pagamento_2, condizione_pagamento_3, condizione_pagamento_4, imponibile_offerta, iva_offerta, iva_percentuale, prezzo_offerta, anni_finanziamento, esposizione, area_m2_impianto, scheda_tecnica_moduli, scheda_tecnica_inverter, scheda_tecnica_batterie, scheda_tecnica_ottimizzatori, numero_colonnina_74kw, numero_colonnina_22kw, numero_ottimizzatori, marca_ottimizzatori, numero_linea_vita, prezzo_energia, numero_rate_mensili,  anni_finanziamento, tipo_incentivo, durata_incentivo, indirizzo, ragione_sociale, garanzia_moduli, garanzia_inverter, garanzia_batterie, acconto_diretto, tilt, azimuth,
        
        percentuale_risparmio_energetico: datiTecniciData.percentuale_risparmio_energetico,
        anni_ritorno_investimento: datiTecniciData.anni_ritorno_investimento,
        utile_25_anni: datiTecniciData.utile_25_anni,
        percentuale_autoconsumo: datiTecniciData.percentuale_autoconsumo,
        media_vendita: datiTecniciData.media_vendita,
        detrazione: datiTecniciData.detrazione,
        massimale: datiTecniciData.massimale,
        rata_mensile: datiTecniciData.rata_mensile,
        produzione_primo_anno: datiTecniciData.produzione_primo_anno,

    });
    
    Logger.log('Placeholder Mapping:', mappaturaSegnapostov2);

    // Replace placeholders in the document
    replacePlaceholders(corpo, mappaturaSegnapostov2);


    //add hyperlinks to "schede tecniche"
    addHyperlink(corpo, 'Link scheda tecnica moduli', scheda_tecnica_moduli);
    addHyperlink(corpo, 'Link scheda tecnica inverter', scheda_tecnica_inverter);
    addHyperlink(corpo, 'Link scheda tecnica batterie', scheda_tecnica_batterie);
    addHyperlink(corpo, 'Link scheda tecnica ottimizzatori', scheda_tecnica_ottimizzatori);


    // replace chart placeholders with images
    insertCharts(doc, chartMappings);


    // Save and close the document
    doc.saveAndClose();


    // Keep track of the doc ID for the chart insertion step
    createdDocIds.push(doc.getId());

    Logger.log('Documento generato: ' + dato.nomeFile);
  });
 
  Logger.log('Fine esecuzione funzione stampaOffertaV2');

}