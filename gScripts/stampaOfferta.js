/**
 * ==============================================
 *        SCRIPT DI GENERAZIONE OFFERTA
 * ==============================================
 *
 * Questo script automatizza la generazione di documenti di offerta
 * basandosi su template predefiniti e dati forniti dall'utente.
 * Gestisce la creazione e l'organizzazione di cartelle su Google Drive,
 * l'elaborazione di dati tecnici ed energetici, e la sostituzione
 * di segnaposto nei documenti finali.
 *
 * Autore: Luca
 * Link GPT: https://chatgpt.com/c/33d84469-fd54-4ff4-a0e5-0a3a947a185a
 * Data: 22-08-24
 * Versione: 5.2
 */


/** ===============================
 *       COSTANTI GLOBALI
 *  ===============================
 */

// Definizione globale di TEMPLATES
const TEMPLATES = {
  presentazioneFinanz: '1zMIjekT-K_JWssZidSBjSuog_LfcHZjMLcEbePnP_t8',
  offertaMateriale: '1gMJGZZA7LwdugXKEFTK5LbJU2iiIIs6Ee5zBnlW81es',
  presentazione: '1XYDLbJymoNqU8B1nYqJm0k52-SU5O19G1Xzph_rjShg',
  contratto: '1_PNr5Y6svOADvgKZIjFjKsoDFpNV6TkOxivLIVqcZdA',
  contrattoREDEN: '1mFtXfWCxKv2y4-kbkRLugrDx_Hnih_ZkMmFto0RtSVU',
  contrattoGSE: '1t5S9CYogDPAtKhy2ejMVELKjAkkieqfu31eIFF06GYg',
  contrattoFinanz: '1RCr8lgM98ryQwMiGFqecMHiWgIHsPN0tfV5HN82eYr4'
};

//nome del file dati tecnici creato per il cliente
const nomeFileDatiTecnici = 'dati tecnici 5.8';

// ID del modello di base del foglio "dati tecnici
const DATI_TECNICI_TEMPLATE_ID = '1cmiPctIhVJu3xA8jqnOYO0eXSBBmWkTBTUkEko1GLFs';

// ID del database CRM
const CRMdatabase = SpreadsheetApp.openById('1_QEo5ynx_29j3I3uJJff5g7ZzGZJnPcIarIXfr5O2gQ'); 

// Nome del foglio con le offerte
const sheetOfferte = CRMdatabase.getSheetByName('offerte');


/** ===============================
*       FUNZIONI PRINCIPALI
*  ===============================
*/

/**
* Esegue l'intero processo di generazione dell'offerta.
*
* @param {string} appID - ID dell'applicazione/opportunità.
* @param {Object} inputData - Dati di input necessari per la generazione dell'offerta.
*/


// Funzione principale
function stampaOffertaV2(appID, tipo_opportunita, id, yy, nome, cognome, indirizzo, telefono, email, numero_moduli, numero_inverter, marca_moduli, 
                      marca_inverter, numero_batteria, capacita_batteria, totale_capacita_batterie, marca_batteria, tetto, 
                      potenza_impianto, alberi, testo_aggiuntivo, tipo_pagamento, 
                      condizione_pagamento_1, condizione_pagamento_2, condizione_pagamento_3, condizione_pagamento_4, imponibile_offerta,
                      iva_offerta, prezzo_offerta, cartella, anni_finanziamento, esposizione, area_m2_impianto, 
                      numero_colonnina_74kw, numero_colonnina_22kw, numero_ottimizzatori, marca_ottimizzatori, numero_linea_vita, 
                      scheda_tecnica_moduli, scheda_tecnica_inverter, scheda_tecnica_batterie, scheda_tecnica_ottimizzatori, consumi_annui, 
                      profilo_di_consumo, provincia, prezzo_energia, numero_rate_mensili, anni_finanziamento, durata_incentivo) {

Logger.log('Avvio funzione stampaOffertaV2 per appID: ' + appID);

// Creazione stringa di data nel formato 'dd/mm/yyyy'
const oggi = new Date();
const dataOggi = new Intl.DateTimeFormat('it-IT', { day: '2-digit', month: '2-digit', year: 'numeric' }).format(oggi);
Logger.log('Data odierna: ' + dataOggi);

// Ripristino del metodo originale per ottenere l'ID della cartella di destinazione
var cartellaDestinazioneId = cartella.split('/folders/')[1];
Logger.log('ID della cartella principale: ' + cartellaDestinazioneId);

// Recupero o creazione della cartella "contratto"
var cartellaContrattoId = getOrCreateSubfolder(cartellaDestinazioneId, 'contratto');
Logger.log('Cartella contratto ID: ' + cartellaContrattoId);

// Creazione della cartella con la data odierna
var cartellaDataId = getOrCreateSubfolder(cartellaContrattoId, dataOggi);
Logger.log('Cartella con la data ID: ' + cartellaDataId);

// Determina i template dei documenti da usare in base al tipo di opportunità e pagamento
const datiDocumento = determineDocumentTemplates(tipo_opportunita, tipo_pagamento, nome, cognome, dataOggi, id, yy);
Logger.log('Dati documenti: ' + JSON.stringify(datiDocumento));

// Se l'opportunità non è di tipo "MAT", esegue operazioni aggiuntive
if (tipo_opportunita !== "MAT") {
  Logger.log('Esecuzione operazioni aggiuntive per tipo_opportunita: ' + tipo_opportunita);
  processAdditionalOperations(appID, cartella, consumi_annui, profilo_di_consumo, provincia, esposizione, prezzo_energia);
}


// Prepara il file dati tecnici
var cartellaProgettoId = getOrCreateSubfolder(cartellaDestinazioneId, 'progetto');
  Logger.log('ID cartella progetto: ' + cartellaProgettoId);

var fileDatiTecnici = DriveApp.getFolderById(cartellaProgettoId).getFilesByName(nomeFileDatiTecnici);
var nuovoFileDatiTecnici;
if (fileDatiTecnici.hasNext()) {
      nuovoFileDatiTecnici = SpreadsheetApp.openById(fileDatiTecnici.next().getId());
      Logger.log('File dati tecnici esistente trovato e aperto: ' + nuovoFileDatiTecnici.getId());
} else {
      var modelloDatiTecnici = DriveApp.getFileById(DATI_TECNICI_TEMPLATE_ID).makeCopy(nomeFileDatiTecnici, DriveApp.getFolderById(cartellaProgettoId));
      nuovoFileDatiTecnici = SpreadsheetApp.openById(modelloDatiTecnici.getId());
      Logger.log('Nuovo file dati tecnici creato: ' + nuovoFileDatiTecnici.getId());
}

  
// Chiamata per catturare i risultati dell'analisi energetica
  const energyAnalysisResults = processEnergyAnalysis(nuovoFileDatiTecnici, consumi_annui, profilo_di_consumo, provincia, esposizione, prezzo_energia, appID);

// Compila i documenti dell'offerta sostituendo i segnaposto con i valori corretti
datiDocumento.forEach(dato => {
  const doc = createDocumentFromTemplate(dato.templateId, cartellaDataId, dato.nomeFile);
  const corpo = doc.getBody();
  let mappaturaSegnapostov2 = createPlaceholderMapping({
      tipo_opportunita, id, yy, nome, cognome, indirizzo, telefono, email, dataOggi, numero_moduli, marca_moduli, numero_inverter, 
      marca_inverter, numero_batteria, capacita_batteria, totale_capacita_batterie, marca_batteria, tetto, potenza_impianto, alberi, testo_aggiuntivo, tipo_pagamento, condizione_pagamento_1, condizione_pagamento_2, 
      condizione_pagamento_3, condizione_pagamento_4, imponibile_offerta, iva_offerta, prezzo_offerta, anni_finanziamento, 
      esposizione, area_m2_impianto, scheda_tecnica_moduli, scheda_tecnica_inverter, 
      scheda_tecnica_batterie, scheda_tecnica_ottimizzatori, numero_colonnina_74kw, numero_colonnina_22kw, numero_ottimizzatori, 
      marca_ottimizzatori, numero_linea_vita, prezzo_energia, numero_rate_mensili, anni_finanziamento, durata_incentivo,

      percentuale_risparmio_energetico: energyAnalysisResults.percentuale_risparmio_energetico,
      anni_ritorno_investimento: energyAnalysisResults.anni_ritorno_investimento,
      utile_25_anni: energyAnalysisResults.utile_25_anni,
      percentuale_autoconsumo: energyAnalysisResults.percentuale_autoconsumo,
      media_vendita: energyAnalysisResults.media_vendita,
      detrazione: energyAnalysisResults.detrazione,
      massimale: energyAnalysisResults.massimale,
      rata_mensile: energyAnalysisResults.rata_mensile,
      produzione_primo_anno: energyAnalysisResults.produzione_primo_anno

  });

  replacePlaceholders(corpo, mappaturaSegnapostov2);
  addHyperlink(corpo, 'Link scheda tecnica moduli', scheda_tecnica_moduli);
  addHyperlink(corpo, 'Link scheda tecnica inverter', scheda_tecnica_inverter);
  addHyperlink(corpo, 'Link scheda tecnica batterie', scheda_tecnica_batterie);
  addHyperlink(corpo, 'Link scheda tecnica ottimizzatori', scheda_tecnica_ottimizzatori);

  doc.saveAndClose();
  Logger.log('Documento generato: ' + dato.nomeFile);
});

Logger.log('Fine esecuzione funzione stampaOffertaV2');
}


/** ===============================
*       FUNZIONI HELPER
*  ===============================
*/


/**
* Formats a number using the Italian locale.
*
* @param {number} value - The number to format.
* @param {number} decimals - The number of decimal places.
* @returns {string} The formatted number as a string.
*/
function formatNumberItalian(value, decimals) {
return new Intl.NumberFormat('it-IT', {
  minimumFractionDigits: decimals,
  maximumFractionDigits: decimals,
}).format(value);
}


/**
* Crea o ottiene una sottocartella in una cartella specificata.
*
* @param {string} parentFolderId - ID della cartella principale.
* @param {string} folderName - Nome della sottocartella da creare o ottenere.
* @returns {string} ID della sottocartella.
*/


function getOrCreateSubfolder(parentFolderId, folderName) {
  Logger.log('Verifica ID cartella: ' + parentFolderId);
  try {
      // Estrai l'ID se è un URL
      if (parentFolderId.indexOf('folders/') !== -1) {
          parentFolderId = parentFolderId.split('/folders/')[1];
          Logger.log('Estratto ID della cartella dall\'URL: ' + parentFolderId);
      }
      
      var parentFolder = DriveApp.getFolderById(parentFolderId);
      Logger.log('Cartella trovata: ' + parentFolder.getName());
      var subfolderIterator = parentFolder.getFoldersByName(folderName);
      var subfolderId;
      if (subfolderIterator.hasNext()) {
          subfolderId = subfolderIterator.next().getId();
          Logger.log('Sottocartella esistente trovata, ID: ' + subfolderId);
      } else {
          var newFolder = parentFolder.createFolder(folderName);
          subfolderId = newFolder.getId();
          Logger.log('Nuova sottocartella creata, ID: ' + subfolderId);
      }
      return subfolderId;
  } catch (error) {
      Logger.log('Errore durante la ricerca o creazione della sottocartella: ' + error.message);
      throw error;
  }
}


/**
* Determina i template dei documenti da utilizzare in base al tipo di opportunità e al tipo di pagamento.
*

* @param {string} dataOggi - Data corrente formattata.
* @param {string} id - ID dell'offerta.
* @param {string} yy - Anno.
* @returns {Array} Array di oggetti contenenti templateId e nomeFile per ogni documento da creare.
*/
function determineDocumentTemplates(tipo_opportunita, tipo_pagamento, nome, cognome, dataOggi, id, yy) {
Logger.log('Determinazione dei template per tipo_opportunita: ' + tipo_opportunita + ', tipo_pagamento: ' + tipo_pagamento);
const templates = [];

if (tipo_opportunita === "MAT") {
  templates.push({
    templateId: TEMPLATES.offertaMateriale,
    nomeFile: `Offerta Myenergy ${nome} ${cognome} ${dataOggi}`
  });
} else {
  const presentazioneTemplate = tipo_pagamento === "Finanziamento" ? TEMPLATES.presentazioneFinanz : TEMPLATES.presentazione;
  templates.push({
    templateId: presentazioneTemplate,
    nomeFile: `Presentazione offerta Myenergy ${nome} ${cognome}`
  });

  if (tipo_opportunita === "REDEN") {
    templates.push({
      templateId: TEMPLATES.contrattoREDEN,
      nomeFile: `Offerta ${tipo_opportunita}-${id}-${yy} ${nome} ${cognome} ${dataOggi}`
    }, {
      templateId: TEMPLATES.contrattoGSE,
      nomeFile: `Contratto GSE ${tipo_opportunita}-${id}-${yy} ${nome} ${cognome} ${dataOggi}`
    });
  } else {
    const contrattoTemplate = tipo_pagamento === "Finanziamento" ? TEMPLATES.contrattoFinanz : TEMPLATES.contratto;
    templates.push({
      templateId: contrattoTemplate,
      nomeFile: `Offerta ${tipo_opportunita}-${id}-${yy} ${nome} ${cognome} ${dataOggi}`
    });
  }
}

Logger.log('Template selezionati: ' + JSON.stringify(templates));
return templates;
}


/**
* Esegue operazioni aggiuntive per opportunità non di tipo "MAT".
*
* @param {string} appID - ID specifico dell'offerta in appSheet.
* @param {string} cartella - ID della cartella principale.
*/
function processAdditionalOperations(appID, cartella, consumi_annui, profilo_di_consumo, provincia, esposizione, prezzo_energia) {
Logger.log('Esecuzione delle operazioni aggiuntive per appID: ' + appID);

const cartellaProgettoId = getOrCreateSubfolder(cartella, 'progetto');
Logger.log('ID cartella progetto: ' + cartellaProgettoId);

const nuovoFileDatiTecnici = createOrUpdateTechnicalDataFile(cartellaProgettoId, nomeFileDatiTecnici);

updateTechnicalDataLog(nuovoFileDatiTecnici, appID);

processEnergyAnalysis(nuovoFileDatiTecnici, consumi_annui, profilo_di_consumo, provincia, esposizione, prezzo_energia, appID);

modifyReturnGraph(nuovoFileDatiTecnici);
}


/**
* Crea o aggiorna il file dei dati tecnici nella cartella progetto.
*
* @param {string} cartellaProgettoId - ID della cartella progetto.
* @param {string} nomeFileDatiTecnici - Nome del file dati tecnici.
* @returns {Spreadsheet} Riferimento al file di dati tecnici aperto o creato.
*/
function createOrUpdateTechnicalDataFile(cartellaProgettoId, nomeFileDatiTecnici) {
Logger.log('Creazione o aggiornamento del file dati tecnici: ' + nomeFileDatiTecnici);
const fileDatiTecniciIterator = DriveApp.getFolderById(cartellaProgettoId).getFilesByName(nomeFileDatiTecnici);
const fileId = fileDatiTecniciIterator.hasNext() 
  ? fileDatiTecniciIterator.next().getId()
  : DriveApp.getFileById(DATI_TECNICI_TEMPLATE_ID).makeCopy(nomeFileDatiTecnici, DriveApp.getFolderById(cartellaProgettoId)).getId();
Logger.log('ID file dati tecnici: ' + fileId);
return SpreadsheetApp.openById(fileId);
}


/**
* Aggiorna il log dei dati tecnici con l'ultima offerta generata, mettendo l'appID seguito dal numero della riga.
*
* @param {Spreadsheet} nuovoFileDatiTecnici - Riferimento al file di dati tecnici appena creato.
* @param {string} appID - ID di appSheet relativo all'offerta
*/
function updateTechnicalDataLog(nuovoFileDatiTecnici, appID) {
Logger.log('Aggiornamento log dati tecnici per appID: ' + appID);


const data = sheetOfferte.getDataRange().getValues();  // Ottieni tutti i dati del foglio "offerte"
const appIDColIndex = data[0].indexOf('appID');  // Trova l'indice della colonna appID

if (appIDColIndex === -1) {
  throw new Error('Colonna "appID" non trovata');
}

// Trova la riga con l'appID corrispondente
const selectedRow = data.find(row => row[appIDColIndex] === appID);
if (!selectedRow) {
  throw new Error('Nessuna riga trovata con appID: ' + appID);
}

Logger.log('Riga selezionata per appID: ' + JSON.stringify(selectedRow));

const nuovoSheet = nuovoFileDatiTecnici.getActiveSheet();
const ultimaRigaVuota = nuovoSheet.getLastRow() + 1;  // Calcola l'ultima riga vuota

// Inserisci l'appID seguito dal numero della riga nella prima colonna
selectedRow[0] = `${appID}-${ultimaRigaVuota}`;

// Scrivi la riga selezionata nel nuovo foglio
nuovoSheet.getRange(ultimaRigaVuota, 1, 1, selectedRow.length).setValues([selectedRow]);

Logger.log('Log dati tecnici aggiornato con successo. appID: ' + appID + '-' + ultimaRigaVuota);
}




/**
* Gestisce l'analisi energetica all'interno del file dei dati tecnici.
*
* @param {Spreadsheet} nuovoFileDatiTecnici - Riferimento al file di dati tecnici appena creato.
* @param {string} appID - ID di appSheet relativo all'offerta
*/
function processEnergyAnalysis(nuovoFileDatiTecnici, consumi_annui, profilo_di_consumo, provincia, esposizione, prezzo_energia, appID) {
  Logger.log('Esecuzione dell\'analisi energetica per appID: ' + appID);

  const sheetAnalisiEnergetica = nuovoFileDatiTecnici.getSheetByName('analisi energetica');

  // Format numbers using Italian locale
  const formattedConsumiAnnui = formatNumberItalian(consumi_annui, 2);
  const formattedPrezzoEnergia = formatNumberItalian(prezzo_energia, 2);

  // Set values in the spreadsheet
  sheetAnalisiEnergetica.getRange('consumi_annui').setValue(formattedConsumiAnnui);
  sheetAnalisiEnergetica.getRange('profilo_di_consumo').setValue(profilo_di_consumo);
  sheetAnalisiEnergetica.getRange('provincia').setValue(provincia);
  sheetAnalisiEnergetica.getRange('esposizione').setValue(esposizione);
  sheetAnalisiEnergetica.getRange('prezzo_energia').setValue(formattedPrezzoEnergia);

  // Retrieve calculated results
  const percentuale_autoconsumo = sheetAnalisiEnergetica.getRange('percentuale_autoconsumo').getValue();
  const media_vendita = sheetAnalisiEnergetica.getRange('media_vendita').getValue();
  const anni_ritorno_investimento = sheetAnalisiEnergetica.getRange('anni_ritorno_investimento').getValue();
  const percentuale_risparmio_energetico = sheetAnalisiEnergetica.getRange('percentuale_risparmio_energetico').getValue();
  const utile_25_anni = sheetAnalisiEnergetica.getRange('utile_25_anni').getValue();
  const detrazione = sheetAnalisiEnergetica.getRange('detrazione').getValue();
  const massimale = sheetAnalisiEnergetica.getRange('massimale').getValue();
  const rata_mensile = sheetAnalisiEnergetica.getRange('rata_mensile').getValue();
  const produzione_primo_anno = sheetAnalisiEnergetica.getRange('produzione_primo_anno').getValue();


  Logger.log('Analisi energetica completata.');

  // Return calculated values
  return {
      percentuale_autoconsumo,
      media_vendita,
      anni_ritorno_investimento,
      percentuale_risparmio_energetico,
      utile_25_anni,
      detrazione,
      massimale,
      rata_mensile,
      produzione_primo_anno
  };
}




/**
* Estrae i valori da un intervallo denominato in un foglio di calcolo.
*
* @param {Sheet} sheet - Il foglio di calcolo da cui estrarre i valori.
* @param {string} rangeName - Il nome dell'intervallo denominato.
* @returns {Array} I valori dell'intervallo denominato.
*/
function getNamedRangeValues(sheet, rangeName) {
  var range = sheet.getRange(rangeName);
  if (!range) {
      throw new Error('Intervallo denominato "' + rangeName + '" non trovato nel foglio "' + sheet.getName() + '".');
  }
  return range.getValues();
}

/**
* Modifica il grafico di ritorno annuo nel foglio di calcolo associato.
*
* @param {Spreadsheet} nuovoFileDatiTecnici - Riferimento al file di dati tecnici appena creato.
*/
function modifyReturnGraph(nuovoFileDatiTecnici) {
Logger.log('Modifica del grafico di ritorno annuo.');

const grafico = SpreadsheetApp.openById('1cfLNo1WU-poleX1i4hg9fynUP53i06d_6DpIlbFZcTc');
const foglio1GRAFICO = grafico.getSheetByName('foglio1');

const sheetCalcoliDATITECNICI = nuovoFileDatiTecnici.getSheetByName('calcoli');
const valoriUtile25AnniGrafico = getNamedRangeValues(sheetCalcoliDATITECNICI, 'utile_25_anni_grafico');

foglio1GRAFICO.getRange(1, 2, valoriUtile25AnniGrafico.length, valoriUtile25AnniGrafico[0].length).setValues(valoriUtile25AnniGrafico);

Logger.log('Grafico di ritorno annuo aggiornato.');
}

/**
* Crea una mappatura dei segnaposto con i valori forniti come input.
*
* @param {Object} params - Oggetto contenente tutti i parametri necessari per la mappatura.
* @returns {Object} Mappatura dei segnaposto con i relativi valori.
*/
function createPlaceholderMapping(params) {
Logger.log('Creazione della mappatura dei segnaposto.');
Logger.log('Valore originale di detrazione: ' + params.detrazione + ' (tipo: ' + typeof params.detrazione + ')');
Logger.log('Valore originale di iva_offerta: ' + params.iva_offerta + ' (tipo: ' + typeof params.iva_offerta + ')');
Logger.log('Valore originale di prezzo_offerta: ' + params.prezzo_offerta + ' (tipo: ' + typeof params.prezzo_offerta + ')');

return {
  '{{tipo_opportunità}}': params.tipo_opportunita,
  '{{id}}': params.id,
  '{{yy}}': params.yy,
  '{{nome}}': params.nome,
  '{{cognome}}': params.cognome,
  '{{indirizzo}}': params.indirizzo,
  '{{telefono}}': params.telefono,
  '{{email}}': params.email,
  '{{data ultima modifica}}': params.dataOggi,
  '{{numero_moduli}}': params.numero_moduli,
  '{{marca_moduli}}': params.marca_moduli,
  '{{numero_inverter}}': params.numero_inverter,
  '{{marca_inverter}}': params.marca_inverter,
  '{{numero_batteria}}': params.numero_batteria,
  '{{capacità batteria}}': params.capacita_batteria,
  '{{totale_capacità_batterie}}': params.totale_capacita_batterie,
  '{{marca_batteria}}': params.marca_batteria,
  '{{tetto}}': params.tetto,
  '{{potenza_impianto}}': formatNumber(params.potenza_impianto, 2),
  '{{produzione_primo_anno}}': formatNumber(params.produzione_primo_anno, 0),
  '{{alberi}}': formatNumber(params.alberi, 0),
  '{{testo_aggiuntivo}}': params.testo_aggiuntivo,
  '{{tipo_pagamento}}': params.tipo_pagamento,
  '{{condizione_pagamento_1}}': params.condizione_pagamento_1,
  '{{condizione_pagamento_2}}': params.condizione_pagamento_2,
  '{{condizione_pagamento_3}}': params.condizione_pagamento_3,
  '{{condizione_pagamento_4}}': params.condizione_pagamento_4,
  '{{imponibile_offerta}}': formatCurrency(params.imponibile_offerta),
  '{{iva_offerta}}': formatCurrency(params.iva_offerta, 0),
  '{{prezzo_offerta}}': formatCurrency(params.prezzo_offerta),
  '{{anni_finanziamento}}': formatNumber(params.anni_finanziamento, 0),
  '{{rata_mensile}}': formatCurrency(params.rata_mensile),
  '{{numero_rate_mensili}}': formatNumber(params.numero_rate_mensili, 0),
  '{{esposizione}}': params.esposizione,
  '{{area_m2_impianto}}': formatNumber(params.area_m2_impianto, 2),
  '{{scheda_tecnica_moduli}}': 'Link scheda tecnica moduli',
  '{{scheda_tecnica_inverter}}': 'Link scheda tecnica inverter',
  '{{scheda_tecnica_batterie}}': 'Link scheda tecnica batterie',
  '{{scheda_tecnica_ottimizzatori}}': 'Link scheda tecnica ottimizzatori',
  '{{numero_colonnina_74kw}}': params.numero_colonnina_74kw,
  '{{numero_colonnina_22kw}}': params.numero_colonnina_22kw,
  '{{numero_ottimizzatori}}': params.numero_ottimizzatori,
  '{{marca_ottimizzatori}}': params.marca_ottimizzatori,
  '{{numero_linea_vita}}': params.numero_linea_vita,
  '{{detrazione}}': formatCurrency(params.detrazione),
  '{{anni_ritorno_investimento}}': formatNumber(params.anni_ritorno_investimento, 1),
  '{{utile_25_anni}}': formatCurrency(params.utile_25_anni),
  '{{percentuale_autoconsumo}}': formatPercentage(params.percentuale_autoconsumo),
  '{{media_vendita}}': formatCurrency(params.media_vendita),
  '{{prezzo_energia}}': formatCurrency(params.prezzo_energia),
  '{{percentuale_risparmio_energetico}}': formatPercentage(params.percentuale_risparmio_energetico),
  '{{durata_incentivo}}': formatNumber(params.durata_incentivo, 0)
};
}


/**
* Formatta un numero con il numero specificato di decimali.
*
* @param {number} value - Numero da formattare.
* @param {number} decimals - Numero di decimali.
* @returns {string} Numero formattato.
*/
function formatNumber(value, decimals) {
Logger.log('Valore ricevuto in formatNumber: ' + value + ' (tipo: ' + typeof value + ')');

if (value == null || value === '') {
  return '-';
}

// Rimuovi tutti i caratteri non numerici, tranne punto e virgola
value = value.toString().replace(/[^\d.,-]/g, '');

Logger.log('Valore dopo pulizia: ' + value);

// Gestione dei separatori decimali
if (value.includes(',') && value.includes('.')) {
  value = value.replace(/\./g, '');
  value = value.replace(',', '.');
} else if (value.includes(',')) {
  value = value.replace(',', '.');
}

Logger.log('Valore prima della conversione numerica: ' + value);

value = parseFloat(value);

Logger.log('Valore numerico: ' + value);

if (isNaN(value)) {
  return '-';
}

// Formattazione personalizzata
let formattedValue = value.toFixed(decimals);
let parts = formattedValue.split('.');
parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, '.');
formattedValue = parts.join(',');

Logger.log('Valore formattato personalizzato: ' + formattedValue);

return formattedValue;
}


/**
* Formatta un valore come valuta.
*
* @param {number} value - Valore da formattare.
* @returns {string} Valore formattato come valuta.
*/
function formatCurrency(value) {
Logger.log('Valore ricevuto in formatCurrency: ' + value + ' (tipo: ' + typeof value + ')');

if (value == null || value === '') {
  return '-';
}

// Rimuovi tutti i caratteri non numerici, tranne punto e virgola
value = value.toString().replace(/[^\d.,-]/g, '');

Logger.log('Valore dopo pulizia: ' + value);

// Gestione dei separatori decimali
if (value.includes(',') && value.includes('.')) {
  // Caso in cui ci sono sia punto che virgola (es. "1.231,82")
  value = value.replace(/\./g, '');
  value = value.replace(',', '.');
} else if (value.includes(',')) {
  // Caso in cui c'è solo la virgola (es. "1231,82")
  value = value.replace(',', '.');
}

Logger.log('Valore prima della conversione numerica: ' + value);

value = parseFloat(value);

Logger.log('Valore numerico: ' + value);

if (isNaN(value)) {
  return '-';
}

// Formattazione personalizzata
let parts = value.toFixed(2).split('.');
parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, '.');
let formattedValue = parts.join(',');
formattedValue += ' €';

Logger.log('Valore formattato personalizzato: ' + formattedValue);

return formattedValue;
}


/**
* Formatta un valore come percentuale.
*
* @param {number} value - Valore da formattare.
* @returns {string} Valore formattato come percentuale.
*/
function formatPercentage(value) {
return new Intl.NumberFormat('it-IT', { style: 'percent', minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(value);
}


/**
* Crea un documento da un template e lo salva in una cartella specificata.
*
* @param {string} templateId - ID del template da usare.
* @param {string} destinationFolderId - ID della cartella di destinazione.
* @param {string} fileName - Nome del file da creare.
* @returns {Document} Riferimento al documento creato.
*/
function createDocumentFromTemplate(templateId, destinationFolderId, fileName) {
Logger.log('Creazione del documento da template ID: ' + templateId);
const documentCopy = DriveApp.getFileById(templateId).makeCopy(fileName, DriveApp.getFolderById(destinationFolderId));
Logger.log('Documento creato: ' + fileName);
return DocumentApp.openById(documentCopy.getId());
}


/**
* Funzione per sostituire i segnaposto con i valori nel documento.
* Sostituisce i valori vuoti con un trattino (-).
*
* @param {Body} body - Il corpo del documento Google Docs.
* @param {Object} placeholders - Oggetto mappa che contiene i segnaposto e i loro valori.
*/
function replacePlaceholders(body, placeholders) {
  Object.keys(placeholders).forEach(function(placeholder) {
      let value = placeholders[placeholder];
      if (value === null || value === undefined || value === "") {
          value = "-"; // Sostituisci valori vuoti, null o undefined con un trattino
      }
      body.replaceText(placeholder, value.toString());
  });
}


/**
* Aggiunge un hyperlink a un testo specificato all'interno di un documento.
*
* @param {Body} corpo - Corpo del documento.
* @param {string} searchText - Testo da cercare e sostituire con un link.
* @param {string} url - URL da collegare al testo.
*/
function addHyperlink(corpo, searchText, url) {
Logger.log('Aggiunta di un hyperlink al testo: ' + searchText);
let foundElement = corpo.findText(searchText);
while (foundElement) {
  const foundText = foundElement.getElement().asText();
  const startOffset = foundElement.getStartOffset();
  const endOffset = foundElement.getEndOffsetInclusive();
  foundText.setLinkUrl(startOffset, endOffset, url);
  foundElement = corpo.findText(searchText, foundElement);
}
Logger.log('Hyperlink aggiunto.');
}


/** ===============================
*       DEBUG SCRIPT
*  ===============================
*/

// Dati di esempio per l'esecuzione
const exampleInputData = {
appID: '5bd33768',
id: 'RES-6741-24',
nome: 'Mario',
cognome: 'Rossi',
indirizzo: 'Via Roma 1, Milano',
telefono: '1234567890',
email: 'mario.rossi@example.com',
tipo_opportunita: 'RES',
tipo_pagamento: 'Acquisto diretto',
// Aggiungi altri campi necessari...
};

// Esecuzione dello script con i dati di esempio
function runStampaOffertaV2() {
stampaOffertaV2(exampleInputData.appID, exampleInputData);
}