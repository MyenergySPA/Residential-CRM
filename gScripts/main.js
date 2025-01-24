/* Questo script automatizza la generazione di documenti di offerta basandosi su template predefiniti e dati forniti dall'utente.
Gestisce la creazione e l'organizzazione di cartelle su Google Drive, l'elaborazione di dati tecnici ed energetici, e la sostituzione
di segnaposto nei documenti finali.
 COSTANTI GLOBALI */

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
const nomeFileDatiTecnici = 'dati tecnici 5.9';

// ID del modello di base del foglio "dati tecnici"
const DATI_TECNICI_TEMPLATE_ID = '1cmiPctIhVJu3xA8jqnOYO0eXSBBmWkTBTUkEko1GLFs';

// ID del database CRM
const CRMdatabase = SpreadsheetApp.openById('1_QEo5ynx_29j3I3uJJff5g7ZzGZJnPcIarIXfr5O2gQ'); 

// Nome del foglio con le offerte
const sheetOfferte = CRMdatabase.getSheetByName('offerte');


/** FUNZIONE PRINCIPALE */

function main(appID, tipo_opportunita, id, yy, nome, cognome, indirizzo, telefono, email, numero_moduli, numero_inverter, marca_moduli, 
                      marca_inverter, numero_batteria, capacita_batteria, totale_capacita_batterie, marca_batteria, tetto, 
                      potenza_impianto, alberi, testo_aggiuntivo, tipo_pagamento, 
                      condizione_pagamento_1, condizione_pagamento_2, condizione_pagamento_3, condizione_pagamento_4, imponibile_offerta,
                      iva_offerta, prezzo_offerta, cartella, anni_finanziamento, esposizione, area_m2_impianto, 
                      numero_colonnina_74kw, numero_colonnina_22kw, numero_ottimizzatori, marca_ottimizzatori, numero_linea_vita, 
                      scheda_tecnica_moduli, scheda_tecnica_inverter, scheda_tecnica_batterie, scheda_tecnica_ottimizzatori, consumi_annui, 
                      profilo_di_consumo, provincia, prezzo_energia, numero_rate_mensili, anni_finanziamento, durata_incentivo, indirizzo) {

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
const datiDocumento = determineDocumentTemplates(tipo_opportunita, tipo_pagamento, nome, cognome, dataOggi, id, yy);
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

/* Se l'opportunità non è di tipo "MAT", esegue operazioni aggiuntive
if (tipo_opportunita !== "MAT") { 
*/
  //creazione nuova riga in sheet "log" in "dati tecnici", contenente info offerta
  newLogDatiTecnici(nuovoFileDatiTecnici, appID);

  // elaborazione dati e grafici su "dati tecnici" sheet
  const datiTecniciData = processDatiTecnici(nuovoFileDatiTecnici, consumi_annui, profilo_di_consumo, provincia, esposizione, prezzo_energia, appID, cartellaProgettoId);
  Logger.log('Energy Analysis Results: ' + JSON.stringify(datiTecniciData));
// }
  

// --- Upload charts using CHART_DEFINITIONS from newCharts.gs ---
// Simply reference CHART_DEFINITIONS (it's globally accessible)
const chartMappings = newCharts(nuovoFileDatiTecnici, cartellaProgettoId, CHART_DEFINITIONS);


// crea un array con tutti i doc d'offerta:
const createdDocIds = [];
datiDocumento.forEach(dato => {
  const doc = createDocumentFromTemplate(dato.templateId, cartellaDataId, dato.nomeFile);
  const corpo = doc.getBody();
  
let mappaturaSegnapostov2 = createPlaceholderMapping({
    tipo_opportunita, id, yy, nome, cognome, indirizzo, telefono, email, dataOggi, numero_moduli, marca_moduli, numero_inverter, marca_inverter, numero_batteria, capacita_batteria, totale_capacita_batterie, marca_batteria, tetto, potenza_impianto, alberi, testo_aggiuntivo, tipo_pagamento, condizione_pagamento_1, condizione_pagamento_2, condizione_pagamento_3, condizione_pagamento_4, imponibile_offerta, iva_offerta, prezzo_offerta, anni_finanziamento, esposizione, area_m2_impianto, scheda_tecnica_moduli, scheda_tecnica_inverter, scheda_tecnica_batterie, scheda_tecnica_ottimizzatori, numero_colonnina_74kw, numero_colonnina_22kw, numero_ottimizzatori, marca_ottimizzatori, numero_linea_vita, prezzo_energia, numero_rate_mensili, anni_finanziamento, durata_incentivo, indirizzo,
      
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