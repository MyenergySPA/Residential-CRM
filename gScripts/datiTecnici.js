/**
 * Aggiorna il log dei dati tecnici con l'ultima offerta generata, mettendo l'appID seguito dal numero della riga.
 * Questa funzione si aspetta che la variabile 'sheetOfferte' sia definita globalmente.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} nuovoFileDatiTecnici - Riferimento al file di dati tecnici appena creato.
 * @param {string} appID - ID di appSheet relativo all'offerta
 */
function newLogDatiTecnici(nuovoFileDatiTecnici, appID) {
  Logger.log('Aggiornamento log dati tecnici per appID: ' + appID);

  // La funzione ora utilizza la variabile globale 'sheetOfferte', che viene
  // inizializzata dallo script principale 'main'.
  if (typeof sheetOfferte === 'undefined' || !sheetOfferte) {
    throw new Error("La variabile globale 'sheetOfferte' non è definita. Assicurati che lo script 'main' la inizializzi correttamente prima di chiamare questa funzione.");
  }

  const data = sheetOfferte.getDataRange().getValues();  // Ottieni tutti i dati del foglio "offerte"
  const appIDColIndex = data[0].indexOf('appID');  // Trova l'indice della colonna appID
  
  if (appIDColIndex === -1) {
    throw new Error('Colonna "appID" non trovata nel foglio "offerte"');
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


/** * Incolla sul file i valori specificati, chiama l'API PVGIS e raccoglie i risultati.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} nuovoFileDatiTecnici
 * @param {number} consumi_annui
 * @param {string} profilo_di_consumo
 * @param {string} provincia
 * @param {string} esposizione
 * @param {number} prezzo_energia
 * @param {string} indirizzo
 * @param {string} coordinate
 * @param {string} appID
 * @returns {object} I risultati calcolati dal foglio.
 */
function processDatiTecnici(nuovoFileDatiTecnici, consumi_annui, profilo_di_consumo, provincia, esposizione, prezzo_energia, indirizzo, coordinate, appID) {
  Logger.log('Esecuzione dell\'analisi energetica per appID: ' + appID);

  if (!nuovoFileDatiTecnici) {
    throw new Error('Oggetto "nuovoFileDatiTecnici" non valido.');
  }

  // Apri il foglio "analisi energetica"
  const sheetAnalisiEnergetica = nuovoFileDatiTecnici.getSheetByName('analisi energetica');
  if (!sheetAnalisiEnergetica) {
    throw new Error('Foglio "analisi energetica" non trovato.');
  }
  Logger.log('Impostazione dei valori nel foglio "analisi energetica".');


  // Format numbers using formatting functions
  let formattedConsumiAnnui = formatNumberItalian(consumi_annui, 2);
  let formattedPrezzoEnergia = formatNumberItalian(prezzo_energia, 2);


  // format values and get pvgis data
  sheetAnalisiEnergetica.getRange('consumi_annui').setValue(formattedConsumiAnnui);

  sheetAnalisiEnergetica.getRange('profilo_di_consumo').setValue(profilo_di_consumo);

  sheetAnalisiEnergetica.getRange('provincia').setValue(provincia);

  sheetAnalisiEnergetica.getRange('esposizione').setValue(esposizione);

  sheetAnalisiEnergetica.getRange('prezzo_energia').setValue(formattedPrezzoEnergia);



  // Recupera i valori calcolati dal foglio per la chiamata API

  let tilt = sheetAnalisiEnergetica.getRange("tilt").getValue();

  let azimuth = -sheetAnalisiEnergetica.getRange("azimuth").getValue();
  
  let tipoInstallazione = sheetAnalisiEnergetica.getRange("tipo_Installazione").getValue().toString().trim().toLowerCase();

  let perdita = sheetAnalisiEnergetica.getRange("perdita").getValue();

  Logger.log(`Valori per API: tilt=${tilt}, azimuth=${azimuth}, tipoInstallazione="${tipoInstallazione}", perdita=${perdita}`);
  

  /** * Inizio chiamata API a PVGIS
   */

  // Dividi le coordinate in Latitudine e Longitudine
  let latLongSplit = coordinate.split(",");
  let lat = latLongSplit[0].trim();
  let lon = latLongSplit[1].trim();

  // Converte "Tipo installazione" nel formato compatibile con l'API
  let mountingPlace = (tipoInstallazione === "tetto") ? "building" : "free";
  Logger.log(`Valore 'mountingplace' per API: "${mountingPlace}"`);

  // Costruisci l'URL dell'API. Il valore 'perdita' viene passato direttamente,
  // perché PVGIS si aspetta un valore percentuale (es. 14 per 14%).
  let apiUrl = `https://re.jrc.ec.europa.eu/api/v5_3/PVcalc?format=json&lat=${lat}&lon=${lon}&usehorizon=0&raddatabase=PVGIS-SARAH3&peakpower=1&pvtechchoice=crystSi&mountingplace=${mountingPlace}&loss=${perdita}&fixed=1&angle=${tilt}&aspect=${azimuth}&outputformat=json&browser=0`;

  Logger.log("API URL: " + apiUrl);

  // Esegui la chiamata API
  let response = UrlFetchApp.fetch(apiUrl, { 'muteHttpExceptions': true });
  let responseData = response.getContentText();
  
  // Parsing della risposta JSON
  let pvgisData;
  try {
    pvgisData = JSON.parse(responseData);
  } catch (e) {
    Logger.log("Errore durante il parsing del JSON da PVGIS: " + responseData);
    throw new Error("La risposta da PVGIS non è un JSON valido: " + responseData);
  }


  // Controlla se ci sono errori nella risposta dell'API
  if (response.getResponseCode() !== 200 || !pvgisData.outputs || !pvgisData.outputs.monthly) {
      Logger.log("Errore dalla API PVGIS: " + responseData);
      throw new Error("La chiamata a PVGIS ha restituito un errore: " + responseData);
  }

  // Naviga la struttura JSON per ottenere i valori di produzione mensile (E_m)
  let monthlyOutputs = pvgisData.outputs.monthly.fixed;
  
  // Estrai solo il valore 'E_m' (produzione energetica mensile) da ogni oggetto del mese
  let monthValues = monthlyOutputs.map(monthData => monthData.E_m);

  Logger.log("Valori mensili di produzione (E_m) estratti dal JSON: " + JSON.stringify(monthValues));
  
  if (monthValues.length !== 12) {
      throw new Error('La risposta da PVGIS non conteneva 12 valori mensili. Dati ricevuti: ' + JSON.stringify(monthValues));
  }

  // *** NUOVA LOGICA DI SCRITTURA DATI ***
  Logger.log("Inizio scrittura dei dati nel foglio Impostazioni.");

  const sheet_impostazioni = nuovoFileDatiTecnici.getSheetByName('Impostazioni');
  if (!sheet_impostazioni) {
      throw new Error('Foglio "Impostazioni" non trovato.');
  }

  // Imposta il tipo di esposizione su "personalizzata".
  // Si presume che la cella per il valore sia B3. Per una maggiore robustezza, si potrebbe usare un intervallo denominato.
  try {
      sheet_impostazioni.getRange("B3").setValue("personalizzata");
      Logger.log("Valore 'personalizzata' scritto nella cella B3.");
  } catch (e) {
      Logger.log("ATTENZIONE: Impossibile scrivere 'personalizzata' nella cella B3. Errore: " + e.message);
  }

  // Scrive i valori di produzione mensile in modo efficiente a partire dall'intervallo "prod_gennaio"
  try {
    const startRange = nuovoFileDatiTecnici.getRangeByName("prod_gennaio");
    if (!startRange) {
        throw new Error("L'intervallo denominato 'prod_gennaio' è obbligatorio e non è stato trovato nel foglio 'Impostazioni'.");
    }

    // Converte l'array di valori 1D in un array 2D (una colonna) per setValues()
    const valuesToWrite = monthValues.map(value => [value]);

    // Ottieni il range di destinazione e scrivi tutti i valori in una sola operazione
    sheet_impostazioni.getRange(startRange.getRow(), startRange.getColumn(), 12, 1).setValues(valuesToWrite);
    
    Logger.log("Tutti i 12 valori di produzione sono stati scritti con successo a partire dalla cella " + startRange.getA1Notation());

  } catch (e) {
    Logger.log("Errore critico durante la scrittura dei valori di produzione. Errore: " + e.message);
    throw new Error("Impossibile scrivere i dati di produzione. Assicurarsi che l'intervallo 'prod_gennaio' esista. Dettagli: " + e.message);
  }


  // retrieve calculated results for "analisi energetica"
  let results = {
    percentuale_autoconsumo: sheetAnalisiEnergetica.getRange('percentuale_autoconsumo').getValue(),
    media_vendita: sheetAnalisiEnergetica.getRange('media_vendita').getValue(),
    anni_ritorno_investimento: sheetAnalisiEnergetica.getRange('anni_ritorno_investimento').getValue(),
    percentuale_risparmio_energetico: sheetAnalisiEnergetica.getRange('percentuale_risparmio_energetico').getValue(),
    utile_25_anni: sheetAnalisiEnergetica.getRange('utile_25_anni').getValue(),
    detrazione: sheetAnalisiEnergetica.getRange('detrazione').getValue(),
    massimale: sheetAnalisiEnergetica.getRange('massimale').getValue(),
    rata_mensile: sheetAnalisiEnergetica.getRange('rata_mensile').getValue(),
    produzione_primo_anno: sheetAnalisiEnergetica.getRange('produzione_primo_anno').getValue()
  };

  Logger.log('Risultati finali dell\'analisi energetica: ' + JSON.stringify(results));

  // Restituisci i risultati
  return results;
}
