/**
 * Aggiorna il log dei dati tecnici con l'ultima offerta generata, mettendo l'appID seguito dal numero della riga.
 *
 * @param {Spreadsheet} nuovoFileDatiTecnici - Riferimento al file di dati tecnici appena creato.
 * @param {string} appID - ID di appSheet relativo all'offerta
 */

function newLogDatiTecnici(nuovoFileDatiTecnici, appID) {
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
 * incolla sul file i valori specificati e ne raccoglie altri
 */

  function processDatiTecnici(nuovoFileDatiTecnici, consumi_annui, profilo_di_consumo, provincia, esposizione, prezzo_energia, appID) {
  Logger.log('Esecuzione dell\'analisi energetica per appID: ' + appID);

  if (!nuovoFileDatiTecnici) {
    throw new Error('Invalid "nuovoFileDatiTecnici" object.');
  }


  // Open the "analisi energetica" sheet
  const sheetAnalisiEnergetica = nuovoFileDatiTecnici.getSheetByName('analisi energetica');
  if (!sheetAnalisiEnergetica) {
    throw new Error('"analisi energetica" sheet not found.');
  }

  Logger.log('Setting values in "analisi energetica" sheet.');


  // Format numbers using formatting functions
  const formattedConsumiAnnui = formatNumberItalian(consumi_annui, 2);
  const formattedPrezzoEnergia = formatNumberItalian(prezzo_energia, 2);


  // Set values
  sheetAnalisiEnergetica.getRange('consumi_annui').setValue(formattedConsumiAnnui);
  sheetAnalisiEnergetica.getRange('profilo_di_consumo').setValue(profilo_di_consumo);
  sheetAnalisiEnergetica.getRange('provincia').setValue(provincia);
  sheetAnalisiEnergetica.getRange('esposizione').setValue(esposizione);
  sheetAnalisiEnergetica.getRange('prezzo_energia').setValue(formattedPrezzoEnergia);


  // Retrieve calculated results
  const results = {
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

  Logger.log('Energy Analysis Results: ' + JSON.stringify(results));


  // Return the file ID
  return { ...results};

}