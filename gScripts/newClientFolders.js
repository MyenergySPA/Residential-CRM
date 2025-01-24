/**
 * Script per la creazione di cartelle Google Drive legate a opportunità specifiche
 * L'obiettivo dello script è creare una cartella principale e delle sottocartelle (progetto, documenti, allegati, contratto) 
 * per ogni cliente/opportunità, e poi inserire i link di tali cartelle in un foglio Google Sheets.
 *
 * Ordine del codice:
 * 1. Funzione principale: creaCartelle() - Crea le cartelle e aggiorna Google Sheets
 * 2. Funzione secondaria: aggiornaFoglioConURL() - Aggiorna il foglio di calcolo con i link delle cartelle create
 */



/**
 * Funzione principale che crea la cartella principale e le sottocartelle,
 * e aggiorna il foglio Google con i rispettivi link.
 * 
 * @param {string} tipoOpportunita - Il tipo di opportunità (es. "residenziale", "aziendale").
 * @param {string} id - ID univoco per l'opportunità/cliente.
 * @param {string} yy - Anno corrente in formato YY.
 * @param {string} nome - Nome del cliente.
 * @param {string} cognome - Cognome del cliente.
 * @returns {Object} - Ritorna gli URL della cartella principale e delle sottocartelle.
 */
function newClientFolders(tipoOpportunita, id, yy, nome, cognome) {
  Logger.log('Inizio della funzione newClientFolders');
  
  //da modificare sostituendo con variabile globale
  const parentFolderId = "1kpBsmlPAaeCFWvgCEIw38tEk5Q-xQpH_"; // main folder Id, containing all clients' offers
  const parentFolder = DriveApp.getFolderById(parentFolderId);
  
  const folderName = `${tipoOpportunita}-${id}-${yy} ${nome} ${cognome}`; // Nome della cartella
  let mainFolderUrl;
  let subfolderUrls = {};

  // Verifica se la cartella esiste già
  const existingFolders = parentFolder.getFoldersByName(folderName);
  if (existingFolders.hasNext()) {
    Logger.log('La cartella esiste già');
    
    // Recupera la cartella esistente e i link delle sottocartelle
    const existingFolder = existingFolders.next();
    mainFolderUrl = existingFolder.getUrl();
    const existingSubfolders = existingFolder.getFolders();
    while (existingSubfolders.hasNext()) {
      const subfolder = existingSubfolders.next();
      subfolderUrls[subfolder.getName()] = subfolder.getUrl();
    }
  } else {
    Logger.log('La cartella non esiste, creazione di una nuova cartella');
    
    // Crea la cartella principale e le sottocartelle
    const mainFolder = parentFolder.createFolder(folderName);
    mainFolderUrl = mainFolder.getUrl();
    const subfolders = ['progetto', 'documenti', 'allegati', 'contratto'];
    
    // Crea le sottocartelle e memorizza i link
    subfolders.forEach((folder) => {
      const createdFolder = mainFolder.createFolder(folder);
      subfolderUrls[folder] = createdFolder.getUrl();
    });
  }
  
  // Aggiorna il foglio di calcolo con i link 
  // (da modificare con variabili globali al posto di id e nome foglio)
  foldersUrl("1_QEo5ynx_29j3I3uJJff5g7ZzGZJnPcIarIXfr5O2gQ", "cronologia", id, mainFolderUrl, subfolderUrls);
  
  return { mainFolderUrl: mainFolderUrl, subfolderUrls: subfolderUrls };
}

/**
 * Funzione secondaria che aggiorna il foglio di Google con i link delle cartelle create.
 * 
 * @param {string} sheetId - L'ID del foglio Google.
 * @param {string} sheetName - Il nome della pagina del foglio Google.
 * @param {string} id - L'ID univoco per l'opportunità/cliente.
 * @param {string} mainFolderUrl - L'URL della cartella principale.
 * @param {Object} subfolderUrls - Gli URL delle sottocartelle (progetto, documenti, allegati, contratto).
 */
function foldersUrl(sheetId, sheetName, id, mainFolderUrl, subfolderUrls) {
  Logger.log('Inizio della funzione aggiornaFoglioConURL');
  
  const sheet = SpreadsheetApp.openById(sheetId).getSheetByName(sheetName);
  const values = sheet.getDataRange().getValues();
  
  // Trova gli indici delle colonne richieste
  const idColumnIndex = values[0].indexOf("id") + 1;
  const folderColumnIndex = values[0].indexOf("cartella") + 1;
  const progettoColumnIndex = values[0].indexOf("sottocartella progetto") + 1;
  const documentiColumnIndex = values[0].indexOf("sottocartella documenti") + 1;
  const allegatiColumnIndex = values[0].indexOf("sottocartella allegati") + 1;
  const contrattoColumnIndex = values[0].indexOf("sottocartella contratto") + 1;

  // Cerca la riga corrispondente all'ID
  let targetRow;
  for (let i = 1; i < values.length; i++) {
    if (values[i][idColumnIndex - 1].toString() === id.toString()) {
      targetRow = i + 1;
      break;
    }
  }

  // Aggiorna la riga con i link delle cartelle
  if (targetRow) {
    sheet.getRange(targetRow, folderColumnIndex).setValue(mainFolderUrl);
    sheet.getRange(targetRow, progettoColumnIndex).setValue(subfolderUrls['progetto']);
    sheet.getRange(targetRow, documentiColumnIndex).setValue(subfolderUrls['documenti']);
    sheet.getRange(targetRow, allegatiColumnIndex).setValue(subfolderUrls['allegati']);
    sheet.getRange(targetRow, contrattoColumnIndex).setValue(subfolderUrls['contratto']);
  } else {
    throw new Error("Non è stato possibile trovare una riga con l'ID specificato.");
  }
}
