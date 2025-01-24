function newId() {
  Logger.log('Inizio della funzione generaID');

  // Imposta l'ID del foglio "database CRM" e il nome del foglio
  var sheetID = '1_QEo5ynx_29j3I3uJJff5g7ZzGZJnPcIarIXfr5O2gQ';
  var sheetName = 'cronologia';
  Logger.log('sheet "database CRM" id: ' + sheetID);
  Logger.log('Nome del foglio: ' + sheetName);
  
  var sheet = SpreadsheetApp.openById(sheetID).getSheetByName(sheetName);
  Logger.log('Foglio aperto correttamente');

  // Trova la colonna "id"
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  Logger.log('Headers trovati: ' + headers);
  
  var idColumn = headers.indexOf('id') + 1;
  Logger.log('Indice della colonna "id": ' + idColumn);

  if (idColumn === 0) {
    Logger.log('Errore: Header "id" non trovato.');
    throw new Error('Header "id" non trovato.');
  }

  // Ottieni tutti gli ID esistenti
  var data = sheet.getRange(2, idColumn, sheet.getLastRow() - 1).getValues();
  var existingIDs = data.flat();
  Logger.log('ID esistenti: ' + existingIDs);

  // Funzione per generare un ID a 4 cifre
  function generateID() {
    var id = '';
    for (var i = 0; i < 4; i++) {
      id += Math.floor(Math.random() * 10);
    }
    return id;
  }

  // Genera un ID univoco
  var newID;
  do {
    newID = generateID();
    Logger.log('Generato nuovo ID: ' + newID);
  } while (existingIDs.includes(newID));

  Logger.log('ID univoco generato: ' + newID);
  return newID;
}

function impostaID() {
  Logger.log('Inizio della funzione impostaID');

  // Genera un nuovo ID univoco
  var newID = generaID();
  Logger.log('Nuovo ID generato: ' + newID);

  // Restituisci il nuovo ID
  return newID;
}
