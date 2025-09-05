/**
 * v2.1 — Log rapido: cerca l’ultima riga per appID leggendo solo la colonna appID.
 * - Niente getDataRange (evita di caricare tutto il foglio "offerte")
 * - Scansione dal fondo per prendere l’ultima occorrenza (la più recente)
 * - Carica una sola riga completa quando trova la corrispondenza
 */
function newLogDatiTecnici(nuovoFileDatiTecnici, appID) {
  Logger.log('Aggiornamento log dati tecnici per appID: ' + appID);

  if (typeof sheetOfferte === 'undefined' || !sheetOfferte) {
    throw new Error("La variabile globale 'sheetOfferte' non è definita. Inizializzala in 'main' prima di chiamare questa funzione.");
  }

  const sh = sheetOfferte;
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2) throw new Error('Il foglio "offerte" non contiene dati.');

  // 1) Trova l’indice colonna "appID" dall’intestazione (riga 1)
  const header = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  const appIDColIndex = header.indexOf('appID') + 1; // 1-based
  if (appIDColIndex <= 0) throw new Error('Colonna "appID" non trovata nel foglio "offerte".');

  // 2) Leggi SOLO la colonna appID (dalla riga 2 in giù) e cerca DAL FONDO
  const colValues = sh.getRange(2, appIDColIndex, lastRow - 1, 1).getValues(); // [[val], [val], ...]
  let foundRowNum = 0; // 1-based riga intera del foglio
  for (let i = colValues.length - 1; i >= 0; i--) {
    if (String(colValues[i][0]).trim() === String(appID).trim()) {
      foundRowNum = i + 2; // +2 perché partiamo dalla riga 2
      break;
    }
  }
  if (!foundRowNum) throw new Error('Nessuna riga trovata con appID: ' + appID);

  // 3) Carica SOLO la riga completa trovata
  const selectedRow = sh.getRange(foundRowNum, 1, 1, lastCol).getValues()[0];

  Logger.log('Riga selezionata per appID alla riga ' + foundRowNum + ': ' + JSON.stringify(selectedRow));

  // 4) Scrivi nel nuovo file "dati tecnici"
  const nuovoSheet = nuovoFileDatiTecnici.getActiveSheet();
  const ultimaRigaVuota = nuovoSheet.getLastRow() + 1;

  // Prima cella = appID-rigaLog (mantieni la tua semantica)
  selectedRow[0] = appID + '-' + ultimaRigaVuota;

  // Un solo setValues per tutta la riga
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


  // valori numerici puri
  const nConsumi = parseItNumber(consumi_annui);
  const nPrezzo  = parseItNumber(prezzo_energia);


  // scrivi NUMERI, non stringhe formattate
  sheetAnalisiEnergetica.getRange('consumi_annui').setValue(nConsumi);
  sheetAnalisiEnergetica.getRange('profilo_di_consumo').setValue(profilo_di_consumo);
  sheetAnalisiEnergetica.getRange('provincia').setValue(provincia);
  sheetAnalisiEnergetica.getRange('esposizione').setValue(esposizione);
  sheetAnalisiEnergetica.getRange('prezzo_energia').setValue(nPrezzo);

  // forza il formato di visualizzazione (evita % / interpretazioni strane)
  sheetAnalisiEnergetica.getRange('consumi_annui').setNumberFormat('#,##0');       // kWh
  sheetAnalisiEnergetica.getRange('prezzo_energia').setNumberFormat('0.00 "€"');  // €/kWh

  // (opzionale) diagnostica: intercetta numeri improbabili
  if (nPrezzo > 2) Logger.log('[WARN] prezzo_energia anomalo: ' + nPrezzo);

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

// Esegui la chiamata API con CACHE (fino a 6 ore)
let pvgisData = fetchWithCache_(apiUrl, 21600);

// Validazione struttura
if (!pvgisData || !pvgisData.outputs || !pvgisData.outputs.monthly) {
  throw new Error("Struttura risposta PVGIS inattesa: " + JSON.stringify(pvgisData).slice(0, 200));
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
    incentivo_effettivo: sheetAnalisiEnergetica.getRange('incentivo_effettivo').getValue(),
    massimale: sheetAnalisiEnergetica.getRange('massimale').getValue(),
    rata_mensile: sheetAnalisiEnergetica.getRange('rata_mensile').getValue(),
    produzione_primo_anno: sheetAnalisiEnergetica.getRange('produzione_primo_anno').getValue()
  };

  Logger.log('Risultati finali dell\'analisi energetica: ' + JSON.stringify(results));

  // Restituisci i risultati
  return results;
}

function fetchWithCache_(url, ttlSeconds) {
  var cache = CacheService.getScriptCache();
  var key = 'PV_' + Utilities.base64Encode(url).slice(0, 80);

  // cache hit
  var hit = cache.get(key);
  if (hit) {
    try { return JSON.parse(hit); } catch (_) { /* ignora e rifai fetch */ }
  }

  // fetch live
  var res  = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
  var code = res.getResponseCode();
  var text = res.getContentText();

  if (code >= 400) {
    // NON mettere in cache le risposte di errore
    throw new Error('HTTP ' + code + ' da PVGIS. Body: ' + text.slice(0, 200));
  }

  var json;
  try {
    json = JSON.parse(text);
  } catch (e) {
    // NON mettere in cache risposte non-JSON
    throw new Error('Risposta PVGIS non JSON. Body: ' + text.slice(0, 200));
  }

  // metti in cache massimo 6 ore
  cache.put(key, JSON.stringify(json), Math.min(ttlSeconds || 21600, 21600));
  return json;
}
