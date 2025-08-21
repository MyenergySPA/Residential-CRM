/***** CONFIG MINIMA (legge il nome dal Main se presente) *****/
const NOME_FILE_DATI_TECNICI = (typeof nomeFileDatiTecnici !== 'undefined' && nomeFileDatiTecnici)
  ? nomeFileDatiTecnici                  // es. 'dati tecnici 6.5' dal tuo Main
  : 'dati tecnici 6.5';

const NOME_RANGE_BUDGET        = 'budget';
const RANGE_OFFERTA_ANALIZZATA = 'offerta_analizzata';

/** HANDLER EXPORT: salva in Drive ma scarica direttamente dal browser */
function handleExport(p) {
  try {
    const projectFolderId = _extractFolderId(p.projectFolderId);
    const appID           = (p.appID || '').trim();

    if (!projectFolderId) throw new Error('Parametro "projectFolderId" mancante o non valido (passa l’URL/ID della cartella "progetto").');
    if (!appID)           throw new Error('Parametro "appID" mancante.');

    const res = exportBudgetCsvForOffer({ projectFolderId, appID }); // -> crea file in Drive + ritorna {csv,name,url,...}

    // Pagina che avvia il download via Blob + <a download> (niente passaggio su Drive)
    const html = `
<!DOCTYPE html><html lang="it"><head><meta charset="utf-8">
<title>Download CSV</title></head><body style="font-family:Verdana">
<script>
  (function(){
    const csv = ${JSON.stringify(res.csv)};         // già con BOM per Excel
    const name = ${JSON.stringify(res.name)};
    const blob = new Blob([csv], {type: 'text/csv;charset=utf-8;'});
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement('a');
    a.href = url;
    a.download = name;
    document.body.appendChild(a);
    a.click();
    setTimeout(function(){ URL.revokeObjectURL(url); }, 1000);
  })();
</script>
<p>Download avviato. Il file è stato salvato anche in Drive:<br>
<a href="${res.url}" target="_blank" rel="noopener">apri in Drive</a></p>
</body></html>`;
    return HtmlService.createHtmlOutput(html);
  } catch (err) {
    return ContentService.createTextOutput('Errore: ' + String(err))
      .setMimeType(ContentService.MimeType.TEXT);
  }
}

/** FUNZIONE PRINCIPALE (percorso unico, nessun fallback) */
function exportBudgetCsvForOffer({ projectFolderId, appID }) {
  // 1) Apri direttamente la cartella "progetto"
  let progettoFolder;
  try {
    progettoFolder = DriveApp.getFolderById(projectFolderId);
  } catch (e) {
    throw new Error('Cartella "progetto" non accessibile o inesistente (ID: ' + projectFolderId + ').');
  }

  // 2) Trova il file "dati tecnici <versione>" esatto
  const datiTecniciFile = _findSingleFileByName(progettoFolder, NOME_FILE_DATI_TECNICI);
  if (!datiTecniciFile) {
    throw new Error('File "' + NOME_FILE_DATI_TECNICI + '" non trovato nella cartella "' +
                    progettoFolder.getName() + '" (ID: ' + progettoFolder.getId() + ').');
  }
  const ss = SpreadsheetApp.openById(datiTecniciFile.getId());

  // 3) Range necessari
  const rngOfferta = ss.getRangeByName(RANGE_OFFERTA_ANALIZZATA);
  if (!rngOfferta) throw new Error('Named range "' + RANGE_OFFERTA_ANALIZZATA + '" non trovato (foglio "analisi energetica").');
  const rngBudget  = ss.getRangeByName(NOME_RANGE_BUDGET);
  if (!rngBudget)  throw new Error('Named range "' + NOME_RANGE_BUDGET + '" non trovato (foglio "analisi energetica").');

  // 4) Ricava il token dal log usando appID (ultimo "APP-xxxx-<n>")
  const tokenDaUsare = _resolveLogTokenByAppId(ss, appID);
  if (!tokenDaUsare) {
    throw new Error('Nessuna riga trovata in "log!A:A" con prefisso "' + appID + '-".');
  }

  // 5) Forza temporaneamente offerta_analizzata → export → ripristina
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  const formulaOriginale = rngOfferta.getFormula();
  const valoreOriginale  = rngOfferta.getValue();

  try {
    rngOfferta.setFormula('');
    rngOfferta.setValue(tokenDaUsare);
    SpreadsheetApp.flush();
    Utilities.sleep(250);

    // ---- CREA CSV (con BOM) ----
    const values = rngBudget.getDisplayValues();
    const ts     = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
    const name   = 'Budget_' + tokenDaUsare + '_' + ts + '.csv';
    const body   = values.map(r => r.map(_csvEscape).join(',')).join('\r\n');

    // BOM per compatibilità Excel / caratteri accentati
    const csvWithBom = '\uFEFF' + body;

    // Salva in Drive (audit) e ritorna anche csv+name per il download diretto
    const file = progettoFolder.createFile(Utilities.newBlob(csvWithBom, 'text/csv', name));
    const url  = file.getUrl();

    return { url, fileId: file.getId(), name, logToken: tokenDaUsare, csv: csvWithBom };

  } finally {
    if (formulaOriginale) rngOfferta.setFormula(formulaOriginale);
    else rngOfferta.setValue(valoreOriginale);
    SpreadsheetApp.flush();
    lock.releaseLock();
  }
}

/***** HELPERS MINIMI *****/
function _extractFolderId(input) {
  if (!input) return null;
  const s = String(input);
  if (/^[A-Za-z0-9_-]{20,}$/.test(s)) return s; // già ID
  // supporta: /folders/ID  |  open?id=ID  |  file/d/ID  |  uc?id=ID
  const m = s.match(/(?:\/folders\/|[?&]id=|file\/d\/|uc\?id=)([A-Za-z0-9_-]{20,})/);
  return m ? m[1] : null;
}

function _findSingleFileByName(folder, name) {
  const files = folder.getFilesByName(name);
  return files.hasNext() ? files.next() : null;
}

function _csvEscape(val) {
  const s = String(val ?? '');
  const needsQuotes = /[",\r\n]/.test(s);
  const escaped = s.replace(/"/g, '""');
  return needsQuotes ? `"${escaped}"` : escaped;
}

function _resolveLogTokenByAppId(ss, appID) {
  const shLog = ss.getSheetByName('log');
  if (!shLog) throw new Error('Foglio "log" non trovato nel file dati tecnici.');
  const lastRow = shLog.getLastRow();
  if (lastRow < 1) return null;

  const colA = shLog.getRange(1, 1, lastRow, 1).getValues();
  const prefix = appID + '-';
  let found = null;
  for (let r = 0; r < colA.length; r++) {
    const v = String(colA[r][0] ?? '');
    if (v.startsWith(prefix)) found = v; // ultima occorrenza
  }
  return found;
}
