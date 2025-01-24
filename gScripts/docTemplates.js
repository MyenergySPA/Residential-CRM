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