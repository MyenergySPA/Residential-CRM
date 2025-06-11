// Definizione dei templates doc di offerta
 TEMPLATES = {
  presentazioneFinanz: '1zMIjekT-K_JWssZidSBjSuog_LfcHZjMLcEbePnP_t8',
  offertaMAT: '1gMJGZZA7LwdugXKEFTK5LbJU2iiIIs6Ee5zBnlW81es',
  presentazione: '1XYDLbJymoNqU8B1nYqJm0k52-SU5O19G1Xzph_rjShg',
  offerta: '1_PNr5Y6svOADvgKZIjFjKsoDFpNV6TkOxivLIVqcZdA',
  offertaREDEN: '1Pl16i3gROkfvxXYSIi8C1q9paqZ6xo7LF5W1isZcXFA',
  presentazioneREDEN: '1tGpVVMZL1Yuw5JIlytc-KiqxeQxd7Wegj74fLj_snMg',
  offertaGSE: '1t5S9CYogDPAtKhy2ejMVELKjAkkieqfu31eIFF06GYg',
  offertaFinanz: '1RCr8lgM98ryQwMiGFqecMHiWgIHsPN0tfV5HN82eYr4',
  presentazioneCOND: '1iIom4lvX5ymUtQYucba2QyPiwOY94ySErcX19BowMrg',
  presentazioneFinanzCOND: '1pUcvT8cJtJgXs_ax-vExw5nfKraRqSsBl6gfn63dX7k',
  presentazione_PNRR: '1l-2iiXSI5mtPuEE6f5dWSUSKkHacHxNQKxtrDXQ0NMQ' ,
  offerta_PNRR: '1dEfWpHHEAyUjz8eJEHhdGP0hqrpnHpZV05KYad0EMq8'

  }

/**
 * Configurazione per i documenti da creare in base al tipo di opportunità.
 * Per ogni tipo di opportunità (ad es. "MAT", "COND", …) viene definito un array di documenti.
 * Ogni documento ha:
 *  - templateId: un valore fisso o una funzione che, in base ai parametri, restituisce l'ID del template.
 *  - nomeFile: una funzione che genera il nome del file.
 */
const DOCUMENT_CONFIG = {

  MAT: [
    {
      // Per "MAT" si crea solo l'offerta (senza presentazione)
      templateId: TEMPLATES.offertaMAT,
      nomeFile: ({ nome, cognome, dataOggi }) => `Offerta Myenergy ${nome} ${cognome} ${dataOggi}`
    }
  ],

  ADD: [
    {
      // Per "ADD", uguale a "MAT"
      templateId: TEMPLATES.offertaMAT,
      nomeFile: ({ nome, cognome, dataOggi }) => `Offerta Myenergy ${nome} ${cognome} ${dataOggi}`
    }
  ], 

  COND: [
    {
      // Per "COND": documento di presentazione
      templateId: ({ tipo_pagamento }) =>
        tipo_pagamento === "Finanziamento" ? TEMPLATES.presentazioneFinanzCOND : TEMPLATES.presentazioneCOND,
      nomeFile: ({ nome, cognome }) => `Presentazione offerta Myenergy ${nome} ${cognome}`
    },
    {
      // Per "COND": documento offerta
      templateId: ({ tipo_pagamento }) =>
        tipo_pagamento === "Finanziamento" ? TEMPLATES.offertaFinanz : TEMPLATES.offerta,
      nomeFile: ({ tipo_opportunita, id, yy, nome, cognome, dataOggi }) =>
        `Offerta ${tipo_opportunita}-${id}-${yy} ${nome} ${cognome} ${dataOggi}`
    }
  ],

  REDEN: [
    {
      // Documento di presentazione
      templateId: TEMPLATES.presentazioneREDEN,
      nomeFile: ({ nome, cognome }) => `Presentazione offerta Myenergy ${nome} ${cognome}`
    },
    {
      // Documento offerta
      templateId: TEMPLATES.offertaREDEN,
      nomeFile: ({ tipo_opportunita, id, yy, nome, cognome, dataOggi }) =>
        `Offerta ${tipo_opportunita}-${id}-${yy} ${nome} ${cognome} ${dataOggi}`
    }
  ],

  // Configurazione per RES e varianti simili
  // se la condizione non viene soddisfatta, viene scelto il modello RES standard
   DEFAULT: [
    {
      templateId: ({ tipo_incentivo, tipo_pagamento }) => {
        // 1) se è PNRR → PNRR
        if (tipo_incentivo === "PNRR") {
          return TEMPLATES.presentazione_PNRR;
        }
        // 2) se è finanziamento → finanziamento
        if (tipo_pagamento === "Finanziamento") {
          return TEMPLATES.presentazioneFinanz;
        }
        // 3) altrimenti → standard
        return TEMPLATES.presentazione;
      },
      nomeFile: ({ nome, cognome }) => `Presentazione offerta Myenergy ${nome} ${cognome}`
    },
    {
      templateId: ({ tipo_incentivo, tipo_pagamento }) => {
        if (tipo_incentivo === "PNRR") {
          return TEMPLATES.offerta_PNRR;
        }
        if (tipo_pagamento === "Finanziamento") {
          return TEMPLATES.offertaFinanz;
        }
        return TEMPLATES.offerta;
      },
      nomeFile: ({ tipo_opportunita, id, yy, nome, cognome, dataOggi }) =>
        `Offerta ${tipo_opportunita}-${id}-${yy} ${nome} ${cognome} ${dataOggi}`
    }
  ]
};

/**
 * Determina i template dei documenti da utilizzare in base al tipo di opportunità e al tipo di pagamento.
 * La logica è interamente definita in DOCUMENT_CONFIG, rendendo il codice future-proof.
 *
 * @param {string} tipo_opportunita - Tipo di opportunità (es. "MAT", "COND", …).
 * @param {string} tipo_pagamento - Tipo di pagamento (es. "Finanziamento", …).
 * @param {string} nome - Nome del cliente.
 * @param {string} cognome - Cognome del cliente.
 * @param {string} dataOggi - Data corrente formattata.
 * @param {string} id - ID dell'offerta.
 * @param {string} yy - Anno.
 * @returns {Array} Array di oggetti contenenti templateId e nomeFile per ogni documento da creare.
 */
function determineDocumentTemplates(tipo_opportunita, tipo_pagamento, tipo_incentivo, nome, cognome, dataOggi, id, yy) {
  Logger.log(`Determinazione dei template per tipo_opportunita: ${tipo_opportunita}, tipo_pagamento: ${tipo_pagamento}, tipo_incentivo: ${tipo_incentivo}`);
  
  // Raggruppiamo tutti i parametri in un unico oggetto
  const params = { tipo_opportunita, tipo_pagamento, tipo_incentivo, nome, cognome, dataOggi, id, yy };

  // Se esiste una configurazione specifica per il tipo_opportunita, la usiamo, altrimenti usiamo quella DEFAULT
  const config = DOCUMENT_CONFIG[tipo_opportunita] || DOCUMENT_CONFIG.DEFAULT;

  // Generiamo l'array dei documenti da creare, valutando le funzioni se presenti
  const templates = config.map(docConfig => {
    const templateId = typeof docConfig.templateId === 'function'
      ? docConfig.templateId(params)
      : docConfig.templateId;
    const nomeFile = typeof docConfig.nomeFile === 'function'
      ? docConfig.nomeFile(params)
      : docConfig.nomeFile;
    return { templateId, nomeFile };
  });

  Logger.log(`Template selezionati: ${JSON.stringify(templates)}`);
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

