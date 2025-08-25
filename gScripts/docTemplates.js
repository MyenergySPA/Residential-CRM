// versione 2.0
// Definizione dei templates doc di offerta
 TEMPLATES = {
  presentazione_Finanz: '1zMIjekT-K_JWssZidSBjSuog_LfcHZjMLcEbePnP_t8',
  offerta_MAT: '1gMJGZZA7LwdugXKEFTK5LbJU2iiIIs6Ee5zBnlW81es',
  presentazione: '1XYDLbJymoNqU8B1nYqJm0k52-SU5O19G1Xzph_rjShg',
  offerta: '1_PNr5Y6svOADvgKZIjFjKsoDFpNV6TkOxivLIVqcZdA',
  offerta_REDEN: '1Pl16i3gROkfvxXYSIi8C1q9paqZ6xo7LF5W1isZcXFA',
  presentazione_REDEN: '1tGpVVMZL1Yuw5JIlytc-KiqxeQxd7Wegj74fLj_snMg',
  offerta_GSE: '1t5S9CYogDPAtKhy2ejMVELKjAkkieqfu31eIFF06GYg',
  offerta_Finanz: '1RCr8lgM98ryQwMiGFqecMHiWgIHsPN0tfV5HN82eYr4',
  presentazione_COND: '1iIom4lvX5ymUtQYucba2QyPiwOY94ySErcX19BowMrg',
  presentazione_Finanz_COND: '1pUcvT8cJtJgXs_ax-vExw5nfKraRqSsBl6gfn63dX7k',
  presentazione_PNRR: '1l-2iiXSI5mtPuEE6f5dWSUSKkHacHxNQKxtrDXQ0NMQ' ,
  offerta_PNRR: '1dEfWpHHEAyUjz8eJEHhdGP0hqrpnHpZV05KYad0EMq8',
  presentazione_finanz_PNRR: '1LWzIEdTpzKVlq80FvO8JJ9XBSoEkVLcanC4mvC5_se8',
  offerta_finanz_PNRR: '17S0KGjkbExL_vCBKevTUwpfAcBl_J7hkJkgdDroQdeA',
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
      templateId: TEMPLATES.offerta_MAT,
      nomeFile: ({ nome_referente, cognome_referente, dataOggi }) => `Offerta Myenergy ${nome_referente} ${cognome_referente} ${dataOggi}`
    }
  ],

  ADD: [
    {
      // Per "ADD", uguale a "MAT"
      templateId: TEMPLATES.offerta_MAT,
      nomeFile: ({ nome_referente, cognome_referente, dataOggi }) => `Offerta Myenergy ${nome_referente} ${cognome_referente} ${dataOggi}`
    }
  ], 

  COND: [
    {
      // Per "COND": documento di presentazione
      templateId: ({ tipo_pagamento }) =>
        tipo_pagamento === "Finanziamento" ? TEMPLATES.presentazione_Finanz_COND : TEMPLATES.presentazione_COND,
      nomeFile: ({ nome_referente, cognome_referente }) => `Presentazione offerta Myenergy ${nome_referente} ${cognome_referente}`
    },
    {
      // Per "COND": documento offerta
      templateId: ({ tipo_pagamento }) =>
        tipo_pagamento === "Finanziamento" ? TEMPLATES.offerta_Finanz : TEMPLATES.offerta,
      nomeFile: ({ tipo_opportunita, id, yy, nome_referente, cognome_referente, dataOggi }) =>
        `Offerta ${tipo_opportunita}-${id}-${yy} ${nome_referente} ${cognome_referente} ${dataOggi}`
    }
  ],

  // configuration for RES subtypes
  // se la condizione non viene soddisfatta, viene scelto il modello RES standard
   DEFAULT: [
    {
      templateId: ({ tipo_incentivo, tipo_pagamento }) => {
        
        // "ifs for 'presentazione'"
        
        if (tipo_incentivo === "PNRR" && tipo_pagamento === "Finanziamento") {
          return TEMPLATES.presentazione_finanz_PNRR;
        }

        if (tipo_incentivo === "PNRR") {
          return TEMPLATES.presentazione_PNRR;
        }

        if (tipo_incentivo === "REDEN") {
          return TEMPLATES.presentazione_REDEN;
        }

        // 2) se è finanziamento → finanziamento
        if (tipo_pagamento === "Finanziamento") {
          return TEMPLATES.presentazione_Finanz;
        }

        // 3) altrimenti → standard
        return TEMPLATES.presentazione;
      },
      nomeFile: ({ nome_referente, cognome_referente }) => `Presentazione offerta Myenergy ${nome_referente} ${cognome_referente}`
    },
    {
      templateId: ({ tipo_incentivo, tipo_pagamento }) => {

        // "ifs for 'offerta'"

        if (tipo_incentivo === "PNRR" && tipo_pagamento === "Finanziamento") {
          return TEMPLATES.offerta_finanz_PNRR;
        }

        if (tipo_incentivo === "PNRR") {
          return TEMPLATES.offerta_PNRR;
        }

        if (tipo_incentivo === "REDEN") {
          return TEMPLATES.offerta_REDEN;
        }

        if (tipo_pagamento === "Finanziamento") {
          return TEMPLATES.offerta_Finanz;
        }

      return TEMPLATES.offerta;
      },
      nomeFile: ({ tipo_opportunita, id, yy, nome_referente, cognome_referente, dataOggi }) =>
        `Offerta ${tipo_opportunita}-${id}-${yy} ${nome_referente} ${cognome_referente} ${dataOggi}`
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
function determineDocumentTemplates(tipo_opportunita, tipo_pagamento, tipo_incentivo, nome_referente, cognome_referente, dataOggi, id, yy) {
  Logger.log(`Determinazione dei template per tipo_opportunita: ${tipo_opportunita}, tipo_pagamento: ${tipo_pagamento}, tipo_incentivo: ${tipo_incentivo}`);
  
  // Raggruppiamo tutti i parametri in un unico oggetto
  const params = { tipo_opportunita, tipo_pagamento, tipo_incentivo, nome_referente, cognome_referente, dataOggi, id, yy };

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
