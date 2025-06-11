  /**
 * Crea una mappatura dei segnaposto con i valori forniti come input.
 *
 * @param {Object} params - Oggetto contenente tutti i parametri necessari per la mappatura.
 * @returns {Object} Mappatura dei segnaposto con i relativi valori.
 */
function createPlaceholderMapping(params) {
  Logger.log('Creazione della mappatura dei segnaposto.');
  Logger.log('Valore originale di detrazione: ' + params.detrazione + ' (tipo: ' + typeof params.detrazione + ')');
  Logger.log('Valore originale di iva_offerta: ' + params.iva_offerta + ' (tipo: ' + typeof params.iva_offerta + ')');
  Logger.log('Valore originale di prezzo_offerta: ' + params.prezzo_offerta + ' (tipo: ' + typeof params.prezzo_offerta + ')');

  return {
    '{{tipo_opportunità}}': params.tipo_opportunita,
    '{{id}}': params.id,
    '{{yy}}': params.yy,
    '{{nome}}': params.nome,
    '{{cognome}}': params.cognome,
    '{{indirizzo}}': params.indirizzo,
    '{{telefono}}': params.telefono,
    '{{email}}': params.email,
    '{{data ultima modifica}}': params.dataOggi,
    '{{numero_moduli}}': params.numero_moduli,
    '{{marca_moduli}}': params.marca_moduli,
    '{{numero_inverter}}': params.numero_inverter,
    '{{marca_inverter}}': params.marca_inverter,
    '{{numero_batteria}}': params.numero_batteria,
    '{{capacità batteria}}': params.capacita_batteria,
    '{{totale_capacità_batterie}}': params.totale_capacita_batterie,
    '{{marca_batteria}}': params.marca_batteria,
    '{{tetto}}': params.tetto,
    '{{potenza_impianto}}': formatNumber(params.potenza_impianto, 2),
    '{{produzione_primo_anno}}': formatNumber(params.produzione_primo_anno, 0),
    '{{alberi}}': formatNumber(params.alberi, 0),
    '{{testo_aggiuntivo}}': params.testo_aggiuntivo,
    '{{tipo_pagamento}}': params.tipo_pagamento,
    '{{condizione_pagamento_1}}': params.condizione_pagamento_1,
    '{{condizione_pagamento_2}}': params.condizione_pagamento_2,
    '{{condizione_pagamento_3}}': params.condizione_pagamento_3,
    '{{condizione_pagamento_4}}': params.condizione_pagamento_4,
    '{{imponibile_offerta}}': formatCurrency(params.imponibile_offerta),
    '{{iva_offerta}}': formatCurrency(params.iva_offerta, 0),
    '{{iva_percentuale}}': params.iva_percentuale *= 100,
    '{{prezzo_offerta}}': formatCurrency(params.prezzo_offerta),
    '{{anni_finanziamento}}': formatNumber(params.anni_finanziamento, 0),
    '{{rata_mensile}}': formatCurrency(params.rata_mensile),
    '{{numero_rate_mensili}}': formatNumber(params.numero_rate_mensili, 0),
    '{{esposizione}}': params.esposizione,
    '{{area_m2_impianto}}': formatNumber(params.area_m2_impianto, 2),
    '{{scheda_tecnica_moduli}}': 'Link scheda tecnica moduli',
    '{{scheda_tecnica_inverter}}': 'Link scheda tecnica inverter',
    '{{scheda_tecnica_batterie}}': 'Link scheda tecnica batterie',
    '{{scheda_tecnica_ottimizzatori}}': 'Link scheda tecnica ottimizzatori',
    '{{numero_colonnina_74kw}}': params.numero_colonnina_74kw,
    '{{numero_colonnina_22kw}}': params.numero_colonnina_22kw,
    '{{numero_ottimizzatori}}': params.numero_ottimizzatori,
    '{{marca_ottimizzatori}}': params.marca_ottimizzatori,
    '{{numero_linea_vita}}': params.numero_linea_vita,
    '{{detrazione}}': formatCurrency(params.detrazione),
    '{{anni_ritorno_investimento}}': formatNumber(params.anni_ritorno_investimento, 1),
    '{{utile_25_anni}}': formatCurrency(params.utile_25_anni),
    '{{percentuale_autoconsumo}}': formatPercentage(params.percentuale_autoconsumo),
    '{{media_vendita}}': formatCurrency(params.media_vendita),
    '{{prezzo_energia}}': formatCurrency(params.prezzo_energia),
    '{{percentuale_risparmio_energetico}}': formatPercentage(params.percentuale_risparmio_energetico),
    '{{durata_incentivo}}': formatNumber(params.durata_incentivo, 0),
    '{{acconto_diretto}}' : formatCurrency(params.acconto_diretto),
    '{{condominio}}' : params.condominio,
    '{{garanzia_moduli}}' : params.garanzia_moduli,
    '{{garanzia_inverter}}' : params.garanzia_inverter,
    '{{garanzia_batterie}}' : params.garanzia_batterie,
    '{{tipo_incentivo}}' : params.tipo_incentivo,
    '{{massimale}}': formatCurrency(params.massimale),

  };
}


/*
 * Funzione per sostituire i segnaposto con i valori nel documento.
 * Sostituisce i valori vuoti con un carattere indicato.
 *
 * @param {Body} docBody - Il corpo del documento Google Docs.
 * @param {Object} placeholders - Oggetto mappa che contiene i segnaposto e i loro valori.
 */
function replacePlaceholders(docBody, placeholders) {
  Logger.log('Replacing text placeholders in the document.');

  // Replace text placeholders
  for (const key in placeholders) {
    let value = placeholders[key];

    // Handle undefined or null values
    if (value === undefined || value === null) {
      value = ""; // Replace with an empty string
    }

    // Log each key-value pair for debugging
    Logger.log(`Replacing placeholder: ${key} with value: ${value}`);

    try {
      docBody.replaceText(key, value);
    } catch (error) {
      Logger.log(`Error replacing placeholder ${key} with value ${value}: ${error.message}`);
      throw error; // Rethrow the error for visibility
    }
  }
}