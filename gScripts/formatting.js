/**
 * formatting.gs v1
 * Formats a number using the Italian locale.
 * @param {number} value - The number to format.
 * @param {number} decimals - The number of decimal places.
 * @returns {string} The formatted number as a string.
 */
function formatNumberItalian(value, decimals) {
  return new Intl.NumberFormat('it-IT', {
    minimumFractionDigits: decimals,
    maximumFractionDigits: decimals,
  }).format(value);
}



/**
 * Formatta un numero con il numero specificato di decimali.
 *
 * @param {number} value - Numero da formattare.
 * @param {number} decimals - Numero di decimali.
 * @returns {string} Numero formattato.
 */
function formatNumber(value, decimals) {

  if (value == null || value === '') {
    return '-';
  }

  // Rimuovi tutti i caratteri non numerici, tranne punto e virgola
  value = value.toString().replace(/[^\d.,-]/g, '');


  // Gestione dei separatori decimali
  if (value.includes(',') && value.includes('.')) {
    value = value.replace(/\./g, '');
    value = value.replace(',', '.');
  } else if (value.includes(',')) {
    value = value.replace(',', '.');
  }

  value = parseFloat(value);


  if (isNaN(value)) {
    return '-';
  }

  // Formattazione personalizzata
  let formattedValue = value.toFixed(decimals);
  let parts = formattedValue.split('.');
  parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, '.');
  formattedValue = parts.join(',');


  return formattedValue;
}



/**
 * Formatta un valore come valuta.
 *
 * @param {number} value - Valore da formattare.
 * @returns {string} Valore formattato come valuta.
 */
function formatCurrency(value) {

  if (value == null || value === '') {
    return '-';
  }

  // Rimuovi tutti i caratteri non numerici, tranne punto e virgola
  value = value.toString().replace(/[^\d.,-]/g, '');

  // Gestione dei separatori decimali
  if (value.includes(',') && value.includes('.')) {
    // Caso in cui ci sono sia punto che virgola (es. "1.231,82")
    value = value.replace(/\./g, '');
    value = value.replace(',', '.');
  } else if (value.includes(',')) {
    // Caso in cui c'è solo la virgola (es. "1231,82")
    value = value.replace(',', '.');
  }


  value = parseFloat(value);


  if (isNaN(value)) {
    return '-';
  }

  // Formattazione personalizzata
  let parts = value.toFixed(2).split('.');
  parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, '.');
  let formattedValue = parts.join(',');
  formattedValue += ' €';

  return formattedValue;
}



/**
 * Formatta un valore come percentuale.
 *
 * @param {number} value - Valore da formattare.
 * @returns {string} Valore formattato come percentuale.
 */
function formatPercentage(value) {
  return new Intl.NumberFormat('it-IT', { style: 'percent', minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(value);
}



/**
 * Aggiunge un hyperlink a un testo specificato all'interno di un documento.
 *
 * @param {Body} corpo - Corpo del documento.
 * @param {string} searchText - Testo da cercare e sostituire con un link.
 * @param {string} url - URL da collegare al testo.
 */
function addHyperlink(corpo, searchText, url) {
  Logger.log('Aggiunta di un hyperlink al testo: ' + searchText);
  let foundElement = corpo.findText(searchText);
  while (foundElement) {
    const foundText = foundElement.getElement().asText();
    const startOffset = foundElement.getStartOffset();
    const endOffset = foundElement.getEndOffsetInclusive();
    foundText.setLinkUrl(startOffset, endOffset, url);
    foundElement = corpo.findText(searchText, foundElement);
  }
  Logger.log('Hyperlink aggiunto.');
}

//numeri in → numero JS out
function parseItNumber(v) {
  if (typeof v === 'number') return v;
  const s = String(v || '').replace(/\./g, '').replace(',', '.').replace(/[^\d.-]/g, '');
  const n = Number(s);
  if (!isFinite(n)) throw new Error('Valore numerico non valido: ' + v);
  return n;
}


// numeri + format
function writeNumberWithFormat(sh, a1, value, numberFormat) {
  const n = parseItNumber(value);
  const r = sh.getRange(a1);
  r.setValue(n);
  if (numberFormat) r.setNumberFormat(numberFormat);
  return n;
}
