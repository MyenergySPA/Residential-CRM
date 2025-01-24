/**
 * Global array of chart definitions, specifying which charts
 * to upload from the "dati tecnici" spreadsheet and how they
 * map to placeholders in your offer docs.
 */

var CHART_DEFINITIONS = [
  {
    // e.g. from 'analisi energetica' sheet
    sheetName: 'calcoli',
    chartIndex: 0,
    placeholder: '{{CHART_RITORNO_25_ANNI}}',
    customFileName: 'chart_ritorno_25_anni'
  },
  {
    // e.g. from 'analisi energetica' sheet
    sheetName: 'analisi energetica',
    chartIndex: 0,
    placeholder: '{{test}}',
    customFileName: 'chart_test'
  },
  // Add as many chart definitions as needed
];



/**
 * Uploads all charts from all sheets in the 'nuovoFileDatiTecnici' spreadsheet
 * to the 'cartellaProgettoId' folder in Drive.
 *
 * @param {Spreadsheet} nuovoFileDatiTecnici - The "dati tecnici" Spreadsheet object.
 * @param {string} cartellaProgettoId - The folder ID where charts will be stored.
 * @param {Array} chartDefinitions - (Optional) Array of definitions for each chart placeholder.
 *        Each entry can look like:
 *        {
 *          sheetName: 'NameOfSheet',
 *          chartIndex: 0, // or 1, 2, ...
 *          placeholder: '{{CHART_EXAMPLE}}',
 *          customFileName: 'My Special Chart'
 *        }
 *        If provided, only these specific charts will be exported. Otherwise, all charts in all sheets are exported.
 *
 * @return {Array} Array of objects describing each exported chart, with fields:
 *         {
 *           placeholder: string,   // e.g. '{{CHART_EXAMPLE}}'
 *           fileId: string,        // The ID of the uploaded PNG in Drive
 *           blob: Blob,           // The chart image blob (you can insert this directly into a Doc)
 *           sheetName: string,
 *           chartIndex: number
 *         }
 */

function newCharts(nuovoFileDatiTecnici, cartellaProgettoId, chartDefinitions) {
  const folderProgetto = DriveApp.getFolderById(cartellaProgettoId);
  const sheets = nuovoFileDatiTecnici.getSheets();
  let results = [];

  // If chartDefinitions is provided, we assume you only want *specific* charts.
  // If not provided, we will export *all* charts from all sheets.
  const hasDefinitions = Array.isArray(chartDefinitions) && chartDefinitions.length > 0;

  if (hasDefinitions) {
    // Loop over each chart definition
    chartDefinitions.forEach(def => {
      const sheetName = def.sheetName;
      const chartIndex = def.chartIndex;
      const placeholder = def.placeholder;
      const customFileName = def.customFileName || (sheetName + ' - Chart ' + chartIndex);

      const sheet = nuovoFileDatiTecnici.getSheetByName(sheetName);
      if (!sheet) {
        Logger.log(`Sheet "${sheetName}" not found. Skipping.`);
        return;
      }
      const charts = sheet.getCharts();
      if (chartIndex < 0 || chartIndex >= charts.length) {
        Logger.log(`Chart index ${chartIndex} out of range in sheet "${sheetName}". Skipping.`);
        return;
      }

      const chart = charts[chartIndex];
      const blob = chart.getAs('image/png');
      blob.setName(customFileName + '.png');

      // Upload the file into Drive
      const file = folderProgetto.createFile(blob);
      Logger.log('Uploaded chart file: ' + file.getName() + ' (ID: ' + file.getId() + ')');

      results.push({
        placeholder: placeholder,
        fileId: file.getId(),
        blob: blob,
        sheetName: sheetName,
        chartIndex: chartIndex
      });
    });

  } else {
    // No chart definitions provided: export *all* charts from every sheet
    sheets.forEach(sheet => {
      const sheetName = sheet.getName();
      const charts = sheet.getCharts();

      charts.forEach((chart, index) => {
        const defaultName = sheetName + ' - Chart ' + (index + 1);
        const blob = chart.getAs('image/png').setName(defaultName + '.png');

        // Upload to Drive
        const file = folderProgetto.createFile(blob);
        Logger.log('Uploaded chart file: ' + file.getName() + ' (ID: ' + file.getId() + ')');

        // In a full app, you'd map this to a placeholder or keep track of it
        results.push({
          placeholder: '', // or something like '{{CHART_' + sheetName + '_' + (index+1) + '}}'
          fileId: file.getId(),
          blob: blob,
          sheetName: sheetName,
          chartIndex: index
        });
      });
    });
  }

  return results;
}