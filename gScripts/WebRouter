/** Router unico della Web App */
function doGet(e) {
  try {
    const p = e.parameter || {};
    const mode = (p.mode || '').toLowerCase();

    if (mode === 'export' || (!mode && p.projectFolderId)) {
      return handleExport(p);
    }
    if (mode === 'email' || (!mode && p.recipiente)) {
      return handleEmail(p);
    }

    return ContentService.createTextOutput(
      'Errore: specifica mode=export oppure mode=email'
    ).setMimeType(ContentService.MimeType.TEXT);

  } catch (err) {
    Logger.log('Router doGet error: ' + (err && err.stack ? err.stack : err));
    return ContentService.createTextOutput('Errore: ' + String(err))
      .setMimeType(ContentService.MimeType.TEXT);
  }
}
