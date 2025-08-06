// -----------------------------------------------------------------------------
// Yousign Integration Script (Sandbox)
// Milan, Myenergy Solutions
// Pulito, modulare, commentato per facilità di manutenzione
// -----------------------------------------------------------------------------

// Credenziali e endpoint
const YOUSIGN_API_KEY = 'cYEnc4l5hDoIylY6vvsNCco0Gasg9WlW';
const YOUSIGN_API_BASE = 'https://api.yousign.app/v3';

/**
 * Wrapper esposto per test manuale.
 */
function eSignOfferta(
  nome, cognome, telefono, email, locale, folderLink,
  page = 4, x = 350, y = 650, width = 200, height = 50  // modificato: coordinate corrette per "FIRMA CLIENTE"
) {
  return sendToYousign(
    { nome, cognome, telefono, email, locale, folderLink },
    { page, x, y, width, height }
  );
}

/**
 * Crea e avvia una Signature Request configurata per 15 giorni,
 * nome dinamico personalizzato e body email personalizzato.
 */
function sendToYousign(
  { nome, cognome, telefono, email, locale, folderLink },
  { page, x, y, width, height }
) {
  // 1. Recupera il PDF e ottieni il nome base del file
  const pdfBlob = getOfferPdfBlob(folderLink);
  const originalPdfName = pdfBlob.getName();
  const documentBaseName = originalPdfName.replace(/\.pdf$/i, '');

  // 2. Costruisci il corpo dell'email
  const emailBody =
    `Buongiorno, ${nome} ${cognome}!\n\n` +
    `La tua offerta personalizzata è ora pronta!\n\n` +
    `Se l'offerta risponde alle tue aspettative, potrai confermarla firmando digitalmente, tramite il pulsante presente in questa email. Per qualsiasi dubbio o domanda, o per apportare modifiche alla tua offerta, non esitare a contattarci ai seguenti riferimenti:\n` +
    `Telefono: 3792610174\n` +
    `E-mail: residenziale@myenergy.it\n\n` +
    `Grazie di aver considerato Myenergy!\n\n`;

  // 3. Calcola la data di scadenza (solo YYYY-MM-DD)
  const expirationDate = new Date(
    Date.now() + 15 * 24 * 60 * 60 * 1000
  ).toISOString().slice(0, 10);

  // 4. Crea la richiesta in bozza con nome personalizzato e custom email
  const createResp = UrlFetchApp.fetch(
    `${YOUSIGN_API_BASE}/signature_requests`, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        name: `Offerta per ${nome} ${cognome}`,
        delivery_mode: 'email',
        timezone: 'Europe/Rome',
        expiration_date: expirationDate,
        signers: [{
          info: {
            first_name: nome,
            last_name: cognome,
            email: email,
            phone_number: telefono,
            locale: locale
          },
          signature_level: 'electronic_signature',
          signature_authentication_mode: 'otp_email',
          custom_text: {
            request_subject: `Offerta ${documentBaseName}`,
            request_body: emailBody
          }
        }]
      }),
      headers: { 'Authorization': `Bearer ${YOUSIGN_API_KEY}` },
      muteHttpExceptions: true
    }
  );
  const jsonReq = JSON.parse(createResp.getContentText());
  if (!jsonReq.id) throw new Error(`Errore creazione request: ${createResp.getContentText()}`);

  const requestId = jsonReq.id;
  const signerId = jsonReq.signers[0].id;

  // 5. Upload del documento con nome originale e creazione del campo firma
  //    Assicura che il blob abbia il nome originale del file PDF
  pdfBlob.setName(originalPdfName);  // modificato: preserva il nome originale del PDF nell'upload
  const uploadResp = UrlFetchApp.fetch(
    `${YOUSIGN_API_BASE}/signature_requests/${requestId}/documents`, {
      method: 'post',
      payload: { file: pdfBlob, nature: 'signable_document' },
      headers: { 'Authorization': `Bearer ${YOUSIGN_API_KEY}` },
      muteHttpExceptions: true
    }
  );
  const docJson = JSON.parse(uploadResp.getContentText());
  const documentId = docJson.id;

  UrlFetchApp.fetch(
    `${YOUSIGN_API_BASE}/signature_requests/${requestId}/documents/${documentId}/fields`, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ page, x, y, width, height, type: 'signature', signer_id: signerId }),
      headers: { 'Authorization': `Bearer ${YOUSIGN_API_KEY}` },
      muteHttpExceptions: true
    }
  );

  // 6. Attiva e invia email
  const actResp = UrlFetchApp.fetch(
    `${YOUSIGN_API_BASE}/signature_requests/${requestId}/activate`, {
      method: 'post',
      headers: { 'Authorization': `Bearer ${YOUSIGN_API_KEY}` },
      muteHttpExceptions: true
    }
  );
  return JSON.parse(actResp.getContentText());
}

/**
 * Trova Blob PDF "offerta" in Google Drive.
 */
function getOfferPdfBlob(folderLink = '') {
  if (!folderLink) throw new Error('folderLink mancante');
  // Estrai ID della cartella padre
  const match = folderLink.match(/[-\w]{25,}/);
  if (!match) throw new Error('ID cartella non valido');
  let folder = DriveApp.getFolderById(match[0]);

  // Individua la sottocartella nominata con la data di oggi (dd/MM/yyyy)
  const today = new Date();
  const dd = String(today.getDate()).padStart(2, '0');
  const mm = String(today.getMonth() + 1).padStart(2, '0');
  const yyyy = today.getFullYear();
  const todayName = `${dd}/${mm}/${yyyy}`;
  const subFolders = folder.getFolders();
  let found = false;
  while (subFolders.hasNext()) {
    const sub = subFolders.next();
    if (sub.getName() === todayName) {
      folder = sub;
      found = true;
      break;
    }
  }
  if (!found) {
    throw new Error(`Sottocartella "${todayName}" non trovata in ${folderLink}`);
  }

  // Cerca PDF contenente "offerta"
  const pdfs = folder.getFilesByType(MimeType.PDF);
  while (pdfs.hasNext()) {
    const f = pdfs.next();
    if (f.getName().toLowerCase().includes('offerta')) return f.getBlob();
  }
  // Se non trovato, cerca Google Docs e converte
  const docs = folder.getFilesByType(MimeType.GOOGLE_DOCS);
  while (docs.hasNext()) {
    const d = docs.next();
    if (d.getName().toLowerCase().includes('offerta')) {
      const blob = DriveApp.getFileById(d.getId()).getAs('application/pdf');
      blob.setName(d.getName() + '.pdf');
      return blob;
    }
  }
  throw new Error('Nessun documento "offerta" trovato');
}
