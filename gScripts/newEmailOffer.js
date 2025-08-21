/** HANDLER EMAIL */
function handleEmail(p) {
  // 1) Parametri normalizzati
  var recipiente    = (p.recipiente || '').trim();
  var nomeReferente = (p.nome || '').trim();
  var id            = (p.id || '').trim();

  Logger.log('Parametri email: recipiente=' + recipiente + ', nome=' + nomeReferente + ', id=' + id);

  // 2) Validazioni minime
  if (!recipiente) {
    throw new Error('Parametro "recipiente" mancante.');
  }
  if (!isValidEmail(recipiente)) {
    throw new Error('Indirizzo email cliente non valido: ' + recipiente);
  }

  // 3) CREA la bozza
  sendCustomEmail(recipiente, nomeReferente, id);

  // 4) Pagina di conferma
  var html = `
  <!DOCTYPE html>
  <html lang="it">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>Bozza creata - Myenergy</title>
    <style>
      body { margin:0; font-family: Verdana, sans-serif; background:#f4f4f4; }
      .container { max-width:600px; margin:40px auto; background:#fff; border-radius:16px; box-shadow:0 0 20px rgba(0,0,0,.08); padding:30px 20px; text-align:center; }
      .logo { width:180px; margin-bottom:20px; }
      h2 { color:#2270a8; font-size:20px; margin-bottom:15px; }
      .btn { display:inline-flex; align-items:center; justify-content:center; gap:8px; padding:12px 20px; font-size:16px; border:none; border-radius:8px; text-decoration:none; color:#fff; font-weight:bold; margin:10px auto; transition:.1s; box-sizing:border-box; max-width:90%; }
      .btn:active { transform:scale(.97); box-shadow:inset 0 3px 5px rgba(0,0,0,.2); }
      .icon { font-size:20px; }
      .btn-gmail { background:#D93025; }
      .btn-apple { background:#000; }
      .btn-outlook { background:#0072C6; }
    </style>
  </head>
  <body>
    <div class="container">
      <img src="https://i.imgur.com/KyIh0P7.png" alt="Logo Myenergy" class="logo" />
      <h2>✅ Bozza creata con successo</h2>
      <p>La tua bozza email è ora pronta nella tua casella di posta.</p>

      <a href="https://mail.google.com/mail/u/0/#drafts" target="_blank" class="btn btn-gmail">
        <span class="icon">📧</span><span class="label">Apri con Gmail</span>
      </a>
      <a href="mailto:" class="btn btn-apple">
        <span class="icon">📬</span><span class="label">Apri con Mail (Apple)</span>
      </a>
      <a href="ms-outlook://compose" class="btn btn-outlook">
        <span class="icon">📨</span><span class="label">Apri con Outlook</span>
      </a>
    </div>
  </body>
  </html>
  `;
  return HtmlService.createHtmlOutput(html);
}

/** CREA LA BOZZA (funzione separata) */
function sendCustomEmail(recipiente, nomeReferente, id) {
  var htmlBody = `
  <!DOCTYPE html>
  <html>
  <head><meta charset="UTF-8"></head>
  <body style="margin:0;padding:0;background-color:#f4f4f4;">
    <center>
      <div style="max-width:600px;margin:auto;background:#fff;padding:20px;box-shadow:0 0 10px rgba(0,0,0,.1);text-align:justify;font-family:Verdana,sans-serif;font-size:16px;">
        <img src="https://i.imgur.com/KyIh0P7.png" alt="Myenergy solutions Logo" style="width:100%;height:auto;border:0;">
        <br><br>
        <p style="font-weight:bold;color:#2270a8;margin-top:20px;">Buongiorno, ${nomeReferente || ''}!</p>
        <p>Siamo lieti di annunciarti che la tua <b>offerta personalizzata</b> è ora pronta! Puoi trovarla in <b>allegato</b> a questa email.</p>
        <br><hr><br>
        <p>Con oltre <b style="color:#DA6418;">200 MW</b> di impianti installati e <b style="color:#DA6418;">1000 impianti fotovoltaici</b> residenziali realizzati, Myenergy group offre <b style="color:#2270ad;">qualità e attenzione al cliente</b> ante, durante e post vendita.</p>
        <img src="https://i.imgur.com/uU7uUXY.png" alt="Impianto solare Myenergy" style="width:45%;height:auto;display:block;margin:0 auto;border:0;">
        <hr><br>
        <p>Se l'offerta risponde alle tue aspettative, potrai <b>confermarla</b> inviando il contratto firmato a <a href="mailto:residenziale@myenergy.it">residenziale@myenergy.it</a>.</p>
        <p>Per domande o modifiche:</p>
        <p>☎ <b>Telefono:</b> 3792610174</p>
        <p>✉ <b>E-mail:</b> <a href="mailto:residenziale@myenergy.it">residenziale@myenergy.it</a></p>
        <br><hr><br>
        <p>Validità dell'offerta: <b>15 giorni</b>.</p>
        <p style="margin-top:30px;"><i>Team Myenergy</i></p>
        <br><hr><br>
        <p style="text-align:center;">
          <a href="https://www.facebook.com/myenergy.residenziale/">Facebook</a> | 
          <a href="https://www.instagram.com/myenergy_solutions/">Instagram</a> | 
          <a href="https://www.myenergy.it/realizzazioni/residenziale">Sito web</a>
        </p>
        <p style="text-align:center;font-size:10px;">
          <a href="https://storyset.com/">attribuzioni illustrazioni storyset</a>
        </p>
        <img src="https://i.imgur.com/nvp6lzL.png" alt="blu closure" style="width:100%;height:60px;border:0;">
      </div>
    </center>
  </body>
  </html>`;

  GmailApp.createDraft(
    recipiente,
    "Offerta personalizzata impianto fotovoltaico" + (id ? " #" + id : ""),
    "",
    { htmlBody: htmlBody, bcc: 'residenziale@myenergy.it' }
  );
}

/** VALIDATORE EMAIL */
function isValidEmail(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(email || '').trim());
}
