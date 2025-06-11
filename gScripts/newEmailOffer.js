//robe per webapp

function doGet(e) {
  try {
    // Accedi ai parametri direttamente dall'oggetto e.parameter
    var recipiente = e.parameter.recipiente;
    var nome = e.parameter.nome;
    var id = e.parameter.id;

    Logger.log('Parametri ricevuti: recipiente = ' + recipiente + ', nome = ' + nome + ', id = ' + id);

    // Verifica se il parametro recipiente è un indirizzo email valido
    if (!isValidEmail(recipiente)) {
      throw new Error('Indirizzo email cliente non valido: ' + recipiente);
    }

    // Chiama la tua funzione per inviare la email personalizzata
    sendCustomEmail(recipiente, nome, id);

    // Restituisci un HTML con un pulsante per chiudere la scheda
    var htmlOutput = HtmlService.createHtmlOutput(
      '<html><body>' +
      "<p> La bozza email è stata creata nella tua casella di posta. Puoi chiudere questa scheda. </p>" +
      '</body></html>'
    );

    return htmlOutput;
  } catch (error) {
    Logger.log('Errore durante il doGet: ' + error.toString());
    return ContentService.createTextOutput('Errore: ' + error.toString());
  }
}

// Funzione per validare l'indirizzo email
function isValidEmail(email) {
  var emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email);
}


//script

function sendCustomEmail(recipiente, nome, id) {
  var htmlBody = `
  <!DOCTYPE html>
<html>

<head>
<meta charset="UTF-8">
</head>

<body style="margin: 0; padding: 0; background-color: #f4f4f4;">
  <center>
    <div style="max-width: 600px; margin: auto; background-color: white; padding: 20px; box-shadow: 0 0 10px rgba(0, 0, 0, 0.1); text-align: justify; font-family: Verdana, sans-serif; font-size: 16px;">

      <!-- Logo Myenergy -->
      <img src="https://i.imgur.com/KyIh0P7.png" alt="Myenergy solutions Logo" style="width:100%; height:auto; border: 0;">
    
    <br><br>

      <!-- Saluto personalizzato -->
      <p style="font-weight: bold; color: #2270a8; margin-top: 20px;">Buongiorno, ${nome}!</p>

      <p>Siamo lieti di annunciarti che la tua <b> offerta personalizzata </b> è ora pronta! Puoi trovarla in <b> allegato </b> a questa email.</p>

    <br><hr><br>

      <p> Con oltre <b style="color:#DA6418;"> 200 MW </b> di impianti installati e <b style="color:#DA6418;"> 1000 impianti fotovoltaici </b> residenziali realizzati, Myenergy group offre <b style="color:#2270ad;">qualità e attenzione al cliente </b> ante, durante e post vendita.</p>
    
      <!-- Immagine dell'impianto solare -->
        <img src="https://i.imgur.com/uU7uUXY.png" alt="Impianto solare Myenergy" style="width:45%; height:auto; display:block; margin-left:auto; margin-right:auto; border:0;">

    <hr><br>

      <p>Se l'offerta risponde alle tue aspettative, potrai <b> confermarla </b> inviando il contratto firmato a <a href="mailto:residenziale@myenergy.it">residenziale@myenergy.it</a>.</p>

      <p>Per qualsiasi dubbio o domanda, o per apportare modifiche alla tua offerta, non esitare a contattarci ai seguenti riferimenti:</p>
      <p> ☎ <b>Telefono:</b> <span>3792610174</span></p>
      <p> ✉ <b>E-mail:</b> <a href="mailto:residenziale@myenergy.it">residenziale@myenergy.it</a></p>

    <br><hr><br>

      <p>Desideriamo sottolineare che l'offerta ha una validità di <b>15 giorni.</b></p>

      <p>Grazie di aver considerato Myenergy, siamo pronti a intraprendere insieme a voi questo viaggio verso un futuro più sostenibile!</p>

      <p style="margin-top: 30px;"><i>Team Myenergy</i></p>

    <br><hr><br>

      <p style="text-align: center;">
        <a href="https://www.facebook.com/myenergy.residenziale/">Facebook</a> | <a href="https://www.instagram.com/myenergy_solutions/">Instagram</a> | <a href="https://www.myenergy.it/realizzazioni/residenziale]">Sito web</a>
      </p>
      <p style="text-align: center; font-size: 10px;">
        <a href="https://storyset.com/">attribuzioni illustrazioni storyset</a>
      </p>

            <!-- Blu chiusura -->
      <img src="https://i.imgur.com/nvp6lzL.png" alt="blu closure" style="width:100%; height:60px; border: 0;">

    </div>
  </center>
</body>
</html>`;

var draft = GmailApp.createDraft(recipiente, "Offerta personalizzata impianto fotovoltaico", "", {
    htmlBody: htmlBody,
    bcc: 'residenziale@myenergy.it'
  });
}