function duplicateCheck() {
  const ss     = CRMdatabase;
  const sheet  = sheetCronologia;
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) {
    console.log('Nessun dato utile (meno di 2 righe totali).');
    return;
  }

  // Estrai header e trova gli indici
  const header    = values.shift().map(h => h.toString().toLowerCase());
  const idxNome   = header.indexOf('nome');
  const idxCogn   = header.indexOf('cognome');
  const idxTipoOp = header.indexOf('tipo_opportunità');
  const idxId     = header.indexOf('id');
  const idxyy     = header.indexOf('yy');  
  if (idxNome < 0 || idxCogn < 0 || idxTipoOp < 0 || idxId < 0 || idxyy < 0) {
    throw new Error('Colonne "nome", "cognome", "tipo opportunità" e/o "id" non trovate');
  }

  // Conta occorrenze del quadrupletto chiave, saltando le righe vuote
  const counts = {};
  values.forEach(row => {
    // se tutte e quattro le celle sono vuote, salto
    if ([idxNome, idxCogn, idxTipoOp, idxId, idxyy].every(i => row[i].toString().trim() === '')) {
      return;
    }
    const key = `${row[idxNome]}|${row[idxCogn]}|${row[idxTipoOp]}|${row[idxId]}|${row[idxyy]}`;
    counts[key] = (counts[key] || 0) + 1;
  });

  // Trova i duplicati
  const duplicates = Object.keys(counts).filter(k => counts[k] > 1);
  if (duplicates.length === 0) {
    console.log('Nessun duplicato trovato.');
    return;
  }

  // Log dettagliato per debug (opzionale)
  duplicates.forEach(key => {
    const [nome, cognome, tipo, id, yy] = key.split('|');
    console.log(`Duplicato: ${nome} ${cognome} — ${tipo} — id: ${id} ${yy} (${counts[key]} volte)`);
  });

  // Costruisci il corpo HTML con la colonna id
  let html = '<p>❗️ Attenzione: rilevati possibili duplicati nelle opportunità CRM:</p>';
  html += '<table border="1" cellpadding="4" style="border-collapse:collapse;">';
  html += '<tr><th>Nome</th><th>Cognome</th><th>Tipo Opportunità</th><th>ID</th><th>Occorrenze</th></tr>';
  duplicates.forEach(key => {
    const [nome, cognome, tipo, id, yy] = key.split('|');
    html += `<tr>
      <td>${nome}</td>
      <td>${cognome}</td>
      <td>${tipo}</td>
      <td>${id}-${yy}</td>
      <td>${counts[key]}</td>
    </tr>`;
  });
  html += '</table>';
  console.log('HTML email:\n' + html);

  // Manda l'email e logga eventuali errori
  try {
  GmailApp.sendEmail(
    'residenziale@myenergy.it',
    '⚠️ Duplicati opportunità CRM',
    '',                                     // corpo testuale (vuoto perché usiamo htmlBody)
    {
      htmlBody: html,
      from:     'residenziale@myenergy.it', // deve essere un alias confermato
      name:     'CRM residenziale'
    }
  );
  console.log('Email inviata via GmailApp da alias residenziale@myenergy.it');
} catch (e) {
  console.error('Errore invio email via GmailApp:', e.message);
}


  // Log finale
  console.log(`Totale righe duplicate individuate: ${duplicates.length}`);
}
