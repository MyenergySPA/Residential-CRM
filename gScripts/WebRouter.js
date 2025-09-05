 opportunitàREF contiene l'appID della riga in 'cronologia'
  const leadAppID = String(off[co('opportunitàREF') - 1] || '').trim();
  if (!leadAppID) throw new Error('opportunitàREF vuoto per offerta=' + offerId);

  const { row: lead, col: cl } = _getLeadRowByAppID_(leadAppID);
  const inc = _lookupIncentivoFromOfferta_(off, co);

  const getO = n => off[co(n) - 1];
  const getL = n => lead[cl(n) - 1];

  const cartellaUrl = String(getL('cartella') || '');

  const args = [
    offerId,
    String(getL('tipo_opportunità') || getO('tipo_opportunita_helper') || ''),
    String(getL('id') || ''),
    String(getL('yy') || ''),
    String(getL('nome referente') || ''),
    String(getL('cognome referente') || ''),
    String(getL('indirizzo') || ''),
    String(getL('telefono') || ''),
    String(getL('email') || ''),

    Number(getO('numero moduli') || 0),
    Number(getO('numero inverter') || 1),
    String(getO('tipo moduli') || ''),
    String(getO('modello inverter') || ''),
    Number(getO('numero batterie') || 0),
    Number(getO('capacità batteria [kWh]') || 0),
    Number(getO('totale capacità batteria [kWh]') || 0),
    String(getO('modello batteria') || ''),
    String(getO('struttura') || ''),
    Number(getO('Potenza [kWp]') || 0),
    Number(getO('alberi in 25 anni') || 0),
    String(getO('testo_aggiuntivo') || ''),
    String(getO('tipo_pagamento') || ''),

    String(getO('condizione_pagamento_1') || ''),
    String(getO('condizione_pagamento_2') || ''),
    String(getO('condizione_pagamento_3') || ''),
    String(getO('condizione_pagamento_4') || ''),
    Number(getO('imponibile offerta - tutto incluso') || 0),
    Number(getO('iva €') || 0),
    Number(getO('iva %') || 0),
    Number(getO('prezzo offerta appros. - tutto incluso') || 0),

    cartellaUrl,

    Number(getO('anni finanziamento') || 0),
    String(getO('esposizione') || ''),
    Number(getO('area m2 impianto') || 0),
    Number(getO('numero colonnina 7,4 kw') || 0),
    Number(getO('numero colonnina 22 kw') || 0),
    Number(getO('numero ottimizzatori') || 0),
    String(getO('modello ottimizzatori') || ''),
    Number(getO('numero linea vita') || 0),

    String(getO('Scheda tecnica Moduli') || ''),
    String(getO('Scheda tecnica Inverter') || ''),
    String(getO('Scheda tecnica batterie') || ''),
    String(getO('Scheda tecnica ottimizzatori') || ''),

    Number(getL('kwh annui') || 0),                // <-- da cronologia
    String(getL('profilo di consumo') || ''),       // <-- da cronologia (fix)
    String(getL('provincia') || ''),                // da cronologia
    Number(getL('prezzo energia') || 0),            // da cronologia

    String(getO('tipo incentivo') || ''),
    Number(getO('anni durata incentivo') || inc.durata_incentivo || 0),
    String(getL('coordinate') || ''),               // da cronologia
    String(getL('ragione sociale') || ''),          // da cronologia
    Number(getO('garanzia moduli') || 0),
    Number(getO('garanzia inverter') || 0),
    Number(getO('garanzia batterie') || 0),
    Number(getO('Acconto diretto') || 0),
    Number(getO('tilt') || 0),
    Number(getO('azimuth') || 0),

    String(inc.nome_incentivo || ''),
    String(inc.descrizione_offerta || '')
  ];

    return { argsForMain: args, cartellaUrl };
}
