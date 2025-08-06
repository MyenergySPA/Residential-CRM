/**
 * Restituisce le coordinate in formato "lat, lng".
 *
 * @param {string} address L’indirizzo da geocodificare.
 * @return {string} Coordinate formattate, es. "45.4641, 9.1919".
 * @throws {Error} se la geocodifica fallisce.
 */
function geolocation(address) {
  if (!address || typeof address !== 'string') {
    throw new Error('Indirizzo non valido.');
  }

  // --- Cache 24 h ------------------------------------------------------
  const cache = CacheService.getScriptCache();
  const key   = `geo_str:${address.trim().toLowerCase()}`;
  const hit   = cache.get(key);
  if (hit) {
    return hit;                    // la stringa è già pronta
  }

  // --- Geocodifica -----------------------------------------------------
  const geocoder = Maps.newGeocoder()
                       .setLanguage('it')
                       .setRegion('it');
  const res = geocoder.geocode(address);

  if (res.status !== 'OK' || !res.results.length) {
    throw new Error(`Geocodifica fallita: ${res.status}`);
  }

  const { lat, lng } = res.results[0].geometry.location;
  const coordStr     = `${lat}, ${lng}`;   // <- output richiesto

  // --- Salva in cache --------------------------------------------------
  cache.put(key, coordStr, 60 * 60 * 24);

  return coordStr;
}
