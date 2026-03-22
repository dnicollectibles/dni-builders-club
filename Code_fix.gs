/**
 * DNI Builders Club — Google Apps Script Backend
 * ================================================
 * Deploy: "Web app" → Esegui come "Me" → Accesso "Chiunque"
 * Dopo ogni modifica: Distribuisci → Gestisci distribuzioni → Nuova versione
 *
 * v6 — marzo 2026
 * ─────────────────────────────────────────────────────────────────────────────
 * MODIFICHE v6:
 *
 * [SKU]
 *   - Rimossa colonna "transcodifica" dal Catalogo
 *   - Rinominata "external sku" → "vendor_sku"
 *   - sku DNI ora calcolato on-the-fly: "DNI-" + vendor_sku (non letto dal foglio)
 *
 * [CATALOGO]
 *   - Aggiunta colonna "anno" (estratta dal campo mese: "2026-03" → anno=2026)
 *   - Il tier rimane colonna manuale (Alberto decide per ogni prodotto)
 *
 * [FOGLI NUOVI]
 *   - Notizie       : post per la home pubblica
 *   - Prodotti_Shop : prodotti pubblici con prezzo_shop
 *   - Spotlight     : prodotti in evidenza nella home
 *   - Banner        : messaggio globale temporaneo
 *
 * [ENDPOINT NUOVI — tutti richiedono admin_key]
 *   GET  : catalogo_admin, notizie, prodotti_shop, spotlight, banner,
 *          utenti_admin, prenotazioni_admin
 *   POST : salva_notizia, elimina_notizia,
 *          salva_prodotto_shop, elimina_prodotto_shop,
 *          salva_spotlight,
 *          salva_banner,
 *          aggiorna_utente, aggiorna_prenotazione,
 *          aggiungi_prodotto_catalogo, aggiorna_prodotto_catalogo
 *
 * [INVARIATI]
 *   - login, ordini, conferma, acconto_extra, sblocca
 *   - Logica depositi, credito, Log_Attività, email
 * ─────────────────────────────────────────────────────────────────────────────
 */

const SPREADSHEET_ID = 'INCOLLA_QUI_ID_DEL_FOGLIO_V6.2';  // ← AGGIORNA CON IL TUO ID
const ADMIN_EMAIL    = 'dni.infoshop@gmail.com';
const ADMIN_KEY      = 'Dni2026!Admin@1984_Sup3rVisor';

// ═══════════════════════════════════════════════════════════════════════════════
// OUTPUT & CORS
// ═══════════════════════════════════════════════════════════════════════════════

function corsOutput(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function errOutput(msg, code) {
  return corsOutput({ ok: false, error: msg, errore: msg, code: code || 400 });
}

// ═══════════════════════════════════════════════════════════════════════════════
// ROUTER
// ═══════════════════════════════════════════════════════════════════════════════

function doGet(e) {
  try {
    const p = e.parameter || {};
    switch (p.action) {
      // ── Soci ──────────────────────────────────────────────────────────────
      case 'login':               return actionLogin(p);
      case 'ordini':              return actionOrdini(p);
      case 'sblocca':             return actionSblocca(p);
      // ── Admin: lettura ────────────────────────────────────────────────────
      case 'catalogo_admin':      return actionCatalogoAdmin(p);
      case 'magazzino':           return actionMagazzino(p);
      case 'notizie':             return actionNotizie(p);
      case 'prodotti_shop':       return actionProdottiShop(p);
      case 'spotlight':           return actionSpotlight(p);
      case 'banner':              return actionBanner(p);
      case 'utenti_admin':        return actionUtentiAdmin(p);
      case 'prenotazioni_admin':  return actionPrenotazioniAdmin(p);
      // ── Pubblico: home e shop (no admin_key) ──────────────────────────────
      case 'home_pubblica':       return actionHomePubblica(p);
      case 'shop_pubblico':       return actionShopPubblico(p);
      default:                    return errOutput('Azione non valida', 400);
    }
  } catch (err) {
    return errOutput('Errore interno: ' + err.message, 500);
  }
}

function doPost(e) {
  try {
    let body = {};
    if (e.postData && e.postData.contents) {
      body = JSON.parse(e.postData.contents);
    }
    const action = (e.parameter && e.parameter.action) || body.action || '';
    switch (action) {
      // ── Soci ──────────────────────────────────────────────────────────────
      case 'conferma':                    return actionConferma(body);
      case 'acconto_extra':               return actionAcExtra(body);
      // ── Admin: catalogo ───────────────────────────────────────────────────
      case 'aggiungi_prodotto_catalogo':  return actionAggiungiProdottoCatalogo(body);
      case 'aggiorna_prodotto_catalogo':  return actionAggiornaProdottoCatalogo(body);
      case 'elimina_prodotto_catalogo':   return actionEliminaProdottoCatalogo(body);
      case 'aggiorna_stock':              return actionAggiornaStock(body);
      // ── Admin: notizie ────────────────────────────────────────────────────
      case 'salva_notizia':               return actionSalvaNotizia(body);
      case 'elimina_notizia':             return actionEliminaNotizia(body);
      // ── Admin: shop ───────────────────────────────────────────────────────
      case 'salva_prodotto_shop':         return actionSalvaProdottoShop(body);
      case 'elimina_prodotto_shop':       return actionEliminaProdottoShop(body);
      // ── Admin: spotlight & banner ─────────────────────────────────────────
      case 'salva_spotlight':             return actionSalvaSpotlight(body);
      case 'salva_banner':                return actionSalvaBanner(body);
      // ── Admin: soci & prenotazioni ────────────────────────────────────────
      case 'aggiorna_utente':             return actionAggiornaUtente(body);
      case 'aggiorna_prenotazione':       return actionAggiornaPrenotazione(body);
      default:                            return errOutput('Azione POST non valida', 400);
    }
  } catch (err) {
    return errOutput('Errore interno: ' + err.message, 500);
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
// UTILITY
// ═══════════════════════════════════════════════════════════════════════════════

function getSheet(name) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const ws = ss.getSheetByName(name);
  if (!ws) throw new Error('Scheda "' + name + '" non trovata');
  return ws;
}

/**
 * Converte i dati di un foglio in array di oggetti.
 * Riga 1 = header, Riga 2 = descrizioni (saltata), Riga 3+ = dati.
 * Fogli con solo riga header (no riga descrizioni): passare skipDescRow=false.
 */
function sheetToObjects(ws, skipDescRow) {
  if (skipDescRow === undefined) skipDescRow = true;
  const rows = ws.getDataRange().getValues();
  const dataStart = skipDescRow ? 2 : 1;
  if (rows.length <= dataStart) return [];
  const headers = rows[0].map(h => String(h).trim());
  const result = [];
  for (let i = dataStart; i < rows.length; i++) {
    const row = rows[i];
    if (row.every(c => c === '' || c === null || c === undefined)) continue;
    const obj = { _row: i + 1 };
    headers.forEach((h, idx) => { if (h) obj[h] = row[idx]; });
    result.push(obj);
  }
  return result;
}

function findUtente(id) {
  const ws = getSheet('Utenti');
  const rows = ws.getDataRange().getValues();
  const headers = rows[0].map(h => String(h).trim());
  const idCol = headers.indexOf('id');
  for (let i = 2; i < rows.length; i++) {
    if (String(rows[i][idCol]).trim() === String(id).trim()) {
      const obj = {};
      headers.forEach((h, idx) => { obj[h] = rows[i][idx]; });
      return { rowIndex: i + 1, data: obj, ws, headers };
    }
  }
  return null;
}

function nextId(ws, prefix) {
  const rows = ws.getDataRange().getValues();
  let max = 0;
  for (let i = 2; i < rows.length; i++) {
    const val = String(rows[i][0] || '');
    if (val.startsWith(prefix)) {
      const num = parseInt(val.slice(prefix.length), 10);
      if (!isNaN(num) && num > max) max = num;
    }
  }
  return prefix + String(max + 1).padStart(4, '0');
}

/** Genera id per fogli senza riga descrizione (Notizie, Prodotti_Shop, ecc.) */
function nextIdSimple(ws, prefix) {
  const rows = ws.getDataRange().getValues();
  let max = 0;
  for (let i = 1; i < rows.length; i++) {
    const val = String(rows[i][0] || '');
    if (val.startsWith(prefix)) {
      const num = parseInt(val.slice(prefix.length), 10);
      if (!isNaN(num) && num > max) max = num;
    }
  }
  return prefix + String(max + 1).padStart(4, '0');
}

function isoNow() {
  return Utilities.formatDate(new Date(), 'Europe/Rome', "yyyy-MM-dd'T'HH:mm:ss");
}

function meseCorrente() {
  return Utilities.formatDate(new Date(), 'Europe/Rome', 'yyyy-MM');
}

/** Controlla admin_key. Lancia errore se non valida. */
function requireAdmin(params) {
  const key = params.admin_key || params.adminKey || '';
  if (String(key).trim() !== String(ADMIN_KEY).trim())
    throw new Error('Admin key non valida');
}

/**
 * Calcola lo SKU DNI a partire dal vendor_sku.
 * Formato: "DNI-" + vendor_sku
 * Es.: vendor_sku = "SNAA-SC-008" → sku = "DNI-SNAA-SC-008"
 */
function calcolaSku(vendorSku) {
  if (!vendorSku || String(vendorSku).trim() === '') return '';
  return 'DNI-' + String(vendorSku).trim();
}

// ═══════════════════════════════════════════════════════════════════════════════
// LOG ATTIVITÀ
// ═══════════════════════════════════════════════════════════════════════════════

function logAttivita(params) {
  try {
    const ws    = getSheet('Log_Attività');
    const logId = nextId(ws, 'A');
    ws.appendRow([
      logId,
      params.idPrenotazione  || '',
      params.idUtente        || '',
      params.nomeUtente      || '',
      params.timestamp       || isoNow(),
      params.nomeProdotto    || '',
      params.sku             || '',
      params.tierProdotto    || '',
      params.tipoEvento      || '',
      params.importo         || 0,
      params.accontoBase     || 0,
      params.accontoExtraTot || 0,
      params.prezzoTotale    || 0,
      params.residuo         || 0,
      params.note            || '',
    ]);
  } catch(err) {
    Logger.log('Log_Attività fallito (non blocca): ' + err.message);
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
// EMAIL
// ═══════════════════════════════════════════════════════════════════════════════

function emailConfermaFullpay(params) {
  try {
    const subject = '[DNI Club] 💳 Full Payment — ' + params.nomeSocio + ' — ' + params.nomeProdotto;
    const htmlBody =
      '<div style="font-family:Arial,sans-serif;max-width:620px;margin:0 auto;color:#222">' +
      '<div style="background:#0b0b0f;padding:20px 28px">' +
        '<span style="font-size:20px;font-weight:bold;color:#fff">DNI </span>' +
        '<span style="font-size:20px;font-weight:bold;color:#b4ff00">BUILDERS</span>' +
        '<span style="font-size:20px;font-weight:bold;color:#fff"> CLUB — Full Payment</span></div>' +
      '<div style="background:#eaf4fb;padding:16px 28px;border-left:4px solid #1a5276">' +
        '<p style="margin:0;font-size:15px">💳 <strong>' + params.nomeSocio + '</strong> ha effettuato un acquisto a saldo immediato.</p>' +
        '<p style="margin:6px 0 0;font-size:13px;color:#555">ID: <code>' + params.idPrenotazione + '</code> — ' + params.now + '</p>' +
      '</div>' +
      '<table style="width:100%;border-collapse:collapse;font-size:13px">' +
        '<tr style="background:#f4f4f4"><td style="padding:10px 16px;font-weight:bold;width:45%">Prodotto</td><td style="padding:10px 16px">' + params.nomeProdotto + '</td></tr>' +
        '<tr><td style="padding:10px 16px;font-weight:bold;background:#f4f4f4">SKU</td><td style="padding:10px 16px;font-family:monospace">' + params.sku + '</td></tr>' +
        '<tr style="background:#f4f4f4"><td style="padding:10px 16px;font-weight:bold">Tier</td><td style="padding:10px 16px">' + String(params.tierProd).toUpperCase() + '</td></tr>' +
        '<tr><td style="padding:10px 16px;font-weight:bold;background:#f4f4f4">Importo pagato</td><td style="padding:10px 16px;color:#1a5276;font-weight:bold">€' + params.acconto + '</td></tr>' +
        '<tr style="background:#f4f4f4"><td style="padding:10px 16px;font-weight:bold">Prezzo totale</td><td style="padding:10px 16px">€' + params.prezzoTotale + '</td></tr>' +
      '</table>' +
      '<div style="background:#0b0b0f;padding:14px 28px;text-align:center">' +
        '<span style="font-size:11px;color:#666">Conferma l\'ordine nel foglio e aggiorna lo stato.</span></div></div>';
    MailApp.sendEmail({ to: ADMIN_EMAIL, subject: subject, htmlBody: htmlBody });
  } catch(mailErr) {
    Logger.log('Email fullpay fallita [' + params.idPrenotazione + ']: ' + mailErr.message);
  }
}

function emailSaldo(params) {
  try {
    const subject = '[DNI Club] ✅ Ordine saldato — ' + params.nomeSocio + ' — ' + params.nomeProdotto;
    const htmlBody =
      '<div style="font-family:Arial,sans-serif;max-width:620px;margin:0 auto;color:#222">' +
      '<div style="background:#0b0b0f;padding:20px 28px">' +
        '<span style="font-size:20px;font-weight:bold;color:#fff">DNI </span>' +
        '<span style="font-size:20px;font-weight:bold;color:#b4ff00">BUILDERS</span>' +
        '<span style="font-size:20px;font-weight:bold;color:#fff"> CLUB — Ordine Saldato</span></div>' +
      '<div style="background:#eafaf1;padding:16px 28px;border-left:4px solid #27ae60">' +
        '<p style="margin:0;font-size:15px">✅ <strong>' + params.nomeSocio + '</strong> ha saldato un ordine.</p>' +
        '<p style="margin:6px 0 0;font-size:13px;color:#555">ID: <code>' + params.idPrenotazione + '</code> — ' + params.now + '</p>' +
        (params.opId ? '<p style="margin:4px 0 0;font-size:11px;color:#aaa">Op ID: <code>' + params.opId + '</code></p>' : '') +
      '</div>' +
      '<table style="width:100%;border-collapse:collapse;font-size:13px">' +
        '<tr style="background:#f4f4f4"><td style="padding:10px 16px;font-weight:bold;width:45%">Prodotto</td><td style="padding:10px 16px">' + params.nomeProdotto + '</td></tr>' +
        '<tr><td style="padding:10px 16px;font-weight:bold;background:#f4f4f4">SKU</td><td style="padding:10px 16px;font-family:monospace">' + params.sku + '</td></tr>' +
        '<tr style="background:#f4f4f4"><td style="padding:10px 16px;font-weight:bold">Tier</td><td style="padding:10px 16px">' + String(params.tierProd).toUpperCase() + '</td></tr>' +
        '<tr><td style="padding:10px 16px;font-weight:bold;background:#f4f4f4">Tipo ordine originale</td><td style="padding:10px 16px">' + params.tipo + '</td></tr>' +
        '<tr style="background:#f4f4f4"><td style="padding:10px 16px;font-weight:bold">Acconto base versato</td><td style="padding:10px 16px">€' + params.accontoBase + '</td></tr>' +
        '<tr><td style="padding:10px 16px;font-weight:bold;background:#f4f4f4">Acconti extra totali</td><td style="padding:10px 16px;color:#27ae60;font-weight:bold">€' + params.nuovoAccontoExtra + '</td></tr>' +
        '<tr style="background:#f4f4f4"><td style="padding:10px 16px;font-weight:bold">Versato ora</td><td style="padding:10px 16px;color:#27ae60">€' + params.delta + '</td></tr>' +
        '<tr><td style="padding:10px 16px;font-weight:bold;background:#f4f4f4">Prezzo totale</td><td style="padding:10px 16px;font-weight:bold">€' + params.prezzoTotale + '</td></tr>' +
        '<tr style="background:#f4f4f4"><td style="padding:10px 16px;font-weight:bold">Data conferma ordine</td><td style="padding:10px 16px">' + params.dataConf + '</td></tr>' +
      '</table>' +
      '<div style="background:#0b0b0f;padding:14px 28px;text-align:center">' +
        '<span style="font-size:11px;color:#666">Inserisci il numero di tracking quando spedisci.</span></div></div>';
    MailApp.sendEmail({ to: ADMIN_EMAIL, subject: subject, htmlBody: htmlBody });
  } catch(mailErr) {
    Logger.log('Email saldo fallita [' + params.idPrenotazione + ']: ' + mailErr.message);
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
// HELPER: LETTURA CATALOGO
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Helper: converte un valore Sheets in booleano.
 */
function toBool(v) {
  return v === true || String(v || '').trim().toUpperCase() === 'TRUE';
}

/**
 * Calcola i conteggi stock per un prodotto a partire dalle Prenotazioni attive.
 * "Attiva" = confermata=TRUE e stato_ordine non in ('saldato','annullato').
 * Restituisce { prenotato_vip }.
 */
function getStockPrenotato(prodottoId) {
  try {
    const ws  = getSheet('Prenotazioni');
    const all = sheetToObjects(ws, true);
    const pid = String(prodottoId).trim();
    const attivi = all.filter(r => {
      const stessa  = String(r.prodotto_id || '').trim() === pid;
      const conf    = toBool(r.confermata);
      const stato   = String(r.stato_ordine || '').trim();
      const nonChiuso = stato !== 'saldato' && stato !== 'annullato';
      return stessa && conf && nonChiuso;
    });
    return { prenotato_vip: attivi.length };
  } catch(e) {
    return { prenotato_vip: 0 };
  }
}

/**
 * Mappa un record raw del Catalogo in oggetto normalizzato.
 * Condiviso da getCatalogo, getCatalogoAdmin, getShopPubblico.
 */
function mapProdottoCatalogo(r) {
  const isTbd     = toBool(r.is_tbd);
  const vendorSku = String(r.vendor_sku || r['external sku'] || '').trim();
  const anno      = String(r.mese || '').split('-')[0] || '';

  const stockTotale = Number(r.stock_totale) || 0;
  const quotaVip    = Number(r.quota_vip)    || 0;
  const quotaShop   = Number(r.quota_shop)   || 0;

  return {
    id:                    r.id,
    mese:                  String(r.mese || '').trim(),
    anno,
    tier:                  String(r.tier      || '').trim().toLowerCase(),
    categoria:             String(r.categoria || '').trim().toLowerCase(),
    name:                  String(r.nome_prodotto || ''),
    vendor_sku:            vendorSku,
    sku:                   calcolaSku(vendorSku),
    price:                 isTbd ? null : (Number(r.prezzo_dbc)      || null),
    prezzo_pubblico:       isTbd ? null : (Number(r.prezzo_pubblico) || null),
    acconto:               Number(r.acconto_base) || 0,
    isTbd,
    // stock
    stock_totale:          stockTotale,
    quota_vip:             quotaVip,
    quota_shop:            quotaShop,
    stock_in_arrivo:       toBool(r.stock_in_arrivo),
    data_arrivo_prevista:  r.data_arrivo_prevista ? String(r.data_arrivo_prevista) : null,
    // visibilità
    attivo_vip:            toBool(r.attivo_vip),
    attivo_shop:           toBool(r.attivo_shop),
    attivo:                toBool(r.attivo) !== false && String(r.attivo || '').toUpperCase() !== 'FALSE',
    // media — array già splittato; frontend: img_catalogo/[vendor_sku]/[file]
    imgs:                  String(r.img || '').split(';').map(s => s.trim()).filter(Boolean),
    imgBase:               vendorSku,
    // note (solo admin)
    note:                  String(r.note || ''),
  };
}

/**
 * Legge il Catalogo per il portale soci (login).
 * Filtra per mese corrente, attivo=TRUE, attivo_vip=TRUE.
 * Aggiunge disponibile_vip calcolato live dalle Prenotazioni.
 *
 * @param {string} mese      Es. "2026-03"
 * @param {string} tierSocio Tier del socio loggato (scout/veteran/elite)
 */
function getCatalogo(mese, tierSocio) {
  const TIER_RANK = { scout: 1, veteran: 2, elite: 3 };
  const ws  = getSheet('Catalogo');
  const all = sheetToObjects(ws, true);

  return all
    .filter(r => {
      if (String(r.mese || '').trim() !== mese) return false;
      if (!toBool(r.attivo))     return false; // override globale
      if (!toBool(r.attivo_vip)) return false; // non attivo per VIP
      // cross-tier: il socio vede il suo tier e i tier inferiori
      const tierProd = String(r.tier || '').trim().toLowerCase();
      const rankProd = TIER_RANK[tierProd] || 0;
      const rankSocio= TIER_RANK[tierSocio] || 0;
      return rankSocio >= rankProd;
    })
    .map(r => {
      const prod   = mapProdottoCatalogo(r);
      const stockP = getStockPrenotato(r.id);
      const disponibile_vip = Math.max(0, prod.quota_vip - stockP.prenotato_vip);
      return {
        ...prod,
        prenotato_vip:    stockP.prenotato_vip,
        disponibile_vip,
        esaurito_vip:     prod.quota_vip > 0 && disponibile_vip === 0,
      };
    });
}

// ═══════════════════════════════════════════════════════════════════════════════
// HELPER: LETTURA PRENOTAZIONI
// ═══════════════════════════════════════════════════════════════════════════════

function getPrenotazioni(idUtente) {
  const ws  = getSheet('Prenotazioni');
  const all = sheetToObjects(ws, true);
  return all
    .filter(r => String(r.id_utente).trim() === String(idUtente).trim())
    .map(r => ({
      id:               r.id,
      prodotto_id:      r.prodotto_id,
      nome_prodotto:    String(r.nome_prodotto    || ''),
      sku:              String(r.sku              || ''),
      tier_prodotto:    String(r.tier_prodotto    || '').trim(),
      tipo:             String(r.tipo             || 'deposito'),
      acconto:          Number(r.acconto)          || 0,
      acconto_extra:    Number(r.acconto_extra)    || 0,
      prezzo_totale:    Number(r.prezzo_totale)    || 0,
      residuo:          Number(r.residuo)          || 0,
      confermata:       r.confermata   === true || String(r.confermata).trim().toUpperCase()   === 'TRUE',
      modificabile:     r.modificabile === true || String(r.modificabile).trim().toUpperCase() === 'TRUE',
      stato_ordine:     String(r.stato_ordine     || 'in_preparazione'),
      commento_alberto: String(r.commento_alberto || ''),
      tracking:         String(r.tracking         || ''),
      data_conferma:    String(r.data_conferma    || ''),
      mese:             String(r.mese             || ''),
    }));
}

// ═══════════════════════════════════════════════════════════════════════════════
// AZIONI SOCI (invariate dalla v5)
// ═══════════════════════════════════════════════════════════════════════════════

function actionLogin(p) {
  if (!p.id || !p.code) return errOutput('Parametri mancanti: id, code', 400);

  const found = findUtente(p.id);
  if (!found) return errOutput('Utente non trovato', 404);
  const u = found.data;
  if (String(u.codice).trim() !== String(p.code).trim()) return errOutput('Codice non valido', 401);

  const mese         = meseCorrente();
  const tierSocio    = String(u.tier || 'scout').trim().toLowerCase();
  const catalogo     = getCatalogo(mese, tierSocio);
  const prenotazioni = getPrenotazioni(p.id);

  let prossimaDataRinnovo = null;
  if (u.prossima_data_rinnovo) {
    try {
      const d = new Date(u.prossima_data_rinnovo);
      if (!isNaN(d.getTime()))
        prossimaDataRinnovo = Utilities.formatDate(d, 'Europe/Rome', 'yyyy-MM-dd');
    } catch(e) {}
  }

  return corsOutput({
    ok: true, mese,
    utente: {
      id:                    u.id,
      nome:                  u.nome,
      tier:                  String(u.tier).trim().toLowerCase(),
      credito:               Number(u.credito) || 0,
      stato:                 String(u.stato).trim(),
      pagamento:             String(u.pagamento).trim(),
      prossima_data_rinnovo: prossimaDataRinnovo,
      accontoSpecialePct: {
        scout:   Number(u.acconto_speciale_pct_scout)   || 50,
        veteran: Number(u.acconto_speciale_pct_veteran) || 40,
        elite:   Number(u.acconto_speciale_pct_elite)   || 30,
      },
      creditoBloccatoPct: {
        scout:   Number(u.credito_bloccato_pct_scout)   || 50,
        veteran: Number(u.credito_bloccato_pct_veteran) || 40,
        elite:   Number(u.credito_bloccato_pct_elite)   || 30,
      },
    },
    catalogo,
    prenotazioni,
  });
}

function actionOrdini(p) {
  if (!p.id || !p.code) return errOutput('Parametri mancanti: id, code', 400);
  const found = findUtente(p.id);
  if (!found) return errOutput('Utente non trovato', 404);
  const u = found.data;
  if (String(u.codice).trim() !== String(p.code).trim()) return errOutput('Codice non valido', 401);
  return corsOutput({
    ok: true,
    stato:        String(u.stato).trim(),
    credito:      Number(u.credito) || 0,
    prenotazioni: getPrenotazioni(p.id),
  });
}

function actionConferma(body) {
  if (!body.id_utente || !body.codice || !body.mese || !Array.isArray(body.selezioni))
    return errOutput('Dati mancanti: id_utente, codice, mese, selezioni[]', 400);
  if (body.selezioni.length === 0)
    return errOutput('Nessuna selezione da confermare', 400);

  const found = findUtente(body.id_utente);
  if (!found) return errOutput('Utente non trovato', 404);
  const u = found.data;
  if (String(u.codice).trim() !== String(body.codice).trim()) return errOutput('Codice non valido', 401);

  const wsBk  = getSheet('Prenotazioni');
  const now   = isoNow();
  let totale  = 0;
  const idsBk = [];

  body.selezioni.forEach(sel => {
    const newId   = nextId(wsBk, 'P');
    const tipo    = String(sel.tipo || 'deposito');
    const acconto = Number(sel.acconto)       || 0;
    const prezzo  = Number(sel.prezzo_totale) || 0;
    const residuo = Math.max(0, prezzo - acconto);

    const confermata   = tipo === 'deposito';
    const modificabile = tipo === 'deposito';
    const stato        = tipo === 'deposito' ? 'da_saldare'
                       : (residuo === 0 ? 'saldato' : 'da_saldare');

    // Lo SKU salvato in Prenotazioni è sempre quello DNI calcolato
    const skuDni = sel.sku || calcolaSku(sel.vendor_sku || '');

    wsBk.appendRow([
      newId, body.id_utente, u.nome, body.mese,
      sel.prodotto_id, sel.nome_prodotto || '', skuDni, sel.tier_prodotto || '',
      tipo, acconto, 0, prezzo, residuo, now,
      confermata, modificabile, stato, '', '',
    ]);

    totale += acconto;
    idsBk.push(newId);

    const tipoEvento = tipo === 'deposito' ? 'conferma_deposito' : 'conferma_fullpay';
    logAttivita({
      idPrenotazione:  newId,
      idUtente:        body.id_utente,
      nomeUtente:      u.nome,
      timestamp:       now,
      nomeProdotto:    sel.nome_prodotto || '',
      sku:             skuDni,
      tierProdotto:    sel.tier_prodotto || '',
      tipoEvento,
      importo:         acconto,
      accontoBase:     acconto,
      accontoExtraTot: 0,
      prezzoTotale:    prezzo,
      residuo,
      note:            'mese=' + body.mese,
    });

    if (tipo === 'saldo' && ADMIN_EMAIL) {
      emailConfermaFullpay({
        nomeSocio:      u.nome,
        idUtente:       body.id_utente,
        idPrenotazione: newId,
        nomeProdotto:   sel.nome_prodotto || '',
        sku:            skuDni,
        tierProd:       sel.tier_prodotto || '',
        acconto,
        prezzoTotale:   prezzo,
        now,
      });
    }
  });

  const creditoCol = found.headers.indexOf('credito') + 1;
  found.ws.getRange(found.rowIndex, creditoCol)
    .setValue(Math.max(0, (Number(found.data.credito) || 0) - totale));
  const tsCol = found.headers.indexOf('ultimo_aggiornamento') + 1;
  if (tsCol > 0) found.ws.getRange(found.rowIndex, tsCol).setValue(now);

  return corsOutput({
    ok: true,
    messaggio:        'Ordine confermato.',
    id_prenotazioni:  idsBk,
    totale_acconti:   totale,
  });
}

function actionAcExtra(body) {
  if (!body.id_utente || !body.codice || !body.id_prenotazione || body.importo === undefined)
    return errOutput('Dati mancanti: id_utente, codice, id_prenotazione, importo', 400);

  const found = findUtente(body.id_utente);
  if (!found) return errOutput('Utente non trovato', 404);
  if (String(found.data.codice).trim() !== String(body.codice).trim())
    return errOutput('Codice non valido', 401);

  const wsBk    = getSheet('Prenotazioni');
  const rows    = wsBk.getDataRange().getValues();
  const headers = rows[0].map(h => String(h).trim());
  const col     = name => headers.indexOf(name);

  let targetRow = -1;
  for (let i = 2; i < rows.length; i++) {
    if (String(rows[i][col('id')]).trim()        === String(body.id_prenotazione).trim() &&
        String(rows[i][col('id_utente')]).trim() === String(body.id_utente).trim()) {
      targetRow = i + 1; break;
    }
  }
  if (targetRow < 0) return errOutput('Prenotazione non trovata', 404);

  const rowData      = rows[targetRow - 1];
  const modificabile = rowData[col('modificabile')] === true ||
                       String(rowData[col('modificabile')]).trim().toUpperCase() === 'TRUE';
  if (!modificabile) return errOutput('Questa prenotazione non è modificabile', 403);

  const delta              = Number(body.importo)                 || 0;
  const accontoBase        = Number(rowData[col('acconto')])       || 0;
  const prezzoTotale       = Number(rowData[col('prezzo_totale')]) || 0;
  const accontoExtraFoglio = Number(rowData[col('acconto_extra')]) || 0;
  const nomeProdotto       = String(rowData[col('nome_prodotto')] || '');
  const sku                = String(rowData[col('sku')]           || '');
  const tierProd           = String(rowData[col('tier_prodotto')] || '');
  const tipo               = String(rowData[col('tipo')]          || '');
  const dataConf           = String(rowData[col('data_conferma')] || '');

  const nuovoAccontoExtra = accontoExtraFoglio + delta;

  const frontendTot = Number(body.acconto_extra_tot) || 0;
  const discrepanza = Math.abs(nuovoAccontoExtra - frontendTot) > 0.01;
  if (discrepanza) {
    Logger.log('[DISCREPANZA acconto_extra] pren=' + body.id_prenotazione +
      ' foglio+delta=' + nuovoAccontoExtra + ' frontend=' + frontendTot);
  }

  const nuovoResiduo = Math.max(0, prezzoTotale - accontoBase - nuovoAccontoExtra);
  const isSaldato    = nuovoResiduo === 0;
  const now          = isoNow();

  wsBk.getRange(targetRow, col('acconto_extra') + 1).setValue(nuovoAccontoExtra);
  wsBk.getRange(targetRow, col('residuo')       + 1).setValue(nuovoResiduo);
  if (isSaldato) {
    wsBk.getRange(targetRow, col('stato_ordine')     + 1).setValue('saldato');
    wsBk.getRange(targetRow, col('modificabile')     + 1).setValue(false);
    wsBk.getRange(targetRow, col('commento_alberto') + 1).setValue('Spedizione in preparazione');
  }

  const creditoAttuale = Number(found.data.credito) || 0;
  const nuovoCredito   = Math.max(0, creditoAttuale - delta);
  const creditoCol     = found.headers.indexOf('credito') + 1;
  const tsCol          = found.headers.indexOf('ultimo_aggiornamento') + 1;
  found.ws.getRange(found.rowIndex, creditoCol).setValue(nuovoCredito);
  if (tsCol > 0) found.ws.getRange(found.rowIndex, tsCol).setValue(now);

  const tipoEvento = isSaldato ? 'saldo' : 'acconto_extra';
  const noteLog    = discrepanza ? 'DISCREPANZA frontend=' + frontendTot
                   : (body.op_id ? 'op_id=' + body.op_id : '');

  logAttivita({
    idPrenotazione:  body.id_prenotazione,
    idUtente:        body.id_utente,
    nomeUtente:      found.data.nome || '',
    timestamp:       now,
    nomeProdotto,
    sku,
    tierProdotto:    tierProd,
    tipoEvento,
    importo:         delta,
    accontoBase,
    accontoExtraTot: nuovoAccontoExtra,
    prezzoTotale,
    residuo:         nuovoResiduo,
    note:            noteLog,
  });

  if (isSaldato && ADMIN_EMAIL) {
    emailSaldo({
      nomeSocio:         found.data.nome || body.id_utente,
      idUtente:          body.id_utente,
      idPrenotazione:    body.id_prenotazione,
      nomeProdotto,
      sku,
      tierProd,
      tipo,
      accontoBase,
      nuovoAccontoExtra,
      delta,
      prezzoTotale,
      dataConf,
      opId:              body.op_id || '',
      now,
    });
  }

  return corsOutput({
    ok:              true,
    id_prenotazione: body.id_prenotazione,
    op_id:           body.op_id || '',
    acconto_extra:   nuovoAccontoExtra,
    residuo:         nuovoResiduo,
    saldato:         isSaldato,
    credito:         nuovoCredito,
  });
}

function actionSblocca(p) {
  if (!p.id || !p.admin_key) return errOutput('Parametri mancanti: id, admin_key', 400);
  if (String(p.admin_key).trim() !== String(ADMIN_KEY).trim()) return errOutput('Admin key non valida', 401);

  const found = findUtente(p.id);
  if (!found) return errOutput('Utente non trovato', 404);

  const statoCol = found.headers.indexOf('stato')                + 1;
  const tsCol    = found.headers.indexOf('ultimo_aggiornamento') + 1;
  const now      = isoNow();

  found.ws.getRange(found.rowIndex, statoCol).setValue('libero');
  if (tsCol > 0) found.ws.getRange(found.rowIndex, tsCol).setValue(now);

  return corsOutput({ ok: true, messaggio: 'Utente ' + p.id + ' sbloccato.', stato: 'libero', timestamp: now });
}

// ═══════════════════════════════════════════════════════════════════════════════
// AZIONI PUBBLICHE (home e shop senza login)
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Restituisce: notizie pubblicate + spotlight + banner attivo.
 * Usata dalla home page pubblica.
 */
function actionHomePubblica(p) {
  try {
    const notizie   = getNotiziePubbliche();
    const spotlight = getSpotlightPubblico();
    const banner    = getBannerAttivo();
    return corsOutput({ ok: true, notizie, spotlight, banner });
  } catch(err) {
    return errOutput('Errore home pubblica: ' + err.message, 500);
  }
}

/**
 * Shop pubblico — restituisce prodotti dove attivo_shop=TRUE.
 * I dati prodotto vengono letti dal Catalogo tramite Prodotti_Shop.prodotto_id.
 * Aggiunge prezzo_shop dal foglio Prodotti_Shop e disponibile_shop calcolato.
 */
function actionShopPubblico(p) {
  try {
    const wsCat  = getSheet('Catalogo');
    const catAll = sheetToObjects(wsCat, true);
    const wsSh   = getSheet('Prodotti_Shop');
    const shAll  = sheetToObjects(wsSh, false);

    const catFiltro = (p.categoria || '').trim().toLowerCase();

    // Indice Catalogo per id rapido
    const catMap = {};
    catAll.forEach(r => { catMap[String(r.id).trim()] = r; });

    // Costruisce lista prodotti shop
    const prodotti = [];
    catAll
      .filter(r => toBool(r.attivo) && toBool(r.attivo_shop))
      .filter(r => !catFiltro || String(r.categoria || '').trim().toLowerCase() === catFiltro)
      .forEach(r => {
        const prod = mapProdottoCatalogo(r);

        // Cerca entry in Prodotti_Shop per prezzo_shop e descrizione_extra
        const shopEntry = shAll.find(s => String(s.prodotto_id || '').trim() === String(r.id).trim());
        const prezzoShop    = shopEntry ? (Number(shopEntry.prezzo_shop)    || prod.prezzo_pubblico || 0) : (prod.prezzo_pubblico || 0);
        const descExtra     = shopEntry ? String(shopEntry.descrizione_extra || '') : '';

        // Calcola disponibile_shop contando ordini shop confermati (futura logica)
        // Per ora: quota_shop come disponibile
        const disponibile_shop = Math.max(0, prod.quota_shop);

        prodotti.push({
          id:                prod.id,
          nome:              prod.name,
          vendor_sku:        prod.vendor_sku,
          sku:               prod.sku,
          categoria:         prod.categoria,
          tier:              prod.tier,
          prezzo_shop:       prezzoShop,
          prezzo_pubblico:   prod.prezzo_pubblico,
          isTbd:             prod.isTbd,
          imgs:              prod.imgs,
          imgBase:           prod.imgBase,
          stock_in_arrivo:   prod.stock_in_arrivo,
          data_arrivo_prevista: prod.data_arrivo_prevista,
          disponibile_shop,
          esaurito_shop:     prod.quota_shop > 0 && disponibile_shop === 0,
          descrizione_extra: descExtra,
          mese:              prod.mese,
        });
      });

    return corsOutput({ ok: true, prodotti });
  } catch(err) {
    return errOutput('Errore shop pubblico: ' + err.message, 500);
  }
}

// ── Helper lettura pubblica ───────────────────────────────────────────────────

function getNotiziePubbliche() {
  try {
    const ws  = getSheet('Notizie');
    const all = sheetToObjects(ws, false);
    return all
      .filter(r => r.pubblicato === true || String(r.pubblicato).trim().toUpperCase() === 'TRUE')
      .map(r => ({
        id:           r.id,
        titolo:       String(r.titolo       || ''),
        testo:        String(r.testo        || ''),
        immagine_url: String(r.immagine_url || ''),
        data:         String(r.data_pubblicazione || ''),
        autore:       String(r.autore       || 'DNI'),
      }))
      .sort((a, b) => b.data.localeCompare(a.data));
  } catch(e) { return []; }
}

function getSpotlightPubblico() {
  try {
    const wsSp = getSheet('Spotlight');
    const spRows = sheetToObjects(wsSp, false);
    const attivi = spRows
      .filter(r => r.attivo === true || String(r.attivo).trim().toUpperCase() === 'TRUE')
      .sort((a, b) => (Number(a.ordine) || 0) - (Number(b.ordine) || 0));

    if (attivi.length === 0) return [];

    // Recupera dati prodotto dal catalogo per ogni spotlight
    const wsCat   = getSheet('Catalogo');
    const catAll  = sheetToObjects(wsCat, true);
    const wsShop  = getSheet('Prodotti_Shop');
    const shopAll = sheetToObjects(wsShop, false);

    return attivi.map(sp => {
      const pid    = String(sp.prodotto_id || '').trim();
      const fonte  = String(sp.fonte || 'catalogo').trim(); // 'catalogo' | 'shop'
      let prodotto = null;

      if (fonte === 'shop') {
        prodotto = shopAll.find(p => String(p.id || '').trim() === pid ||
                                     String(p.prodotto_id || '').trim() === pid);
      } else {
        prodotto = catAll.find(p => String(p.id || '').trim() === pid);
      }

      if (!prodotto) return null;

      const vendorSku = String(prodotto.vendor_sku || prodotto['external sku'] || '').trim();
      return {
        spotlight_id: sp.id,
        ordine:       Number(sp.ordine) || 0,
        prodotto_id:  pid,
        fonte,
        nome:         String(prodotto.nome_prodotto || prodotto.nome || ''),
        sku:          calcolaSku(vendorSku),
        vendor_sku:   vendorSku,
        tier:         String(prodotto.tier || '').toLowerCase(),
        prezzo:       fonte === 'shop' ? (Number(prodotto.prezzo_shop) || 0)
                                       : (Number(prodotto.prezzo_dbc)  || 0),
        img:          String(prodotto.img || ''),
        mese:         String(prodotto.mese || ''),
      };
    }).filter(Boolean);
  } catch(e) {
    Logger.log('getSpotlightPubblico errore: ' + e.message);
    return [];
  }
}

function getBannerAttivo() {
  try {
    const ws  = getSheet('Banner');
    const all = sheetToObjects(ws, false);
    const now = new Date();
    const attivi = all.filter(r => {
      if (r.attivo !== true && String(r.attivo).trim().toUpperCase() !== 'TRUE') return false;
      if (r.data_scadenza) {
        const scad = new Date(r.data_scadenza);
        if (!isNaN(scad.getTime()) && scad < now) return false;
      }
      return true;
    });
    if (attivi.length === 0) return null;
    const b = attivi[attivi.length - 1]; // usa l'ultimo inserito
    return {
      id:             b.id,
      testo:          String(b.testo   || ''),
      colore:         String(b.colore  || 'info'),
      data_scadenza:  b.data_scadenza ? String(b.data_scadenza) : null,
    };
  } catch(e) { return null; }
}

// ═══════════════════════════════════════════════════════════════════════════════
// AZIONI ADMIN — LETTURA
// ═══════════════════════════════════════════════════════════════════════════════

function actionCatalogoAdmin(p) {
  try {
    requireAdmin(p);
    const mese = (p.mese || '').trim();
    const ws   = getSheet('Catalogo');
    const all  = sheetToObjects(ws, true);

    const result = all
      .filter(r => !mese || String(r.mese || '').trim() === mese)
      .map(r => {
        const prod   = mapProdottoCatalogo(r);
        const stockP = getStockPrenotato(r.id);
        return {
          _row:             r._row,
          ...prod,
          prenotato_vip:    stockP.prenotato_vip,
          disponibile_vip:  Math.max(0, prod.quota_vip  - stockP.prenotato_vip),
          disponibile_shop: Math.max(0, prod.quota_shop),
        };
      });

    return corsOutput({ ok: true, catalogo: result });
  } catch(err) {
    return errOutput(err.message, err.message.includes('Admin') ? 401 : 500);
  }
}

/**
 * Aggiorna le quote stock di un prodotto del catalogo.
 * Usato dalla pagina Magazzino dell'admin.
 * body: { admin_key, id, stock_totale, quota_vip, quota_shop,
 *         attivo_vip, attivo_shop, stock_in_arrivo, data_arrivo_prevista }
 */
function actionAggiornaStock(body) {
  try {
    requireAdmin(body);
    if (!body.id) return errOutput('id prodotto mancante', 400);

    const ws      = getSheet('Catalogo');
    const rows    = ws.getDataRange().getValues();
    const headers = rows[0].map(h => String(h).trim());
    const idCol   = headers.indexOf('id');

    let targetRow = -1;
    for (let i = 2; i < rows.length; i++) {
      if (String(rows[i][idCol]).trim() === String(body.id).trim()) {
        targetRow = i + 1; break;
      }
    }
    if (targetRow < 0) return errOutput('Prodotto non trovato', 404);

    const campi = [
      'stock_totale','quota_vip','quota_shop',
      'attivo_vip','attivo_shop',
      'stock_in_arrivo','data_arrivo_prevista',
    ];
    campi.forEach(campo => {
      if (body[campo] === undefined) return;
      const colIdx = headers.indexOf(campo);
      if (colIdx >= 0) ws.getRange(targetRow, colIdx + 1).setValue(body[campo]);
    });

    // Ricalcola quota_shop se non passata esplicitamente
    if (body.stock_totale !== undefined && body.quota_vip !== undefined && body.quota_shop === undefined) {
      const qShopCol = headers.indexOf('quota_shop');
      if (qShopCol >= 0) {
        const nuovaQuotaShop = Math.max(0, Number(body.stock_totale) - Number(body.quota_vip));
        ws.getRange(targetRow, qShopCol + 1).setValue(nuovaQuotaShop);
      }
    }

    // Aggiorna Prodotti_Shop automaticamente se attivo_shop cambia
    if (body.attivo_shop !== undefined) {
      try {
        const wsSh    = getSheet('Prodotti_Shop');
        const shRows  = wsSh.getDataRange().getValues();
        const shHdrs  = shRows[0].map(h => String(h).trim());
        const pidCol  = shHdrs.indexOf('prodotto_id');
        let   found   = false;

        for (let i = 1; i < shRows.length; i++) {
          if (String(shRows[i][pidCol] || '').trim() === String(body.id).trim()) {
            found = true;
            break;
          }
        }

        // Se attivo_shop=TRUE e non esiste ancora in Prodotti_Shop, crea entry automatica
        if (body.attivo_shop === true && !found) {
          const newShId = nextIdSimple(wsSh, 'S');
          const prezzoShop = Number(body.prezzo_shop) || 0;
          wsSh.appendRow([newShId, body.id, prezzoShop, '']);
          Logger.log('Prodotti_Shop: aggiunta entry automatica per prodotto ' + body.id);
        }
      } catch(shErr) {
        Logger.log('Sync Prodotti_Shop fallita (non blocca): ' + shErr.message);
      }
    }

    return corsOutput({ ok: true, messaggio: 'Stock aggiornato per prodotto ' + body.id });
  } catch(err) {
    return errOutput(err.message, err.message.includes('Admin') ? 401 : 500);
  }
}

/**
 * Magazzino admin — vista completa stock con semafori.
 * Restituisce tutti i prodotti (tutti i mesi) con dati stock live.
 * Filtrabile per mese, tier, categoria, attivo_vip, attivo_shop.
 */
function actionMagazzino(p) {
  try {
    requireAdmin(p);
    const ws  = getSheet('Catalogo');
    const all = sheetToObjects(ws, true);

    const meseF     = (p.mese      || '').trim();
    const tierF     = (p.tier      || '').trim().toLowerCase();
    const catF      = (p.categoria || '').trim().toLowerCase();

    const result = all
      .filter(r => {
        if (meseF && String(r.mese || '').trim() !== meseF) return false;
        if (tierF && String(r.tier || '').trim().toLowerCase() !== tierF) return false;
        if (catF  && String(r.categoria || '').trim().toLowerCase() !== catF) return false;
        return true;
      })
      .map(r => {
        const prod   = mapProdottoCatalogo(r);
        const stockP = getStockPrenotato(r.id);

        const disponibile_vip  = Math.max(0, prod.quota_vip  - stockP.prenotato_vip);
        const disponibile_shop = Math.max(0, prod.quota_shop);
        const pct_vip          = prod.quota_vip  > 0 ? Math.round((stockP.prenotato_vip / prod.quota_vip) * 100) : 0;

        // Semaforo: 'green' > 50%, 'yellow' 1-50%, 'red' = 0
        const semaforo_vip  = disponibile_vip  === 0 ? 'red'   : pct_vip  >= 50 ? 'yellow' : 'green';
        const semaforo_shop = disponibile_shop === 0 ? 'red'   : 'green';

        return {
          id:                prod.id,
          mese:              prod.mese,
          tier:              prod.tier,
          categoria:         prod.categoria,
          name:              prod.name,
          vendor_sku:        prod.vendor_sku,
          sku:               prod.sku,
          attivo_vip:        prod.attivo_vip,
          attivo_shop:       prod.attivo_shop,
          attivo:            prod.attivo,
          stock_in_arrivo:   prod.stock_in_arrivo,
          data_arrivo_prevista: prod.data_arrivo_prevista,
          // stock numerico
          stock_totale:      prod.stock_totale,
          quota_vip:         prod.quota_vip,
          quota_shop:        prod.quota_shop,
          prenotato_vip:     stockP.prenotato_vip,
          disponibile_vip,
          disponibile_shop,
          // semafori visivi
          semaforo_vip,
          semaforo_shop,
          pct_prenotato_vip: pct_vip,
        };
      });

    return corsOutput({ ok: true, magazzino: result });
  } catch(err) {
    return errOutput(err.message, err.message.includes('Admin') ? 401 : 500);
  }
}

function actionNotizie(p) {
  try {
    requireAdmin(p);
    const ws  = getSheet('Notizie');
    const all = sheetToObjects(ws, false);
    return corsOutput({ ok: true, notizie: all });
  } catch(err) {
    return errOutput(err.message, err.message.includes('Admin') ? 401 : 500);
  }
}

function actionProdottiShop(p) {
  try {
    requireAdmin(p);
    const ws  = getSheet('Prodotti_Shop');
    const all = sheetToObjects(ws, false);
    const result = all.map(r => ({
      _row:        r._row,
      id:          r.id,
      prodotto_id: r.prodotto_id,
      nome:        String(r.nome        || ''),
      vendor_sku:  String(r.vendor_sku  || ''),
      sku:         calcolaSku(r.vendor_sku || ''),
      categoria:   String(r.categoria   || ''),
      tier:        String(r.tier        || ''),
      prezzo_dbc:  Number(r.prezzo_dbc) || 0,
      prezzo_shop: Number(r.prezzo_shop)|| 0,
      img:         String(r.img         || ''),
      descrizione: String(r.descrizione || ''),
      attivo:      r.attivo === true || String(r.attivo).trim().toUpperCase() === 'TRUE',
      mese:        String(r.mese        || ''),
    }));
    return corsOutput({ ok: true, prodotti: result });
  } catch(err) {
    return errOutput(err.message, err.message.includes('Admin') ? 401 : 500);
  }
}

function actionSpotlight(p) {
  try {
    requireAdmin(p);
    const ws  = getSheet('Spotlight');
    const all = sheetToObjects(ws, false);
    return corsOutput({ ok: true, spotlight: all });
  } catch(err) {
    return errOutput(err.message, err.message.includes('Admin') ? 401 : 500);
  }
}

function actionBanner(p) {
  try {
    requireAdmin(p);
    const ws  = getSheet('Banner');
    const all = sheetToObjects(ws, false);
    return corsOutput({ ok: true, banner: all });
  } catch(err) {
    return errOutput(err.message, err.message.includes('Admin') ? 401 : 500);
  }
}

function actionUtentiAdmin(p) {
  try {
    requireAdmin(p);
    const ws  = getSheet('Utenti');
    const all = sheetToObjects(ws, true);
    // Non esporre il PIN (codice) nell'elenco admin
    const result = all.map(r => {
      const u = Object.assign({}, r);
      delete u.codice;
      return u;
    });
    return corsOutput({ ok: true, utenti: result });
  } catch(err) {
    return errOutput(err.message, err.message.includes('Admin') ? 401 : 500);
  }
}

function actionPrenotazioniAdmin(p) {
  try {
    requireAdmin(p);
    const ws  = getSheet('Prenotazioni');
    const all = sheetToObjects(ws, true);
    // filtro opzionale per mese o utente
    const meseF  = (p.mese     || '').trim();
    const utenteF= (p.id_utente|| '').trim();
    const result = all.filter(r =>
      (!meseF   || String(r.mese      || '').trim() === meseF) &&
      (!utenteF || String(r.id_utente || '').trim() === utenteF)
    );
    return corsOutput({ ok: true, prenotazioni: result });
  } catch(err) {
    return errOutput(err.message, err.message.includes('Admin') ? 401 : 500);
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
// AZIONI ADMIN — CATALOGO
// ═══════════════════════════════════════════════════════════════════════════════

function actionAggiungiProdottoCatalogo(body) {
  try {
    requireAdmin(body);
    const ws        = getSheet('Catalogo');
    const rows      = ws.getDataRange().getValues();
    const headers   = rows[0].map(h => String(h).trim());

    // Genera id nel formato AAAAMM + slot (es. 202603301)
    const mese = String(body.mese || meseCorrente()).trim();
    const [anno, meseNum] = mese.split('-');
    // Trova il numero più alto di slot per questo mese
    let maxSlot = 100;
    for (let i = 2; i < rows.length; i++) {
      const idVal = String(rows[i][headers.indexOf('id')] || '');
      const prefix = anno + meseNum;
      if (idVal.startsWith(prefix)) {
        const slot = parseInt(idVal.slice(prefix.length), 10);
        if (!isNaN(slot) && slot > maxSlot) maxSlot = slot;
      }
    }
    const newId = anno + meseNum + String(maxSlot + 1);

    const vendorSku = String(body.vendor_sku || '').trim();
    const isTbd     = body.is_tbd === true || String(body.is_tbd || '').toUpperCase() === 'TRUE';

    // Costruisce la riga rispettando l'ordine degli header
    const row = headers.map(h => {
      switch(h) {
        case 'mese':          return mese;
        case 'id':            return newId;
        case 'tier':          return String(body.tier || 'scout').toLowerCase();
        case 'categoria':     return String(body.categoria || 'model kit').toLowerCase();
        case 'nome_prodotto': return String(body.nome_prodotto || '');
        case 'vendor_sku':    return vendorSku;
        case 'sku':           return calcolaSku(vendorSku);
        case 'prezzo_dbc':    return isTbd ? '' : (Number(body.prezzo_dbc) || '');
        case 'prezzo_pubblico':return isTbd ? '' : (Number(body.prezzo_pubblico) || '');
        case 'is_tbd':        return isTbd;
        case 'acconto_base':  return Number(body.acconto_base) || 0;
        case 'img':           return String(body.img || '');
        case 'attivo':        return body.attivo !== false;
        case 'note':          return String(body.note || '');
        default:              return '';
      }
    });

    ws.appendRow(row);
    return corsOutput({ ok: true, id: newId, sku: calcolaSku(vendorSku), messaggio: 'Prodotto aggiunto al catalogo.' });
  } catch(err) {
    return errOutput(err.message, err.message.includes('Admin') ? 401 : 500);
  }
}

function actionAggiornaProdottoCatalogo(body) {
  try {
    requireAdmin(body);
    if (!body.id) return errOutput('id prodotto mancante', 400);

    const ws      = getSheet('Catalogo');
    const rows    = ws.getDataRange().getValues();
    const headers = rows[0].map(h => String(h).trim());
    const idCol   = headers.indexOf('id');

    let targetRow = -1;
    for (let i = 2; i < rows.length; i++) {
      if (String(rows[i][idCol]).trim() === String(body.id).trim()) {
        targetRow = i + 1; break;
      }
    }
    if (targetRow < 0) return errOutput('Prodotto non trovato nel catalogo', 404);

    const campiAggiornabili = ['tier','categoria','nome_prodotto','vendor_sku','prezzo_dbc',
                               'prezzo_pubblico','is_tbd','acconto_base',
                               'img','attivo','note'];

    campiAggiornabili.forEach(campo => {
      if (body[campo] === undefined) return;
      const colIdx = headers.indexOf(campo);
      if (colIdx < 0) return;
      let val = body[campo];
      // Se aggiorno vendor_sku, aggiorno anche sku
      if (campo === 'vendor_sku') {
        const skuCol = headers.indexOf('sku');
        if (skuCol >= 0) ws.getRange(targetRow, skuCol + 1).setValue(calcolaSku(String(val).trim()));
      }
      ws.getRange(targetRow, colIdx + 1).setValue(val);
    });

    return corsOutput({ ok: true, messaggio: 'Prodotto catalogo aggiornato.' });
  } catch(err) {
    return errOutput(err.message, err.message.includes('Admin') ? 401 : 500);
  }
}

function actionEliminaProdottoCatalogo(body) {
  try {
    requireAdmin(body);
    if (!body.id) return errOutput('id prodotto mancante', 400);

    const ws      = getSheet('Catalogo');
    const rows    = ws.getDataRange().getValues();
    const headers = rows[0].map(h => String(h).trim());
    const idCol   = headers.indexOf('id');
    const attivoCol = headers.indexOf('attivo');

    for (let i = 2; i < rows.length; i++) {
      if (String(rows[i][idCol]).trim() === String(body.id).trim()) {
        // Soft delete: mette attivo=FALSE invece di cancellare la riga
        if (attivoCol >= 0) ws.getRange(i + 1, attivoCol + 1).setValue(false);
        return corsOutput({ ok: true, messaggio: 'Prodotto disattivato dal catalogo.' });
      }
    }
    return errOutput('Prodotto non trovato', 404);
  } catch(err) {
    return errOutput(err.message, err.message.includes('Admin') ? 401 : 500);
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
// AZIONI ADMIN — NOTIZIE
// ═══════════════════════════════════════════════════════════════════════════════

function actionSalvaNotizia(body) {
  try {
    requireAdmin(body);
    const ws = getSheet('Notizie');

    if (body.id) {
      // Aggiorna notizia esistente
      const rows    = ws.getDataRange().getValues();
      const headers = rows[0].map(h => String(h).trim());
      const idCol   = headers.indexOf('id');
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][idCol]).trim() === String(body.id).trim()) {
          const campi = ['titolo','testo','immagine_url','data_pubblicazione','pubblicato','autore'];
          campi.forEach(c => {
            if (body[c] !== undefined) {
              const col = headers.indexOf(c);
              if (col >= 0) ws.getRange(i + 1, col + 1).setValue(body[c]);
            }
          });
          return corsOutput({ ok: true, messaggio: 'Notizia aggiornata.' });
        }
      }
      return errOutput('Notizia non trovata', 404);
    } else {
      // Crea nuova notizia
      const newId = nextIdSimple(ws, 'N');
      ws.appendRow([
        newId,
        String(body.titolo       || ''),
        String(body.testo        || ''),
        String(body.immagine_url || ''),
        body.data_pubblicazione  || isoNow().split('T')[0],
        body.pubblicato !== false,
        String(body.autore       || 'Admin'),
      ]);
      return corsOutput({ ok: true, id: newId, messaggio: 'Notizia creata.' });
    }
  } catch(err) {
    return errOutput(err.message, err.message.includes('Admin') ? 401 : 500);
  }
}

function actionEliminaNotizia(body) {
  try {
    requireAdmin(body);
    if (!body.id) return errOutput('id notizia mancante', 400);
    const ws      = getSheet('Notizie');
    const rows    = ws.getDataRange().getValues();
    const headers = rows[0].map(h => String(h).trim());
    const idCol   = headers.indexOf('id');
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][idCol]).trim() === String(body.id).trim()) {
        ws.deleteRow(i + 1);
        return corsOutput({ ok: true, messaggio: 'Notizia eliminata.' });
      }
    }
    return errOutput('Notizia non trovata', 404);
  } catch(err) {
    return errOutput(err.message, err.message.includes('Admin') ? 401 : 500);
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
// AZIONI ADMIN — PRODOTTI SHOP
// ═══════════════════════════════════════════════════════════════════════════════

function actionSalvaProdottoShop(body) {
  try {
    requireAdmin(body);
    const ws = getSheet('Prodotti_Shop');

    if (body.id) {
      // Aggiorna prodotto esistente
      const rows    = ws.getDataRange().getValues();
      const headers = rows[0].map(h => String(h).trim());
      const idCol   = headers.indexOf('id');
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][idCol]).trim() === String(body.id).trim()) {
          const campi = ['nome','vendor_sku','categoria','tier','prezzo_dbc',
                         'prezzo_shop','img','descrizione','attivo','mese'];
          campi.forEach(c => {
            if (body[c] !== undefined) {
              const col = headers.indexOf(c);
              if (col >= 0) ws.getRange(i + 1, col + 1).setValue(body[c]);
            }
          });
          return corsOutput({ ok: true, messaggio: 'Prodotto shop aggiornato.' });
        }
      }
      return errOutput('Prodotto shop non trovato', 404);
    } else {
      const newId     = nextIdSimple(ws, 'S');
      const vendorSku = String(body.vendor_sku || '').trim();
      ws.appendRow([
        newId,
        body.prodotto_id    || '',
        String(body.nome    || ''),
        vendorSku,
        String(body.categoria   || ''),
        String(body.tier        || ''),
        Number(body.prezzo_dbc) || 0,
        Number(body.prezzo_shop)|| 0,
        String(body.img         || ''),
        String(body.descrizione || ''),
        body.attivo !== false,
        String(body.mese        || meseCorrente()),
      ]);
      return corsOutput({ ok: true, id: newId, sku: calcolaSku(vendorSku), messaggio: 'Prodotto shop creato.' });
    }
  } catch(err) {
    return errOutput(err.message, err.message.includes('Admin') ? 401 : 500);
  }
}

function actionEliminaProdottoShop(body) {
  try {
    requireAdmin(body);
    if (!body.id) return errOutput('id prodotto mancante', 400);
    const ws      = getSheet('Prodotti_Shop');
    const rows    = ws.getDataRange().getValues();
    const headers = rows[0].map(h => String(h).trim());
    const idCol   = headers.indexOf('id');
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][idCol]).trim() === String(body.id).trim()) {
        ws.deleteRow(i + 1);
        return corsOutput({ ok: true, messaggio: 'Prodotto shop eliminato.' });
      }
    }
    return errOutput('Prodotto shop non trovato', 404);
  } catch(err) {
    return errOutput(err.message, err.message.includes('Admin') ? 401 : 500);
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
// AZIONI ADMIN — SPOTLIGHT & BANNER
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Salva la lista spotlight completa (sostituisce tutto).
 * body.items = [ { prodotto_id, fonte, ordine, attivo }, ... ]
 */
function actionSalvaSpotlight(body) {
  try {
    requireAdmin(body);
    if (!Array.isArray(body.items)) return errOutput('items[] mancante', 400);

    const ws      = getSheet('Spotlight');
    const lastRow = ws.getLastRow();
    // Cancella tutti i dati (mantieni solo header)
    if (lastRow > 1) ws.getRange(2, 1, lastRow - 1, ws.getLastColumn()).clearContent();

    body.items.forEach((item, idx) => {
      const id = 'SP' + String(idx + 1).padStart(4, '0');
      ws.appendRow([
        id,
        String(item.prodotto_id || ''),
        String(item.fonte       || 'catalogo'),
        Number(item.ordine)     || (idx + 1),
        item.attivo !== false,
      ]);
    });

    return corsOutput({ ok: true, messaggio: 'Spotlight salvato (' + body.items.length + ' elementi).' });
  } catch(err) {
    return errOutput(err.message, err.message.includes('Admin') ? 401 : 500);
  }
}

/**
 * Crea o sostituisce il banner attivo.
 * Per disattivare: body.attivo = false
 */
function actionSalvaBanner(body) {
  try {
    requireAdmin(body);
    const ws      = getSheet('Banner');
    const lastRow = ws.getLastRow();

    if (body.id) {
      // Aggiorna banner esistente
      const rows    = ws.getDataRange().getValues();
      const headers = rows[0].map(h => String(h).trim());
      const idCol   = headers.indexOf('id');
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][idCol]).trim() === String(body.id).trim()) {
          const campi = ['testo','colore','data_scadenza','attivo'];
          campi.forEach(c => {
            if (body[c] !== undefined) {
              const col = headers.indexOf(c);
              if (col >= 0) ws.getRange(i + 1, col + 1).setValue(body[c]);
            }
          });
          return corsOutput({ ok: true, messaggio: 'Banner aggiornato.' });
        }
      }
    }

    // Crea nuovo banner
    const newId = 'BN' + String(lastRow).padStart(4, '0');
    ws.appendRow([
      newId,
      String(body.testo          || ''),
      String(body.colore         || 'info'),
      body.data_scadenza         || '',
      body.attivo !== false,
    ]);
    return corsOutput({ ok: true, id: newId, messaggio: 'Banner creato.' });
  } catch(err) {
    return errOutput(err.message, err.message.includes('Admin') ? 401 : 500);
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
// AZIONI ADMIN — SOCI & PRENOTAZIONI
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Aggiorna campi di un utente.
 * Non aggiorna mai il campo "codice" (PIN) per sicurezza — operazione separata.
 */
function actionAggiornaUtente(body) {
  try {
    requireAdmin(body);
    if (!body.id) return errOutput('id utente mancante', 400);

    const found = findUtente(body.id);
    if (!found) return errOutput('Utente non trovato', 404);

    const campiAggiornabili = [
      'nome','tier','credito','stato','pagamento',
      'acconto_speciale_pct_scout','acconto_speciale_pct_veteran','acconto_speciale_pct_elite',
      'credito_bloccato_pct_scout','credito_bloccato_pct_veteran','credito_bloccato_pct_elite',
      'note','data_iscrizione','data_ultimo_rinnovo','prossima_data_rinnovo',
    ];

    const now = isoNow();
    campiAggiornabili.forEach(campo => {
      if (body[campo] === undefined) return;
      const colIdx = found.headers.indexOf(campo);
      if (colIdx < 0) return;
      found.ws.getRange(found.rowIndex, colIdx + 1).setValue(body[campo]);
    });

    // Aggiorna timestamp
    const tsCol = found.headers.indexOf('ultimo_aggiornamento') + 1;
    if (tsCol > 0) found.ws.getRange(found.rowIndex, tsCol).setValue(now);

    return corsOutput({ ok: true, messaggio: 'Utente ' + body.id + ' aggiornato.', timestamp: now });
  } catch(err) {
    return errOutput(err.message, err.message.includes('Admin') ? 401 : 500);
  }
}

/**
 * Aggiorna campi di una prenotazione.
 * Usato dall'admin per: cambiare stato_ordine, commento_alberto, tracking,
 * modificabile, confermata.
 */
function actionAggiornaPrenotazione(body) {
  try {
    requireAdmin(body);
    if (!body.id) return errOutput('id prenotazione mancante', 400);

    const ws      = getSheet('Prenotazioni');
    const rows    = ws.getDataRange().getValues();
    const headers = rows[0].map(h => String(h).trim());
    const idCol   = headers.indexOf('id');

    let targetRow = -1;
    for (let i = 2; i < rows.length; i++) {
      if (String(rows[i][idCol]).trim() === String(body.id).trim()) {
        targetRow = i + 1; break;
      }
    }
    if (targetRow < 0) return errOutput('Prenotazione non trovata', 404);

    const campiAggiornabili = [
      'stato_ordine','commento_alberto','tracking',
      'modificabile','confermata',
      'acconto','acconto_extra','prezzo_totale','residuo',
    ];

    campiAggiornabili.forEach(campo => {
      if (body[campo] === undefined) return;
      const colIdx = headers.indexOf(campo);
      if (colIdx < 0) return;
      ws.getRange(targetRow, colIdx + 1).setValue(body[campo]);
    });

    return corsOutput({ ok: true, messaggio: 'Prenotazione ' + body.id + ' aggiornata.' });
  } catch(err) {
    return errOutput(err.message, err.message.includes('Admin') ? 401 : 500);
  }
}
