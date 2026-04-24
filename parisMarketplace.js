/*******************************
 * Paris Marketplace (STG) - Google Apps Script
 *
 * Script Properties (MANUAL):
 *  - PARIS_API_KEY_PROD || PARIS_API_KEY_STG    (obligatorio)
 *  - PARIS_URL_PROD    (opcional) ej: https://api-developers.ecomm.cencosud.com
 *  - PARIS_URL_STG     (opcional) ej: https://api-developers.ecomm-stg.cencosud.com
 *  - PARIS_WAREHOUSE   (opcional) default: ffparis
 *
 * Script Properties (AUTO):
 *  - PARIS_ACCESS_TOKEN
 *  - PARIS_ACCESS_EXPIRES_MS
 *******************************/

// ---- Endpoints confirmados ----
const PARIS_ENDPOINTS = {
  AUTH_APIKEY: '/v1/auth/apiKey',
  PRODUCTS_SEARCH: '/v2/products/search',
  PRICE_TYPES: '/v2/price-types',
  STOCK_SKU_SELLER: '/v1/stock/sku-seller',
  PRICES_UPSERT_V2: '/v2/prices/product', // + '/{productIdPadre}'
  STOCK_GET_V2: '/v2/stock'
};

// ---- Ambientes ----
const PARIS_ENV = 'PROD'; // 'STG' o 'PROD'

// ---- Price types (confirmados) ----
const PARIS_PRICE_TYPE_LISTA  = '6503baaf-16d0-4590-a4d6-494719593a12'; // "Precio"
const PARIS_PRICE_TYPE_OFERTA = 'c25aaf10-fd85-416d-b2bd-d64ffa174ba7'; // "Precio oferta"

const PARIS_LOG_TIMING = true;     // prende/apaga logs de tiempo // Cuando ya esté estable, deja: PARIS_LOG_TIMING = false y prendes solo cuando sientas lentitud.
const PARIS_LOG_BODY_ON_ERR = true; // loguea body solo si hay error
const PARIS_TIMING_STATS = {}; // { key: {count,totalMs,maxMs} }

// ---- storePrice fijo (según doc/colección) ----
const PARIS_STORE_PRICE_ID = '8678fdf5-86f9-4530-aaee-dd67b9843976';

// ---- Sheets ----
// const SHEET_PRODS = 'Copia de Prods. PM';
// const SHEET_PRECIOS = 'Copia de Precios. PM';
const SHEET_PRODS = 'Prods. PM';
const SHEET_PRECIOS = 'Precios. PM';

// ===== Timing / Stats (por endpoint normalizado) =====
function paris_normalizePathForStats_(path) {
  const p = String(path || '');
  const base = p.split('?')[0]; // quita query string

  // Normalizaciones clave (para que no quede por parentId)
  if (base.indexOf('/v2/prices/product/') === 0) return '/v2/prices/product/*';
  if (base.indexOf('/v2/products/search') === 0) return '/v2/products/search';

  return base;
}

function paris_timingAdd_(key, ms) {
  const s = PARIS_TIMING_STATS[key] || (PARIS_TIMING_STATS[key] = { count: 0, totalMs: 0, maxMs: 0 });
  s.count++;
  s.totalMs += ms;
  if (ms > s.maxMs) s.maxMs = ms;
}

function paris_timingReport_() {
  const keys = Object.keys(PARIS_TIMING_STATS).sort();
  for (const k of keys) {
    const s = PARIS_TIMING_STATS[k];
    const avg = Math.round(s.totalMs / Math.max(1, s.count));
    Logger.log(`[PARIS][STATS] ${k} count=${s.count} avg=${avg}ms max=${s.maxMs}ms total=${s.totalMs}ms`);
  }
}
// ===== FIN Timing / Stats (por endpoint normalizado) =====

// ---- Helpers: base url / warehouse ----
function paris_getApiKey_() {
  const props = PropertiesService.getScriptProperties();
  return PARIS_ENV === 'PROD'
    ? props.getProperty('PARIS_API_KEY_PROD') 
    : props.getProperty('PARIS_API_KEY_STG');
}

function paris_getBaseUrl_() {
  const props = PropertiesService.getScriptProperties();
  return PARIS_ENV === 'PROD'
    ? props.getProperty('PARIS_URL_PROD') || 'https://api-developers.ecomm.cencosud.com'
    : props.getProperty('PARIS_URL_STG') || 'https://api-developers.ecomm-stg.cencosud.com';
}

function paris_getWarehouse_() {
  const props = PropertiesService.getScriptProperties();
  return props.getProperty('PARIS_WAREHOUSE') || 'ffparis';
}

// ---- Token on-demand (API Key -> accessToken 4h) ----
function paris_getAccessToken_() {
  const props = PropertiesService.getScriptProperties();
  const now = Date.now();

  const token = props.getProperty('PARIS_ACCESS_TOKEN');
  const expMs = Number(props.getProperty('PARIS_ACCESS_EXPIRES_MS') || 0);

  if (token && expMs && now < (expMs - 60000)) return token; // margen 60s

  const fresh = paris_fetchNewAccessToken_();
  props.setProperty('PARIS_ACCESS_TOKEN', fresh.accessToken);
  props.setProperty('PARIS_ACCESS_EXPIRES_MS', String(now + fresh.expiresInMs));
  return fresh.accessToken;
}

function paris_fetchNewAccessToken_() {
  const props = PropertiesService.getScriptProperties();
  const apiKey = paris_getApiKey_();

  if (!apiKey) {
    const need = (PARIS_ENV === 'PROD') ? 'PARIS_API_KEY_PROD' : 'PARIS_API_KEY_STG';
    throw new Error('Paris: falta Script Property ' + need);
  }

  const url = paris_getBaseUrl_() + PARIS_ENDPOINTS.AUTH_APIKEY;

  const resp = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: { 'Authorization': 'Bearer ' + apiKey },
    muteHttpExceptions: true
  });

  const code = resp.getResponseCode();
  const text = resp.getContentText();
  if (code !== 200) throw new Error('Paris AUTH error ' + code + ': ' + text);

  const data = JSON.parse(text);
  if (!data.accessToken || !data.expiresIn) throw new Error('Paris AUTH response inesperada: ' + text);

  return { accessToken: data.accessToken, expiresInMs: Number(data.expiresIn) * 1000 };
}

// ---- Fetch JSON helper (renueva token si 401/403) ----
function paris_fetchJson_(path, options) {
  const url = paris_getBaseUrl_() + path;

  options = options || {};
  options.headers = options.headers || {};
  options.muteHttpExceptions = true;

  let token = paris_getAccessToken_();
  options.headers.Authorization = 'Bearer ' + token;
  if (!options.headers.Accept) options.headers.Accept = 'application/json';

  const fetchFn = (path === PARIS_ENDPOINTS.STOCK_SKU_SELLER)
    ? paris_fetchWithRetryFast_
    : paris_fetchWithRetry_;

  const method = String(options.method || 'get').toLowerCase();
  const statKey = method + ' ' + paris_normalizePathForStats_(path); // agrupa /v2/prices/product/*, etc. 

  const t0 = Date.now();

  // --- 1er intento ---
  let out = fetchFn(url, options);

  // Soporta ambos retornos:
  // - HTTPResponse directo
  // - { resp: HTTPResponse, retries: n, fast: true }
  let resp = out && out.resp ? out.resp : out;
  let retries = out && out.resp ? Number(out.retries || 0) : 0;
  let fastTag = out && out.resp && out.fast ? ' [FAST]' : '';

  let code = resp.getResponseCode();

  // --- Refresh token si 401/403 ---
  if (code === 401 || code === 403) {
    const props = PropertiesService.getScriptProperties();
    props.deleteProperty('PARIS_ACCESS_TOKEN');
    props.deleteProperty('PARIS_ACCESS_EXPIRES_MS');

    token = paris_getAccessToken_();
    options.headers.Authorization = 'Bearer ' + token;

    out = fetchFn(url, options);
    resp = out && out.resp ? out.resp : out;
    retries += (out && out.resp ? Number(out.retries || 0) : 0);
    fastTag = fastTag || (out && out.resp && out.fast ? ' [FAST]' : '');

    code = resp.getResponseCode();
  }

  const ms = Date.now() - t0;

  // stats por endpoint normalizado
  paris_timingAdd_(statKey, ms);

  if (PARIS_LOG_TIMING) {
    Logger.log(`[PARIS] ${method} ${path} -> ${code} in ${ms}ms (retries=${retries})${fastTag}`);
  }

  const text = resp.getContentText();
  if (code < 200 || code >= 300) {
    throw new Error('Paris API error ' + code + ' @ ' + path + ': ' + text);
  }

  return text ? JSON.parse(text) : null;
}

// =========================
//  MAPEO SKU -> PRODUCT ID PADRE (para precios)
// =========================
function paris_findParentIdBySku_(skuMarketplaceOrParentId) {
  const s = String(skuMarketplaceOrParentId || '').trim();
  if (!s) throw new Error('SKU vacío');

  // Si no tiene '-', normalmente ya es el ID padre (ej: MK87G8VE1K)
  if (s.indexOf('-') === -1) return s;

  const variantSku = s;

  // 0) cache rápido (si ya existe)
  const cache = CacheService.getScriptCache();
  const cacheKey = 'PARIS_PARENTID_' + variantSku;
  const cached = cache.get(cacheKey);
  if (cached) return String(cached);

  // 1) paginado real
  const limit = 100;
  let offset = 0;

  while (true) {
    const data = paris_fetchJson_(
      PARIS_ENDPOINTS.PRODUCTS_SEARCH + '?limit=' + limit + '&offset=' + offset,
      { method: 'get' }
    );

    const results = (data && data.results) ? data.results : [];
    if (!results.length) break;

    for (let i = 0; i < results.length; i++) {
      const r = results[i];
      const vars = Array.isArray(r.variants) ? r.variants : [];
      for (let j = 0; j < vars.length; j++) {
        if (String(vars[j].sku || '').trim() === variantSku) {
          const parentId = String(r.id || '').trim();
          if (!parentId) break;

          // cache 6h
          try { cache.put(cacheKey, parentId, 21600); } catch (e) {}
          return parentId;
        }
      }
    }

    offset += limit;
    if (data.total !== undefined && offset >= Number(data.total)) break;
  }

  // 2) fallback final: prefijo antes del '-'
  // (si esto te preocupa, en vez de devolver, puedes throw para “fail safe”)
  const parentPrefix = variantSku.split('-')[0];
  if (parentPrefix) {
    try { cache.put(cacheKey, parentPrefix, 21600); } catch (e) {}
    return parentPrefix;
  }

  throw new Error('No se pudo mapear SKU variante a ID padre: ' + variantSku);
}

// =========================
//  SHEETS UTIL
// =========================
function sheet_getHeaderMap_(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach((h, idx) => {
    const key = String(h || '').trim();
    if (key) map[key] = idx + 1; // 1-based
  });
  return map;
}

function toNumber_(v) {
  if (v === null || v === '') return null;
  if (typeof v === 'number') return v;
  const s = String(v).replace(/\./g, '').replace(',', '.').trim(); // por si viene con separadores
  const n = Number(s);
  return isFinite(n) ? n : null;
}

function toIsoZ_(v) {
  // Acepta Date o string
  if (!v) return null;
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime())) {
    return Utilities.formatDate(v, 'UTC', "yyyy-MM-dd'T'HH:mm:ss.SSS'Z'");
  }
  // Si viene string, intentar Date.parse
  const d = new Date(v);
  if (!isNaN(d.getTime())) {
    return Utilities.formatDate(d, 'UTC', "yyyy-MM-dd'T'HH:mm:ss.SSS'Z'");
  }
  return null;
}

// =========================
//  1) ACTUALIZAR PRECIOS (solo lista)
//  - Efecto: deja lista y "baja" oferta (no la envía)
//  - Hoja: "Copia de Precios. PM"
//  - Respeta: NO tocar L:V (solo lee A:K)
//  - Control: columna A "Actualizar" (TRUE/FALSE)
// =========================
function paris_updatePriceList_fromSheet() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_PRECIOS);
  if (!sh) throw new Error('No existe hoja: ' + SHEET_PRECIOS);

  const hm = sheet_getHeaderMap_(sh);
  const colActualizar = hm['Actualizar'];
  const colSku = hm['SKU(*)'];
  const colPrecio = hm['Precio(*)'];

  if (!colActualizar || !colSku || !colPrecio) {
    throw new Error('Headers faltantes en ' + SHEET_PRECIOS + '. Revisa: Actualizar, SKU(*), Precio(*)');
  }

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  const numRows = lastRow - 1;

  const readCols = Math.min(11, sh.getLastColumn()); // A:K
  const values = sh.getRange(2, 1, numRows, readCols).getValues();

  const actualizarRange = sh.getRange(2, colActualizar, numRows, 1);
  const actualizarVals = actualizarRange.getValues();

  // Indicadores (misma columna "Actualizar")
  const bgRange = sh.getRange(2, colActualizar, numRows, 1);
  const noteRange = sh.getRange(2, colActualizar, numRows, 1);

  // 1) SKUs marcados para construir índice 1 sola vez
  const neededSkus = [];
  for (let r = 0; r < numRows; r++) {
    if (!!actualizarVals[r][0]) {
      const skuInSheet = String(values[r][colSku - 1] || '').trim();
      if (skuInSheet && skuInSheet.indexOf('-') !== -1) neededSkus.push(skuInSheet);
    }
  }

  const parentIdx = paris_buildParentIdIndexFromSearch_(neededSkus, {
    limit: 100,
    maxPages: 200,
    stopWhenAllFound: true
  });

  // pacing anti-spikes
  const pace = { sleepMs: 150, slowMs: 6000, maxSleepMs: 600 };

  // Circuit breaker / time budget
  const tStart = Date.now();
  const BUDGET_MS = 320000; // ~5m20s para evitar "Exceeded maximum execution time"
  let consecTimeouts = 0;
  const MAX_CONSEC_TIMEOUTS = 5;

  let updated = 0;
  let failed = 0;
  let needsCacheClear = false;

  // Acumuladores indicadores
  const touched = [];
  const okMask = [];
  const errMsgs = [];

  for (let r = 0; r < numRows; r++) {
    const doUpdate = !!actualizarVals[r][0];
    if (!doUpdate) continue;

    try {
      const row = values[r];
      const skuInSheet = String(row[colSku - 1] || '').trim();
      const precioLista = toNumber_(row[colPrecio - 1]);

      if (!skuInSheet) throw new Error('SKU(*) vacío');
      if (precioLista === null) throw new Error('Precio(*) inválido');

      const parentId = paris_resolveParentId_(skuInSheet, parentIdx);

      const payload = {
        prices: [
          { value: precioLista, storePrice: PARIS_STORE_PRICE_ID, type: PARIS_PRICE_TYPE_LISTA }
        ]
      };

      const path = PARIS_ENDPOINTS.PRICES_UPSERT_V2 + '/' + encodeURIComponent(parentId);

      // ✅ upsert con pacing + retry inteligente para ETIMEDOUT
      paris_priceUpsertWithTimeoutRetry_(path, payload, pace);

      // ✅ OK: desmarcar
      actualizarVals[r][0] = false;
      updated++;
      needsCacheClear = true;

      consecTimeouts = 0;

      touched.push(r); okMask.push(true); errMsgs.push('');

    } catch (e) {
      failed++;
      Logger.log('❌ LISTA fila ' + (r + 2) + ' SKU=' + (values[r][colSku - 1] || '') + ' -> ' + e.message);

      if (paris_isTimeoutError_(e)) consecTimeouts++;
      else consecTimeouts = 0;

      // Mantener checkbox TRUE para reintentar
      touched.push(r); okMask.push(false); errMsgs.push(e.message);
    }

    // Stop elegante por tiempo o por backend inestable
    if (paris_shouldStopByTime_(tStart, BUDGET_MS)) {
      Logger.log('⏹️ Stop: presupuesto de tiempo alcanzado. Quedan filas marcadas para reintento.');
      break;
    }
    if (consecTimeouts >= MAX_CONSEC_TIMEOUTS) {
      Logger.log('⏹️ Stop: demasiados ETIMEDOUT seguidos. Backend inestable. Reintenta más tarde.');
      break;
    }
  }

  // 1 sola escritura de checkboxes (aplica también los TRUE que quedaron)
  actualizarRange.setValues(actualizarVals);

  // Indicadores en columna Actualizar: color + nota
  paris_applyIndicatorsOnColA_(sh, colActualizar, bgRange, noteRange, touched, okMask, errMsgs);

  // Invalida cache de consultas (1 sola vez)
  if (needsCacheClear) {
    try { paris_clearCache(); } catch (e) {}
  }

  ss.toast('Lista: OK=' + updated + ' | FAIL=' + failed, 'Paris', 5);
  Logger.log('Precio lista actualizado: OK=' + updated + ' FAIL=' + failed);
  paris_timingReport_();
}


// =========================
//  2) ACTUALIZAR PRECIOS OFERTA (lista + oferta + fechas)
//  - Siempre envía lista (Precio(*)) + oferta (Precio oferta) con fechas
//  - Hoja: "Copia de Precios. PM"
//  - Respeta: NO tocar L:V (solo lee A:K)
//  - Control: columna A "Actualizar" (TRUE/FALSE)
// =========================
function paris_updatePriceOffer_fromSheet() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_PRECIOS);
  if (!sh) throw new Error('No existe hoja: ' + SHEET_PRECIOS);

  const hm = sheet_getHeaderMap_(sh);

  const colActualizar = hm['Actualizar'];
  const colSku = hm['SKU(*)'];
  const colPrecio = hm['Precio(*)'];
  const colOferta = hm['Precio oferta'];
  const colDesde = hm['Fecha desde'];
  const colHasta = hm['Fecha hasta'];

  if (!colActualizar || !colSku || !colPrecio || !colOferta || !colDesde || !colHasta) {
    throw new Error('Headers faltantes en ' + SHEET_PRECIOS +
      '. Revisa: Actualizar, SKU(*), Precio(*), Precio oferta, Fecha desde, Fecha hasta');
  }

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  const numRows = lastRow - 1;

  const readCols = Math.min(11, sh.getLastColumn()); // A:K
  const values = sh.getRange(2, 1, numRows, readCols).getValues();

  const actualizarRange = sh.getRange(2, colActualizar, numRows, 1);
  const actualizarVals = actualizarRange.getValues();

  // Indicadores (misma columna "Actualizar")
  const bgRange = sh.getRange(2, colActualizar, numRows, 1);
  const noteRange = sh.getRange(2, colActualizar, numRows, 1);

  // 1) SKUs marcados para índice
  const neededSkus = [];
  for (let r = 0; r < numRows; r++) {
    if (!!actualizarVals[r][0]) {
      const skuInSheet = String(values[r][colSku - 1] || '').trim();
      if (skuInSheet && skuInSheet.indexOf('-') !== -1) neededSkus.push(skuInSheet);
    }
  }

  const parentIdx = paris_buildParentIdIndexFromSearch_(neededSkus, {
    limit: 100,
    maxPages: 200,
    stopWhenAllFound: true
  });

  const pace = { sleepMs: 150, slowMs: 6000, maxSleepMs: 600 };

  // Circuit breaker / time budget
  const tStart = Date.now();
  const BUDGET_MS = 320000;
  let consecTimeouts = 0;
  const MAX_CONSEC_TIMEOUTS = 5;

  let updated = 0;
  let failed = 0;
  let needsCacheClear = false;

  const touched = [];
  const okMask = [];
  const errMsgs = [];

  for (let r = 0; r < numRows; r++) {
    const doUpdate = !!actualizarVals[r][0];
    if (!doUpdate) continue;

    try {
      const row = values[r];

      const skuInSheet = String(row[colSku - 1] || '').trim();
      const precioLista = toNumber_(row[colPrecio - 1]);
      const precioOferta = toNumber_(row[colOferta - 1]);
      const desdeIso = toIsoZ_(row[colDesde - 1]);
      const hastaIso = toIsoZ_(row[colHasta - 1]);

      if (!skuInSheet) throw new Error('SKU(*) vacío');
      if (precioLista === null) throw new Error('Precio(*) inválido');
      if (precioOferta === null) throw new Error('Precio oferta inválido');
      if (!desdeIso || !hastaIso) throw new Error('Fechas inválidas (desde/hasta)');

      const parentId = paris_resolveParentId_(skuInSheet, parentIdx);

      const payload = {
        prices: [
          { value: precioLista, storePrice: PARIS_STORE_PRICE_ID, type: PARIS_PRICE_TYPE_LISTA },
          { value: precioOferta, storePrice: PARIS_STORE_PRICE_ID, type: PARIS_PRICE_TYPE_OFERTA, showFrom: desdeIso, showTo: hastaIso }
        ]
      };

      const path = PARIS_ENDPOINTS.PRICES_UPSERT_V2 + '/' + encodeURIComponent(parentId);

      paris_priceUpsertWithTimeoutRetry_(path, payload, pace);

      actualizarVals[r][0] = false;
      updated++;
      needsCacheClear = true;

      consecTimeouts = 0;

      touched.push(r); okMask.push(true); errMsgs.push('');

    } catch (e) {
      failed++;
      Logger.log('❌ OFERTA fila ' + (r + 2) + ' SKU=' + (values[r][colSku - 1] || '') + ' -> ' + e.message);

      if (paris_isTimeoutError_(e)) consecTimeouts++;
      else consecTimeouts = 0;

      touched.push(r); okMask.push(false); errMsgs.push(e.message);
    }

    if (paris_shouldStopByTime_(tStart, BUDGET_MS)) {
      Logger.log('⏹️ Stop: presupuesto de tiempo alcanzado. Quedan filas marcadas para reintento.');
      break;
    }
    if (consecTimeouts >= MAX_CONSEC_TIMEOUTS) {
      Logger.log('⏹️ Stop: demasiados ETIMEDOUT seguidos. Backend inestable. Reintenta más tarde.');
      break;
    }
  }

  actualizarRange.setValues(actualizarVals);

  paris_applyIndicatorsOnColA_(sh, colActualizar, bgRange, noteRange, touched, okMask, errMsgs);

  if (needsCacheClear) {
    try { paris_clearCache(); } catch (e) {}
  }

  ss.toast('Oferta: OK=' + updated + ' | FAIL=' + failed, 'Paris', 5);
  Logger.log('Precio lista + oferta: OK=' + updated + ' FAIL=' + failed);
  paris_timingReport_();
}


// =========================
//  ACTUALIZAR STOCK (desde "Prods. PM")
//  - Usa: columna A "Actualizar" (checkbox)
//  - Input: columna "stock"
//  - Si OK: desmarca Actualizar
// =========================
function paris_updateStock_fromSheet() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_PRODS);
  if (!sh) throw new Error('No existe hoja: ' + SHEET_PRODS);

  const hm = sheet_getHeaderMap_(sh);

  const colActualizar = hm['Actualizar'];
  const colSkuMkp = hm['sku_mkp'];
  const colSkuSeller = hm['sku_seller'];
  const colStock = hm['stock'];

  if (!colActualizar || !colSkuMkp || !colSkuSeller || !colStock) {
    throw new Error('Headers faltantes en ' + SHEET_PRODS + '. Revisa: Actualizar, sku_mkp, sku_seller, stock');
  }

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  const readCols = Math.min(8, sh.getLastColumn()); // A:H, sin nuevo_stock
  const range = sh.getRange(2, 1, lastRow - 1, readCols);
  const values = range.getValues();

  const warehouse = paris_getWarehouse_();

  const items = [];
  const skusPayload = [];

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const doUpdate = !!row[colActualizar - 1];
    if (!doUpdate) continue;

    const skuMkp = row[colSkuMkp - 1];
    const skuSeller = row[colSkuSeller - 1];
    const stock = toNumber_(row[colStock - 1]);

    if (!skuMkp || !skuSeller) continue;
    if (stock === null) throw new Error('Fila ' + (i + 2) + ': stock inválido');

    Logger.log(
      '[PARIS][STOCK] SKU=' + skuMkp +
      ' SELLER=' + skuSeller +
      ' new=' + stock
    );

    items.push({idx: i});

    skusPayload.push({
      sku: String(skuMkp),
      sku_seller: String(skuSeller),
      quantity: stock,
      warehouse: warehouse
    });
  }

  if (!skusPayload.length) {
    Logger.log('No hay filas marcadas para actualizar.');
    return;
  }

  paris_postStockInChunks_(skusPayload);

  for (let k = 0; k < items.length; k++) {
    const { idx } = items[k];
    values[idx][colActualizar - 1] = false;
  }

  range.setValues(values);

  paris_timingReport_();
  ss.toast('Stock actualizado en "' + SHEET_PRODS + '": ' + items.length + ' variantes. Puede tardar algunos minutos en reflejarse para consultar.', 'Paris', 5);
}

// =========================
//  EJECUCIÓN TOTAL (si quieres “un botón”)
// =========================
function paris_runAll_updates() {
  paris_updatePriceList_fromSheet();
  paris_updateStock_fromSheet();
}

// =========================
// Ejecuta fetchProducts_mapSkus() y fetchProducts_mapSkus()
// =========================
function paris_runAll_fetchs() {
    paris_fetchPrices_toSheet();
    
}

// =========================
//  CONSULTAR PRODUCTOS (mapear SKU mkp variante + skuSeller) -> "Prods. PM"
//  - Inserta/actualiza filas por sku_mkp
//  - Mantiene stock real desde Paris en columna stock
//  - Ya no usa columna nuevo_stock
// =========================
function paris_fetchProducts_mapSkus() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_PRODS);
  if (!sh) throw new Error('No existe hoja: ' + SHEET_PRODS);

  const hm = sheet_getHeaderMap_(sh);

  const colActualizar  = hm['Actualizar'];
  const colSkuMkp      = hm['sku_mkp'];
  const colSkuSeller   = hm['sku_seller'];
  const colTitulo      = hm['titulo'];
  const colTalla       = hm['talla'];
  const colEstado      = hm['estado_de_variante_de_producto'];
  const colFulfillment = hm['fulfillment'];
  const colStock       = hm['stock'];

  const need = [
    'Actualizar',
    'sku_mkp',
    'sku_seller',
    'titulo',
    'talla',
    'estado_de_variante_de_producto',
    'fulfillment',
    'stock'
  ];

  const missing = need.filter(k => !hm[k]);
  if (missing.length) {
    throw new Error('Headers faltantes en "' + SHEET_PRODS + '": ' + missing.join(', '));
  }

  const lastRow = sh.getLastRow();
  const readCols = 8; // A:H, sin nuevo_stock

  const existing = (lastRow >= 2)
    ? sh.getRange(2, 1, lastRow - 1, readCols).getValues()
    : [];

  const keepBySku = {};

  for (let i = 0; i < existing.length; i++) {
    const skuRaw = existing[i][colSkuMkp - 1];
    if (!skuRaw) continue;

    const k = paris_normSkuMkp_(skuRaw);
    const stockVal = existing[i][colStock - 1];

    if (!keepBySku[k]) {
      keepBySku[k] = { stock: stockVal };
    } else {
      if ((keepBySku[k].stock === '' || keepBySku[k].stock == null) && stockVal !== '' && stockVal != null) {
        keepBySku[k].stock = stockVal;
      }
    }
  }

  const outRows = [];
  let offset = 0;
  const limit = 100;

  while (true) {
    const path = PARIS_ENDPOINTS.PRODUCTS_SEARCH + '?limit=' + limit + '&offset=' + offset;
    const data = paris_fetchJson_(path, { method: 'get' });

    const results = (data && data.results) ? data.results : [];
    if (!results.length) break;

    for (const prod of results) {
      const prodName = prod.name || '';
      const variants = Array.isArray(prod.variants) ? prod.variants : [];

      for (const v of variants) {
        const skuMkp = v.sku;
        if (!skuMkp) continue;

        const skuSeller = v.skuSeller;
        const fulfillment = paris_pickFulfillment_(prod, v);

        let talla = '';
        if (Array.isArray(v.attributes)) {
          const aTalla = v.attributes.find(a => String(a.name || '').toLowerCase() === 'talla');
          if (aTalla) talla = aTalla.optionName || aTalla.value || '';
        }

        const estado = v.statusApproval || v.status || '';

        const keepKey = paris_normSkuMkp_(skuMkp);
        const keep = keepBySku[keepKey] || { stock: '' };

        const row = new Array(8).fill('');
        row[colActualizar - 1]  = false;
        row[colSkuMkp - 1]      = skuMkp;
        row[colSkuSeller - 1]   = skuSeller || '';
        row[colTitulo - 1]      = prodName;
        row[colTalla - 1]       = talla;
        row[colEstado - 1]      = estado;
        row[colFulfillment - 1] = fulfillment;
        row[colStock - 1]       = keep.stock;

        outRows.push(row);
      }
    }

    offset += limit;
    if (data.total !== undefined && offset >= Number(data.total)) break;
  }

  const stockBySku = paris_fetchAllStockMap_(200);

  for (let i = 0; i < outRows.length; i++) {
    const skuMkp = outRows[i][colSkuMkp - 1];
    const k = paris_normSkuMkp_(skuMkp);
    if (k in stockBySku) outRows[i][colStock - 1] = stockBySku[k];
  }

  if (lastRow > 1) {
    sh.getRange(2, 1, lastRow - 1, readCols).clearContent();
  }

  if (outRows.length) {
    sh.getRange(2, 1, outRows.length, readCols).setValues(outRows);
    sh.getRange(2, colActualizar, outRows.length, 1).insertCheckboxes();
  }

  paris_timingReport_();
  ss.toast('Actualización de Productos en "' + SHEET_PRODS + '": ' + outRows.length + ' variantes.', 'Paris', 5);
  Logger.log('Productos/variantes escritos: ' + outRows.length);
}

// =========================
//  CONSULTAR PRECIOS -> "Copia de Precios. PM"
//  - Llena SOLO A:I (para NO tocar L:V)
//  - No marca "Actualizar"
//  - Usa precios que vienen dentro de /v2/products/search (variants[].prices[])
// =========================
function paris_fetchPrices_toSheet() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_PRECIOS);
  if (!sh) throw new Error('No existe hoja: ' + SHEET_PRECIOS);

  const hm = sheet_getHeaderMap_(sh);

  const colActualizar = hm['Actualizar'];
  const colSku        = hm['SKU(*)'];
  const colSellerSku  = hm['SELLER SKU'];
  const colNombre     = hm['NOMBRE'];
  const colTienda     = hm['Tienda(*)'];
  const colPrecio     = hm['Precio(*)'];
  const colOferta     = hm['Precio oferta'];
  const colDesde      = hm['Fecha desde'];
  const colHasta      = hm['Fecha hasta'];

  if (!colActualizar || !colSku || !colSellerSku || !colNombre || !colTienda ||
      !colPrecio || !colOferta || !colDesde || !colHasta) {
    throw new Error('Headers faltantes en ' + SHEET_PRECIOS);
  }

  // Limpia SOLO A:I (contenido). Mantengo tu lógica.
  const lastRow = sh.getLastRow();
  if (lastRow > 1) sh.getRange(2, 1, lastRow - 1, 9).clearContent();

  const tiendaValue = paris_getWarehouse_();

  const rows = [];
  const limit = 100;
  const ttlSeconds = 600;
  let offset = 0;

  while (true) {
    const page = paris_getProductsSearchPageCached_(limit, offset, ttlSeconds);
    const results = (page && page.results) ? page.results : [];
    if (!results.length) break;

    for (const p of results) {
      const { pLista, pOferta, fDesde, fHasta } = paris_extractPricesFromSlim_(p);

      const row = new Array(9).fill('');
      row[colActualizar - 1] = false;
      row[colSku - 1]        = p.id || '';
      row[colSellerSku - 1]  = p.sellerSku || '';
      row[colNombre - 1]     = p.name || '';
      row[colTienda - 1]     = tiendaValue;
      row[colPrecio - 1]     = pLista;
      row[colOferta - 1]     = pOferta;
      row[colDesde - 1]      = fDesde;
      row[colHasta - 1]      = fHasta;

      rows.push(row);
    }

    offset += limit;
    if (page.total !== undefined && offset >= Number(page.total)) break;
  }

  if (rows.length) {
    sh.getRange(2, 1, rows.length, 9).setValues(rows);

    const aRange = sh.getRange(2, colActualizar, rows.length, 1);
    aRange.insertCheckboxes();

    // ✅ Reset visual "top excellence": fondo blanco + sin notas
    aRange.setBackground(null); // o '#ffffff'
    aRange.setNote('');         // limpia notas (si quedaron de runs anteriores)
  }

  paris_timingReport_();
  ss.toast('Consultar precios: ' + rows.length + ' productos cargados (A:I).', 'Paris', 5);
}

function paris_fetchWithRetryFast_(url, options) {
  const maxRetries = 2;
  let delayMs = 200;

  for (let attempt = 0; attempt <= maxRetries; attempt++) {
    const resp = UrlFetchApp.fetch(url, options);
    const code = resp.getResponseCode();

    if (code !== 429 && code !== 502 && code !== 503) {
      return { resp, retries: attempt };
    }

    Utilities.sleep(delayMs + Math.floor(Math.random() * 150));
    delayMs = Math.min(delayMs * 2, 1200);
  }

  const resp = UrlFetchApp.fetch(url, options);
  return { resp, retries: maxRetries + 1 };
}

function paris_fetchWithRetry_(url, options) {
  const maxRetries = 6;
  let delayMs = 800;

  for (let attempt = 0; attempt <= maxRetries; attempt++) {
    const resp = UrlFetchApp.fetch(url, options);
    const code = resp.getResponseCode();

    if (code !== 429 && code !== 502 && code !== 503) {
      return { resp, retries: attempt };
    }

    Utilities.sleep(delayMs + Math.floor(Math.random() * 250));
    delayMs = Math.min(delayMs * 2, 8000);
  }

  // último intento “extra” como estaba antes
  const resp = UrlFetchApp.fetch(url, options);
  return { resp, retries: maxRetries + 1 };
}

function paris_cacheKey_productsPage_(limit, offset) {
  return 'PARIS_PRODS_V2_SEARCH_LIMIT_' + limit + '_OFFSET_' + offset;
}

function paris_getProductsSearchPageCached_(limit, offset, ttlSeconds) {
  const cache = CacheService.getScriptCache();
  const key = paris_cacheKey_productsPage_(limit, offset);

  const cached = cache.get(key);
  if (cached) return JSON.parse(cached);

  const path = PARIS_ENDPOINTS.PRODUCTS_SEARCH + '?limit=' + limit + '&offset=' + offset;
  const data = paris_fetchJson_(path, { method: 'get' });

  // --- SLIM: guardar SOLO lo necesario para precios ---
  const results = (data && data.results) ? data.results : [];
  const slimResults = results.map(p => {
    const variants = Array.isArray(p.variants) ? p.variants : [];
    // Solo necesitamos precios lista/oferta con fechas. Tomamos precios de la primera variante que tenga prices.
    let pricesSlim = [];
    for (const v of variants) {
      const pr = Array.isArray(v.prices) ? v.prices : [];
      if (!pr.length) continue;

      const lista = pr.find(x => x.type && x.type.id === PARIS_PRICE_TYPE_LISTA);
      const oferta = pr.find(x => x.type && x.type.id === PARIS_PRICE_TYPE_OFERTA);

      if (lista) {
        pricesSlim.push({ typeId: PARIS_PRICE_TYPE_LISTA, value: lista.value });
      }
      if (oferta) {
        pricesSlim.push({
          typeId: PARIS_PRICE_TYPE_OFERTA,
          value: oferta.value,
          showFrom: oferta.showFrom || '',
          showTo: oferta.showTo || ''
        });
      }
      break; // ya tenemos precios desde una variante
    }

    return {
      id: p.id || '',
      name: p.name || '',
      sellerSku: p.sellerSku || '',
      pricesSlim: pricesSlim
    };
  });

  const slimData = {
    total: data.total,
    offset: data.offset,
    limit: data.limit,
    results: slimResults
  };

  // Cache “safe”: si es muy grande, simplemente no cacheamos.
  try {
    cache.put(key, JSON.stringify(slimData), ttlSeconds);
  } catch (e) {
    // No rompemos la ejecución por cache
    Logger.log('Cache omitido (value demasiado grande) para ' + key + ': ' + e);
  }

  return slimData;
}

function paris_clearCache() {
  const cache = CacheService.getScriptCache();

  const limit = 100;        // el mismo que usas en prices
  const maxPages = 60;      // ajusta si necesitas más (60 páginas = 6000 items)
  const keys = [];

  for (let i = 0; i < maxPages; i++) {
    const offset = i * limit;
    keys.push(paris_cacheKey_productsPage_(limit, offset));
  }

  cache.removeAll(keys);

  SpreadsheetApp.getActive().toast('Cache Paris limpiado (hasta ' + (maxPages * limit) + ' items).', 'Paris', 5);
}

function paris_extractPricesFromProduct_(prod) {
  let pLista = '';
  let pOferta = '';
  let fDesde = '';
  let fHasta = '';

  const variants = Array.isArray(prod && prod.variants) ? prod.variants : [];

  for (let i = 0; i < variants.length; i++) {
    const v = variants[i];
    const prices = Array.isArray(v && v.prices) ? v.prices : [];
    if (!prices.length) continue;

    // Recorremos UNA vez prices[]
    for (let j = 0; j < prices.length; j++) {
      const p = prices[j];
      const sp = p && p.storePrice;
      if (!sp || sp.code !== 'paris') continue;

      const t = p.type;
      const typeId = t && t.id;
      if (!typeId) continue;

      if (typeId === PARIS_PRICE_TYPE_LISTA) {
        if (p.value != null) pLista = p.value;
      } else if (typeId === PARIS_PRICE_TYPE_OFERTA) {
        if (p.value != null) {
          pOferta = p.value;
          fDesde = p.showFrom || '';
          fHasta = p.showTo || '';
        }
      }

      // Si ya tenemos algo, no sigas escaneando
      if (pLista !== '' && pOferta !== '') break;
    }

    // Si ya obtuvimos algo, no sigas con más variants
    if (pLista !== '' && pOferta !== '') break;
  }

  return { pLista, pOferta, fDesde, fHasta };
}

function paris_findProductBySellerSku_(sellerSku) {
  sellerSku = String(sellerSku || '').trim();
  if (!sellerSku) throw new Error('sellerSku vacío');

  const limit = 100;
  const ttlSeconds = 600; // 10 min
  let offset = 0;

  while (true) {
    const data = paris_getProductsSearchPageCached_(limit, offset, ttlSeconds); // ya existe 
    const results = (data && data.results) ? data.results : [];
    if (!results.length) break;

    // En tu SLIM, sellerSku está a nivel padre (p.sellerSku) 
    for (const p of results) {
      if (String(p.sellerSku || '').trim() === sellerSku) return p;
    }

    offset += limit;
    if (data.total !== undefined && offset >= Number(data.total)) break;
  }

  throw new Error('No se encontró producto para SELLER SKU=' + sellerSku);
}

function paris_extractPricesFromSlim_(p) {
  let pLista = '';
  let pOferta = '';
  let fDesde = '';
  let fHasta = '';

  const arr = Array.isArray(p && p.pricesSlim) ? p.pricesSlim : [];
  for (const it of arr) {
    if (it.typeId === PARIS_PRICE_TYPE_LISTA && it.value != null) pLista = it.value;
    if (it.typeId === PARIS_PRICE_TYPE_OFERTA && it.value != null) {
      pOferta = it.value;
      fDesde = it.showFrom || '';
      fHasta = it.showTo || '';
    }
  }
  return { pLista, pOferta, fDesde, fHasta };
}


function paris_getParentIdCached_(sku, localMap) {
  sku = String(sku || '').trim();
  if (!sku) throw new Error('SKU vacío');

  // 1) Caché en memoria (por ejecución)
  if (localMap && localMap.has(sku)) return localMap.get(sku);

  // 2) Caché persistente (entre ejecuciones)
  const cache = CacheService.getScriptCache();
  const key = 'PARIS_PARENTID_' + sku;

  const cached = cache.get(key);
  if (cached) {
    if (localMap) localMap.set(sku, cached);
    return cached;
  }

  // 3) Fuente de verdad
  const parentId = paris_findParentIdBySku_(sku);
  if (!parentId) throw new Error('No se encontró parentId para SKU=' + sku);

  const pid = String(parentId);

  // Guardar en CacheService (6 horas)
  cache.put(key, pid, 21600);

  // Guardar también en Map local
  if (localMap) localMap.set(sku, pid);

  // 4) Registrar key para poder limpiarla luego desde menú
  const props = PropertiesService.getScriptProperties();
  let keys;
  try {
    keys = JSON.parse(props.getProperty('PARIS_PARENTID_KEYS') || '[]');
    if (!Array.isArray(keys)) keys = [];
  } catch (e) {
    keys = [];
  }

  if (keys.indexOf(key) === -1) {
    keys.push(key);

    // Evitar crecimiento infinito del índice (máx 5000)
    if (keys.length > 5000) {
      keys = keys.slice(keys.length - 5000);
    }

    props.setProperty('PARIS_PARENTID_KEYS', JSON.stringify(keys));
  }

  return pid;
}

function paris_clearParentIdCache_() {
  const cache = CacheService.getScriptCache();

  // Keys típicas que usamos: PARIS_PARENTID_<SKU>
  // CacheService NO permite listar keys, así que guardamos un índice en PropertiesService.

  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty('PARIS_PARENTID_KEYS') || '[]';

  let keys;
  try {
    keys = JSON.parse(raw);
    if (!Array.isArray(keys)) keys = [];
  } catch (e) {
    keys = [];
  }

  if (!keys.length) {
    SpreadsheetApp.getActive().toast('No hay caché PARIS_PARENTID_ para limpiar.', 'Paris', 5);
    return;
  }

  // Borrar del cache
  cache.removeAll(keys);

  // Limpiar índice
  props.deleteProperty('PARIS_PARENTID_KEYS');

  SpreadsheetApp.getActive().toast('Caché PARIS_PARENTID_ limpiado: ' + keys.length + ' keys.', 'Paris', 5);
}

function paris_buildParentIdIndexFromSearch_(neededSkus, opts) {
  opts = opts || {};
  const limit = opts.limit || 100;
  const maxPages = opts.maxPages || 200; // 200*100 = 20.000
  const stopWhenAllFound = opts.stopWhenAllFound !== false; // default true

  const wanted = new Set(
    (neededSkus || [])
      .map(s => String(s || '').trim())
      .filter(Boolean)
  );

  const map = new Map(); // sku(any) -> parentId
  let offset = 0;

  for (let page = 0; page < maxPages; page++) {
    const data = paris_fetchJson_(
      PARIS_ENDPOINTS.PRODUCTS_SEARCH + '?limit=' + limit + '&offset=' + offset,
      { method: 'get' }
    );

    const results = Array.isArray(data && data.results) ? data.results : [];

    for (let i = 0; i < results.length; i++) {
      const prod = results[i] || {};

      // parentId "real" (ajusta si tu API usa otra propiedad)
      const parentId = (prod.id != null ? String(prod.id) :
                       (prod.parentId != null ? String(prod.parentId) :
                       (prod.sku != null ? String(prod.sku) : ''))).trim();
      if (!parentId) continue;

      // 1) skuSeller del producto
      const pSkuSeller = String(prod.skuSeller || '').trim();
      if (pSkuSeller) {
        if (!map.has(pSkuSeller)) map.set(pSkuSeller, parentId);
        if (wanted.has(pSkuSeller)) wanted.delete(pSkuSeller);
      }

      // 2) variantes: variantSkuSeller (y opcional skuSeller dentro de variant)
      const variants = Array.isArray(prod.variants) ? prod.variants : [];
      for (let j = 0; j < variants.length; j++) {
        const v = variants[j] || {};
        const vSku = String(v.variantSkuSeller || v.skuSeller || '').trim();
        if (!vSku) continue;

        if (!map.has(vSku)) map.set(vSku, parentId);
        if (wanted.has(vSku)) wanted.delete(vSku);
      }
    }

    if (stopWhenAllFound && wanted.size === 0) break;

    offset += limit;

    // Si viene total, corta al final
    if (data && data.total !== undefined && offset >= Number(data.total)) break;

    // Si no hay resultados, no sigas paginando
    if (!results.length) break;
  }

  return map;
}

// Resolver rápido usando el índice (fallback a tu función actual si falta)
function paris_resolveParentId_(skuMarketplaceOrParentId, indexMap) {
  const s = String(skuMarketplaceOrParentId || '').trim();
  if (!s) throw new Error('SKU vacío');

  // Si no tiene '-', normalmente ya es ID padre.
  if (s.indexOf('-') === -1) return s;

  if (indexMap && indexMap.has(s)) return indexMap.get(s);

  // fallback (por si algo no estaba en el índice)
  return paris_findParentIdBySku_(s);
}



function paris_postStockInChunks_(skusPayload) {
  Logger.log("Llegamos paris_postStockInChunks_ con: "+skusPayload);
  let chunkSize = 80;
  const SLOW_MS = 6000; // umbral para considerar request "lenta"

  for (let start = 0; start < skusPayload.length; ) {
    const chunk = skusPayload.slice(start, start + chunkSize);

    try {
      const t0 = Date.now();

      paris_fetchJson_(PARIS_ENDPOINTS.STOCK_SKU_SELLER, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({ skus: chunk })
      });

      const ms = Date.now() - t0;

      // ⏱️ Si fue demasiado lento, bajar tamaño para las siguientes
      if (ms > SLOW_MS && chunkSize > 20) {
        chunkSize = Math.max(20, Math.floor(chunkSize / 2));
        Logger.log('Request lenta (' + ms + 'ms). Bajando chunkSize a ' + chunkSize);
      }

      // 💤 Micro pausa para no saturar el backend
      Utilities.sleep(150);

      // avanzar exactamente lo enviado
      start += chunk.length;

    } catch (e) {
      if (chunkSize <= 10) throw e;
      chunkSize = Math.max(10, Math.floor(chunkSize / 2));
      Logger.log('Stock chunk falló, bajando chunkSize a ' + chunkSize + ' -> ' + e.message);
    }
  }
}

function paris_priceUpsertWithPacing_(path, payloadObj, state) {
  state = state || {};
  if (state.sleepMs == null) state.sleepMs = 150;   // base //spikes grandes, sube base a 250 
  if (state.slowMs == null) state.slowMs = 6000;    // umbral “spike”
  if (state.maxSleepMs == null) state.maxSleepMs = 600;

  const t0 = Date.now();

  paris_fetchJson_(path, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payloadObj)
  });

  const ms = Date.now() - t0;

  // Si fue lento, aumentar pausa para suavizar backend
  if (ms > state.slowMs) {
    state.sleepMs = Math.min(state.maxSleepMs, Math.max(state.sleepMs, Math.floor(state.sleepMs * 2)));
    Logger.log('Upsert lento (' + ms + 'ms). Aumentando sleepMs a ' + state.sleepMs);
  }

  // micro pausa siempre
  Utilities.sleep(state.sleepMs);

  return ms;
}

function paris_applyIndicatorsOnColA_(sh, colA, bgRange, noteRange, touchedRows0, okMask, errMsgs) {
  // touchedRows0: índices 0-based dentro del rango (fila 2 => 0)
  if (!touchedRows0 || !touchedRows0.length) return;

  const bgs = bgRange.getBackgrounds(); // [[color],...]
  const notes = noteRange.getNotes();   // [[note],...]

  for (let i = 0; i < touchedRows0.length; i++) {
    const r0 = touchedRows0[i];
    const ok = !!okMask[i];

    bgs[r0][0] = ok ? '#b7e1cd' : '#f4c7c3'; // verde / rojo
    notes[r0][0] = ok ? '' : String(errMsgs[i] || 'Error').slice(0, 4500);
  }

  bgRange.setBackgrounds(bgs);
  noteRange.setNotes(notes);
}

function paris_isTimeoutError_(e) {
  const msg = String(e && e.message ? e.message : e);
  return msg.indexOf('ETIMEDOUT') !== -1 || msg.indexOf('timed out') !== -1;
}

function paris_shouldStopByTime_(tStartMs, budgetMs) {
  return (Date.now() - tStartMs) > budgetMs;
}

function paris_priceUpsertWithTimeoutRetry_(path, payloadObj, pace) {
  const MAX_TRIES = 2;      // 1 intento + 1 reintento
  let backoffMs = 1000;

  for (let attempt = 0; attempt <= MAX_TRIES; attempt++) {
    try {
      return paris_priceUpsertWithPacing_(path, payloadObj, pace);
    } catch (e) {
      if (!paris_isTimeoutError_(e) || attempt === MAX_TRIES) throw e;

      Logger.log('⏳ ETIMEDOUT en upsert, reintentando en ' + backoffMs + 'ms -> ' + path);
      Utilities.sleep(backoffMs);
      backoffMs = Math.min(5000, backoffMs * 2);
    }
  }
}

function paris_pickFulfillment_(prod, v) {
  const p = prod || {};
  const vv = v || {};

  // 1) string directo
  if (typeof p.fulfillment === 'string' && p.fulfillment) return p.fulfillment;
  if (typeof vv.fulfillment === 'string' && vv.fulfillment) return vv.fulfillment;

  // 2) objeto {code/name/...}
  if (p.fulfillment && typeof p.fulfillment === 'object') {
    return p.fulfillment.code || p.fulfillment.name || JSON.stringify(p.fulfillment);
  }
  if (vv.fulfillment && typeof vv.fulfillment === 'object') {
    return vv.fulfillment.code || vv.fulfillment.name || JSON.stringify(vv.fulfillment);
  }

  // 3) nombres alternativos comunes
  return p.fulfillmentType || vv.fulfillmentType || '';
}

function paris_normSkuMkp_(sku) {
  const s = String(sku || '').trim().toUpperCase();
  const parts = s.split('-');
  if (parts.length === 2 && /^\d+$/.test(parts[1])) {
    return parts[0] + '-' + String(Number(parts[1])); // 01 -> 1
  }
  return s;
}

function paris_fetchStockMapBySkus_(skuList) {
  const out = {}; // normSku -> quantity
  const uniq = Array.from(new Set((skuList || []).map(s => paris_normSkuMkp_(s)).filter(Boolean)));

  for (let i = 0; i < uniq.length; i++) {
    const sku = uniq[i];

    const path =
      PARIS_ENDPOINTS.STOCK_GET_V2 +
      '?limit=1&offset=0&sku=' + encodeURIComponent(sku);

    const data = paris_fetchJson_(path, { method: 'get' });
    const skus = (data && data.skus) ? data.skus : [];

    for (const it of skus) {
      const k = paris_normSkuMkp_(it.sku);
      if (!k) continue;

      const q = (it.quantity != null) ? Number(it.quantity)
              : (it.availableStock != null) ? Number(it.availableStock)
              : null;

      if (q != null && !Number.isNaN(q)) out[k] = q;
    }
  }

  return out;
}

function paris_fetchAllStockMap_(limit) {
  const out = {}; // normSkuMkp -> quantity
  const pageSize = Math.min(Math.max(Number(limit || 200), 50), 500);
  let offset = 0;

  while (true) {
    const path = PARIS_ENDPOINTS.STOCK_GET_V2
      + '?limit=' + pageSize
      + '&offset=' + offset;

    const data = paris_fetchJson_(path, { method: 'get' });
    const skus = (data && data.skus) ? data.skus : [];
    if (!skus.length) break;

    for (const it of skus) {
      const k = paris_normSkuMkp_(it.sku);
      if (!k) continue;

      const q = (it.quantity != null) ? Number(it.quantity)
              : (it.availableStock != null) ? Number(it.availableStock)
              : null;

      if (q != null && !Number.isNaN(q)) out[k] = q;
    }

    offset += pageSize;

    // usa pagging.quantity si viene; si no, corta cuando ya no llena páginas
    const total = data && data.pagging && data.pagging.quantity != null ? Number(data.pagging.quantity) : null;
    if (total != null && offset >= total) break;
    if (skus.length < pageSize) break;
  }

  return out;
}

// ===================================================================
// VENTAS PM — Ingesta de Órdenes Paris Marketplace
// Etapa 1: Ingesta (Histórica e Incremental)
// ===================================================================

// Constantes del módulo
const PM_SHEET_VENTAS   = 'Ventas PM';
const PM_COMISION_RATE  = 0.20;
const PM_PAGE_LIMIT     = 100;
const PM_PACING_MS      = 300;
const PM_LOG_HIST       = '[PM-HIST]';
const PM_LOG_INC        = '[PM-INC]';
const PM_MAX_MS         = 5 * 60 * 1000; // 5 min — flush condicional

// Headers exactos de la hoja — 29 columnas
const PM_HEADERS = [
  'fecha_ingesta','canal','pm_order_id','pm_order_number',
  'pm_suborder_number','pm_suborder_id','id_venta_canal','fecha_venta',
  'sku_canal','sku_maestro','mpn','cantidad',
  'precio_unitario','monto_total','cargo_venta','cargo_envio',
  'monto_neto','fulfillment_type','cliente_nombre','region',
  'comuna','direccion_resumen','estado_pm','estado_conciliacion',
  'fecha_conciliacion','mensaje_conciliacion','clave_unica',
  'existe_en_ventas','json_raw'
];

// Estados de conciliación válidos (referencia)
// PENDIENTE | PENDIENTE_HISTORICO | CANCELADA POR EL COMPRADOR |
// DEVOLUCION CON REEMBOLSO | CONCILIADA | DIFERENCIA |
// DIFERENCIA_CONOCIDA | NO ENCONTRADA EN VENTAS


// ── Función pública menú: Ingesta Histórica ─────────────────────
// Ventanas de 30 días desde PARIS_HISTORICAL_FROM hasta hoy.
function pm_ingestaHistorica() {
  const props          = PropertiesService.getScriptProperties();
  const historicalFrom = String(props.getProperty('PARIS_HISTORICAL_FROM') || '').trim();

  if (!historicalFrom) {
    SpreadsheetApp.getActive().toast('Falta PARIS_HISTORICAL_FROM en Script Properties.', 'PM', 5);
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh   = ss.getSheetByName(PM_SHEET_VENTAS);
  if (!sh) sh = pm_crearHojaVentasPM_();

  const existingKeys = pm_getExistingKeys_(sh);
  const startTime    = Date.now();

  let desde          = new Date(historicalFrom);
  const hasta        = new Date();
  const VENTANA_DIAS = 30;
  let totalOrders    = 0;
  let totalInserted  = 0;
  let ventana        = 0;
  let allRows        = [];

  SpreadsheetApp.getActive().toast('Ingesta histórica PM iniciada...', 'EHI', 5);

  function flushRows_() {
    if (!allRows.length) return;
    const startRow = Math.max(sh.getLastRow() + 1, 2);
    sh.getRange(startRow, 1, allRows.length, PM_HEADERS.length).setValues(allRows);
    totalInserted  += allRows.length;
    allRows.length  = 0;
    Logger.log(PM_LOG_HIST + ' Flush: totalInserted=' + totalInserted);
  }

  while (desde < hasta) {
    const fin     = new Date(desde);
    fin.setDate(fin.getDate() + VENTANA_DIAS);
    const finReal = fin > hasta ? hasta : fin;

    const gteCreatedAt = pm_formatDate_(desde);
    const lteCreatedAt = pm_formatDate_(finReal);

    ventana++;
    Logger.log(PM_LOG_HIST + ' Ventana ' + ventana + ': ' + gteCreatedAt + ' → ' + lteCreatedAt);
    SpreadsheetApp.getActive().toast(
      'Ventana ' + ventana + ' | Órdenes acumuladas: ' + totalOrders + ' | Buffer: ' + allRows.length,
      'PM Histórico ⏳', 8
    );

    let offset         = 0;
    let shouldContinue = true;

    while (shouldContinue) {
      const data = pm_fetchOrders_(gteCreatedAt, lteCreatedAt, offset);
      const orders = (data && data.data) ? data.data : [];
      if (!orders.length) break;

      totalOrders += orders.length;

      for (var i = 0; i < orders.length; i++) {
        const rows = pm_buildRows_(orders[i], 'HISTORICO', existingKeys);
        for (var j = 0; j < rows.length; j++) {
          allRows.push(rows[j]);
        }
      }

      if (orders.length < PM_PAGE_LIMIT) {
        shouldContinue = false;
      } else {
        offset += PM_PAGE_LIMIT;
        Utilities.sleep(PM_PACING_MS);
      }
    }

    // Flush condicional si se acerca el timeout
    if (Date.now() - startTime > PM_MAX_MS) {
      Logger.log(PM_LOG_HIST + ' Timeout inminente — flush forzado en ventana ' + ventana);
      flushRows_();
    }

    desde = fin;
    Utilities.sleep(PM_PACING_MS);
  }

  flushRows_(); // flush final

  const resumen = { ventanas: ventana, fetchedOrders: totalOrders, insertedRows: totalInserted };
  Logger.log(PM_LOG_HIST + ' Resumen: ' + JSON.stringify(resumen));
  SpreadsheetApp.getActive().toast(
    'PM histórico: ventanas=' + ventana + ' | órdenes=' + totalOrders + ' | filas=' + totalInserted,
    'EHI', 10
  );
  return resumen;
}


// ── Función pública menú: Ingesta Incremental ───────────────────
// Ventanas de 7 días desde PARIS_LAST_SYNC hasta hoy.
function pm_ingestaIncremental() {
  const props    = PropertiesService.getScriptProperties();
  const lastSync = String(props.getProperty('PARIS_LAST_SYNC') || '').trim();

  if (!lastSync) throw new Error('Falta PARIS_LAST_SYNC en Script Properties.');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh   = ss.getSheetByName(PM_SHEET_VENTAS);
  if (!sh) sh = pm_crearHojaVentasPM_();

  const existingKeys = pm_getExistingKeys_(sh);
  const startTime    = Date.now();

  let desde          = new Date(lastSync);
  const hasta        = new Date();
  const VENTANA_DIAS = 7;
  let totalOrders    = 0;
  let totalInserted  = 0;
  let ventana        = 0;
  let allRows        = [];

  SpreadsheetApp.getActive().toast('Ingesta incremental PM iniciada...', 'EHI', 5);

  function flushRows_() {
    if (!allRows.length) return;
    const startRow = Math.max(sh.getLastRow() + 1, 2);
    sh.getRange(startRow, 1, allRows.length, PM_HEADERS.length).setValues(allRows);
    totalInserted  += allRows.length;
    allRows.length  = 0;
    Logger.log(PM_LOG_INC + ' Flush: totalInserted=' + totalInserted);
  }

  while (desde < hasta) {
    const fin     = new Date(desde);
    fin.setDate(fin.getDate() + VENTANA_DIAS);
    const finReal = fin > hasta ? hasta : fin;

    const gteCreatedAt = pm_formatDate_(desde);
    const lteCreatedAt = pm_formatDate_(finReal);

    ventana++;
    Logger.log(PM_LOG_INC + ' Ventana ' + ventana + ': ' + gteCreatedAt + ' → ' + lteCreatedAt);
    SpreadsheetApp.getActive().toast(
      'Ventana ' + ventana + ' | Órdenes acumuladas: ' + totalOrders + ' | Buffer: ' + allRows.length,
      'PM Incremental ⏳', 8
    );

    let offset         = 0;
    let shouldContinue = true;

    while (shouldContinue) {
      const data = pm_fetchOrders_(gteCreatedAt, lteCreatedAt, offset);
      const orders = (data && data.data) ? data.data : [];
      if (!orders.length) break;

      totalOrders += orders.length;

      for (var i = 0; i < orders.length; i++) {
        const rows = pm_buildRows_(orders[i], 'INCREMENTAL', existingKeys);
        for (var j = 0; j < rows.length; j++) {
          allRows.push(rows[j]);
        }
      }

      if (orders.length < PM_PAGE_LIMIT) {
        shouldContinue = false;
      } else {
        offset += PM_PAGE_LIMIT;
        Utilities.sleep(PM_PACING_MS);
      }
    }

    if (Date.now() - startTime > PM_MAX_MS) {
      Logger.log(PM_LOG_INC + ' Timeout inminente — flush forzado en ventana ' + ventana);
      flushRows_();
    }

    desde = fin;
    Utilities.sleep(PM_PACING_MS);
  }

  flushRows_();

  props.setProperty('PARIS_LAST_SYNC', hasta.toISOString());

  const resumen = { ventanas: ventana, fetchedOrders: totalOrders, insertedRows: totalInserted };
  Logger.log(PM_LOG_INC + ' Resumen: ' + JSON.stringify(resumen));
  SpreadsheetApp.getActive().toast(
    'PM incremental: ventanas=' + ventana + ' | órdenes=' + totalOrders + ' | filas=' + totalInserted,
    'EHI', 7
  );
  return resumen;
}


// ── Fetch órdenes con ventana de fechas ─────────────────────────
function pm_fetchOrders_(gteCreatedAt, lteCreatedAt, offset) {
  try {
    var path = '/v1/orders'
      + '?limit=' + PM_PAGE_LIMIT
      + '&offset=' + offset
      + '&gteCreatedAt=' + encodeURIComponent(gteCreatedAt)
      + '&lteCreatedAt=' + encodeURIComponent(lteCreatedAt);

    return paris_fetchJson_(path, { method: 'get' });

  } catch(e) {
    Logger.log('[PM] pm_fetchOrders_ excepción: ' + e.message);
    return { data: [] };
  }
}


// ── Construir filas desde una orden (1 fila por SKU único por subOrden) ──
function pm_buildRows_(order, modo, existingKeys) {
  const rows      = [];
  const subOrders = Array.isArray(order.subOrders) ? order.subOrders : [];

  const orderNumber = String(order.originOrderNumber || order.id || '');
  const orderId     = String(order.id || '');
  const fechaVenta  = String(order.originOrderDate || order.createdAt || '');
  const cliente     = String(order.customer && order.customer.name ? order.customer.name : '');

  for (var s = 0; s < subOrders.length; s++) {
    const sub         = subOrders[s];
    const subNumber   = String(sub.subOrderNumber || '');
    const subId       = String(sub.id || '');
    const fulfillment = String(sub.fulfillment || '');
    const statusName  = String(sub.status && sub.status.name ? sub.status.name : '');

    const addr   = sub.shippingAddress || {};
    const region = String(addr.stateCode || '');
    const comuna = String(addr.city || '');
    const direc  = String((addr.address1 || '') + (addr.address2 ? ' ' + addr.address2 : '')).trim();

    const items = Array.isArray(sub.items) ? sub.items : [];

    // ── Agrupar ítems por SKU dentro de la subOrden ──────────
    // Paris puede devolver múltiples objetos con mismo SKU = cantidad > 1
    const skuMap = {};  // skuMaestro → { item, cantidad }
    for (var it = 0; it < items.length; it++) {
      const item       = items[it];
      const skuMaestro = String(item.sellerSku || '');
      if (!skuMaestro) continue;

      if (skuMap[skuMaestro]) {
        skuMap[skuMaestro].cantidad++;
      } else {
        skuMap[skuMaestro] = { item: item, cantidad: 1 };
      }
    }

    // ── Construir una fila por SKU agrupado ──────────────────
    const skuKeys = Object.keys(skuMap);
    for (var k = 0; k < skuKeys.length; k++) {
      const skuMaestro = skuKeys[k];
      const grupo      = skuMap[skuMaestro];
      const item       = grupo.item;
      const cantidad   = grupo.cantidad;

      // Clave interna de deduplicación
      const claveInterna = orderNumber + '|' + skuMaestro;
      if (existingKeys.has(claveInterna)) continue;

      const claveUnica  = orderNumber + '-' + subNumber;
      const precioUnit  = parseFloat(item.basePrice           || 0);
      const montoTotal  = parseFloat((precioUnit * cantidad).toFixed(0));
      const cargoVenta  = parseFloat((montoTotal * PM_COMISION_RATE).toFixed(0));
      const montoNeto   = parseFloat((montoTotal - cargoVenta).toFixed(0));

      // Estado conciliación
      let estadoConciliacion = '';
      const st = statusName.toLowerCase();
      if (st === 'cancelled' || st === 'canceled') {
        estadoConciliacion = 'CANCELADA POR EL COMPRADOR';
      } else if (st === 'returned') {
        estadoConciliacion = 'DEVOLUCION CON REEMBOLSO';
      } else if (modo === 'HISTORICO') {
        estadoConciliacion = 'PENDIENTE_HISTORICO';
      } else {
        estadoConciliacion = 'PENDIENTE';
      }

      existingKeys.add(claveInterna);

      rows.push([
        new Date(),                                        // A fecha_ingesta
        'PM',                                              // B canal
        orderId,                                           // C pm_order_id
        orderNumber,                                       // D pm_order_number
        subNumber,                                         // E pm_suborder_number
        subId,                                             // F pm_suborder_id
        orderNumber + '-' + subNumber,                     // G id_venta_canal
        fechaVenta,                                        // H fecha_venta
        String(item.sku || ''),                            // I sku_canal
        skuMaestro,                                        // J sku_maestro
        skuMaestro.length > 2 ? skuMaestro.slice(0, -2) : skuMaestro, // K mpn
        cantidad,                                          // L cantidad
        precioUnit,                                        // M precio_unitario
        montoTotal,                                        // N monto_total
        cargoVenta,                                        // O cargo_venta
        0,                                                 // P cargo_envio
        montoNeto,                                         // Q monto_neto
        fulfillment,                                       // R fulfillment_type
        cliente,                                           // S cliente_nombre
        region,                                            // T region
        comuna,                                            // U comuna
        direc,                                             // V direccion_resumen
        statusName,                                        // W estado_pm
        estadoConciliacion,                                // X estado_conciliacion
        '',                                                // Y fecha_conciliacion
        '',                                                // Z mensaje_conciliacion
        claveUnica,                                        // AA clave_unica
        false,                                             // AB existe_en_ventas
        JSON.stringify({ order_id: orderId, sub: sub, item: item, cantidad: cantidad }) // AC json_raw
      ]);
    }
  }
  return rows;
}


// ── Helper: claves únicas existentes ───────────────────────────
// Usa clave interna orderNumber|skuMaestro para deduplicación
function pm_getExistingKeys_(sh) {
  const set     = new Set();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return set;

  // Lee cols D (pm_order_number=4) y J (sku_maestro=10)
  const orderNums  = sh.getRange(2, 4, lastRow - 1, 1).getValues();
  const skuMaestros = sh.getRange(2, 10, lastRow - 1, 1).getValues();

  for (var i = 0; i < orderNums.length; i++) {
    const on  = String(orderNums[i][0]  || '').trim();
    const sku = String(skuMaestros[i][0] || '').trim();
    if (on && sku) set.add(on + '|' + sku);
  }
  return set;
}


// ── Crear hoja Ventas PM si no existe ──────────────────────────
function pm_crearHojaVentasPM_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh   = ss.getSheetByName(PM_SHEET_VENTAS);
  if (sh) return sh;

  sh = ss.insertSheet(PM_SHEET_VENTAS);
  sh.getRange(1, 1, 1, PM_HEADERS.length).setValues([PM_HEADERS]);
  sh.setFrozenRows(1);

  // Formato fecha cols A(1) y H(8)
  sh.getRange(2, 1, sh.getMaxRows() - 1, 1).setNumberFormat('dd/mm/yyyy hh:mm:ss');
  sh.getRange(2, 8, sh.getMaxRows() - 1, 1).setNumberFormat('dd/mm/yyyy hh:mm:ss');

  // Formato número cols M(13),N(14),O(15),R(18)
  [13, 14, 15, 17].forEach(function(c) {
    sh.getRange(2, c, sh.getMaxRows() - 1, 1).setNumberFormat('#,##0');
  });

  return sh;
}


// ── Normaliza fecha a formato YYYY-MM-DD para API Paris ────────
function pm_formatDate_(date) {
  const d = new Date(date);
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return y + '-' + m + '-' + day;
}

// ===================================================================
// VENTAS PM — Etapa 2: Conciliación
// Cruza hoja "Ventas PM" contra hoja "Ventas" (fuente de verdad).
// ===================================================================

function pm_conciliar() {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const shPm     = ss.getSheetByName(PM_SHEET_VENTAS);
  const shVentas = ss.getSheetByName('Ventas');

  if (!shPm)     throw new Error('No existe hoja "Ventas PM".');
  if (!shVentas) throw new Error('No existe hoja "Ventas".');

  SpreadsheetApp.getActive().toast('Conciliación PM iniciada...', 'EHI', 5);

  // ── 1. Construir mapas hoja Ventas ────────────────────────────
  const hVentas     = shVentas.getDataRange().getValues();
  const hVentasHead = hVentas[0].map(function(h) {
    return String(h).replace(/\n/g, '').trim();
  });

  const colNumVenta   = hVentasHead.indexOf('# Venta');
  const colSku        = hVentasHead.indexOf('SKU');
  const colPrecioV    = hVentasHead.indexOf('PrecioVenta');
  const colCargosV    = hVentasHead.indexOf('Cargos porVenta');
  const colCargoEnvio = hVentasHead.indexOf('Cargo porEnvío');
  const colRecaudado  = hVentasHead.indexOf('Recaudado');

  if ([colNumVenta, colSku, colPrecioV, colCargosV, colCargoEnvio, colRecaudado].includes(-1)) {
    throw new Error('Faltan columnas requeridas en hoja Ventas. Verifica: # Venta, SKU, PrecioVenta, Cargos porVenta, Cargo porEnvío, Recaudado.');
  }

  const mapaVentasSimple    = {};  // # Venta → datos
  const mapaVentasCompuesto = {};  // # Venta|SKU → datos

  for (var r = 1; r < hVentas.length; r++) {
    const numVenta = String(hVentas[r][colNumVenta]).trim();
    if (!numVenta) continue;

    const precioVenta = pm_parseNumber_(hVentas[r][colPrecioV]);
    if (precioVenta < 0) continue;  // notas de crédito — se omiten

    const sku   = String(hVentas[r][colSku]).trim();
    const datos = {
      precioVenta: precioVenta,
      cargosVenta: pm_parseNumber_(hVentas[r][colCargosV]),
      cargoEnvio:  pm_parseNumber_(hVentas[r][colCargoEnvio]),
      recaudado:   pm_parseNumber_(hVentas[r][colRecaudado])
    };

    mapaVentasSimple[numVenta]                = datos;
    mapaVentasCompuesto[numVenta + '|' + sku] = datos;
  }

  // ── 2. Leer hoja Ventas PM ────────────────────────────────────
  const hPm   = shPm.getDataRange().getValues();
  const mapPm = {};
  hPm[0].forEach(function(h, i) { if (h) mapPm[String(h).trim()] = i; });

  const COL_ORDER_NUMBER = mapPm['pm_suborder_number'];
  const COL_SKU_MAESTRO  = mapPm['sku_maestro'];
  const COL_MONTO_TOTAL  = mapPm['monto_total'];
  const COL_CARGO_VENTA  = mapPm['cargo_venta'];
  const COL_CARGO_ENVIO  = mapPm['cargo_envio'];
  const COL_MONTO_NETO   = mapPm['monto_neto'];
  const COL_ESTADO_CONC  = mapPm['estado_conciliacion'];
  const COL_FECHA_CONC   = mapPm['fecha_conciliacion'];
  const COL_MENSAJE_CONC = mapPm['mensaje_conciliacion'];

  const ESTADOS_OMITIR = new Set([
    'CANCELADA POR EL COMPRADOR', 'CONCILIADA',
    'DIFERENCIA', 'DIFERENCIA_CONOCIDA', 'NO ENCONTRADA EN VENTAS'
  ]);

  const TOLERANCIA = 1;
  let procesadas           = 0;
  let conciliadas          = 0;
  let diferencias          = 0;
  let diferenciasConocidas = 0;
  let noEncontradas        = 0;
  const ahora              = new Date();

  // ── 3. Procesar cada fila ─────────────────────────────────────
  for (var i = 1; i < hPm.length; i++) {
    const estadoConc = String(hPm[i][COL_ESTADO_CONC]).trim();
    if (ESTADOS_OMITIR.has(estadoConc)) continue;

    procesadas++;
    const orderNumber = String(hPm[i][COL_ORDER_NUMBER]).trim();
    const skuMaestro  = String(hPm[i][COL_SKU_MAESTRO]).trim();

    // Búsqueda 2 niveles: compuesto → simple
    const claveComp = orderNumber + '|' + skuMaestro;
    const ventaRow  = mapaVentasCompuesto[claveComp] || mapaVentasSimple[orderNumber];

    // ── No encontrada ─────────────────────────────────────────
    if (!ventaRow) {
      shPm.getRange(i + 1, COL_ESTADO_CONC  + 1).setValue('NO ENCONTRADA EN VENTAS');
      shPm.getRange(i + 1, COL_FECHA_CONC   + 1).setValue(ahora);
      shPm.getRange(i + 1, COL_MENSAJE_CONC + 1).setValue('# Sub-orden ' + orderNumber + ' no existe en hoja Ventas');
      noEncontradas++;
      continue;
    }

    // ── Comparar 4 campos con tolerancia ±$1 ─────────────────
    const montoTotal = pm_parseNumber_(hPm[i][COL_MONTO_TOTAL]);
    const cargoVenta = pm_parseNumber_(hPm[i][COL_CARGO_VENTA]);
    const cargoEnvio = pm_parseNumber_(hPm[i][COL_CARGO_ENVIO]);
    const montoNeto  = pm_parseNumber_(hPm[i][COL_MONTO_NETO]);

    const diffs = [];
    if (Math.abs(montoTotal - ventaRow.precioVenta) > TOLERANCIA)
      diffs.push('monto_total: '  + montoTotal + ' vs ' + ventaRow.precioVenta);
    if (Math.abs(cargoVenta - ventaRow.cargosVenta) > TOLERANCIA)
      diffs.push('cargo_venta: '  + cargoVenta + ' vs ' + ventaRow.cargosVenta);
    if (Math.abs(cargoEnvio - ventaRow.cargoEnvio)  > TOLERANCIA)
      diffs.push('cargo_envio: '  + cargoEnvio + ' vs ' + ventaRow.cargoEnvio);
    if (Math.abs(montoNeto  - ventaRow.recaudado)   > TOLERANCIA)
      diffs.push('monto_neto: '   + montoNeto  + ' vs ' + ventaRow.recaudado);

    let nuevoEstado, mensaje;

    if (!diffs.length) {
      nuevoEstado = 'CONCILIADA';
      mensaje     = 'OK';
      conciliadas++;
    } else {
      // Solo monto_neto difiere → cargo envío seller no disponible en API Paris
      const soloNeto = diffs.every(function(d) { return d.indexOf('monto_neto') === 0; });
      if (soloNeto) {
        nuevoEstado = 'DIFERENCIA_CONOCIDA';
        mensaje     = 'Posible cargo logístico PM (no disponible en API). ' + diffs.join(' | ');
        diferenciasConocidas++;
      } else {
        nuevoEstado = 'DIFERENCIA';
        mensaje     = diffs.join(' | ');
        diferencias++;
      }
    }

    shPm.getRange(i + 1, COL_ESTADO_CONC  + 1).setValue(nuevoEstado);
    shPm.getRange(i + 1, COL_FECHA_CONC   + 1).setValue(ahora);
    shPm.getRange(i + 1, COL_MENSAJE_CONC + 1).setValue(mensaje);
  }

  const resumen = {
    procesadas: procesadas, conciliadas: conciliadas,
    diferencias: diferencias, diferenciasConocidas: diferenciasConocidas,
    noEncontradas: noEncontradas
  };
  Logger.log('[PM-CONC] Resumen: ' + JSON.stringify(resumen));
  SpreadsheetApp.getActive().toast(
    'PM conciliación: ' + procesadas + ' procesadas | ' +
    conciliadas + ' ok | ' + diferencias + ' diff | ' +
    diferenciasConocidas + ' diff_conocida | ' + noEncontradas + ' no encontradas',
    'EHI', 10
  );
  return resumen;
}


// ── Helper: parsea número formato chileno o número directo ──────
function pm_parseNumber_(val) {
  if (val === null || val === undefined || val === '') return 0;
  if (typeof val === 'number') return Math.round(val);
  return parseInt(String(val).replace(/\./g, '').replace(/,/g, ''), 10) || 0;
}