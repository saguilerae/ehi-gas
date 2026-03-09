// ===================================================================
// actualizar.gs
// Ajustado EHI - sincronización multicanal robusta
// ===================================================================


// ===================================================================
// Config de canales EHI
// ===================================================================
const EHI_CHANNELS = {
  T:  'Tienda Presencial',
  ML: 'Mercado Libre',
  FS: 'Falabella Seller Center',
  SH: 'Shopify',
  MR: 'Marketplace Ripley',
  PM: 'Marketplace Paris'
};


// ===================================================================
// Función: actualiza todos los stock
// Regla:
// - actualiza todos los canales EXCEPTO el canal origen
// - si origen = T, empuja a todos los canales online
// - si un SKU no existe en un marketplace, ese marketplace se omite
// - no se detiene el proceso completo por error en un canal
// ===================================================================
function actualizarTodosStock(canal, sku, valor) {
  const canalOrigen = String(canal || '').trim().toUpperCase();
  const skuNorm = String(sku || '').trim();
  const stockNuevo = Number(valor);

  if (!skuNorm) {
    throw new Error('actualizarTodosStock: SKU vacío.');
  }

  if (isNaN(stockNuevo)) {
    throw new Error('actualizarTodosStock: stock inválido para SKU ' + skuNorm);
  }

  if (stockNuevo < 0) {
    throw new Error('actualizarTodosStock: stock negativo bloqueado para SKU ' + skuNorm + ' -> ' + stockNuevo);
  }

  if (!EHI_CHANNELS[canalOrigen]) {
    throw new Error('actualizarTodosStock: canal no reconocido -> ' + canalOrigen);
  }

  Logger.log(
    '[EHI] actualizarTodosStock inicio | canalOrigen=%s | sku=%s | stock=%s',
    canalOrigen, skuNorm, stockNuevo
  );

  const dispatchers = {
    ML: function() {
      actualizarStockML_desdeParametros(skuNorm, stockNuevo);
      return 'ML actualizado';
    },
    FS: function() {
      actualizaProdFS(generaXmlFS(skuNorm, stockNuevo, null, null, null, 'stock'));
      return 'FS actualizado';
    },
    SH: function() {
      actualizarStockSH_desdeParametros(skuNorm, stockNuevo);
      return 'SH actualizado';
    },
    PM: function() {
      actualizarStockPM_desdeParametros_(skuNorm, stockNuevo);
      return 'PM actualizado';
    },
    MR: function() {
      throw new Error('MR pendiente de implementación');
    }
  };

  const canalesDestino = Object.keys(dispatchers).filter(function(destino) {
    return destino !== canalOrigen;
  });

  const resultados = [];

  canalesDestino.forEach(function(destino) {
    try {
      if (!skuExisteEnCanal_(destino, skuNorm)) {
        const detalleOmitido = 'SKU no existe en canal destino';
        resultados.push({ canal: destino, ok: true, estado: 'OMITIDO', msg: detalleOmitido });
        logSyncStock_(canalOrigen, destino, skuNorm, stockNuevo, 'OMITIDO', detalleOmitido);
        Logger.log('[EHI] %s OMITIDO | sku=%s | stock=%s | %s', destino, skuNorm, stockNuevo, detalleOmitido);
        return;
      }

      const msgOk = dispatchers[destino]();
      resultados.push({ canal: destino, ok: true, estado: 'OK', msg: msgOk || 'OK' });
      logSyncStock_(canalOrigen, destino, skuNorm, stockNuevo, 'OK', msgOk || 'OK');
      Logger.log('[EHI] %s OK | sku=%s | stock=%s', destino, skuNorm, stockNuevo);
    } catch (e) {
      const detalleError = e && e.message ? e.message : String(e);
      resultados.push({ canal: destino, ok: false, estado: 'ERROR', msg: detalleError });
      logSyncStock_(canalOrigen, destino, skuNorm, stockNuevo, 'ERROR', detalleError);
      Logger.log('[EHI] %s ERROR | sku=%s | stock=%s | %s', destino, skuNorm, stockNuevo, detalleError);
    }
  });

  const okCount = resultados.filter(function(r) { return r.estado === 'OK'; }).length;
  const omitCount = resultados.filter(function(r) { return r.estado === 'OMITIDO'; }).length;
  const errCount = resultados.filter(function(r) { return r.estado === 'ERROR'; }).length;

  SpreadsheetApp.getActive().toast(
    'Sync stock SKU ' + skuNorm + ': ' + okCount + ' OK / ' + omitCount + ' omitido(s) / ' + errCount + ' error(es).',
    'EHI',
    5
  );

  return resultados;
}


// ===================================================================
// Determina si el SKU existe en el canal destino
// Nota: usa búsqueda en memoria para mejorar performance
// ===================================================================
function skuExisteEnCanal_(canal, sku) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const skuNorm = String(sku || '').trim();

  if (!skuNorm) return false;

  switch (String(canal || '').trim().toUpperCase()) {
    case 'ML':
      return existeSkuEnHoja_(ss.getSheetByName('Prods. MLC'), 2, skuNorm);

    case 'FS':
      return existeSkuEnHoja_(ss.getSheetByName('Prods. FSC'), 2, skuNorm);

    case 'SH':
      return existeSkuEnHoja_(ss.getSheetByName('Prods. Shopify'), 2, skuNorm);

    case 'PM': {
      const sh = ss.getSheetByName((typeof SHEET_PRODS !== 'undefined' && SHEET_PRODS) ? SHEET_PRODS : 'Prods. PM');
      if (!sh) return false;
      if (typeof sheet_getHeaderMap_ !== 'function') return false;
      const hm = sheet_getHeaderMap_(sh);
      const colSkuSeller = hm['sku_seller'];
      if (!colSkuSeller) return false;
      return existeSkuEnHoja_(sh, colSkuSeller, skuNorm);
    }

    case 'MR':
      return false;

    default:
      return false;
  }
}


function existeSkuEnHoja_(sheet, col, sku) {
  if (!sheet) return false;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;

  const values = sheet.getRange(2, col, lastRow - 1, 1).getValues();
  const skuNorm = String(sku || '').trim();
  const skuSet = new Set(
    values.map(function(row) {
      return String(row[0] || '').trim();
    })
  );

  return skuSet.has(skuNorm);
}


// ===================================================================
// Log central de sincronización de stock
// ===================================================================
function logSyncStock_(canalOrigen, canalDestino, sku, stock, estado, detalle) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = 'Log Sync Stock';
  let sh = ss.getSheetByName(sheetName);

  if (!sh) {
    sh = ss.insertSheet(sheetName);
    sh.appendRow(['timestamp', 'canal_origen', 'canal_destino', 'sku', 'stock', 'estado', 'detalle']);
  }

  sh.appendRow([
    new Date(),
    canalOrigen || '',
    canalDestino || '',
    sku || '',
    stock,
    estado || '',
    detalle || ''
  ]);
}


// ===================================================================
// Paris / Marketplace Paris (PM)
// Busca sku_seller == sku y actualiza usando sku_mkp
// Requiere helpers ya existentes en tu bloque Paris
// ===================================================================
function actualizarStockPM_desdeParametros_(skuSeller, stockNuevo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = (typeof SHEET_PRODS !== 'undefined' && SHEET_PRODS) ? SHEET_PRODS : 'Prods. PM';
  const sh = ss.getSheetByName(sheetName);

  if (!sh) {
    throw new Error("Paris: no existe hoja '" + sheetName + "'.");
  }

  if (typeof sheet_getHeaderMap_ !== 'function') {
    throw new Error('Paris: no existe helper sheet_getHeaderMap_.');
  }
  if (typeof paris_getWarehouse_ !== 'function') {
    throw new Error('Paris: no existe helper paris_getWarehouse_.');
  }
  if (typeof paris_postStockInChunks_ !== 'function') {
    throw new Error('Paris: no existe helper paris_postStockInChunks_.');
  }

  const hm = sheet_getHeaderMap_(sh);
  const colSkuMkp = hm['sku_mkp'];
  const colSkuSeller = hm['sku_seller'];

  if (!colSkuMkp || !colSkuSeller) {
    throw new Error("Paris: faltan headers 'sku_mkp' o 'sku_seller'.");
  }

  const lastRow = sh.getLastRow();
  if (lastRow < 2) {
    throw new Error("Paris: la hoja '" + sheetName + "' no tiene datos.");
  }

  const readCols = Math.max(colSkuMkp, colSkuSeller);
  const values = sh.getRange(2, 1, lastRow - 1, readCols).getValues();

  let skuMkp = '';
  for (let i = 0; i < values.length; i++) {
    const rowSkuSeller = String(values[i][colSkuSeller - 1] || '').trim();
    if (rowSkuSeller === String(skuSeller).trim()) {
      skuMkp = String(values[i][colSkuMkp - 1] || '').trim();
      break;
    }
  }

  if (!skuMkp) {
    throw new Error('Paris: no encontré sku_mkp para sku_seller=' + skuSeller);
  }

  const warehouse = paris_getWarehouse_();

  paris_postStockInChunks_([{
    sku: skuMkp,
    sku_seller: String(skuSeller),
    quantity: Number(stockNuevo),
    warehouse: warehouse
  }]);
}


// ===================================================================
// Shopify: actualiza stock por SKU
// ===================================================================
function actualizarStockSH_desdeParametros(skuObjetivo, stockNuevo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Prods. Shopify');

  if (!sheet) {
    throw new Error("No existe la hoja 'Prods. Shopify'.");
  }

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow <= 1) {
    throw new Error("La hoja 'Prods. Shopify' no tiene datos.");
  }

  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  let rowIndex = -1;
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][1] || '').trim() === String(skuObjetivo).trim()) {
      rowIndex = i;
      break;
    }
  }

  if (rowIndex === -1) {
    throw new Error("Shopify: SKU no encontrado en 'Prods. Shopify': " + skuObjetivo);
  }

  const inventoryItemId = data[rowIndex][12]; // col M
  if (!inventoryItemId) {
    throw new Error('Shopify: falta Inventory Item ID para SKU ' + skuObjetivo);
  }

  const scriptProps = PropertiesService.getScriptProperties();
  const locationIdStr = scriptProps.getProperty('SHOPIFY_LOCATION_ID');

  if (!locationIdStr) {
    throw new Error('Falta SHOPIFY_LOCATION_ID en Script Properties.');
  }

  actualizarStockShopify_(inventoryItemId, Number(locationIdStr), stockNuevo);
}


// ===================================================================
// Shopify API call
// ===================================================================
function actualizarStockShopify_(inventoryItemId, locationId, disponible) {
  const scriptProps = PropertiesService.getScriptProperties();
  const storeDomain = scriptProps.getProperty('SHOPIFY_STORE_DOMAIN');
  const accessToken = scriptProps.getProperty('SHOPIFY_ADMIN_TOKEN');
  const apiVersion = '2024-10';

  if (!storeDomain || !accessToken) {
    throw new Error('Faltan SHOPIFY_STORE_DOMAIN o SHOPIFY_ADMIN_TOKEN en Script Properties.');
  }

  const url = 'https://' + storeDomain + '/admin/api/' + apiVersion + '/inventory_levels/set.json';

  const payload = {
    location_id: Number(locationId),
    inventory_item_id: Number(inventoryItemId),
    available: Number(disponible)
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'X-Shopify-Access-Token': accessToken
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const resp = UrlFetchApp.fetch(url, options);
  const code = resp.getResponseCode();

  if (code !== 200) {
    throw new Error('Shopify stock error (' + code + '): ' + resp.getContentText());
  }

  Logger.log('Shopify stock OK: ' + resp.getContentText());
}


// ===================================================================
// Función: actualiza algún atributo
// ===================================================================
function actualizarAtributo(canal, sku, valor, valor2, fecDesde, fecHasta, atributo) {
  switch (String(canal || '').trim().toUpperCase()) {
    case 'ML':
      actualizaProdML(generaJsonML(sku, valor, valor2, fecDesde, fecHasta, atributo));
      break;

    case 'FS':
      actualizaProdFS(generaXmlFS(sku, valor, valor2, fecDesde, fecHasta, atributo));
      break;

    case 'MR':
      Logger.log('actualizarAtributo: MR aún no implementado.');
      break;

    case 'PM':
      Logger.log('actualizarAtributo: PM aún no implementado para atributo=' + atributo);
      break;

    case 'SH':
      Logger.log('actualizarAtributo: SH no usa este router genérico para atributo=' + atributo);
      break;

    default:
      Logger.log('actualizarAtributo: canal no soportado -> ' + canal);
      break;
  }
}
