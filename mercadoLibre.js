/**
 * ===================================================================
 * mercadoLibre.gs
 * EHI - Importadora Elizalde
 *
 * Ajustes incluidos:
 * - Config central desde Script Properties
 * - Refresh automático de token
 * - Fetch autenticado reutilizable
 * - Corrección de obtenerIdYVariation()
 * - Escritura limpia en hoja Prods. MLC con checkboxes
 * - Mantiene estructura original y compatibilidad con actualizarAtributo()
 * ===================================================================
 */


/**
 * ===================================================================
 * Configuración MELI
 * ===================================================================
 */
function getMeliConfig_() {
  const props = PropertiesService.getScriptProperties();

  const cfg = {
    accessToken: String(props.getProperty('MELI_ACCESS_TOKEN') || '').trim(),
    clientId: String(props.getProperty('MELI_CLIENT_ID') || '').trim(),
    clientSecret: String(props.getProperty('MELI_CLIENT_SECRET') || '').trim(),
    nickname: String(props.getProperty('MELI_NICKNAME') || '').trim(),
    refreshToken: String(props.getProperty('MELI_REFRESH_TOKEN') || '').trim(),
    siteId: String(props.getProperty('MELI_SITEID') || '').trim(),
    userId: String(props.getProperty('MELI_USERID') || '').trim(),
    alertEmail: String(props.getProperty('MELI_ALERT_EMAIL') || 'contacto.qds@importadoraelizalde.com').trim()
  };

  if (!cfg.userId) throw new Error('Falta Script Property: MELI_USERID');
  if (!cfg.clientId) throw new Error('Falta Script Property: MELI_CLIENT_ID');
  if (!cfg.clientSecret) throw new Error('Falta Script Property: MELI_CLIENT_SECRET');
  if (!cfg.refreshToken) throw new Error('Falta Script Property: MELI_REFRESH_TOKEN');

  return cfg;
}


/**
 * ===================================================================
 * Helpers auth / fetch
 * ===================================================================
 */
function meliBuildHeaders_(accessToken, extraHeaders) {
  const headers = {
    Authorization: 'Bearer ' + accessToken,
    Accept: 'application/json'
  };

  if (extraHeaders) {
    Object.keys(extraHeaders).forEach(function(key) {
      headers[key] = extraHeaders[key];
    });
  }

  return headers;
}


function meliEnsureAccessToken_() {
  const cfg = getMeliConfig_();

  if (cfg.accessToken) {
    return cfg.accessToken;
  }

  refreshMeliTokens();

  const refreshed = getMeliConfig_().accessToken;
  if (!refreshed) {
    throw new Error('No fue posible obtener MELI_ACCESS_TOKEN después del refresh.');
  }

  return refreshed;
}


function meliApiFetch_(url, options, retryOn401) {
  const shouldRetry = retryOn401 !== false;
  let accessToken = meliEnsureAccessToken_();

  const finalOptions = Object.assign({}, options || {});
  finalOptions.method = finalOptions.method || 'get';
  finalOptions.muteHttpExceptions = true;
  finalOptions.headers = meliBuildHeaders_(accessToken, finalOptions.headers || {});

  let response = UrlFetchApp.fetch(url, finalOptions);

  if (response.getResponseCode() === 401 && shouldRetry) {
    Logger.log('[MELI] 401 detectado. Intentando refresh token...');
    refreshMeliTokens();

    accessToken = meliEnsureAccessToken_();
    finalOptions.headers = meliBuildHeaders_(accessToken, (options && options.headers) || {});
    response = UrlFetchApp.fetch(url, finalOptions);
  }

  return response;
}


function meliApiFetchAll_(requests) {
  const accessToken = meliEnsureAccessToken_();

  const finalRequests = (requests || []).map(function(req) {
    const cloned = Object.assign({}, req);
    cloned.method = cloned.method || 'get';
    cloned.muteHttpExceptions = true;
    cloned.headers = meliBuildHeaders_(accessToken, cloned.headers || {});
    return cloned;
  });

  return UrlFetchApp.fetchAll(finalRequests);
}


/**
 * ===================================================================
 * Hoja Prods. MLC
 * A: checkbox
 * B: sellerSku
 * C: gtin
 * D: itemId
 * E: variationId
 * F: brand
 * G: title
 * H: color
 * I: price
 * J: available_quantity
 * K: status
 * L: comentario
 * ===================================================================
 */
function getMeliSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName('Prods. MLC') || ss.insertSheet('Prods. MLC');
}


function getMeliSheetData_() {
  const sheet = getMeliSheet_();
  const lastRow = sheet.getLastRow();
  const lastCol = Math.max(sheet.getLastColumn(), 12);

  if (lastRow < 2) {
    return {
      sheet: sheet,
      rows: [],
      lastRow: lastRow,
      lastCol: lastCol
    };
  }

  return {
    sheet: sheet,
    rows: sheet.getRange(2, 1, lastRow - 1, lastCol).getValues(),
    lastRow: lastRow,
    lastCol: lastCol
  };
}


/**
 * ===================================================================
 * Consultar Productos (Optimizado: multiget + fetchAll por olas)
 * ===================================================================
 */
function consultaProdsML() {
  const cfg = getMeliConfig_();

  const baseScan = 'https://api.mercadolibre.com/users/' + cfg.userId + '/items/search?search_type=scan';

  let productIds = [];
  let scrollId = null;
  let hasMore = true;

  while (hasMore) {
    const url = scrollId
      ? (baseScan + '&scroll_id=' + encodeURIComponent(scrollId))
      : baseScan;

    const resp = meliApiFetch_(url, { method: 'get' }, true);

    if (resp.getResponseCode() !== 200) {
      throw new Error('Error scan: ' + resp.getResponseCode() + ' ' + resp.getContentText());
    }

    const body = JSON.parse(resp.getContentText());

    if (body && body.results && body.results.length) {
      productIds = productIds.concat(body.results);
    }

    if (body.scroll_id && body.results && body.results.length > 0) {
      scrollId = body.scroll_id;
    } else {
      hasMore = false;
    }
  }

  if (productIds.length === 0) {
    SpreadsheetApp.getActive().toast('No hay publicaciones para sincronizar.', 'Info', 3);
    return;
  }

  const IDS_PER_REQ = 20;
  const REQS_PER_WAVE = 10;
  const ATTRS = 'id,title,status,price,available_quantity,attributes,variations';
  const requests = [];

  for (let i = 0; i < productIds.length; i += IDS_PER_REQ) {
    const ids = productIds.slice(i, i + IDS_PER_REQ).join(',');
    const multiUrl = 'https://api.mercadolibre.com/items'
      + '?ids=' + encodeURIComponent(ids)
      + '&include_attributes=all'
      + '&attributes=' + encodeURIComponent(ATTRS);

    requests.push({
      url: multiUrl,
      method: 'get'
    });
  }

  const data = [];

  function getAttr(attrs, name) {
    if (!attrs) return null;
    for (let k = 0; k < attrs.length; k++) {
      if (attrs[k].id === name) return attrs[k].value_name || attrs[k].value_id || null;
    }
    return null;
  }

  function getVarAttr(attrsOrCombs, name) {
    if (!attrsOrCombs) return null;
    for (let k = 0; k < attrsOrCombs.length; k++) {
      if (attrsOrCombs[k].id === name) return attrsOrCombs[k].value_name || attrsOrCombs[k].value_id || null;
    }
    return null;
  }

  function colorFromVariation(v) {
    return (
      getVarAttr(v.attribute_combinations, 'COLOR') ||
      getVarAttr(v.attribute_combinations, 'LENS_COLOR') ||
      getVarAttr(v.attributes, 'COLOR') ||
      getVarAttr(v.attributes, 'LENS_COLOR') ||
      null
    );
  }

  function pushRow(rowArr) {
    data.push(rowArr);
  }

  for (let w = 0; w < requests.length; w += REQS_PER_WAVE) {
    const wave = requests.slice(w, w + REQS_PER_WAVE);
    const waveResps = meliApiFetchAll_(wave);

    for (let r = 0; r < waveResps.length; r++) {
      const httpCode = waveResps[r].getResponseCode();

      if (httpCode !== 200) {
        Logger.log('[ML] multiget http=%s body=%s', httpCode, waveResps[r].getContentText());
        continue;
      }

      const arr = JSON.parse(waveResps[r].getContentText());
      if (!Array.isArray(arr)) continue;

      for (let x = 0; x < arr.length; x++) {
        const item = arr[x];
        if (!item || item.code !== 200 || !item.body) continue;

        const it = item.body;
        const itemId = it.id;
        const title = it.title;
        const status = it.status;
        const price = it.price;
        const availableQuantity = it.available_quantity;
        const brand = getAttr(it.attributes, 'BRAND');

        if (!it.variations || it.variations.length === 0) {
          const gtin = getAttr(it.attributes, 'GTIN');
          const color = getAttr(it.attributes, 'COLOR');
          const sellerSku = getAttr(it.attributes, 'SELLER_SKU');
          const comentario = 'Catalogo';
          const variationId = '';

          pushRow([
            sellerSku, gtin, itemId, variationId, brand, title, color,
            price, availableQuantity, status, comentario
          ]);
          continue;
        }

        for (let j = 0; j < it.variations.length; j++) {
          const v = it.variations[j];

          const vPrice = (v.price != null ? v.price : price);
          const vQty = (v.available_quantity != null ? v.available_quantity : availableQuantity);
          const vGtin = getVarAttr(v.attributes, 'GTIN') || getAttr(it.attributes, 'GTIN');
          const vSku = getVarAttr(v.attributes, 'SELLER_SKU') || getAttr(it.attributes, 'SELLER_SKU');
          const vColor = colorFromVariation(v);
          const vItemId = (v.item_relations && v.item_relations.length > 0 && v.item_relations[0].id)
            ? v.item_relations[0].id
            : itemId;
          const vId = v.id;
          const comentarioV = '';

          pushRow([
            vSku, vGtin, vItemId, vId, brand, title, vColor,
            vPrice, vQty, status, comentarioV
          ]);
        }
      }
    }

    Utilities.sleep(150);
  }

  data.sort(function(a, b) {
    const brandA = a[4] == null ? '' : String(a[4]);
    const brandB = b[4] == null ? '' : String(b[4]);
    const byBrand = brandA.localeCompare(brandB);
    return byBrand !== 0 ? byBrand : String(a[0] || '').localeCompare(String(b[0] || ''));
  });

  const sheet = getMeliSheet_();
  const lastRow = sheet.getLastRow();

  if (lastRow > 1) {
    sheet.getRange(2, 2, lastRow - 1, 11).clearContent();
    sheet.getRange(2, 1, lastRow - 1, 1).clearContent().clearDataValidations();
  }

  if (data.length > 0) {
    sheet.getRange(2, 2, data.length, 11).setValues(data);

    const checkRange = sheet.getRange(2, 1, data.length, 1);
    const rule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
    checkRange.setDataValidation(rule);
    checkRange.setValue(false);
  }

  try { limpiaFiltro('Prods. MLC'); } catch (e) {}
  try { setPage(); } catch (e) {}
  try { creaFiltro('Prods. MLC', 1); } catch (e) {}

  SpreadsheetApp.getActive().toast(
    'Consulta de productos finalizada (' + productIds.length + ' ítems).',
    'Éxito',
    3
  );
}


/**
 * ===================================================================
 * Guardar la información en una fila de datos
 * ===================================================================
 */
function dataPush(data, sellerSku, gtin, itemId, variationId, brand, title, color, price, availableQuantity, status, comentario) {
  data.push([
    sellerSku,
    gtin,
    itemId,
    variationId,
    brand,
    title,
    color,
    price,
    availableQuantity,
    status,
    comentario
  ]);
}


/**
 * ===================================================================
 * Función auxiliar para obtener un valor de los atributos de una variación
 * ===================================================================
 */
function getVariationAttributeValue(attributes, attributeName) {
  if (!attributes) return null;

  for (var i = 0; i < attributes.length; i++) {
    if (attributes[i].id === attributeName) {
      return attributes[i].value_name || attributes[i].value_id;
    }
  }
  return null;
}


/**
 * ===================================================================
 * Función auxiliar para obtener un valor de los atributos de un producto
 * ===================================================================
 */
function getAttributeValue(attributes, attributeName) {
  if (!attributes) return null;

  for (var i = 0; i < attributes.length; i++) {
    if (attributes[i].id === attributeName) {
      return attributes[i].value_name || attributes[i].value_id;
    }
  }
  return null;
}


/**
 * ===================================================================
 * Función para obtener GTIN desde resultados de búsqueda
 * ===================================================================
 */
function getGtinFromSearch(results, itemId) {
  for (var i = 0; i < results.length; i++) {
    if (results[i].id === itemId) {
      var attributes = results[i].attributes || [];
      for (var j = 0; j < attributes.length; j++) {
        if (attributes[j].id === 'GTIN') {
          return attributes[j].value_name || '';
        }
      }
    }
  }
  return '';
}


/**
 * ===================================================================
 * setPage
 * ===================================================================
 */
function setPage() {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getSheetByName('Prods. MLC');

  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).activate();
  spreadsheet.getActiveRangeList().setFontFamily('Helvetica Neue')
    .setFontSize(8)
    .setVerticalAlignment('middle');

  spreadsheet.getRange('A1').activate();
  var currentCell = spreadsheet.getCurrentCell();

  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.NEXT).activate();
  currentCell.activateAsCurrentCell();

  currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.NEXT).activate();
  currentCell.activateAsCurrentCell();

  spreadsheet.getActiveRangeList().setFontWeight('bold')
    .setBackground('#ffff00');

  spreadsheet.getRange('A1').activate();
}


/**
 * ===================================================================
 * Refresca los tokens de Mercado Libre usando el Refresh Token almacenado
 * ===================================================================
 */
function refreshMeliTokens() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const cfg = getMeliConfig_();
  const emailContacto = cfg.alertEmail;

  try {
    const url = 'https://api.mercadolibre.com/oauth/token';
    const payload = {
      grant_type: 'refresh_token',
      client_id: cfg.clientId,
      client_secret: cfg.clientSecret,
      refresh_token: cfg.refreshToken
    };

    const options = {
      method: 'post',
      headers: {
        accept: 'application/json',
        'content-type': 'application/x-www-form-urlencoded'
      },
      payload: payload,
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);
    const bodyText = response.getContentText();
    const json = JSON.parse(bodyText);

    if (response.getResponseCode() === 200) {
      const newAccessToken = json.access_token;
      const newRefreshToken = json.refresh_token;

      if (!newAccessToken || !newRefreshToken) {
        throw new Error('La API respondió 200 pero no devolvió tokens válidos.');
      }

      scriptProperties.setProperty('MELI_ACCESS_TOKEN', newAccessToken);
      scriptProperties.setProperty('MELI_REFRESH_TOKEN', newRefreshToken);

      Logger.log('✅ Tokens de Mercado Libre actualizados correctamente.');
    } else {
      throw new Error('API MELI Error (' + response.getResponseCode() + '): ' + bodyText);
    }
  } catch (e) {
    Logger.log('❌ Fallo en refreshMeliTokens: ' + e.message);

    const asunto = '⚠️ ALERTA EHI: Fallo crítica en Tokens Mercado Libre';
    const cuerpo =
      'Se ha producido un error al intentar refrescar el token de Mercado Libre.\n\n' +
      'Detalle del error: ' + e.message + '\n\n' +
      'Acción requerida: Es posible que debas realizar el flujo de autorización manual nuevamente para generar un nuevo Refresh Token.';

    try {
      MailApp.sendEmail(emailContacto, asunto, cuerpo);
    } catch (mailErr) {
      Logger.log('No se pudo enviar email de alerta MELI: ' + mailErr.message);
    }

    throw e;
  }
}


/**
 * ===================================================================
 * Objetivo: Actualizar el stock de uno o más productos
 * ===================================================================
 */
function actualizarStockML() {
  const atributo = 'stock';
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getActiveSheet();

  if (sheet.getName() !== 'Prods. MLC') {
    Browser.msgBox("Esta acción solo funciona en la hoja 'Prods. MLC'.");
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const arrProds = sheet.getRange(2, 1, lastRow - 1, Math.max(sheet.getLastColumn(), 12)).getValues();
  let huboActualizacion = false;

  for (let i = 0; i < arrProds.length; i++) {
    if (arrProds[i][0] === true) {
      const sku = arrProds[i][1];
      const idML = arrProds[i][3];
      const variationId = arrProds[i][4];
      const nuevoStock = arrProds[i][9];

      if (!idML) continue;

      if (variationId) {
        const variaciones = obtenerVariacionesStockPorIdMl(idML, arrProds, sku, nuevoStock);

        if (variaciones.length > 0) {
          const arrId = [[idML, '']];
          const arrUrlOpt = generaJsonMLMultiple(arrId, variaciones, atributo);
          actualizaProdML(arrUrlOpt);
          huboActualizacion = true;
        }
      } else {
        const arrIdMl = [[idML, '']];
        const arrUrlOpt = generaJsonML(arrIdMl, nuevoStock, null, null, null, atributo);
        actualizaProdML(arrUrlOpt);
        huboActualizacion = true;
      }
    }
  }

  if (!huboActualizacion) {
    Browser.msgBox('Actualizar Stock ML: No se actualizó ningún producto.');
  }
}


/**
 * ===================================================================
 * Actualiza el estado (active/paused) de los productos seleccionados
 * ===================================================================
 */
function actualizarEstadoML() {
  const atributo = 'status';
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getActiveSheet();

  if (sheet.getName() !== 'Prods. MLC') {
    Browser.msgBox("Esta acción solo funciona en la hoja 'Prods. MLC'.");
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const arrProds = sheet.getRange(2, 1, lastRow - 1, Math.max(sheet.getLastColumn(), 12)).getValues();
  let huboActualizacion = false;

  for (let i = 0; i < arrProds.length; i++) {
    if (arrProds[i][0] === true) {
      const idML = arrProds[i][3];
      const nuevoEstado = arrProds[i][10];

      if (!idML || !nuevoEstado) continue;

      const arrIdMl = [[idML, '']];
      const arrUrlOpt = generaJsonML(arrIdMl, nuevoEstado, null, null, null, atributo);

      actualizaProdML(arrUrlOpt);
      huboActualizacion = true;
    }
  }

  if (!huboActualizacion) {
    Browser.msgBox('Actualizar Estado ML: No se actualizó ningún producto.');
  }
}


/**
 * ===================================================================
 * Objetivo: Actualizar el precio de uno o más productos
 * ===================================================================
 */
function actualizarPreciosML() {
  const atributo = 'precio';
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getActiveSheet();

  if (sheet.getName() !== 'Prods. MLC') {
    Browser.msgBox("Esta acción solo funciona en la hoja 'Prods. MLC'.");
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const arrProds = sheet.getRange(2, 1, lastRow - 1, Math.max(sheet.getLastColumn(), 12)).getValues();

  for (let i = 0; i < arrProds.length; i++) {
    if (arrProds[i][0] === true) {
      const sku = arrProds[i][1];
      const idML = arrProds[i][3];
      const valorSeleccionado = arrProds[i][8];

      if (!idML) continue;

      const variaciones = obtenerVariacionesPorIdMl(idML, arrProds, valorSeleccionado);

      if (variaciones.length > 1) {
        const arrId = [[idML, '']];
        const arrUrlOpt = generaJsonMLMultiple(arrId, variaciones, atributo);
        actualizaProdML(arrUrlOpt);
      } else {
        const arrIdMl = obtenerIdYVariation(sku, arrProds);
        if (arrIdMl !== 'No existe') {
          actualizarAtributo('ML', arrIdMl, valorSeleccionado, null, null, null, atributo);
        } else {
          Browser.msgBox('No se actualizó ningún producto para SKU: ' + sku);
        }
      }
    }
  }
}


/**
 * ===================================================================
 * obtenerVariacionesPorIdMl(idML, arrProd)
 * ===================================================================
 */
function obtenerVariacionesPorIdMl(idML, arrProd, precioBase) {
  const resultado = [];

  for (let i = 0; i < arrProd.length; i++) {
    if (arrProd[i][3] === idML && arrProd[i][4] !== '') {
      resultado.push({
        id: arrProd[i][4],
        price: precioBase
      });
    }
  }

  return resultado;
}


/**
 * ===================================================================
 * obtiene Id y Variation
 * ===================================================================
 */
function obtenerIdYVariation(SKU, arrProd) {
  const filasFiltradas = arrProd.filter(function(fila) {
    return fila[1] === SKU;
  });

  if (filasFiltradas.length === 1) {
    const fila = filasFiltradas[0];
    return [[fila[3], fila[4]]];
  }

  if (filasFiltradas.length > 1) {
    for (let i = 0; i < filasFiltradas.length; i++) {
      const fila = filasFiltradas[i];

      // comentario = índice 11 (columna L)
      if (fila[4] !== '' && fila[11] !== 'Catalogo') {
        return [[fila[3], fila[4]]];
      }
    }
  }

  return 'No existe';
}


/**
 * ===================================================================
 * Actualiza el producto MELI con manejo avanzado de errores
 * ===================================================================
 */
function actualizaProdML(arrUrlOpt) {
  const url = arrUrlOpt[0];
  const options = arrUrlOpt[1];

  try {
    const response = meliApiFetch_(url, options, true);
    const responseCode = response.getResponseCode();
    const rawText = response.getContentText();

    let result = {};
    try {
      result = JSON.parse(rawText);
    } catch (e) {
      result = { raw: rawText };
    }

    if (responseCode >= 200 && responseCode < 300) {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        'Actualización finalizada correctamente.',
        'Éxito',
        3
      );
    } else {
      let errorMessage = 'Error desconocido';

      if (result.message) {
        errorMessage = result.message;
      }

      if (
        result.error === 'validation_error' &&
        result.cause &&
        result.cause.length > 0
      ) {
        const causeCode = result.cause[0].code;

        if (causeCode === 'item.variations.price.different') {
          errorMessage =
            'Mercado Libre detectó que este producto tiene variaciones con precios diferentes. ' +
            'Debes enviar todas las variaciones con sus precios.';
        }

        if (causeCode === 'variations.not_updatable') {
          errorMessage =
            'Mercado Libre no permite modificar variaciones en esta publicación de catálogo.';
        }
      }

      SpreadsheetApp.getActiveSpreadsheet().toast(
        'Error ' + responseCode + ': ' + errorMessage,
        'Error',
        8
      );

      throw new Error('Mercado Libre error ' + responseCode + ': ' + errorMessage);
    }

    Logger.log('[actualizaProdML] Respuesta completa: ' + JSON.stringify(result, null, 2));
  } catch (error) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Error de excepción: ' + error.message,
      'Excepción',
      8
    );
    Logger.log('[actualizaProdML] Excepción: ' + error);
    throw error;
  }
}


/**
 * ===================================================================
 * generar los JSON de ML
 * ===================================================================
 */
function generaJsonML(arrId, valor, valor2, fini, ffin, queJSN) {
  let jsn = '';
  const itemId = arrId[0][0];
  const variationId = arrId[0][1];
  const url = 'https://api.mercadolibre.com/items/' + itemId;
  let payload = '';

  switch (queJSN) {
    case 'precio':
      if (variationId !== '') {
        payload = {
          variations: [{
            id: variationId,
            price: valor
          }]
        };
      } else {
        payload = {
          price: valor
        };
      }
      break;

    case 'stock':
      if (variationId !== '') {
        payload = {
          variations: [{
            id: variationId,
            available_quantity: valor
          }]
        };
      } else {
        payload = {
          available_quantity: valor
        };
      }
      break;

    case 'status':
      payload = {
        status: valor
      };
      break;

    case 'preciooferta':
      jsn = '';
      break;

    default:
      jsn = 'otro';
  }

  const options = {
    method: 'put',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  jsn = [url, options];
  return jsn;
}


/**
 * ===================================================================
 * generaJsonMLMultiple(arrId, variaciones, queJSN)
 * ===================================================================
 */
function generaJsonMLMultiple(arrId, variaciones, queJSN) {
  const itemId = arrId[0][0];
  const url = 'https://api.mercadolibre.com/items/' + itemId;

  let payload = {};

  switch (queJSN) {
    case 'precio':
      payload = {
        variations: variaciones.map(function(v) {
          return { id: v.id, price: v.price };
        })
      };
      break;

    case 'stock':
      payload = {
        variations: variaciones.map(function(v) {
          return { id: v.id, available_quantity: v.available_quantity };
        })
      };
      break;

    default:
      payload = {};
  }

  const options = {
    method: 'put',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  return [url, options];
}


/**
 * ===================================================================
 * Armar todas las variaciones con su stock
 * ===================================================================
 */
function obtenerVariacionesStockPorIdMl(idML, arrProd, skuObjetivo, stockNuevo) {
  const resultado = [];

  for (let i = 0; i < arrProd.length; i++) {
    if (arrProd[i][3] === idML && arrProd[i][4] !== '') {
      const esSkuObjetivo = (arrProd[i][1] === skuObjetivo);
      const stockHoja = arrProd[i][9];

      resultado.push({
        id: arrProdsSafeValue_(arrProd[i][4]),
        available_quantity: esSkuObjetivo ? stockNuevo : stockHoja
      });
    }
  }

  return resultado;
}


function arrProdsSafeValue_(value) {
  return value == null ? '' : value;
}


/**
 * ===================================================================
 * Actualiza los stock sin depender de la hoja activa
 * ===================================================================
 */
function actualizarStockML_desdeParametros(skuObjetivo, stockNuevo) {
  const info = getMeliSheetData_();
  const arrProds = info.rows;

  if (arrProds.length === 0) {
    throw new Error("La hoja 'Prods. MLC' no tiene datos.");
  }

  let idML = null;
  let tieneVariaciones = false;

  for (let i = 0; i < arrProds.length; i++) {
    if (arrProds[i][1] === skuObjetivo) {
      idML = arrProds[i][3];
      tieneVariaciones = !!arrProds[i][4];
      break;
    }
  }

  if (!idML) {
    throw new Error('No se encontró Id ML para el SKU: ' + skuObjetivo);
  }

  if (tieneVariaciones) {
    const variaciones = obtenerVariacionesStockPorIdMl(idML, arrProds, skuObjetivo, stockNuevo);

    if (variaciones.length === 0) {
      throw new Error('No se encontraron variaciones para el Id ML: ' + idML);
    }

    const arrId = [[idML, '']];
    const arrUrlOpt = generaJsonMLMultiple(arrId, variaciones, 'stock');
    actualizaProdML(arrUrlOpt);
  } else {
    const arrIdMl = [[idML, '']];
    const arrUrlOpt = generaJsonML(arrIdMl, stockNuevo, null, null, null, 'stock');
    actualizaProdML(arrUrlOpt);
  }
}