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

// ===================================================================
// MLC Ventas - Ingesta histórica / incremental hacia hoja "Ventas MLC"
// Fase actual:
// - descarga histórico o incremental desde Mercado Libre
// - escribe SOLO en hoja "Ventas MLC"
// - NO inserta en hoja "Ventas"
// - NO ejecuta registrarVenta()
// - NO toca stock
//
// Requiere Script Properties:
// MELI_ACCESS_TOKEN
// MELI_USERID
// MELI_LAST_SYNC
//
// Hoja destino:
// "Ventas MLC"
// Encabezados ya creados en fila 1
// Escritura desde fila 2
// ===================================================================

const MLC_SHEET_VENTAS = 'Ventas MLC';
const MLC_PAGE_LIMIT = 50;

// -------------------------------------------------------------------
// Entry point histórico
// -------------------------------------------------------------------
function ingestarVentasMLC_historico() {
  return ingestarVentasMLC_({
    modo: 'HISTORICO',
    actualizarLastSync: false
  });
}

// -------------------------------------------------------------------
// Entry point incremental
// -------------------------------------------------------------------
function ingestarVentasMLC_incremental() {
  return ingestarVentasMLC_({
    modo: 'INCREMENTAL',
    actualizarLastSync: true
  });
}

// -------------------------------------------------------------------
// Motor principal
// -------------------------------------------------------------------
function ingestarVentasMLC_(opts) {
  opts = opts || {};

  const modo = String(opts.modo || 'HISTORICO').toUpperCase();
  const actualizarLastSync = !!opts.actualizarLastSync;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(MLC_SHEET_VENTAS);
  if (!sh) throw new Error("No existe la hoja '" + MLC_SHEET_VENTAS + "'.");

  const props = PropertiesService.getScriptProperties();
  const shipmentCache = {}; 
  const cfg = getMeliConfig_();
  const sellerId = String(cfg.userId || '').trim();
  const lastSyncProp = String(props.getProperty('MELI_LAST_SYNC') || '').trim();
  const historicalFromProp = String(props.getProperty('MELI_HISTORICAL_FROM') || '').trim();

  if (!sellerId) throw new Error('Falta MELI_USERID en Script Properties.');

  const searchFrom = (modo === 'HISTORICO')
    ? (historicalFromProp || MLC_DEFAULT_HISTORICAL_FROM)
    : lastSyncProp;

  if (modo === 'HISTORICO' && !historicalFromProp) {
    throw new Error(
      'Falta definir MELI_HISTORICAL_FROM en Script Properties.'
    );
  }

  if (!searchFrom) {
    throw new Error(
      modo === 'HISTORICO'
        ? 'No se pudo resolver fecha histórica base.'
        : 'Falta MELI_LAST_SYNC en Script Properties.'
    );
  }

  const headerMap = mlc_getHeaderMap_(sh);

  if (!headerMap['clave_unica']) {
    throw new Error("La hoja 'Ventas MLC' debe tener header 'clave_unica'.");
  }

  const existingKeys = mlc_getExistingKeys_(sh, headerMap['clave_unica']);

  let offset = 0;
  let totalFetchedOrders = 0;
  let totalRowsPrepared = 0;
  let totalRowsInserted = 0;
  let maxDateCreatedSeen = lastSyncProp || searchFrom;
  let shouldContinue = true;

  SpreadsheetApp.getActive().toast('Ingesta MLC iniciada (' + modo + ')', 'EHI', 5);

  while (shouldContinue) {
    const searchResp = mlc_fetchOrdersPage_(sellerId, searchFrom, offset, MLC_PAGE_LIMIT);
    const results = (searchResp && searchResp.results) ? searchResp.results : [];

    if (!results.length) break;

    totalFetchedOrders += results.length;

    const rowsToAppend = [];

    // Pre-fetch todos los shipments del lote en paralelo
    var batchShipmentIds = [];
    for (var pi = 0; pi < results.length; pi++) {
      var shippingNode = results[pi] && results[pi].shipping ? results[pi].shipping : {};
      var sid = String(shippingNode.id || '').trim();
      if (sid && !Object.prototype.hasOwnProperty.call(shipmentCache, sid)) {
        batchShipmentIds.push(sid);
      }
    }

    // Mapa shipmentId → cantidad de órdenes que comparten ese shipment en este lote
    var shipmentOrderCount = {};
    for (var ci = 0; ci < results.length; ci++) {
      var cShipping = results[ci] && results[ci].shipping ? results[ci].shipping : {};
      var cSid = String(cShipping.id || '').trim();
      if (cSid) {
        shipmentOrderCount[cSid] = (shipmentOrderCount[cSid] || 0) + 1;
      }
    }

    // Eliminar duplicados
    batchShipmentIds = batchShipmentIds.filter(function(v, idx, arr) { return arr.indexOf(v) === idx; });
    if (batchShipmentIds.length) {
      var fetchedShipments = mlc_fetchShipmentsBatch_(batchShipmentIds);
      Object.keys(fetchedShipments).forEach(function(k) { shipmentCache[k] = fetchedShipments[k]; });
    }

    for (var i = 0; i < results.length; i++) {
      const order = results[i];
      const dateCreated = mlc_safeIso_(order.date_created || order.date_closed || '');

      if (dateCreated && dateCreated > maxDateCreatedSeen) {
        maxDateCreatedSeen = dateCreated;
      }

      const rowsBuilt = mlc_buildRowsFromOrder_(order, headerMap, existingKeys, modo, shipmentCache, shipmentOrderCount);
      totalRowsPrepared += rowsBuilt.length;

      for (var j = 0; j < rowsBuilt.length; j++) {
        rowsToAppend.push(rowsBuilt[j]);
      }
    }

    if (rowsToAppend.length) {
      const startRow = Math.max(sh.getLastRow() + 1, 2);
      sh.getRange(startRow, 1, rowsToAppend.length, rowsToAppend[0].length).setValues(rowsToAppend);
      totalRowsInserted += rowsToAppend.length;
    }

    if (results.length < MLC_PAGE_LIMIT) {
      shouldContinue = false;
    } else {
      offset += MLC_PAGE_LIMIT;
      Utilities.sleep(150);
    }
  }

  if (actualizarLastSync && maxDateCreatedSeen && maxDateCreatedSeen !== lastSyncProp) {
    props.setProperty('MELI_LAST_SYNC', maxDateCreatedSeen);
  }

  const resumen = {
    modo: modo,
    searchFrom: searchFrom,
    fetchedOrders: totalFetchedOrders,
    preparedRows: totalRowsPrepared,
    insertedRows: totalRowsInserted,
    lastSyncAnterior: lastSyncProp,
    lastSyncNuevo: actualizarLastSync ? maxDateCreatedSeen : lastSyncProp
  };

  Logger.log('[MLC] Ingesta resumen: ' + JSON.stringify(resumen));
  SpreadsheetApp.getActive().toast(
    'MLC ' + modo + ': órdenes=' + totalFetchedOrders + ' | filas nuevas=' + totalRowsInserted,
    'EHI',
    7
  );

  return resumen;
}

// -------------------------------------------------------------------
// Mercado Libre - búsqueda paginada de órdenes
// -------------------------------------------------------------------
function mlc_fetchOrdersPage_(sellerId, lastSyncIso, offset, limit) {
  var url =
    'https://api.mercadolibre.com/orders/search'
    + '?seller=' + encodeURIComponent(sellerId)
    + '&order.date_created.from=' + encodeURIComponent(lastSyncIso)
    + '&sort=date_desc'
    + '&offset=' + encodeURIComponent(offset)
    + '&limit=' + encodeURIComponent(limit);

  var resp = meliApiFetch_(url, { method: 'get' }, true);
  var code = resp.getResponseCode();
  var body = resp.getContentText();

  if (code !== 200) {
    throw new Error('ML orders/search error (' + code + '): ' + body);
  }

  return JSON.parse(body);
}

// -------------------------------------------------------------------
// Construye filas destino desde una orden ML
// Una fila por order_item
// -------------------------------------------------------------------
function mlc_buildRowsFromOrder_(order, headerMap, existingKeys, modo, shipmentCache, shipmentOrderCount) {
  const rows = [];
  const items = order && order.order_items ? order.order_items : [];
  if (!items.length) return rows;

  shipmentCache = shipmentCache || {};

  const canal = 'ML';
  const fechaIngesta = new Date();
  const mlOrderId = mlc_toStr_(order.id);
  const mlPackId = mlc_toStr_(order.pack_id);
  const idVentaCanal = mlOrderId;
  const fechaVenta = mlc_safeIso_(order.date_created || order.date_closed || '');
  const estadoMl = mlc_toStr_(order.status);

  const buyer = order && order.buyer ? order.buyer : {};
  const payments = order && order.payments ? order.payments : [];
  const shippingNode = order && order.shipping ? order.shipping : {};
  const shipmentId = mlc_toStr_(shippingNode.id);

  const shipment = shipmentId ? mlc_getShipmentCached_(shipmentId, shipmentCache) : null;
  const region = mlc_extractRegion_(shipment);
  const comuna = mlc_extractCity_(shipment);
  const direccionResumen = mlc_extractAddress_(shipment);
  const clienteNombre = mlc_resolveBuyerName_(buyer, shipment);

  const totalOrderAmount = mlc_sumOrderItemsAmount_(items);
  const cargoEnvioRaw     = mlc_resolveShipmentCost_(shipment, totalOrderAmount);
  var   ordenesEnShipment = (shipmentId && shipmentOrderCount && shipmentOrderCount[shipmentId]) || 1;
  const totalCargoEnvioOrden = (ordenesEnShipment > 1)
    ? Math.round(cargoEnvioRaw / ordenesEnShipment)
    : cargoEnvioRaw;

  for (var i = 0; i < items.length; i++) {
    const oi = items[i] || {};
    const item = oi.item || {};

    const mlItemId = mlc_toStr_(item.id);
    const mlVariationId = mlc_toStr_(item.variation_id);
    const skuCanal = mlc_resolveSkuCanal_(item);
    const skuMaestro = skuCanal;
    const mpn = '';

    const cantidad = Number(oi.quantity || 0);
    const precioUnitario = Number(oi.unit_price || 0);
    const montoTotal = cantidad * precioUnitario;

    const esCancelada  = (order.status === 'cancelled');
    const esDevolucion = esCancelada
      && (order.cancel_detail && order.cancel_detail.group === 'mediations')
      && ((order.tags || []).indexOf('delivered') !== -1);

    const soloAnularMontos = esCancelada && !esDevolucion;

    const cargoVenta = soloAnularMontos ? 0 : mlc_resolveRowSaleFee_(oi, payments, montoTotal, totalOrderAmount);
    const cargoEnvio = soloAnularMontos ? 0 : mlc_prorateAmount_(totalCargoEnvioOrden, montoTotal, totalOrderAmount);
    const montoNeto  = soloAnularMontos ? 0 : montoTotal - cargoVenta - cargoEnvio;

    const claveUnica = 'ML|' + mlOrderId + '|' + skuCanal;

    if (existingKeys.has(claveUnica)) {
      continue;
    }
    existingKeys.add(claveUnica);

    const estadoConciliacion = esDevolucion        ? 'DEVOLUCION CON REEMBOLSO'
                             : esCancelada         ? 'CANCELADA POR EL COMPRADOR'
                             : (modo === 'HISTORICO') ? 'PENDIENTE_HISTORICO'
                             : 'PENDIENTE';
    const fechaConciliacion = '';
    const mensajeConciliacion = '';
    const existeEnVentas = 'NO';
    const jsonRaw = JSON.stringify(order);
    const observaciones = shipmentId && !shipment ? 'Sin detalle shipment' : '';

    const row = mlc_buildEmptyRow_(headerMap);

    mlc_setCell_(row, headerMap, 'fecha_ingesta', fechaIngesta);
    mlc_setCell_(row, headerMap, 'canal', canal);
    mlc_setCell_(row, headerMap, 'ml_order_id', mlOrderId);
    mlc_setCell_(row, headerMap, 'ml_pack_id', mlPackId);
    mlc_setCell_(row, headerMap, 'ml_item_id', mlItemId);
    mlc_setCell_(row, headerMap, 'ml_variation_id', mlVariationId);
    mlc_setCell_(row, headerMap, 'id_venta_canal', idVentaCanal);
    mlc_setCell_(row, headerMap, 'fecha_venta', fechaVenta);
    mlc_setCell_(row, headerMap, 'sku_canal', skuCanal);
    mlc_setCell_(row, headerMap, 'sku_maestro', skuMaestro);
    mlc_setCell_(row, headerMap, 'mpn', mpn);
    mlc_setCell_(row, headerMap, 'cantidad', cantidad);
    mlc_setCell_(row, headerMap, 'precio_unitario', precioUnitario);
    mlc_setCell_(row, headerMap, 'monto_total', montoTotal);
    mlc_setCell_(row, headerMap, 'cargo_venta', cargoVenta);
    mlc_setCell_(row, headerMap, 'cargo_envio', cargoEnvio);
    mlc_setCell_(row, headerMap, 'monto_neto', montoNeto);
    mlc_setCell_(row, headerMap, 'cliente_nombre', clienteNombre);
    mlc_setCell_(row, headerMap, 'region', region);
    mlc_setCell_(row, headerMap, 'comuna', comuna);
    mlc_setCell_(row, headerMap, 'direccion_resumen', direccionResumen);
    mlc_setCell_(row, headerMap, 'estado_ml', estadoMl);
    mlc_setCell_(row, headerMap, 'estado_conciliacion', estadoConciliacion);
    mlc_setCell_(row, headerMap, 'fecha_conciliacion', fechaConciliacion);
    mlc_setCell_(row, headerMap, 'mensaje_conciliacion', mensajeConciliacion);
    mlc_setCell_(row, headerMap, 'clave_unica', claveUnica);
    mlc_setCell_(row, headerMap, 'existe_en_ventas', existeEnVentas);
    mlc_setCell_(row, headerMap, 'json_raw', jsonRaw);
    mlc_setCell_(row, headerMap, 'observaciones', observaciones);

    rows.push(row);
  }

  return rows;
}

// -------------------------------------------------------------------
// Helpers de hoja
// -------------------------------------------------------------------
function mlc_getHeaderMap_(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) throw new Error("La hoja '" + sheet.getName() + "' no tiene encabezados.");

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const map = {};

  for (var i = 0; i < headers.length; i++) {
    const h = String(headers[i] || '').trim();
    if (h) map[h] = i + 1;
  }

  return map;
}

function mlc_getExistingKeys_(sheet, keyCol) {
  const set = new Set();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return set;

  const values = sheet.getRange(2, keyCol, lastRow - 1, 1).getValues();
  for (var i = 0; i < values.length; i++) {
    const key = String(values[i][0] || '').trim();
    if (key) set.add(key);
  }
  return set;
}

function mlc_buildEmptyRow_(headerMap) {
  const lastCol = Math.max.apply(null, Object.keys(headerMap).map(function(k) { return headerMap[k]; }));
  const row = new Array(lastCol);
  for (var i = 0; i < lastCol; i++) row[i] = '';
  return row;
}

function mlc_setCell_(row, headerMap, headerName, value) {
  const col = headerMap[headerName];
  if (!col) return;
  row[col - 1] = value;
}

// -------------------------------------------------------------------
// Helpers ML
// -------------------------------------------------------------------
function mlc_toStr_(v) {
  return String(v == null ? '' : v).trim();
}

function mlc_safeIso_(v) {
  return String(v || '').trim();
}

function mlc_joinName_(firstName, lastName) {
  return [mlc_toStr_(firstName), mlc_toStr_(lastName)].join(' ').trim();
}

function mlc_resolveSkuCanal_(item) {
  if (!item) return '';

  if (item.seller_custom_field) return mlc_toStr_(item.seller_custom_field);
  if (item.seller_sku) return mlc_toStr_(item.seller_sku);
  if (item.sku) return mlc_toStr_(item.sku);

  const attrs = item.variation_attributes || item.attributes || [];
  for (var i = 0; i < attrs.length; i++) {
    const a = attrs[i] || {};
    const id = mlc_toStr_(a.id).toUpperCase();
    const name = mlc_toStr_(a.name).toUpperCase();
    const valueName = mlc_toStr_(a.value_name);
    const valueId = mlc_toStr_(a.value_id);

    if (id === 'SELLER_SKU' || name === 'SELLER SKU') return valueName || valueId;
    if (id === 'SKU' || name === 'SKU') return valueName || valueId;
  }

  return mlc_toStr_(item.id);
}

function mlc_sumFees_(payments) {
  if (!payments || !payments.length) return 0;

  var total = 0;
  for (var i = 0; i < payments.length; i++) {
    var p = payments[i] || {};

    if (p.marketplace_fee != null) {
      total += Number(p.marketplace_fee) || 0;
    } else if (p.fee_amount != null) {
      total += Number(p.fee_amount) || 0;
    }
  }
  return total;
}

function mlc_extractRegion_(shipping) {
  try {
    return mlc_toStr_(
      shipping.receiver_address &&
      shipping.receiver_address.state &&
      shipping.receiver_address.state.name
    );
  } catch (e) {
    return '';
  }
}

function mlc_extractCity_(shipping) {
  try {
    return mlc_toStr_(
      shipping.receiver_address &&
      shipping.receiver_address.city &&
      shipping.receiver_address.city.name
    );
  } catch (e) {
    return '';
  }
}

function mlc_extractAddress_(shipping) {
  try {
    var ra = shipping.receiver_address || {};
    var parts = [
      mlc_toStr_(ra.address_line),
      mlc_toStr_(ra.comment)
    ].filter(function(x) { return x; });

    return parts.join(' | ');
  } catch (e) {
    return '';
  }
}

function mlc_buildVentaRowFromOrder_(order, canal) {

  var orderId = order.id || '';
  var fecha = order.date_created || '';
  var estado = order.status || '';

  var item = (order.order_items && order.order_items[0]) || {};
  var product = item.item || {};

  var sku = product.seller_sku || '';
  var titulo = product.title || '';
  var cantidad = item.quantity || 0;
  var precio = item.unit_price || 0;

  var payments = order.payments || [];
  var shippingCost = '';
  var cargoVenta = '';

  if (payments.length) {
    shippingCost = payments[0].shipping_cost || 0;
  }

  if (item.sale_fee != null) {
    cargoVenta = item.sale_fee;
  }

  var buyer = order.buyer || {};
  var clienteNombre = buyer.nickname || '';

  return [
    fecha,
    '',                 // # Venta
    sku,
    precio,
    cantidad,
    canal || 'MLC',
    cargoVenta,
    shippingCost,
    '',                 // tipo doc
    '',                 // folio
    '',                 // fecha boleta
    orderId,
    estado,
    titulo,
    JSON.stringify(order)   // json_raw
  ];
}

function mlc_resolveBuyerName_(buyer, shipment) {
  buyer   = buyer   || {};
  shipment = shipment || {};

  // 1. Nombre real desde dirección de entrega del shipment
  var receiverName = shipment.receiver_address && shipment.receiver_address.receiver_name
    ? mlc_toStr_(shipment.receiver_address.receiver_name)
    : '';
  if (receiverName) return receiverName;

  // 2. Nombre desde buyer (si ML lo entrega)
  var fullName = mlc_joinName_(buyer.first_name, buyer.last_name);
  if (fullName) return fullName;

  // 3. Fallback: nickname
  if (buyer.nickname) return mlc_toStr_(buyer.nickname);

  return '';
}

function mlc_sumSaleFees_(orderItems) {
  if (!orderItems || !orderItems.length) return 0;

  var total = 0;
  for (var i = 0; i < orderItems.length; i++) {
    var oi = orderItems[i] || {};
    if (oi.sale_fee != null) {
      total += Number(oi.sale_fee) || 0;
    }
  }
  return total;
}

function mlc_sumShippingCost_(payments, order) {
  var total = 0;

  if (payments && payments.length) {
    for (var i = 0; i < payments.length; i++) {
      var p = payments[i] || {};
      if (p.shipping_cost != null) {
        total += Number(p.shipping_cost) || 0;
      }
    }
  }

  if (!total && order && order.shipping_cost != null) {
    total = Number(order.shipping_cost) || 0;
  }

  return total;
}

function mlc_sumOrderItemsAmount_(orderItems) {
  if (!orderItems || !orderItems.length) return 0;

  var total = 0;
  for (var i = 0; i < orderItems.length; i++) {
    var oi = orderItems[i] || {};
    var qty = Number(oi.quantity || 0);
    var price = Number(oi.unit_price || 0);
    total += qty * price;
  }
  return total;
}

function mlc_prorateAmount_(totalAmount, rowAmount, baseAmount) {
  var total = Number(totalAmount || 0);
  var row = Number(rowAmount || 0);
  var base = Number(baseAmount || 0);

  if (!total || !row || !base) return 0;
  return Math.round((total * row / base) * 100) / 100;
}

function mlc_resolveRowSaleFee_(orderItem, payments, rowAmount, totalOrderAmount) {
  orderItem = orderItem || {};

  if (orderItem.sale_fee != null) {
    var fee      = Number(orderItem.sale_fee) || 0;
    var cantidad = Number(orderItem.quantity  || 1);
    return fee * cantidad;
  }

  var totalFees = mlc_sumFees_(payments);
  return mlc_prorateAmount_(totalFees, rowAmount, totalOrderAmount);
}

function mlc_resolveShipmentCost_(shipment, montoTotal) {
  if (!shipment || typeof shipment !== 'object') {
    Logger.log('[MLC shipment cost] shipment vacío o inválido');
    return 0;
  }

  // Si el comprador pagó el envío (venta < 19990), el costo no es del vendedor
  if ((montoTotal || 0) < 19990) {
    Logger.log('[MLC shipment cost] comprador paga envío (monto < 19990), cargo_envio=0');
    return 0;
  }

  // 1. shipping_option.list_cost = cargo real al vendedor en MLC
  var listCost = shipment.shipping_option && shipment.shipping_option.list_cost;
  if (listCost != null && listCost !== '') {
    var lc = Number(listCost) || 0;
    Logger.log('[MLC shipment cost] usando shipping_option.list_cost=' + lc);
    return lc;
  }

  // 2. base_cost = fallback
  if (shipment.base_cost != null && shipment.base_cost !== '') {
    var baseCost = Number(shipment.base_cost) || 0;
    Logger.log('[MLC shipment cost] usando base_cost=' + baseCost);
    return baseCost;
  }

  Logger.log('[MLC shipment cost] no se encontró costo vendedor en shipment.id=' + (shipment.id || ''));
  return 0;
}

// -------------------------------------------------------------------
// Obtiene detalle de shipment desde Mercado Libre
// -------------------------------------------------------------------
function mlc_fetchShipment_(shipmentId) {
  var id = String(shipmentId || '').trim();
  if (!id) return null;

  var url = 'https://api.mercadolibre.com/shipments/' + encodeURIComponent(id);

  var resp = meliApiFetch_(url, { method: 'get' }, true);
  var code = resp.getResponseCode();
  var body = resp.getContentText();

  if (code === 404) {
    Logger.log('[MLC shipment] no encontrado: ' + id);
    return null;
  }

  if (code !== 200) {
    throw new Error('ML shipments/{id} error (' + code + '): ' + body);
  }

  var json = JSON.parse(body || '{}');
  Logger.log('[MLC shipment] cargado: ' + id);

  return json;
}

// -------------------------------------------------------------------
// Obtiene múltiples shipments en paralelo (batches de 10)
// Retorna {id: shipmentObj}
// -------------------------------------------------------------------
function mlc_fetchShipmentsBatch_(shipmentIds) {
  var result = {};
  var ids = (shipmentIds || []).filter(function(id) { return !!String(id || '').trim(); });
  if (!ids.length) return result;

  var BATCH = 10;
  for (var i = 0; i < ids.length; i += BATCH) {
    var batch = ids.slice(i, i + BATCH);
    var requests = batch.map(function(sid) {
      return { url: 'https://api.mercadolibre.com/shipments/' + encodeURIComponent(String(sid).trim()), method: 'get' };
    });

    var responses = meliApiFetchAll_(requests);

    for (var j = 0; j < responses.length; j++) {
      var sid = String(batch[j]).trim();
      var resp = responses[j];
      var code = resp.getResponseCode();
      if (code === 200) {
        try {
          result[sid] = JSON.parse(resp.getContentText() || '{}');
          Logger.log('[MLC shipment batch] cargado: ' + sid);
        } catch(e) {
          result[sid] = null;
        }
      } else if (code === 404) {
        Logger.log('[MLC shipment batch] no encontrado: ' + sid);
        result[sid] = null;
      } else {
        Logger.log('[MLC shipment batch] error (' + code + ') para id=' + sid);
        result[sid] = null;
      }
    }

    if (i + BATCH < ids.length) {
      Utilities.sleep(200);
    }
  }

  return result;
}

// -------------------------------------------------------------------
// Cache local de shipments para no repetir llamadas
// -------------------------------------------------------------------
function mlc_getShipmentCached_(shipmentId, shipmentCache) {
  var id = String(shipmentId || '').trim();
  if (!id) return null;

  shipmentCache = shipmentCache || {};

  if (Object.prototype.hasOwnProperty.call(shipmentCache, id)) {
    return shipmentCache[id];
  }

  var shipment = mlc_fetchShipment_(id);
  shipmentCache[id] = shipment || null;

  return shipmentCache[id];
}

// -------------------------------------------------------------------
// Conciliación: cruza Ventas MLC contra hoja Ventas
// Procesa: PENDIENTE, PENDIENTE_HISTORICO, DEVOLUCION CON REEMBOLSO
// Omite:   CANCELADA POR EL COMPRADOR, CONCILIADA, DIFERENCIA
// -------------------------------------------------------------------
function mlc_conciliar() {
  var ss         = SpreadsheetApp.getActiveSpreadsheet();
  var sheetMLC   = ss.getSheetByName('Ventas MLC');
  var sheetVentas = ss.getSheetByName('Ventas');

  if (!sheetMLC)    throw new Error('Hoja "Ventas MLC" no encontrada');
  if (!sheetVentas) throw new Error('Hoja "Ventas" no encontrada');

  // --- Leer encabezados Ventas MLC ---
  var headersMLC = sheetMLC.getRange(1, 1, 1, sheetMLC.getLastColumn()).getValues()[0];
  var hMLC = {};
  headersMLC.forEach(function(h, i) { hMLC[h] = i; });

  var COL = {
    mlOrderId          : hMLC['ml_order_id'],
    mlPackId           : hMLC['ml_pack_id'],
    skuCanal           : hMLC['sku_canal'],
    montoTotal         : hMLC['monto_total'],
    cargoVenta         : hMLC['cargo_venta'],
    cargoEnvio         : hMLC['cargo_envio'],
    montoNeto          : hMLC['monto_neto'],
    estadoConciliacion : hMLC['estado_conciliacion'],
    fechaConciliacion  : hMLC['fecha_conciliacion'],
    mensajeConciliacion: hMLC['mensaje_conciliacion']
  };

  // --- Leer encabezados hoja Ventas ---
  var headersV = sheetVentas.getRange(1, 1, 1, sheetVentas.getLastColumn()).getValues()[0];
  var hV = {};
  headersV.forEach(function(h, i) { hV[h] = i; });

  var COLV = {
    numVenta   : hV['# Venta'],
    sku        : hV['SKU'],
    precioVenta: hV['Precio\nVenta'],
    cargosVenta: hV['Cargos por\nVenta'],
    cargoEnvio : hV['Cargo por\nEnvío'],
    recaudado  : hV['Recaudado']
  };

  // --- Construir mapas de hoja Ventas ---
  var dataVentas = sheetVentas.getRange(2, 1, sheetVentas.getLastRow() - 1, sheetVentas.getLastColumn()).getValues();
  var mapaVentas          = {};   // # Venta → fila
  var mapaVentasCompuesto = {};   // # Venta|SKU → fila
  dataVentas.forEach(function(row) {
    var id = mlc_toStr_(row[COLV.numVenta]).trim();
    if (!id) return;

    // Ignorar notas de crédito (valores negativos)
    var precioVenta = mlc_parseChileanNumber_(row[COLV.precioVenta]);
    if (precioVenta < 0) return;

    mapaVentas[id] = row;

    var skuV = mlc_toStr_(row[COLV.sku]).trim();
    if (skuV) {
      mapaVentasCompuesto[id + '|' + skuV] = row;
    }
  });

  // --- Leer filas Ventas MLC ---
  var lastRow = sheetMLC.getLastRow();
  if (lastRow < 2) {
    Logger.log('[MLC conciliar] Sin filas para procesar');
    return;
  }

  var dataMLC = sheetMLC.getRange(2, 1, lastRow - 1, sheetMLC.getLastColumn()).getValues();

  var ESTADOS_A_PROCESAR = {
    'PENDIENTE'              : true,
    'PENDIENTE_HISTORICO'    : true,
    'DEVOLUCION CON REEMBOLSO': true
  };

  var TOLERANCIA = 1;
  var fechaHoy   = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');
  var cambios    = 0;

  for (var i = 0; i < dataMLC.length; i++) {
    var fila        = dataMLC[i];
    var estadoActual = mlc_toStr_(fila[COL.estadoConciliacion]).trim();

    if (!ESTADOS_A_PROCESAR[estadoActual]) continue;

    var mlOrderId = mlc_toStr_(fila[COL.mlOrderId]).trim();
    var mlPackId  = mlc_toStr_(fila[COL.mlPackId]).trim();
    var skuCanal  = mlc_toStr_(fila[COL.skuCanal]).trim();

    // --- Buscar en hoja Ventas (jerarquía de 4 niveles) ---
    var filaVenta = mapaVentas[mlOrderId]
      || (mlPackId && skuCanal ? mapaVentasCompuesto[mlPackId + '|' + skuCanal] : null)
      || (mlPackId ? mapaVentas[mlPackId] : null)
      || null;

    var nuevoEstado  = '';
    var nuevoMensaje = '';

    if (!filaVenta) {
      nuevoEstado  = 'NO ENCONTRADA EN VENTAS';
      nuevoMensaje = '';
    } else {
      // Normalizar valores MLC (enteros)
      var mlMontoTotal  = Number(fila[COL.montoTotal])  || 0;
      var mlCargoVenta  = Number(fila[COL.cargoVenta])  || 0;
      var mlCargoEnvio  = Number(fila[COL.cargoEnvio])  || 0;
      var mlMontoNeto   = Number(fila[COL.montoNeto])   || 0;

      // Normalizar valores Ventas (formato "1.234" → 1234)
      var vMontoTotal  = mlc_parseChileanNumber_(filaVenta[COLV.precioVenta]);
      var vCargoVenta  = mlc_parseChileanNumber_(filaVenta[COLV.cargosVenta]);
      var vCargoEnvio  = mlc_parseChileanNumber_(filaVenta[COLV.cargoEnvio]);
      var vRecaudado   = mlc_parseChileanNumber_(filaVenta[COLV.recaudado]);

      var diferencias = [];

      if (Math.abs(mlMontoTotal - vMontoTotal) > TOLERANCIA)
        diferencias.push('monto_total: ' + mlMontoTotal + ' vs ' + vMontoTotal);
      if (Math.abs(mlCargoVenta - vCargoVenta) > TOLERANCIA)
        diferencias.push('cargo_venta: ' + mlCargoVenta + ' vs ' + vCargoVenta);
      if (Math.abs(mlCargoEnvio - vCargoEnvio) > TOLERANCIA)
        diferencias.push('cargo_envio: ' + mlCargoEnvio + ' vs ' + vCargoEnvio);
      if (Math.abs(mlMontoNeto - vRecaudado) > TOLERANCIA)
        diferencias.push('monto_neto: ' + mlMontoNeto + ' vs ' + vRecaudado);

      if (diferencias.length === 0) {
        nuevoEstado  = 'CONCILIADA';
        nuevoMensaje = 'OK';
      } else {
        nuevoEstado  = 'DIFERENCIA';
        nuevoMensaje = diferencias.join(' | ');
      }
    }

    // Escribir resultado en la fila
    dataMLC[i][COL.estadoConciliacion]  = nuevoEstado;
    dataMLC[i][COL.fechaConciliacion]   = fechaHoy;
    dataMLC[i][COL.mensajeConciliacion] = nuevoMensaje;
    cambios++;
  }

  // --- Escribir cambios de vuelta a la hoja ---
  if (cambios > 0) {
    sheetMLC.getRange(2, 1, dataMLC.length, dataMLC[0].length).setValues(dataMLC);
    Logger.log('[MLC conciliar] Filas procesadas: ' + cambios);
  } else {
    Logger.log('[MLC conciliar] Sin filas para conciliar');
  }

  SpreadsheetApp.getActiveSpreadsheet().toast('Conciliación completada: ' + cambios + ' filas procesadas.', 'EHI - Conciliación ML');
}

// -------------------------------------------------------------------
// Convierte string chileno "1.234" o número a entero
// -------------------------------------------------------------------
function mlc_parseChileanNumber_(val) {
  if (val === null || val === undefined || val === '') return 0;
  if (typeof val === 'number') return Math.round(val);
  return parseInt(String(val).replace(/\./g, '').replace(/,.*$/, ''), 10) || 0;
}