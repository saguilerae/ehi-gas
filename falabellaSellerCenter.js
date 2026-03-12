// ===================================================================
// Falabella Seller Center
// Ajustes incluidos:
// - Lectura de credenciales desde Script Properties
// - OperatorCode parametrizado
// - getSellercenterApiResponse() corregido con headers
// - payload solo en POST
// ===================================================================
// ── Variables Globales Falabella ────────────────────────────────
const UserAgent      = "SC1DB59/Java/1.8/PROPIA/FACL";
const ScApiHost      = "https://sellercenter-api.falabella.com/";
const HASH_ALGORITHM = "HmacSHA256";
const CHAR_UTF_8     = "UTF-8";
const CHAR_ASCII     = "ASCII";
// ===================================================================

// Función: consultaProdsFS
// Objetivo: consultar los productos del Falabella Seller Center
function consultaProdsFS() {
  const cfg = getFalabellaConfig_();

  const params = {
    UserID: cfg.userId,
    Timestamp: getCurrentTimestamp(),
    Version: "1.0",
    Action: "GetProducts",
    Filter: "all",
    Format: "XML"
  };

  const xmlString = getSellercenterApiResponse(params, cfg.apiKey, "", "get");

  const doc  = XmlService.parse(xmlString);
  const root = doc.getRootElement();
  const body = safeChild(root, 'Body');
  const productsNode = safeChild(body, 'Products');
  const products = productsNode ? productsNode.getChildren('Product') : [];

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Prods. FSC") || ss.insertSheet("Prods. FSC");

  const rows = [];
  const NUM_COLS = 17; // columnas B..R

  for (var i = 0; i < products.length; i++) {
    const p = products[i];
    const businessUnits = safeChild(p, 'BusinessUnits');
    const businessUnit  = businessUnits ? safeChild(businessUnits, 'BusinessUnit') : null;
    const productData   = safeChild(p, 'ProductData');

    const v  = (node, tag) => getElementValue(node, tag) || "";
    const vn = (node, tag) => toNumber(getElementValue(node, tag));

    const row = new Array(NUM_COLS);
    row[0]  = v(p, 'SellerSku');
    row[1]  = v(p, 'ProductId');
    row[2]  = v(p, 'ShopSku');
    row[3]  = v(p, 'Brand');
    row[4]  = v(p, 'Name');
    row[5]  = v(p, 'ColorBasico') || v(p, 'Variation');

    row[6]  = vn(businessUnit, 'Price');
    row[7]  = vn(businessUnit, 'SpecialPrice');
    row[8]  = v(businessUnit, 'SpecialFromDate');
    row[9]  = v(businessUnit, 'SpecialToDate');
    row[10] = vn(businessUnit, 'Stock');
    row[11] = v(businessUnit, 'Status');
    row[12] = v(businessUnit, 'IsPublished');

    row[13] = vn(productData, 'PackageWidth');
    row[14] = vn(productData, 'PackageLength');
    row[15] = vn(productData, 'PackageHeight');
    row[16] = vn(productData, 'PackageWeight');

    rows.push(row);
  }

  // Orden: Marca (col 3) y luego SellerSku (col 0)
  rows.sort(function(a, b) {
    const byBrand = str(a[3]).localeCompare(str(b[3]));
    return byBrand !== 0 ? byBrand : str(a[0]).localeCompare(str(b[0]));
  });

  // Limpieza previa: datos (B..R) + checkboxes antiguos (A)
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 2, lastRow - 1, NUM_COLS).clearContent();
    sheet.getRange(2, 1, lastRow - 1, 1).clearContent().clearDataValidations();
  }

  // Escritura de nuevos datos desde B
  if (rows.length > 0) {
    sheet.getRange(2, 2, rows.length, NUM_COLS).setValues(rows);

    const checkRange = sheet.getRange(2, 1, rows.length, 1);
    const rule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
    checkRange.setDataValidation(rule);
    checkRange.setValue(false);
  }

  ss.toast('Finalizó la consulta de productos (' + rows.length + ').', 'Éxito', 3);

  function safeChild(node, name) {
    if (!node) return null;
    const child = node.getChild(name);
    return child || null;
  }

  function str(x) {
    return (x == null ? '' : String(x));
  }

  function toNumber(x) {
    if (x == null || x === '') return 0;
    const n = Number(String(x).replace(',', '.'));
    return isNaN(n) ? 0 : n;
  }
}
// Fin de la Función: consultaProdsFS


// Objetivo: Actualizar el stock de uno o más productos (batch)
function actualizarStockFS() {
  const cfg = getFalabellaConfig_();
  const { filas } = getFilasSeleccionadasFS_();

  const productos = [];
  for (var i = 0; i < filas.length; i++) {
    if (filas[i][0] === true) {       // checkbox en col A
      const sku = filas[i][1];        // SellerSku en col B
      const stock = filas[i][11];     // Stock en col L
      if (!sku) continue;
      productos.push({ sku: sku, valor: stock });
    }
  }

  if (productos.length === 0) {
    SpreadsheetApp.getActive().toast('No hay productos seleccionados.', 'Info', 3);
    return;
  }

  let body = '';
  productos.forEach(function(p) {
    body += `
      <Product>
        <SellerSku>${xmlEscape_(p.sku)}</SellerSku>
        <BusinessUnits>
          <BusinessUnit>
            <OperatorCode>${xmlEscape_(cfg.operatorCode)}</OperatorCode>
            <Stock>${xmlEscape_(p.valor)}</Stock>
          </BusinessUnit>
        </BusinessUnits>
      </Product>`;
  });

  const xml = `<?xml version="1.0" encoding="UTF-8"?>
    <Request>${body}
    </Request>`;

  actualizaProdFS(xml);
}


// Objetivo: Actualizar el estado de uno o más productos (batch)
function actualizarStatusFS() {
  const cfg = getFalabellaConfig_();
  const { filas } = getFilasSeleccionadasFS_();

  const productos = [];
  for (var i = 0; i < filas.length; i++) {
    if (filas[i][0] === true) {
      const sku = filas[i][1];        // SellerSku
      const status = filas[i][12];    // Status en col M
      if (!sku) continue;
      productos.push({ sku: sku, valor: status });
    }
  }

  if (productos.length === 0) {
    SpreadsheetApp.getActive().toast('No hay productos seleccionados.', 'Info', 3);
    return;
  }

  let body = '';
  productos.forEach(function(p) {
    body += `
      <Product>
        <SellerSku>${xmlEscape_(p.sku)}</SellerSku>
        <BusinessUnits>
          <BusinessUnit>
            <OperatorCode>${xmlEscape_(cfg.operatorCode)}</OperatorCode>
            <Status>${xmlEscape_(p.valor)}</Status>
          </BusinessUnit>
        </BusinessUnits>
      </Product>`;
  });

  const xml = `<?xml version="1.0" encoding="UTF-8"?>
    <Request>${body}
    </Request>`;

  actualizaProdFS(xml);
}


// Objetivo: Actualizar el precio de uno o más productos (batch)
function actualizarPreciosFS() {
  const cfg = getFalabellaConfig_();
  const { filas } = getFilasSeleccionadasFS_();

  const productos = [];
  for (var i = 0; i < filas.length; i++) {
    if (filas[i][0] === true) {
      const sku = filas[i][1];        // SellerSku
      const precio = filas[i][7];     // Price en col H
      if (!sku) continue;
      productos.push({ sku: sku, valor: precio });
    }
  }

  if (productos.length === 0) {
    SpreadsheetApp.getActive().toast('No hay productos seleccionados.', 'Info', 3);
    return;
  }

  let body = '';
  productos.forEach(function(p) {
    body += `
      <Product>
        <SellerSku>${xmlEscape_(p.sku)}</SellerSku>
        <BusinessUnits>
          <BusinessUnit>
            <OperatorCode>${xmlEscape_(cfg.operatorCode)}</OperatorCode>
            <Price>${xmlEscape_(p.valor)}</Price>
          </BusinessUnit>
        </BusinessUnits>
      </Product>`;
  });

  const xml = `<?xml version="1.0" encoding="UTF-8"?>
    <Request>${body}
    </Request>`;

  actualizaProdFS(xml);
}


// Objetivo: Actualizar precio oferta + fechas de uno o más productos (batch)
function actualizarPreciosOfertaFS() {
  const cfg = getFalabellaConfig_();
  const { filas } = getFilasSeleccionadasFS_();

  const productos = [];
  for (var i = 0; i < filas.length; i++) {
    if (filas[i][0] === true) {
      const sku    = filas[i][1];   // SellerSku
      const precio = filas[i][7];   // Price
      const oferta = filas[i][8];   // SpecialPrice
      let fini     = filas[i][9];   // From
      let ffin     = filas[i][10];  // To

      if (!sku) continue;

      fini = formateaFechaFs_(fini);
      ffin = formateaFechaFs_(ffin);

      productos.push({
        sku: sku,
        valor: precio,
        valor2: oferta,
        fini: fini,
        ffin: ffin
      });
    }
  }

  if (productos.length === 0) {
    SpreadsheetApp.getActive().toast('No hay productos seleccionados.', 'Info', 3);
    return;
  }

  let body = '';
  productos.forEach(function(p) {
    body += `
      <Product>
        <SellerSku>${xmlEscape_(p.sku)}</SellerSku>
        <BusinessUnits>
          <BusinessUnit>
            <OperatorCode>${xmlEscape_(cfg.operatorCode)}</OperatorCode>
            <Price>${xmlEscape_(p.valor)}</Price>
            <SpecialPrice>${xmlEscape_(p.valor2)}</SpecialPrice>
            <SpecialFromDate>${xmlEscape_(p.fini)}</SpecialFromDate>
            <SpecialToDate>${xmlEscape_(p.ffin)}</SpecialToDate>
          </BusinessUnit>
        </BusinessUnits>
      </Product>`;
  });

  const xml = `<?xml version="1.0" encoding="UTF-8"?>
    <Request>${body}
    </Request>`;

  actualizaProdFS(xml);
}


// Helper fecha
function formateaFechaFs_(val) {
  if (!val) return "";
  if (Object.prototype.toString.call(val) === "[object Date]") {
    return Utilities.formatDate(val, 'UTC-4', 'yyyy-MM-dd');
  }
  return String(val);
}


// Objetivo: Actualizar Precio, Precio oferta y fecha inicio oferta, fecha fin oferta, stock Falabella Seller Center
function actualizaProdFS(XML) {
  const cfg = getFalabellaConfig_();

  const params = {
    UserID: cfg.userId,
    Timestamp: getCurrentTimestamp(),
    Version: "1.0",
    Action: "ProductUpdate"
  };

  const responce = getSellercenterApiResponse(params, cfg.apiKey, XML, 'post');
  const resultado = controlDeRetorno(responce);

  if (resultado === 'SuccessResponse') {
    SpreadsheetApp.getActiveSpreadsheet().toast('Actualización finalizada correctamente.', 'Éxito', 3);
  } else {
    SpreadsheetApp.getActiveSpreadsheet().toast("Se produjo un error: " + resultado, 'Error!', 5);
  }
}


// Objetivo: recuperar los datos desde el XML
function getElementValue(parent, tagName) {
  if (!parent) return '';
  const element = parent.getChild(tagName);
  return element ? element.getText() : '';
}


// Function to calculate the signature and send the request
function getSellercenterApiResponse(params, apiKey, XML, metodo) {
  const sortedParams = Object.keys(params).sort().reduce(function(acc, key) {
    acc[key] = params[key];
    return acc;
  }, {});

  let queryString = toQueryString(sortedParams);
  const signature = hmacDigest(queryString, apiKey, HASH_ALGORITHM);
  queryString += `&Signature=${encodeURIComponent(signature)}`;

  const requestUrl = `${ScApiHost}?${queryString}`;

  const options = {
    method: metodo,
    muteHttpExceptions: true
  };

  if (typeof UserAgent !== 'undefined' && UserAgent) {
    options.headers = {
      'User-Agent': UserAgent
    };
  }

  if (String(metodo).toLowerCase() === 'post') {
    options.contentType = 'application/x-www-form-urlencoded';
    options.payload = XML || '';
  }

  const response = UrlFetchApp.fetch(requestUrl, options);
  return response.getContentText();
}


// Function to generate HMAC digest
function hmacDigest(msg, keyString, algo) {
  const key = Utilities.computeHmacSignature(
    Utilities.MacAlgorithm.HMAC_SHA_256,
    msg,
    keyString,
    Utilities.Charset.UTF_8
  );
  return key.map(function(byte) {
    return ('0' + (byte & 0xFF).toString(16)).slice(-2);
  }).join('');
}


// Function to build query string out of params map
function toQueryString(data) {
  return Object.keys(data).map(function(key) {
    return encodeURIComponent(key) + '=' + encodeURIComponent(data[key]);
  }).join('&');
}


// Function to return the current timestamp in ISO 8601 format
function getCurrentTimestamp() {
  const now = new Date();
  return Utilities.formatDate(now, 'UTC-4', "yyyy-MM-dd'T'HH:mm:ss'Z'");
}


// Objectivo: manejar los responces de los request
function controlDeRetorno(resp) {
  var resutadoEjec = "SuccessResponse";
  const doc = XmlService.parse(resp);
  const root = doc.getRootElement();
  if (root.getName() == "ErrorResponse") {
    var ret = root.getChild('Head');
    resutadoEjec = getElementValue(ret, 'ErrorMessage');
  }
  return resutadoEjec;
}


function getFilasSeleccionadasFS_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Prods. FSC") || ss.insertSheet("Prods. FSC");
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow <= 1) return { sheet: sheet, filas: [] };

  const filas = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  return { sheet: sheet, filas: filas };
}


// Objetivo: generar los XML de FS
function generaXmlFS(sku, valor, valor2, fini, ffin, queXML) {
  const cfg = getFalabellaConfig_();
  const operatorCode = xmlEscape_(cfg.operatorCode);

  var xml = ``;

  switch (queXML) {
    case "precio":
      xml = `<?xml version="1.0" encoding="UTF-8"?>
               <Request>
                  <Product>
                     <SellerSku>${xmlEscape_(sku)}</SellerSku>
                     <BusinessUnits>
                        <BusinessUnit>
                           <OperatorCode>${operatorCode}</OperatorCode>
                           <Price>${xmlEscape_(valor)}</Price>
                        </BusinessUnit>
                     </BusinessUnits>
                   </Product>
               </Request>`;
      break;

    case "stock":
      xml = `<?xml version="1.0" encoding="UTF-8"?>
               <Request>
                  <Product>
                     <SellerSku>${xmlEscape_(sku)}</SellerSku>
                     <BusinessUnits>
                        <BusinessUnit>
                           <OperatorCode>${operatorCode}</OperatorCode>
                           <Stock>${xmlEscape_(valor)}</Stock>
                        </BusinessUnit>
                     </BusinessUnits>
                   </Product>
               </Request>`;
      break;

    case "status":
      xml = `<?xml version="1.0" encoding="UTF-8"?>
               <Request>
                  <Product>
                     <SellerSku>${xmlEscape_(sku)}</SellerSku>
                     <BusinessUnits>
                        <BusinessUnit>
                           <OperatorCode>${operatorCode}</OperatorCode>
                           <Status>${xmlEscape_(valor)}</Status>
                        </BusinessUnit>
                     </BusinessUnits>
                   </Product>
               </Request>`;
      break;

    case "preciooferta":
      xml = `<?xml version="1.0" encoding="UTF-8"?>
             <Request>
               <Product>
                  <SellerSku>${xmlEscape_(sku)}</SellerSku>
                  <BusinessUnits>
                    <BusinessUnit>
                        <OperatorCode>${operatorCode}</OperatorCode>
                        <Price>${xmlEscape_(valor)}</Price>
                        <SpecialPrice>${xmlEscape_(valor2)}</SpecialPrice>
                        <SpecialFromDate>${xmlEscape_(fini)}</SpecialFromDate>
                        <SpecialToDate>${xmlEscape_(ffin)}</SpecialToDate>
                    </BusinessUnit>
                  </BusinessUnits>
                </Product>
             </Request>`;
      break;

    default:
      xml = "otro";
  }

  return xml;
}


// Config Falabella desde Script Properties
function getFalabellaConfig_() {
  const props = PropertiesService.getScriptProperties();

  const cfg = {
    userId: props.getProperty('FALABELLA_USER_ID'),
    apiKey: props.getProperty('FALABELLA_API_KEY'),
    operatorCode: props.getProperty('FALABELLA_OPERATOR_CODE') || 'facl'
  };

  if (!cfg.userId) {
    throw new Error('Falta Script Property: FALABELLA_USER_ID');
  }

  if (!cfg.apiKey) {
    throw new Error('Falta Script Property: FALABELLA_API_KEY');
  }

  return cfg;
}


// Escape básico XML
function xmlEscape_(value) {
  if (value === null || value === undefined) return '';
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}


// ===================================================================
// VENTAS FSC — Ingesta de Órdenes Falabella Seller Center
// Etapa 1: Ingesta (Histórica e Incremental)
// ===================================================================

// Constantes del módulo
const FSC_SHEET_VENTAS    = 'Ventas FSC';
const FSC_COMISION_RATE   = 0.20;
const FSC_PAGE_LIMIT      = 100;
const FSC_ITEMS_CHUNK     = 20;
const FSC_PACING_MS       = 500;
const FSC_LOG_HIST        = '[FSC-HIST]';
const FSC_LOG_INC         = '[FSC-INC]';

// Headers exactos de la hoja — 30 columnas
const FSC_HEADERS = [
  'fecha_ingesta','canal','fsc_order_id','fsc_order_number',
  'fsc_order_item_id','fsc_package_id','id_venta_canal','fecha_venta',
  'sku_canal','sku_maestro','mpn','cantidad',
  'precio_unitario','monto_total','cargo_venta','cargo_envio',
  'shipping_type','monto_neto','cliente_nombre','region',
  'comuna','direccion_resumen','estado_fsc','estado_conciliacion',
  'fecha_conciliacion','mensaje_conciliacion','clave_unica',
  'existe_en_ventas','json_raw','observaciones'
];

// Estados de conciliación válidos (referencia)
// PENDIENTE | PENDIENTE_HISTORICO | CANCELADA POR EL COMPRADOR |
// DEVOLUCION CON REEMBOLSO | CONCILIADA | DIFERENCIA |
// DIFERENCIA_CONOCIDA | NO ENCONTRADA EN VENTAS


// ── Función pública menú: Ingesta Histórica ─────────────────────
function fsc_ingestaHistorica() {
  fsc_ingestarVentasFSC_({ modo: 'HISTORICO', actualizarLastSync: false });
}

// ── Función pública menú: Ingesta Incremental ───────────────────
function fsc_ingestaIncremental() {
  fsc_ingestarVentasFSC_({ modo: 'INCREMENTAL', actualizarLastSync: true });
}


// ── Orquestador principal ───────────────────────────────────────
function fsc_ingestarVentasFSC_(opts) {
  opts = opts || {};
  const modo               = String(opts.modo || 'HISTORICO').toUpperCase();
  const actualizarLastSync = !!opts.actualizarLastSync;
  const logPrefix          = modo === 'HISTORICO' ? FSC_LOG_HIST : FSC_LOG_INC;

  const props   = PropertiesService.getScriptProperties();
  const cfg     = getFalabellaConfig_();
  const lastSync = String(props.getProperty('FALABELLA_LAST_SYNC') || '').trim();

  // Resolver fecha de búsqueda
  let searchFrom = '';
  if (modo === 'HISTORICO') {
    searchFrom = lastSync;
    if (!searchFrom) {
      searchFrom = Browser.inputBox(
        'Ingesta Histórica FSC',
        'Ingresa fecha de inicio (YYYY-MM-DD):',
        Browser.Buttons.OK_CANCEL
      );
      if (!searchFrom || searchFrom === 'cancel') {
        SpreadsheetApp.getActive().toast('Ingesta cancelada.', 'FSC', 3);
        return;
      }
    }
  } else {
    if (!lastSync) throw new Error('Falta FALABELLA_LAST_SYNC en Script Properties.');
    searchFrom = lastSync;
  }

  // Obtener / crear hoja
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh   = ss.getSheetByName(FSC_SHEET_VENTAS);
  if (!sh) sh = fsc_crearHojaVentasFSC_();

  const headerMap    = fsc_getHeaderMap_(sh);
  const existingKeys = fsc_getExistingKeys_(sh, headerMap['clave_unica']);

  let offset              = 0;
  let totalOrders         = 0;
  let totalInserted       = 0;
  let shouldContinue      = true;

  SpreadsheetApp.getActive().toast('Ingesta FSC iniciada (' + modo + ')', 'EHI', 5);
  
  // Normalizar formato fecha para API FSC
  searchFrom = fsc_normalizeDate_(searchFrom);

  while (shouldContinue) {
    // 1. Traer página de órdenes
    const orders = fsc_fetchOrders_(searchFrom, modo, offset);
    if (!orders.length) break;
    totalOrders += orders.length;

    // 2. Extraer OrderIds del batch
    const orderIds = orders.map(function(o) { return String(o.OrderId); });

    // 3. Traer items para todos los OrderIds del batch
    const itemsMap = fsc_fetchMultipleOrderItems_(orderIds, cfg);

    // 4. Construir filas
    const rowsToAppend = [];
    for (var i = 0; i < orders.length; i++) {
      const order = orders[i];
      const orderId = String(order.OrderId);
      const items   = itemsMap[orderId] || [];

      for (var j = 0; j < items.length; j++) {
        const clave = orderId + '-' + String(items[j].OrderItemId);
        if (existingKeys.has(clave)) continue;        // evitar duplicados
        const row = fsc_buildRow_(order, items[j], modo);
        rowsToAppend.push(row);
        existingKeys.add(clave);
      }
    }

    // 5. Escribir batch en la hoja
    if (rowsToAppend.length) {
      const startRow = Math.max(sh.getLastRow() + 1, 2);
      sh.getRange(startRow, 1, rowsToAppend.length, FSC_HEADERS.length)
        .setValues(rowsToAppend);
      totalInserted += rowsToAppend.length;
    }

    // 6. Paginación
    if (orders.length < FSC_PAGE_LIMIT) {
      shouldContinue = false;
    } else {
      offset += FSC_PAGE_LIMIT;
      Utilities.sleep(FSC_PACING_MS);
    }
  }

  // 7. Actualizar LAST_SYNC solo en incremental
  if (actualizarLastSync) {
    props.setProperty('FALABELLA_LAST_SYNC', new Date().toISOString());
  }

  const resumen = {
    modo: modo, searchFrom: searchFrom,
    fetchedOrders: totalOrders, insertedRows: totalInserted
  };
  Logger.log(logPrefix + ' Resumen: ' + JSON.stringify(resumen));
  SpreadsheetApp.getActive().toast(
    'FSC ' + modo + ': órdenes=' + totalOrders + ' | filas nuevas=' + totalInserted,
    'EHI', 7
  );
  return resumen;
}


// ── Fetch órdenes paginadas ─────────────────────────────────────
function fsc_fetchOrders_(searchFrom, modo, offset) {
  try {
    const cfg = getFalabellaConfig_();
    const params = {
      Action:       'GetOrders',
      Format:       'JSON',
      Version:      '1.0',
      UserID:       cfg.userId,
      Timestamp:    getCurrentTimestamp(),
      Limit:        String(FSC_PAGE_LIMIT),
      Offset:       String(offset),
      SortDirection:'ASC'
    };

    // HISTORICO usa CreatedAfter, INCREMENTAL usa UpdatedAfter
    if (modo === 'HISTORICO') {
      params.CreatedAfter = searchFrom;
    } else {
      params.UpdatedAfter = searchFrom;
    }

    const raw  = getSellercenterApiResponse(params, cfg.apiKey, '', 'get');
    const resp = JSON.parse(raw);

    if (resp.ErrorResponse) {
      Logger.log('[FSC] Error GetOrders: ' + JSON.stringify(resp.ErrorResponse));
      return [];
    }

    let orders = [];
    try {
      orders = resp.SuccessResponse.Body.Orders.Order || [];
    } catch(e) {
      return [];
    }

    // Normalizar siempre a array
    if (!Array.isArray(orders)) orders = [orders];
    return orders;

  } catch(e) {
    Logger.log('[FSC] fsc_fetchOrders_ excepción: ' + e.message);
    return [];
  }
}


// ── Fetch items para múltiples órdenes en chunks ────────────────
// ── Fetch items para múltiples órdenes en chunks ────────────────
function fsc_fetchMultipleOrderItems_(orderIdList, cfg) {
  cfg = cfg || getFalabellaConfig_();
  const result = {};

  for (var i = 0; i < orderIdList.length; i += FSC_ITEMS_CHUNK) {
    const chunk = orderIdList.slice(i, i + FSC_ITEMS_CHUNK);
    try {
      // Construir params SIN OrderIdList — no debe entrar en la firma
      const params = {
        Action:    'GetMultipleOrderItems',
        Format:    'JSON',
        Version:   '1.0',
        UserID:    cfg.userId,
        Timestamp: getCurrentTimestamp()
      };

      // Construir firma manualmente (OrderIdList se agrega DESPUÉS de la firma)
      const sortedParams = Object.keys(params).sort().reduce(function(acc, key) {
        acc[key] = params[key];
        return acc;
      }, {});
      const qs        = toQueryString(sortedParams);
      const signature = hmacDigest(qs, cfg.apiKey, HASH_ALGORITHM);
      const finalUrl  = ScApiHost + '?' + qs
                      + '&Signature='   + encodeURIComponent(signature)
                      + '&OrderIdList=' + encodeURIComponent('[' + chunk.join(',') + ']');

      const response = UrlFetchApp.fetch(finalUrl, {
        method:             'get',
        muteHttpExceptions: true,
        headers:            { 'User-Agent': UserAgent }
      });
      const raw  = response.getContentText();
      const resp = JSON.parse(raw);

      if (resp.ErrorResponse) {
        Logger.log('[FSC] Error GetMultipleOrderItems chunk: ' + JSON.stringify(resp.ErrorResponse));
        continue;
      }

      let ordersResp = [];
      try {
        ordersResp = resp.SuccessResponse.Body.Orders.Order || [];
      } catch(e) { continue; }

      if (!Array.isArray(ordersResp)) ordersResp = [ordersResp];

      ordersResp.forEach(function(o) {
        const oid = String(o.OrderId);
        let items = [];
        try {
          items = o.OrderItems.OrderItem || [];
        } catch(e) { items = []; }
        if (!Array.isArray(items)) items = [items];
        result[oid] = items;
      });

    } catch(e) {
      Logger.log('[FSC] fsc_fetchMultipleOrderItems_ chunk excepción: ' + e.message);
    }

    if (i + FSC_ITEMS_CHUNK < orderIdList.length) {
      Utilities.sleep(FSC_PACING_MS);
    }
  }
  return result;
}

// ── Construir una fila de 30 columnas ──────────────────────────
function fsc_buildRow_(order, item, modo) {
  const orderId     = String(order.OrderId    || '');
  const orderItemId = String(item.OrderItemId || '');
  const paidPrice   = parseFloat(item.PaidPrice      || 0);
  const shippingAmt = parseFloat(item.ShippingAmount || 0);
  const cargoVenta  = parseFloat((paidPrice * FSC_COMISION_RATE).toFixed(0));
  const montoNeto   = parseFloat((paidPrice - cargoVenta - shippingAmt).toFixed(0));
  const sku         = String(item.Sku || '');
  const mpn         = sku.length > 2 ? sku.slice(0, -2) : sku;
  const shippingType = String(item.ShippingType || '');
  const isFBF       = shippingType === 'Own Warehouse';

  // Estado conciliacion
  let estadoConciliacion = '';
  const estadoFsc = String(item.Status || '').toLowerCase();
  if (estadoFsc === 'canceled') {
    estadoConciliacion = 'CANCELADA POR EL COMPRADOR';
  } else if (modo === 'HISTORICO') {
    estadoConciliacion = 'PENDIENTE_HISTORICO';
  } else {
    estadoConciliacion = 'PENDIENTE';
  }

  return [
    new Date(),                                           // A fecha_ingesta
    'FS',                                                 // B canal
    orderId,                                              // C fsc_order_id
    String(order.OrderNumber    || ''),                   // D fsc_order_number
    orderItemId,                                          // E fsc_order_item_id
    String(item.PackageId       || ''),                   // F fsc_package_id
    orderId + '-' + orderItemId,                          // G id_venta_canal
    String(item.CreatedAt       || ''),                   // H fecha_venta
    String(item.ShopSku         || ''),                   // I sku_canal
    sku,                                                  // J sku_maestro
    mpn,                                                  // K mpn
    1,                                                    // L cantidad
    parseFloat(item.ItemPrice   || 0),                    // M precio_unitario
    paidPrice,                                            // N monto_total
    cargoVenta,                                           // O cargo_venta
    shippingAmt,                                          // P cargo_envio
    shippingType,                                         // Q shipping_type
    montoNeto,                                            // R monto_neto
    String((order.CustomerFirstName || '') + ' ' +
           (order.CustomerLastName  || '')).trim(),       // S cliente_nombre
    fsc_safeAddr_(order, 'State'),                        // T region
    fsc_safeAddr_(order, 'City'),                         // U comuna
    fsc_safeAddr_(order, 'Address1'),                     // V direccion_resumen
    String(item.Status          || ''),                   // W estado_fsc
    estadoConciliacion,                                   // X estado_conciliacion
    '',                                                   // Y fecha_conciliacion
    '',                                                   // Z mensaje_conciliacion
    orderId + '-' + orderItemId,                          // AA clave_unica
    false,                                                // AB existe_en_ventas
    JSON.stringify({ order_id: orderId, item: item }),    // AC json_raw
    isFBF ? 'FBF - no consume stock' : ''                // AD observaciones
  ];
}


// ── Helper dirección segura ─────────────────────────────────────
function fsc_safeAddr_(order, field) {
  try {
    return String(order.AddressShipping[field] || '');
  } catch(e) {
    return '';
  }
}


// ── Helper: mapa de headers ─────────────────────────────────────
function fsc_getHeaderMap_(sh) {
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach(function(h, i) { if (h) map[String(h).trim()] = i; });
  return map;
}


// ── Helper: claves únicas existentes ───────────────────────────
function fsc_getExistingKeys_(sh, colIndex) {
  const set = new Set();
  const lastRow = sh.getLastRow();
  if (lastRow < 2 || colIndex === undefined) return set;
  const vals = sh.getRange(2, colIndex + 1, lastRow - 1, 1).getValues();
  vals.forEach(function(r) { if (r[0]) set.add(String(r[0])); });
  return set;
}


// ── Crear hoja Ventas FSC si no existe ─────────────────────────
function fsc_crearHojaVentasFSC_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh   = ss.getSheetByName(FSC_SHEET_VENTAS);
  if (sh) return sh;

  sh = ss.insertSheet(FSC_SHEET_VENTAS);
  sh.getRange(1, 1, 1, FSC_HEADERS.length).setValues([FSC_HEADERS]);
  sh.setFrozenRows(1);

  // Formato fecha cols A(1) y H(8)
  sh.getRange(2, 1, sh.getMaxRows() - 1, 1)
    .setNumberFormat('dd/mm/yyyy hh:mm:ss');
  sh.getRange(2, 8, sh.getMaxRows() - 1, 1)
    .setNumberFormat('dd/mm/yyyy hh:mm:ss');

  // Formato número cols M(13),N(14),O(15),P(16),R(18)
  [13, 14, 15, 16, 18].forEach(function(c) {
    sh.getRange(2, c, sh.getMaxRows() - 1, 1)
      .setNumberFormat('#,##0');
  });

  return sh;
}

// ── Normaliza fecha ISO para API FSC ────────────────────────────
// Entrada:  2024-01-01T00:00:00.000-03:00
// Salida:   2024-01-01T00:00:00-0300
function fsc_normalizeDate_(isoStr) {
  return String(isoStr)
    .replace(/\.\d+/, '')                          // elimina milisegundos
    .replace(/([+-]\d{2}):(\d{2})$/, '$1$2');      // -03:00 → -0300
}