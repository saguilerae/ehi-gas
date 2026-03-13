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
// Ventanas de 30 días desde FALABELLA_HISTORICAL_FROM hasta hoy.
// Flush condicional: acumula en memoria y escribe en hoja solo al
// final o cuando elapsed > 5 min (antes del timeout de 6 min GAS).
function fsc_ingestaHistorica() {
  const props          = PropertiesService.getScriptProperties();
  const historicalFrom = String(props.getProperty('FALABELLA_HISTORICAL_FROM') || '').trim();

  if (!historicalFrom) {
    SpreadsheetApp.getActive().toast('Falta FALABELLA_HISTORICAL_FROM en Script Properties.', 'FSC', 5);
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh   = ss.getSheetByName(FSC_SHEET_VENTAS);
  if (!sh) sh = fsc_crearHojaVentasFSC_();

  const headerMap    = fsc_getHeaderMap_(sh);
  const existingKeys = fsc_getExistingKeys_(sh, headerMap['clave_unica']);
  const cfg          = getFalabellaConfig_();

  let desde          = new Date(historicalFrom);
  const hasta        = new Date();
  const VENTANA_DIAS = 30;
  const MAX_MS       = 5 * 60 * 1000; // 5 min — margen antes del timeout GAS (6 min)

  let totalOrders   = 0;
  let totalInserted = 0;
  let ventana       = 0;
  let allRows       = [];             // buffer en memoria
  const startTime   = Date.now();

  SpreadsheetApp.getActive().toast('Ingesta histórica FSC iniciada...', 'EHI', 5);

  // ── Helper flush ─────────────────────────────────────────────
  function flushRows_() {
    if (!allRows.length) return;
    const startRow = Math.max(sh.getLastRow() + 1, 2);
    sh.getRange(startRow, 1, allRows.length, FSC_HEADERS.length).setValues(allRows);
    totalInserted  += allRows.length;
    allRows.length  = 0;  // limpia sin reasignar
    Logger.log(FSC_LOG_HIST + ' Flush: totalInserted hasta ahora=' + totalInserted);
  }

  while (desde < hasta) {
    const fin     = new Date(desde);
    fin.setDate(fin.getDate() + VENTANA_DIAS);
    const finReal = fin > hasta ? hasta : fin;

    const updatedAfter  = fsc_normalizeDate_(desde.toISOString());
    const updatedBefore = fsc_normalizeDate_(finReal.toISOString());

    ventana++;
    Logger.log(FSC_LOG_HIST + ' Ventana ' + ventana + ': ' + updatedAfter + ' → ' + updatedBefore);

    // ── Toast de progreso ────────────────────────────────────────
    SpreadsheetApp.getActive().toast(
      'Ventana ' + ventana + ' | Órdenes acumuladas: ' + totalOrders + ' | Filas buffer: ' + allRows.length,
      'FSC Histórico ⏳', 8
    );

    let offset         = 0;
    let shouldContinue = true;

    while (shouldContinue) {
      const orders = fsc_fetchOrdersVentana_(updatedAfter, updatedBefore, offset, cfg);
      if (!orders.length) break;
      totalOrders += orders.length;

      const orderIds = orders.map(function(o) { return String(o.OrderId); });
      const itemsMap = fsc_fetchOrderItems_(orderIds, cfg);

      for (var i = 0; i < orders.length; i++) {
        const order = orders[i];
        const oid   = String(order.OrderId);
        const items = itemsMap[oid] || [];
        for (var j = 0; j < items.length; j++) {
          const clave = oid + '-' + String(items[j].OrderItemId);
          if (existingKeys.has(clave)) continue;
          allRows.push(fsc_buildRow_(order, items[j], 'HISTORICO'));
          existingKeys.add(clave);
        }
      }

      if (orders.length < FSC_PAGE_LIMIT) {
        shouldContinue = false;
      } else {
        offset += FSC_PAGE_LIMIT;
        Utilities.sleep(FSC_PACING_MS);
      }
    }

    // ── Flush condicional al terminar cada ventana ────────────
    if (Date.now() - startTime > MAX_MS) {
      Logger.log(FSC_LOG_HIST + ' Timeout inminente — flush forzado en ventana ' + ventana);
      flushRows_();
    }

    desde = fin;
    Utilities.sleep(FSC_PACING_MS);
  }

  // ── Flush final con lo que quedó en buffer ────────────────────
  flushRows_();

  const resumen = { ventanas: ventana, fetchedOrders: totalOrders, insertedRows: totalInserted };
  Logger.log(FSC_LOG_HIST + ' Resumen: ' + JSON.stringify(resumen));
  SpreadsheetApp.getActive().toast(
    'FSC histórico: ventanas=' + ventana + ' | órdenes=' + totalOrders + ' | filas=' + totalInserted,
    'EHI', 10
  );
  return resumen;
}


// ── Función pública menú: Ingesta Incremental ───────────────────
// Ventanas de 7 días desde FALABELLA_LAST_SYNC hasta hoy.
// Flush condicional: misma estrategia que histórico.
function fsc_ingestaIncremental() {
  const props    = PropertiesService.getScriptProperties();
  const lastSync = String(props.getProperty('FALABELLA_LAST_SYNC') || '').trim();

  if (!lastSync) throw new Error('Falta FALABELLA_LAST_SYNC en Script Properties.');

  const cfg = getFalabellaConfig_();
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  let sh    = ss.getSheetByName(FSC_SHEET_VENTAS);
  if (!sh) sh = fsc_crearHojaVentasFSC_();

  const headerMap    = fsc_getHeaderMap_(sh);
  const existingKeys = fsc_getExistingKeys_(sh, headerMap['clave_unica']);

  let desde          = new Date(lastSync);
  const hasta        = new Date();
  const VENTANA_DIAS = 7;
  const MAX_MS       = 5 * 60 * 1000;

  let totalOrders   = 0;
  let totalInserted = 0;
  let ventana       = 0;
  let allRows       = [];
  const startTime   = Date.now();

  SpreadsheetApp.getActive().toast('Ingesta incremental FSC iniciada...', 'EHI', 5);

  function flushRows_() {
    if (!allRows.length) return;
    const startRow = Math.max(sh.getLastRow() + 1, 2);
    sh.getRange(startRow, 1, allRows.length, FSC_HEADERS.length).setValues(allRows);
    totalInserted  += allRows.length;
    allRows.length  = 0;
    Logger.log(FSC_LOG_INC + ' Flush: totalInserted hasta ahora=' + totalInserted);
  }

  while (desde < hasta) {
    const fin     = new Date(desde);
    fin.setDate(fin.getDate() + VENTANA_DIAS);
    const finReal = fin > hasta ? hasta : fin;

    const updatedAfter  = fsc_normalizeDate_(desde.toISOString());
    const updatedBefore = fsc_normalizeDate_(finReal.toISOString());

    ventana++;
    Logger.log(FSC_LOG_INC + ' Ventana ' + ventana + ': ' + updatedAfter + ' → ' + updatedBefore);
    SpreadsheetApp.getActive().toast(
      'Ventana ' + ventana + ' | Órdenes acumuladas: ' + totalOrders + ' | Filas buffer: ' + allRows.length,
      'FSC Incremental ⏳', 8
    );

    let offset         = 0;
    let shouldContinue = true;

    while (shouldContinue) {
      const orders = fsc_fetchOrdersVentana_(updatedAfter, updatedBefore, offset, cfg);
      if (!orders.length) break;
      totalOrders += orders.length;

      const orderIds = orders.map(function(o) { return String(o.OrderId); });
      const itemsMap = fsc_fetchOrderItems_(orderIds, cfg);

      for (var i = 0; i < orders.length; i++) {
        const order = orders[i];
        const oid   = String(order.OrderId);
        const items = itemsMap[oid] || [];
        for (var j = 0; j < items.length; j++) {
          const clave = oid + '-' + String(items[j].OrderItemId);
          if (existingKeys.has(clave)) continue;
          allRows.push(fsc_buildRow_(order, items[j], 'INCREMENTAL'));
          existingKeys.add(clave);
        }
      }

      if (orders.length < FSC_PAGE_LIMIT) {
        shouldContinue = false;
      } else {
        offset += FSC_PAGE_LIMIT;
        Utilities.sleep(FSC_PACING_MS);
      }
    }

    if (Date.now() - startTime > MAX_MS) {
      Logger.log(FSC_LOG_INC + ' Timeout inminente — flush forzado en ventana ' + ventana);
      flushRows_();
    }

    desde = fin;
    Utilities.sleep(FSC_PACING_MS);
  }

  flushRows_();

  // Actualizar LAST_SYNC al finalizar exitosamente
  props.setProperty('FALABELLA_LAST_SYNC', hasta.toISOString());

  const resumen = { ventanas: ventana, fetchedOrders: totalOrders, insertedRows: totalInserted };
  Logger.log(FSC_LOG_INC + ' Resumen: ' + JSON.stringify(resumen));
  SpreadsheetApp.getActive().toast(
    'FSC incremental: ventanas=' + ventana + ' | órdenes=' + totalOrders + ' | filas=' + totalInserted,
    'EHI', 7
  );
  return resumen;
}


// ── Fetch órdenes con ventana UpdatedAfter [+ UpdatedBefore] ────
function fsc_fetchOrdersVentana_(updatedAfter, updatedBefore, offset, cfg) {
  try {
    cfg = cfg || getFalabellaConfig_();
    const params = {
      Action:        'GetOrders',
      Format:        'JSON',
      Limit:         String(FSC_PAGE_LIMIT),
      Offset:        String(offset),
      SortDirection: 'ASC',
      Timestamp:     getCurrentTimestamp(),
      UpdatedAfter:  updatedAfter,
      UserID:        cfg.userId,
      Version:       '1.0'
    };

    // UpdatedBefore es opcional — solo se usa en histórico
    if (updatedBefore) params.UpdatedBefore = updatedBefore;

    const raw  = getSellercenterApiResponse(params, cfg.apiKey, '', 'get');
    const resp = JSON.parse(raw);

    if (resp.ErrorResponse) {
      Logger.log('[FSC] Error GetOrders: ' + JSON.stringify(resp.ErrorResponse));
      return [];
    }

    let orders = [];
    try { orders = resp.SuccessResponse.Body.Orders.Order || []; } catch(e) { return []; }
    if (!Array.isArray(orders)) orders = [orders];
    return orders;

  } catch(e) {
    Logger.log('[FSC] fsc_fetchOrdersVentana_ excepción: ' + e.message);
    return [];
  }
}


// ── Fetch items por orden — loop de GetOrderItems ───────────────
function fsc_fetchOrderItems_(orderIdList, cfg) {
  cfg = cfg || getFalabellaConfig_();
  const result = {};

  for (var i = 0; i < orderIdList.length; i++) {
    const orderId = orderIdList[i];
    try {
      const params = {
        Action:    'GetOrderItems',
        Format:    'JSON',
        OrderId:   String(orderId),
        Timestamp: getCurrentTimestamp(),
        UserID:    cfg.userId,
        Version:   '1.0'
      };

      const raw  = getSellercenterApiResponse(params, cfg.apiKey, '', 'get');
      const resp = JSON.parse(raw);

      if (resp.ErrorResponse) {
        Logger.log('[FSC] Error GetOrderItems [' + orderId + ']: ' + JSON.stringify(resp.ErrorResponse));
        continue;
      }

      let items = [];
      try { items = resp.SuccessResponse.Body.OrderItems.OrderItem || []; } catch(e) { items = []; }
      if (!Array.isArray(items)) items = [items];
      result[String(orderId)] = items;

    } catch(e) {
      Logger.log('[FSC] fsc_fetchOrderItems_ excepción [' + orderId + ']: ' + e.message);
    }

    Utilities.sleep(FSC_PACING_MS);
  }
  return result;
}


// ── Construir una fila de 30 columnas ──────────────────────────
function fsc_buildRow_(order, item, modo) {
  const orderId      = String(order.OrderId    || '');
  const orderItemId  = String(item.OrderItemId || '');
  const paidPrice    = parseFloat(item.PaidPrice      || 0);
  const shippingAmt  = parseFloat(item.ShippingAmount || 0);
  const cargoVenta   = parseFloat((paidPrice * FSC_COMISION_RATE).toFixed(0));
  const montoNeto    = parseFloat((paidPrice - cargoVenta - shippingAmt).toFixed(0));
  const sku          = String(item.Sku || '');
  const mpn          = sku.length > 2 ? sku.slice(0, -2) : sku;
  const shippingType = String(item.ShippingType || '');
  const isFBF        = shippingType === 'Own Warehouse';

  const estadoFsc = String(item.Status || '').toLowerCase();
  let estadoConciliacion = '';
  if (estadoFsc === 'canceled') {
    estadoConciliacion = 'CANCELADA POR EL COMPRADOR';
  } else if (estadoFsc === 'returned') {
    estadoConciliacion = 'DEVOLUCION CON REEMBOLSO';
  } else if (modo === 'HISTORICO') {
    estadoConciliacion = 'PENDIENTE_HISTORICO';
  } else {
    estadoConciliacion = 'PENDIENTE';
  }

  return [
    new Date(),                                            // A fecha_ingesta
    'FS',                                                  // B canal
    orderId,                                               // C fsc_order_id
    String(order.OrderNumber    || ''),                    // D fsc_order_number
    orderItemId,                                           // E fsc_order_item_id
    String(item.PackageId       || ''),                    // F fsc_package_id
    orderId + '-' + orderItemId,                           // G id_venta_canal
    String(item.CreatedAt       || ''),                    // H fecha_venta
    String(item.ShopSku         || ''),                    // I sku_canal
    sku,                                                   // J sku_maestro
    mpn,                                                   // K mpn
    1,                                                     // L cantidad
    parseFloat(item.ItemPrice   || 0),                     // M precio_unitario
    paidPrice,                                             // N monto_total
    cargoVenta,                                            // O cargo_venta
    shippingAmt,                                           // P cargo_envio
    shippingType,                                          // Q shipping_type
    montoNeto,                                             // R monto_neto
    String((order.CustomerFirstName || '') + ' ' +
           (order.CustomerLastName  || '')).trim(),        // S cliente_nombre
    fsc_safeAddr_(order, 'State'),                         // T region
    fsc_safeAddr_(order, 'City'),                          // U comuna
    fsc_safeAddr_(order, 'Address1'),                      // V direccion_resumen
    String(item.Status          || ''),                    // W estado_fsc
    estadoConciliacion,                                    // X estado_conciliacion
    '',                                                    // Y fecha_conciliacion
    '',                                                    // Z mensaje_conciliacion
    orderId + '-' + orderItemId,                           // AA clave_unica
    false,                                                 // AB existe_en_ventas
    JSON.stringify({ order_id: orderId, item: item }),     // AC json_raw
    isFBF ? 'FBF - no consume stock' : ''                 // AD observaciones
  ];
}


// ── Helper dirección segura ─────────────────────────────────────
function fsc_safeAddr_(order, field) {
  try {
    return String(order.AddressShipping[field] || '');
  } catch(e) { return ''; }
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

  sh.getRange(2, 1, sh.getMaxRows() - 1, 1).setNumberFormat('dd/mm/yyyy hh:mm:ss');
  sh.getRange(2, 8, sh.getMaxRows() - 1, 1).setNumberFormat('dd/mm/yyyy hh:mm:ss');
  [13, 14, 15, 16, 18].forEach(function(c) {
    sh.getRange(2, c, sh.getMaxRows() - 1, 1).setNumberFormat('#,##0');
  });

  return sh;
}


// ── Normaliza fecha ISO para API FSC ────────────────────────────
// Entrada:  2024-01-01T00:00:00.000-03:00
// Salida:   2024-01-01T00:00:00-0300
function fsc_normalizeDate_(isoStr) {
  return String(isoStr)
    .replace(/\.\d+/, '')
    .replace(/([+-]\d{2}):(\d{2})$/, '$1$2');
}

// ===================================================================
// VENTAS FSC — Etapa 2: Conciliación
// Cruza hoja "Ventas FSC" contra hoja "Ventas" (fuente de verdad).
// ===================================================================

function fsc_conciliar() {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const shFsc    = ss.getSheetByName(FSC_SHEET_VENTAS);
  const shVentas = ss.getSheetByName('Ventas');

  if (!shFsc)    throw new Error('No existe hoja "Ventas FSC".');
  if (!shVentas) throw new Error('No existe hoja "Ventas".');

  SpreadsheetApp.getActive().toast('Conciliación FSC iniciada...', 'EHI', 5);

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

    const precioVenta = fsc_parseNumber_(hVentas[r][colPrecioV]);
    if (precioVenta < 0) continue;  // notas de crédito — se omiten

    const sku   = String(hVentas[r][colSku]).trim();
    const datos = {
      precioVenta: precioVenta,
      cargosVenta: fsc_parseNumber_(hVentas[r][colCargosV]),
      cargoEnvio:  fsc_parseNumber_(hVentas[r][colCargoEnvio]),
      recaudado:   fsc_parseNumber_(hVentas[r][colRecaudado])
    };

    mapaVentasSimple[numVenta]               = datos;
    mapaVentasCompuesto[numVenta + '|' + sku] = datos;
  }

  // ── 2. Leer hoja Ventas FSC ───────────────────────────────────
  const hFsc   = shFsc.getDataRange().getValues();
  const mapFsc = fsc_getHeaderMap_(shFsc);

  const COL_ORDER_NUMBER = mapFsc['fsc_order_number'];
  const COL_SKU_MAESTRO  = mapFsc['sku_maestro'];
  const COL_MONTO_TOTAL  = mapFsc['monto_total'];
  const COL_CARGO_VENTA  = mapFsc['cargo_venta'];
  const COL_CARGO_ENVIO  = mapFsc['cargo_envio'];
  const COL_MONTO_NETO   = mapFsc['monto_neto'];
  const COL_ESTADO_CONC  = mapFsc['estado_conciliacion'];
  const COL_FECHA_CONC   = mapFsc['fecha_conciliacion'];
  const COL_MENSAJE_CONC = mapFsc['mensaje_conciliacion'];

  const ESTADOS_OMITIR = new Set([
    'CANCELADA POR EL COMPRADOR', 'CONCILIADA',
    'DIFERENCIA', 'DIFERENCIA_CONOCIDA', 'NO ENCONTRADA EN VENTAS'
  ]);

  const TOLERANCIA = 1;
  let procesadas        = 0;
  let conciliadas       = 0;
  let diferencias       = 0;
  let diferenciasConocidas = 0;
  let noEncontradas     = 0;
  const ahora           = new Date();

  // ── 3. Procesar cada fila ─────────────────────────────────────
  for (var i = 1; i < hFsc.length; i++) {
    const estadoConc = String(hFsc[i][COL_ESTADO_CONC]).trim();
    if (ESTADOS_OMITIR.has(estadoConc)) continue;

    procesadas++;
    const orderNumber = String(hFsc[i][COL_ORDER_NUMBER]).trim();
    const skuMaestro  = String(hFsc[i][COL_SKU_MAESTRO]).trim();

    // ── Búsqueda 2 niveles ──────────────────────────────────
    // Nivel 1: # Venta + SKU (orden con múltiples ítems distintos)
    // Nivel 2: # Venta simple (orden con un solo ítem — fallback)
    const claveComp = orderNumber + '|' + skuMaestro;
    const ventaRow  = mapaVentasCompuesto[claveComp] || mapaVentasSimple[orderNumber];

    // ── No encontrada ───────────────────────────────────────
    if (!ventaRow) {
      shFsc.getRange(i + 1, COL_ESTADO_CONC  + 1).setValue('NO ENCONTRADA EN VENTAS');
      shFsc.getRange(i + 1, COL_FECHA_CONC   + 1).setValue(ahora);
      shFsc.getRange(i + 1, COL_MENSAJE_CONC + 1).setValue('# Venta ' + orderNumber + ' no existe en hoja Ventas');
      noEncontradas++;
      continue;
    }

    // ── Comparar 4 campos con tolerancia ±$1 ───────────────
    const montoTotal = fsc_parseNumber_(hFsc[i][COL_MONTO_TOTAL]);
    const cargoVenta = fsc_parseNumber_(hFsc[i][COL_CARGO_VENTA]);
    const cargoEnvio = fsc_parseNumber_(hFsc[i][COL_CARGO_ENVIO]);
    const montoNeto  = fsc_parseNumber_(hFsc[i][COL_MONTO_NETO]);

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
      // Si la única diferencia es monto_neto → cofinanciamiento logístico FSC
      const soloNeto = diffs.every(function(d) { return d.indexOf('monto_neto') === 0; });
      if (soloNeto) {
        nuevoEstado = 'DIFERENCIA_CONOCIDA';
        mensaje     = 'Posible cofinanciamiento logístico FSC (no disponible en API). ' + diffs.join(' | ');
        diferenciasConocidas++;
      } else {
        nuevoEstado = 'DIFERENCIA';
        mensaje     = diffs.join(' | ');
        diferencias++;
      }
    }

    shFsc.getRange(i + 1, COL_ESTADO_CONC  + 1).setValue(nuevoEstado);
    shFsc.getRange(i + 1, COL_FECHA_CONC   + 1).setValue(ahora);
    shFsc.getRange(i + 1, COL_MENSAJE_CONC + 1).setValue(mensaje);
  }

  const resumen = {
    procesadas: procesadas, conciliadas: conciliadas,
    diferencias: diferencias, diferenciasConocidas: diferenciasConocidas,
    noEncontradas: noEncontradas
  };
  Logger.log('[FSC-CONC] Resumen: ' + JSON.stringify(resumen));
  SpreadsheetApp.getActive().toast(
    'FSC conciliación: ' + procesadas + ' procesadas | ' +
    conciliadas + ' ok | ' + diferencias + ' diff | ' +
    diferenciasConocidas + ' diff_conocida | ' + noEncontradas + ' no encontradas',
    'EHI', 10
  );
  return resumen;
}


// ── Helper: parsea número en formato chileno o número directo ───
function fsc_parseNumber_(val) {
  if (val === null || val === undefined || val === '') return 0;
  if (typeof val === 'number') return Math.round(val);
  return parseInt(String(val).replace(/\./g, '').replace(/,/g, ''), 10) || 0;
}
