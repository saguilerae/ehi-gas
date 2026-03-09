// ===================================================================
// Falabella Seller Center
// Ajustes incluidos:
// - Lectura de credenciales desde Script Properties
// - OperatorCode parametrizado
// - getSellercenterApiResponse() corregido con headers
// - payload solo en POST
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