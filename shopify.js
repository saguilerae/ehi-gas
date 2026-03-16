// ===================================================================
// Obtiene la configuración Shopify
// ===================================================================
function getShopifyConfig() {
  const props = PropertiesService.getScriptProperties();
  return {
    storeDomain: props.getProperty('SHOPIFY_STORE_DOMAIN'), // p.ej. quickdropshipping.myshopify.com
    token: props.getProperty('SHOPIFY_ADMIN_TOKEN'),
    apiVersion: '2024-10', // puedes ajustarlo si Shopify cambia
  };
}

// ===================================================================
// Invoca a los parametros de Shopify
// ===================================================================
function callShopify(method, path, payload) {
  const cfg = getShopifyConfig();
  if (!cfg.storeDomain || !cfg.token) {
    throw new Error('Falta configurar SHOPIFY_STORE_DOMAIN o SHOPIFY_ADMIN_TOKEN en Propiedades del script.');
  }

  const url = `https://${cfg.storeDomain}/admin/api/${cfg.apiVersion}${path}`;
  const options = {
    method,
    muteHttpExceptions: true,
    headers: {
      'X-Shopify-Access-Token': cfg.token,
      'Content-Type': 'application/json'
    }
  };
  if (payload) {
    options.payload = JSON.stringify(payload);
  }

  const res = UrlFetchApp.fetch(url, options);
  const code = res.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error(`Shopify ${method} ${path} → ${code}: ${res.getContentText()}`);
  }
  const text = res.getContentText();
  return text ? JSON.parse(text) : {};
}

// ===================================================================
// ************ 1) CONSULTAR PRODUCTOS → SHEET ************/
// Obtiene los headers
// ===================================================================
function getHeaderIndex_(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const idx = {};
  headers.forEach(function(name, i) {
    idx[name] = i;
  });
  return idx;
}


// ===================================================================
// Hoja donde trabajar
// ===================================================================
const SHEET_NAME = 'Prods. SH';  // cambia al nombre que uses

// ===================================================================
// Orden de columnas requerido
// ===================================================================
const HEADERS = [
  'Actualizar',               // checkbox
  'Variant SKU',
  'Title',
  'Type',
  'Published',
  'Option1 Name',
  'Option1 Value',
  'Variant Price',
  'Variant Compare At Price',
  'Variant Inventory Qty',
  'Status',
  // columnas internas para updates (NO TOCAR ORDEN)
  'Variant ID',
  'Inventory Item ID',
  'Product ID',
  // NUEVO: campos para SEO / análisis (al final, para no romper nada)
  'Handle',
  'SEO Title',
  'SEO Description',
  'Description',   // descripción del producto (texto plano)
  'Frame Color'    // color de la montura (usamos Option1 Value)
];

// ===================================================================
// Trae TODOS los productos (hasta 250 por llamada) y los deja en la hoja
// Trae TODOS los productos y los deja en la hoja
// ===================================================================
function syncShopifyProductsToSheet() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error(`No existe la hoja "${SHEET_NAME}"`);

  sheet.appendRow(HEADERS);

  const products = fetchAllProducts(); // array de productos

  const rows = [];
  products.forEach(prod => {
    const title       = prod.title;
    const type        = prod.product_type || '';
    const published   = !!prod.published_at;
    const status      = prod.status || '';
    const productId   = prod.id;

    const handle      = prod.handle || '';
    const seoTitle    = prod.metafields_global_title_tag || '';
    const seoDesc     = prod.metafields_global_description_tag || '';

    // Descripción en texto plano (sacamos las etiquetas HTML básicas)
    const rawBody     = prod.body_html || '';
    const description = rawBody.replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ').trim();

    const option1Name = (prod.options && prod.options[0]) ? prod.options[0].name : '';

    prod.variants.forEach(variant => {
      const sku             = variant.sku || '';
      const option1Value    = variant.option1 || '';   // color de variante
      const price           = variant.price || '';
      const compareAt       = variant.compare_at_price || '';
      const qty             = variant.inventory_quantity || 0;
      const variantId       = variant.id;
      const inventoryItemId = variant.inventory_item_id;

      const frameColor      = option1Value; // usamos Option1 como color de montura

      rows.push([
        false,           // Actualizar
        sku,
        title,
        type,
        published,
        option1Name,
        option1Value,
        price,
        compareAt,
        qty,
        status,
        variantId,
        inventoryItemId,
        productId,
        handle,
        seoTitle,
        seoDesc,
        description,
        frameColor
      ]);
    });
  });

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, HEADERS.length).setValues(rows).clearContent()
    sheet.getRange(2, 1, rows.length, HEADERS.length).setValues(rows);
  }
}

// ===================================================================
// Paginación sencilla (hasta que no haya "link: rel=next")
// ===================================================================
function fetchAllProducts() {
  const perPage = 250;
  // añadimos body_html (descripción) y campos SEO globales
  let path =
    `/products.json?limit=${perPage}` +
    `&fields=id,title,body_html,product_type,status,published_at,variants,options,` +
    `handle,metafields_global_title_tag,metafields_global_description_tag`;

  const all = [];
  const cfg = getShopifyConfig();

  while (true) {
    const url = `https://${cfg.storeDomain}/admin/api/${cfg.apiVersion}${path}`;
    const res = UrlFetchApp.fetch(url, {
      method: 'get',
      muteHttpExceptions: true,
      headers: {
        'X-Shopify-Access-Token': cfg.token
      }
    });

    const code = res.getResponseCode();
    if (code < 200 || code >= 300) {
      throw new Error(`Shopify GET ${path} → ${code}: ${res.getContentText()}`);
    }

    const data = JSON.parse(res.getContentText());
    if (data.products && data.products.length) {
      Array.prototype.push.apply(all, data.products);
    }

    const link = res.getHeaders()['Link'];
    if (!link || link.indexOf('rel="next"') === -1) break;

    const match = link.match(/<([^>]+)>;\s*rel="next"/);
    if (!match) break;
    const nextUrl = match[1];

    // convertimos la URL absoluta en un path relativo compatible con la lógica anterior
    path = nextUrl.replace(`https://${cfg.storeDomain}/admin/api/${cfg.apiVersion}`, '');
  }

  return all;
}


// ===================================================================
// *********** 2) ACTUALIZAR PRECIOS Y STOCK DESDE LA HOJA ***********/
// Actualiza Variant Price, Compare At Price y Stock para las filas con "Actualizar" = TRUE
// ===================================================================
function updateShopifyFromSheet() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error(`No existe la hoja "${SHEET_NAME}"`);

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const data = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();

  const locationId = getDefaultLocationId(); // una sola vez

  data.forEach((row, i) => {
    const actualizar = row[0];
    if (!actualizar) return;

    const sku = row[1];
    const price = row[7];
    const compareAt = row[8];
    const qty = row[9];

    const variantId = row[11];
    const inventoryItemId = row[12];

    if (!variantId) {
      Logger.log(`Fila ${i + 2} (${sku}) sin Variant ID, se omite.`);
      return;
    }

    // 1) actualizar precio y compare_at_price
    const pathVariant = `/variants/${variantId}.json`;
    const payloadVariant = {
      variant: {
        id: variantId,
        price: price,
        compare_at_price: compareAt || null
      }
    };
    callShopify('PUT', pathVariant, payloadVariant);

    // 2) actualizar stock (si tenemos inventory_item_id y location)
    if (inventoryItemId && locationId && qty !== '') {
      const pathInv = `/inventory_levels/set.json`;
      const payloadInv = {
        location_id: locationId,
        inventory_item_id: inventoryItemId,
        available: Number(qty)
      };
      callShopify('POST', pathInv, payloadInv);
    }
  });

  SpreadsheetApp.getActive().toast('Actualización enviada a Shopify', 'GS-Sync', 5);
}

// ===================================================================
// Obtiene el primer location_id y lo guarda para reutilizarlo
// ===================================================================
function getDefaultLocationId() {
  const props = PropertiesService.getScriptProperties();
  let locId = props.getProperty('SHOPIFY_LOCATION_ID');
  if (locId) return Number(locId);

  const res = callShopify('GET', '/locations.json', null);
  if (!res.locations || !res.locations.length) {
    throw new Error('No se encontraron ubicaciones (locations) en Shopify.');
  }
  locId = res.locations[0].id;
  props.setProperty('SHOPIFY_LOCATION_ID', String(locId));
  return locId;
}

// ===================================================================
// Función para actualizar solo STOCK
// ===================================================================
function updateShopifyStockOnly() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME); // 'Prods. SH'
  if (!sheet) throw new Error(`No existe la hoja '${SHEET_NAME}'`);

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const lastCol = sheet.getLastColumn();
  const idx = getHeaderIndex_(sheet);

  // Validar headers críticos
  const required = ['Actualizar', 'Inventory Item ID', 'Variant Inventory Qty'];
  const missing = required.filter(h => idx[h] == null);
  if (missing.length) {
    throw new Error(
      `Headers faltantes en '${SHEET_NAME}': ${missing.join(', ')}. ` +
      `Asegúrate de ejecutar sobre la hoja correcta y que los encabezados coincidan.`
    );
  }

  const locationIdRaw = PropertiesService.getScriptProperties().getProperty('SHOPIFY_LOCATION_ID');
  const locationId = normalizeShopifyNumericId_(locationIdRaw);
  if (!locationId) throw new Error('Falta/ inválido SHOPIFY_LOCATION_ID (Location ID)');

  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  for (let r = 0; r < data.length; r++) {
    const row = data[r];

    const shouldUpdate = row[idx['Actualizar']] === true;
    if (!shouldUpdate) continue;

    const inventoryItemId = normalizeShopifyNumericId_(row[idx['Inventory Item ID']]);
    const qtyRaw = row[idx['Variant Inventory Qty']];

    // qty puede venir como número o string
    const qty = (qtyRaw === '' || qtyRaw == null) ? null : Number(String(qtyRaw).replace(',', '.'));

    if (!inventoryItemId) {
      Logger.log(`Shopify: inventory_item_id inválido. fila=${r + 2}`);
      continue;
    }
    if (qty == null || !isFinite(qty)) {
      Logger.log(`Shopify: qty inválido. fila=${r + 2}`);
      continue;
    }

    const payload = {
      location_id: Number(locationId),
      inventory_item_id: Number(inventoryItemId),
      available: Number(qty)
    };

    try {
      callShopify('POST', '/inventory_levels/set.json', payload);
      // opcional: desmarcar checkbox al éxito
      // sheet.getRange(r + 2, idx['Actualizar'] + 1).setValue(false);
    } catch (e) {
      // Deja el checkbox marcado para reintentar y loguea error
      Logger.log(`ERROR Shopify stock fila=${r + 2}: ${e.message}`);
      throw e; // si prefieres continuar con otras filas, comenta este throw
    }
  }

  ss.toast('Stock actualizado', 'Shopify Sync', 5);

  // ---- helper local ----
  function normalizeShopifyNumericId_(x) {
    if (x == null) return '';
    const s = String(x).trim();
    if (!s) return '';
    // Si viene como GID: gid://shopify/InventoryItem/123456789
    const m = s.match(/\/(\d+)\s*$/);
    if (m) return m[1];
    // Si ya viene numérico
    if (/^\d+$/.test(s)) return s;
    return '';
  }
}

// ===================================================================
// Función para actualizar solo PRECIO PRINCIPAL
// ===================================================================
function updateShopifyPriceOnly() {
  const sheet   = SpreadsheetApp.getActiveSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const lastCol = sheet.getLastColumn();
  const idx     = getHeaderIndex_(sheet);
  const data    = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  data.forEach(row => {
    const shouldUpdate = row[idx['Actualizar']];
    const variantId    = row[idx['Variant ID']];
    const price        = row[idx['Variant Price']];

    if (!shouldUpdate || !variantId || price === '') return;

    const payload = {
      variant: {
        id: Number(variantId),
        price: String(price).replace(',', '.')
      }
    };

    callShopify('PUT', `/variants/${variantId}.json`, payload);
  });

  SpreadsheetApp.getActive().toast('Precios actualizados', 'Shopify Sync', 5);
}

// ===================================================================
// Actualiza sólo el estado del producto
// ===================================================================
function updateShopifyStatusOnly() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(SHEET_NAME); // "Prods. SH"
  if (!sheet) {
    throw new Error(`No existe la hoja "${SHEET_NAME}"`);
  }

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) return;

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  const colActualizar = headers.indexOf('Actualizar');
  const colStatus     = headers.indexOf('Status');
  const colProductId  = headers.indexOf('Product ID');

  if (colActualizar === -1 || colStatus === -1 || colProductId === -1) {
    throw new Error('Falta alguna columna: "Actualizar" / "Status" / "Product ID"');
  }

  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  data.forEach((row, i) => {
    const marcado   = row[colActualizar];
    const status    = row[colStatus];
    const productId = row[colProductId];

    if (marcado === true && productId && status) {
      Logger.log(`Fila ${i+2}: actualizando Product ${productId} → status "${status}"`);
      try {
        actualizarEstadoProductoShopify_(productId, status);
      } catch (e) {
        Logger.log(`Error en fila ${i+2}: ` + e);
      }
    }
  });
}

function actualizarEstadoProductoShopify_(productId, status) {
  const payload = {
    product: {
      id: Number(productId),
      status: status  // 'active' | 'draft' | 'archived'
    }
  };

  // Usamos callShopify → respeta dominio, token y apiVersion configurados
  const res = callShopify(
    'PUT',
    `/products/${productId}.json`,
    payload
  );

  Logger.log(
    `Estado actualizado para Product ${productId} → "${status}". ` +
    `Respuesta Shopify: ${JSON.stringify(res)}`
  );
}

// Actualiza Title, Handle, SEO Title y SEO Description desde la hoja
function updateShopifySeoFromSheet() {
  const ss    = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(SHEET_NAME); // 'Prods. SH'
  if (!sheet) {
    throw new Error(`No existe la hoja "${SHEET_NAME}"`);
  }

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) {
    Logger.log('No hay filas de datos para procesar.');
    return;
  }

  // Índices por nombre de columna
  const idx = getHeaderIndex_(sheet);

  const requiredCols = [
    'Actualizar',
    'Product ID',
    'Title',
    'Handle',
    'SEO Title',
    'SEO Description'
  ];

  requiredCols.forEach(name => {
    if (idx[name] === undefined) {
      throw new Error(`Falta la columna "${name}" en los encabezados.`);
    }
  });

  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  let updatesCount = 0;

  data.forEach((row, i) => {
    const rowNumber = i + 2; // fila real en la hoja

    const actualizar = row[idx['Actualizar']];
    if (!actualizar) {
      return; // solo filas marcadas con TRUE
    }

    const productId = row[idx['Product ID']];
    if (!productId) {
      Logger.log(`Fila ${rowNumber}: sin Product ID, se omite.`);
      return;
    }

    const title       = row[idx['Title']] || '';
    const handle      = row[idx['Handle']] || '';
    const seoTitle    = row[idx['SEO Title']] || '';
    const seoDesc     = row[idx['SEO Description']] || '';

    // Construimos el objeto product solo con lo que tengamos
    const productPayload = {
      id: Number(productId)
    };

    // Sólo seteamos campos si vienen con algún valor,
    // para no borrar SEO existente accidentalmente.
    if (title) {
      productPayload.title = title;
    }
    if (handle) {
      productPayload.handle = handle;
    }
    if (seoTitle) {
      productPayload.metafields_global_title_tag = seoTitle;
    }
    if (seoDesc) {
      productPayload.metafields_global_description_tag = seoDesc;
    }

    // Si sólo tenemos el id y nada más, no tiene sentido llamar a Shopify
    if (Object.keys(productPayload).length === 1) {
      Logger.log(`Fila ${rowNumber}: no hay campos SEO/Title/Handle para actualizar, se omite.`);
      return;
    }

    try {
      const payload = { product: productPayload };
      const res = callShopify('PUT', `/products/${productId}.json`, payload);

      Logger.log(
        `Fila ${rowNumber}: Product ${productId} actualizado. ` +
        `Respuesta parcial: ${JSON.stringify({
          id: res.product && res.product.id,
          title: res.product && res.product.title,
          handle: res.product && res.product.handle
        })}`
      );

      updatesCount++;
      // Pequeño delay opcional para no saturar el rate limit de Shopify
      Utilities.sleep(300);

    } catch (e) {
      Logger.log(`Error en fila ${rowNumber} (Product ${productId}): ${e}`);
    }
  });

  SpreadsheetApp.getActive().toast(
    `SEO actualizado para ${updatesCount} producto(s).`,
    'Shopify SEO Sync',
    5
  );
}

// ===================================================================
// Actualiza el título principal de los productos en Shopify
// basándose en la columna "Title" de la Google Sheet.
// =================================================================== 
function updateShopifyTitlesFromSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  // Índices de columnas según tu planilla
  const titleIdx = headers.indexOf('Title');
  const productIdIdx = headers.indexOf('Product ID');
  const updateIdx = headers.indexOf('Actualizar'); // Solo actualiza si está marcado

  if (titleIdx === -1 || productIdIdx === -1) {
    throw new Error('No se encontraron las columnas "Title" o "Product ID".');
  }

  let count = 0;

  // Empezamos en i = 1 para saltar el encabezado
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const productId = row[productIdIdx];
    const newTitle = row[titleIdx];
    const shouldUpdate = row[updateIdx];

    // Validamos que tenga ID, Título y que la celda "Actualizar" sea TRUE (o esté marcada)
    if (productId && newTitle && (shouldUpdate === true || shouldUpdate === "TRUE")) {
      try {
        const payload = {
          product: {
            id: productId,
            title: newTitle
          }
        };

        // Llamada a la API de Shopify usando tu función existente
        callShopify('PUT', `/products/${productId}.json`, payload);
        
        Logger.log(`Actualizado ID ${productId}: ${newTitle}`);
        count++;
        
        // Opcional: Desmarcar la casilla de "Actualizar" tras éxito
        if (updateIdx !== -1) {
          sheet.getRange(i + 1, updateIdx + 1).setValue(false);
        }

      } catch (e) {
        Logger.log(`Error en fila ${i + 1} (ID ${productId}): ${e.message}`);
      }
    }
  }

  SpreadsheetApp.getActive().toast(
    `Título actualizado para ${count} producto(s).`,
    'Shopify Título Sync',
    5
  );
}

// =================================================================== 
// UTILITARIOS
// =================================================================== 

// =================================================================== 
// Muestra el ID de Localicación de mi Shotify
// =================================================================== 
function showShopifyLocations() {
  const cfg = getShopifyConfig();
  const res = callShopify('GET', '/locations.json');
  Logger.log(JSON.stringify(res, null, 2));
}

// Muestra metafields_global_title_tag y metafields_global_description_tag
function debugProductSeo() {
  const cfg = getShopifyConfig();
  const productId = 10703153037585; // cambia por el ID que quieras

  const url = `https://${cfg.storeDomain}/admin/api/${cfg.apiVersion}/products/${productId}.json` +
              `?fields=id,title,metafields_global_title_tag,metafields_global_description_tag,handle`;
  const res = UrlFetchApp.fetch(url, {
    method: 'get',
    muteHttpExceptions: true,
    headers: {
      'X-Shopify-Access-Token': cfg.token
    }
  });

  Logger.log(res.getContentText());
}

// =================================================================== 
// Prueba la actualización de SEOs
// =================================================================== 
function testUpdateSeoOne() {
  const productId = 10703153037585; // cambia si quieres otro

  const payload = {
    product: {
      id: productId,
      metafields_global_title_tag: 'Lentes de Sol SH501092 UV400 Mujer Cuadrado Retro | Elizalde',
      metafields_global_description_tag: 'Lentes de sol retro SH501092 para mujer con protección UV400. Marco liviano y detalles metálicos. Envío rápido en Chile.',
      handle: 'sh501092-lentes-uv400-mujer-retro'

    }
  };

  const res = callShopify('PUT', `/products/${productId}.json`, payload);
  Logger.log(JSON.stringify(res));
}
// ========================================================================================
// FIN DE UTILITARIOS
// ========================================================================================

// ===================================================================
// VENTAS SH — Ingesta de Órdenes Shopify
// Etapa 1: Ingesta (Histórica e Incremental)
// ===================================================================

// Constantes del módulo
const SH_SHEET_VENTAS         = 'Ventas SH';
const SH_GATEWAY_RATE         = 0.0284;  // Mercado Pago 2.84%
const SH_SHOPIFY_RATE         = 0.02;    // Shopify 2%
const SH_CARGO_ENVIO          = 5000;    // Estimado fijo CLP
const SH_PAGE_LIMIT           = 250;     // máximo Shopify REST
const SH_PACING_MS            = 300;
const SH_MAX_MS               = 5 * 60 * 1000;
const SH_LOG_HIST             = '[SH-HIST]';
const SH_LOG_INC              = '[SH-INC]';

// Headers exactos — 33 columnas
const SH_HEADERS = [
  'fecha_ingesta','canal','sh_order_id','sh_order_number',
  'sh_order_name','id_venta_canal','fecha_venta','sku_canal',
  'sku_maestro','mpn','cantidad','precio_unitario',
  'monto_total','cargo_gateway','cargo_shopify','cargo_venta_estimado',
  'cargo_envio','monto_neto_estimado','monto_neto_parcial','gateway',
  'financial_status','fulfillment_status','cliente_nombre','region',
  'comuna','direccion_resumen','estado_sh','estado_conciliacion',
  'fecha_conciliacion','mensaje_conciliacion','clave_unica',
  'existe_en_ventas','json_raw'
];

// Estados de conciliación válidos (referencia)
// PENDIENTE | PENDIENTE_HISTORICO | CANCELADA POR EL COMPRADOR |
// DEVOLUCION CON REEMBOLSO | CONCILIADA | DIFERENCIA |
// DIFERENCIA_CONOCIDA | NO ENCONTRADA EN VENTAS


// ── Función pública menú: Ingesta Histórica ─────────────────────
// Ventanas de 30 días desde SHOPIFY_HISTORICAL_FROM hasta hoy.
function sh_ingestaHistorica() {
  const props          = PropertiesService.getScriptProperties();
  const historicalFrom = String(props.getProperty('SHOPIFY_HISTORICAL_FROM') || '').trim();

  if (!historicalFrom) {
    SpreadsheetApp.getActive().toast('Falta SHOPIFY_HISTORICAL_FROM en Script Properties.', 'SH', 5);
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh   = ss.getSheetByName(SH_SHEET_VENTAS);
  if (!sh) sh = sh_crearHojaVentasSH_();

  const existingKeys = sh_getExistingKeys_(sh);
  const startTime    = Date.now();

  let desde          = new Date(historicalFrom);
  const hasta        = new Date();
  const VENTANA_DIAS = 30;
  let totalOrders    = 0;
  let totalInserted  = 0;
  let ventana        = 0;
  let allRows        = [];

  SpreadsheetApp.getActive().toast('Ingesta histórica SH iniciada...', 'EHI', 5);

  function flushRows_() {
    if (!allRows.length) return;
    const startRow = Math.max(sh.getLastRow() + 1, 2);
    sh.getRange(startRow, 1, allRows.length, SH_HEADERS.length).setValues(allRows);
    totalInserted  += allRows.length;
    allRows.length  = 0;
    Logger.log(SH_LOG_HIST + ' Flush: totalInserted=' + totalInserted);
  }

  while (desde < hasta) {
    const fin     = new Date(desde);
    fin.setDate(fin.getDate() + VENTANA_DIAS);
    const finReal = fin > hasta ? hasta : fin;

    const createdAtMin = desde.toISOString();
    const createdAtMax = finReal.toISOString();

    ventana++;
    Logger.log(SH_LOG_HIST + ' Ventana ' + ventana + ': ' + createdAtMin + ' → ' + createdAtMax);
    SpreadsheetApp.getActive().toast(
      'Ventana ' + ventana + ' | Órdenes acumuladas: ' + totalOrders + ' | Buffer: ' + allRows.length,
      'SH Histórico ⏳', 8
    );

    // Paginación por cursor (Link header)
    var path = '/orders.json'
      + '?limit=' + SH_PAGE_LIMIT
      + '&status=any'
      + '&created_at_min=' + encodeURIComponent(createdAtMin)
      + '&created_at_max=' + encodeURIComponent(createdAtMax);

    while (path) {
      const result = sh_fetchOrdersPage_(path);
      const orders = result.orders || [];
      if (!orders.length) break;

      totalOrders += orders.length;

      for (var i = 0; i < orders.length; i++) {
        const rows = sh_buildRows_(orders[i], 'HISTORICO', existingKeys);
        for (var j = 0; j < rows.length; j++) {
          allRows.push(rows[j]);
        }
      }

      path = result.nextPath || null;
      if (path) Utilities.sleep(SH_PACING_MS);
    }

    // Flush condicional si se acerca el timeout
    if (Date.now() - startTime > SH_MAX_MS) {
      Logger.log(SH_LOG_HIST + ' Timeout inminente — flush forzado en ventana ' + ventana);
      flushRows_();
    }

    desde = fin;
    Utilities.sleep(SH_PACING_MS);
  }

  flushRows_();

  const resumen = { ventanas: ventana, fetchedOrders: totalOrders, insertedRows: totalInserted };
  Logger.log(SH_LOG_HIST + ' Resumen: ' + JSON.stringify(resumen));
  SpreadsheetApp.getActive().toast(
    'SH histórico: ventanas=' + ventana + ' | órdenes=' + totalOrders + ' | filas=' + totalInserted,
    'EHI', 10
  );
  return resumen;
}


// ── Función pública menú: Ingesta Incremental ───────────────────
// Ventanas de 7 días desde SHOPIFY_LAST_SYNC hasta hoy.
function sh_ingestaIncremental() {
  const props    = PropertiesService.getScriptProperties();
  const lastSync = String(props.getProperty('SHOPIFY_LAST_SYNC') || '').trim();

  if (!lastSync) throw new Error('Falta SHOPIFY_LAST_SYNC en Script Properties.');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh   = ss.getSheetByName(SH_SHEET_VENTAS);
  if (!sh) sh = sh_crearHojaVentasSH_();

  const existingKeys = sh_getExistingKeys_(sh);
  const startTime    = Date.now();

  let desde          = new Date(lastSync);
  const hasta        = new Date();
  const VENTANA_DIAS = 7;
  let totalOrders    = 0;
  let totalInserted  = 0;
  let ventana        = 0;
  let allRows        = [];

  SpreadsheetApp.getActive().toast('Ingesta incremental SH iniciada...', 'EHI', 5);

  function flushRows_() {
    if (!allRows.length) return;
    const startRow = Math.max(sh.getLastRow() + 1, 2);
    sh.getRange(startRow, 1, allRows.length, SH_HEADERS.length).setValues(allRows);
    totalInserted  += allRows.length;
    allRows.length  = 0;
    Logger.log(SH_LOG_INC + ' Flush: totalInserted=' + totalInserted);
  }

  while (desde < hasta) {
    const fin     = new Date(desde);
    fin.setDate(fin.getDate() + VENTANA_DIAS);
    const finReal = fin > hasta ? hasta : fin;

    const createdAtMin = desde.toISOString();
    const createdAtMax = finReal.toISOString();

    ventana++;
    Logger.log(SH_LOG_INC + ' Ventana ' + ventana + ': ' + createdAtMin + ' → ' + createdAtMax);
    SpreadsheetApp.getActive().toast(
      'Ventana ' + ventana + ' | Órdenes acumuladas: ' + totalOrders + ' | Buffer: ' + allRows.length,
      'SH Incremental ⏳', 8
    );

    var path = '/orders.json'
      + '?limit=' + SH_PAGE_LIMIT
      + '&status=any'
      + '&created_at_min=' + encodeURIComponent(createdAtMin)
      + '&created_at_max=' + encodeURIComponent(createdAtMax);

    while (path) {
      const result = sh_fetchOrdersPage_(path);
      const orders = result.orders || [];
      if (!orders.length) break;

      totalOrders += orders.length;

      for (var i = 0; i < orders.length; i++) {
        const rows = sh_buildRows_(orders[i], 'INCREMENTAL', existingKeys);
        for (var j = 0; j < rows.length; j++) {
          allRows.push(rows[j]);
        }
      }

      path = result.nextPath || null;
      if (path) Utilities.sleep(SH_PACING_MS);
    }

    if (Date.now() - startTime > SH_MAX_MS) {
      Logger.log(SH_LOG_INC + ' Timeout inminente — flush forzado en ventana ' + ventana);
      flushRows_();
    }

    desde = fin;
    Utilities.sleep(SH_PACING_MS);
  }

  flushRows_();

  props.setProperty('SHOPIFY_LAST_SYNC', hasta.toISOString());

  const resumen = { ventanas: ventana, fetchedOrders: totalOrders, insertedRows: totalInserted };
  Logger.log(SH_LOG_INC + ' Resumen: ' + JSON.stringify(resumen));
  SpreadsheetApp.getActive().toast(
    'SH incremental: ventanas=' + ventana + ' | órdenes=' + totalOrders + ' | filas=' + totalInserted,
    'EHI', 7
  );
  return resumen;
}

// =================================================================== 
// ── Fetch página de órdenes con cursor ─────────────────────────
// =================================================================== 
function sh_fetchOrdersPage_(path) {
  const cfg = getShopifyConfig();
  const url = 'https://' + cfg.storeDomain + '/admin/api/' + cfg.apiVersion + path;

  const res = UrlFetchApp.fetch(url, {
    method: 'get',
    muteHttpExceptions: true,
    headers: {
      'X-Shopify-Access-Token': cfg.token,
      'Content-Type': 'application/json'
    }
  });

  const code = res.getResponseCode();
  if (code < 200 || code >= 300) {
    Logger.log('[SH] Error GET ' + path + ' → ' + code + ': ' + res.getContentText());
    return { orders: [], nextPath: null };
  }

  const data = JSON.parse(res.getContentText());

  // Extraer cursor siguiente del Link header
  var nextPath = null;
  const link = res.getHeaders()['Link'] || res.getHeaders()['link'] || '';
  if (link && link.indexOf('rel="next"') !== -1) {
    const match = link.match(/<([^>]+)>;\s*rel="next"/);
    if (match) {
      nextPath = match[1].replace(
        'https://' + cfg.storeDomain + '/admin/api/' + cfg.apiVersion, ''
      );
    }
  }

  return { orders: data.orders || [], nextPath: nextPath };
}

// =================================================================== 
// ── Construir filas desde una orden (1 fila por line_item) ──────
// =================================================================== 
function sh_buildRows_(order, modo, existingKeys) {
  const rows       = [];
  const lineItems  = Array.isArray(order.line_items) ? order.line_items : [];

  const orderId     = String(order.id || '');
  const orderNumber = String(order.order_number || '');
  const orderName   = String(order.name || '');
  const fechaVenta  = String(order.created_at || '');
  const finStatus   = String(order.financial_status || '');
  const fulStatus   = String(order.fulfillment_status || '');
  const gateway     = order.payment_gateway_names && order.payment_gateway_names.length
                      ? String(order.payment_gateway_names[0])
                      : '';

  // Cliente
  const cust    = order.customer || {};
  const cliente = String(((cust.first_name || '') + ' ' + (cust.last_name || '')).trim());

  // Dirección envío
  const addr   = order.shipping_address || order.billing_address || {};
  const region = String(addr.province || '');
  const comuna = String(addr.city || '');
  const direc  = String(addr.address1 || '');

  // Estado sh combinado
  const estadoSh = finStatus + '/' + fulStatus;

  for (var i = 0; i < lineItems.length; i++) {
    const item       = lineItems[i];
    const sku        = String(item.sku || '');
    if (!sku) continue;  // omitir ítems sin SKU (gift cards, etc.)

    const mpn        = sku.length > 2 ? sku.slice(0, -2) : sku;
    const cantidad   = Number(item.quantity || 1);
    const precioUnit = parseFloat(item.price || 0);
    const montoTotal = parseFloat((precioUnit * cantidad).toFixed(0));

    // Clave interna deduplicación: orderNumber|sku
    const claveInterna = orderNumber + '|' + sku;
    if (existingKeys.has(claveInterna)) continue;

    const claveUnica       = orderNumber + '-' + sku;
    const cargoGateway     = parseFloat((montoTotal * SH_GATEWAY_RATE).toFixed(0));
    const cargoShopify     = parseFloat((montoTotal * SH_SHOPIFY_RATE).toFixed(0));
    const cargoVentaEst    = cargoGateway + cargoShopify;
    const montoNetoEst     = parseFloat((montoTotal - cargoVentaEst - SH_CARGO_ENVIO).toFixed(0));
    const montoNetoParcial = parseFloat((montoTotal - cargoGateway - SH_CARGO_ENVIO).toFixed(0));

    // Estado conciliación
    let estadoConciliacion = '';
    if (finStatus === 'refunded' || finStatus === 'partially_refunded') {
      estadoConciliacion = 'DEVOLUCION CON REEMBOLSO';
    } else if (finStatus === 'voided' || order.cancelled_at) {
      estadoConciliacion = 'CANCELADA POR EL COMPRADOR';
    } else if (modo === 'HISTORICO') {
      estadoConciliacion = 'PENDIENTE_HISTORICO';
    } else {
      estadoConciliacion = 'PENDIENTE';
    }

    existingKeys.add(claveInterna);

    rows.push([
      new Date(),       // A fecha_ingesta
      'SH',             // B canal
      orderId,          // C sh_order_id
      orderNumber,      // D sh_order_number
      orderName,        // E sh_order_name
      orderNumber,      // F id_venta_canal
      fechaVenta,       // G fecha_venta
      sku,              // H sku_canal
      sku,              // I sku_maestro (en Shopify son iguales)
      mpn,              // J mpn
      cantidad,         // K cantidad
      precioUnit,       // L precio_unitario
      montoTotal,       // M monto_total
      cargoGateway,     // N cargo_gateway
      cargoShopify,     // O cargo_shopify
      cargoVentaEst,    // P cargo_venta_estimado
      SH_CARGO_ENVIO,   // Q cargo_envio
      montoNetoEst,     // R monto_neto_estimado
      montoNetoParcial, // S monto_neto_parcial
      gateway,          // T gateway
      finStatus,        // U financial_status
      fulStatus,        // V fulfillment_status
      cliente,          // W cliente_nombre
      region,           // X region
      comuna,           // Y comuna
      direc,            // Z direccion_resumen
      estadoSh,         // AA estado_sh
      estadoConciliacion, // AB estado_conciliacion
      '',               // AC fecha_conciliacion
      '',               // AD mensaje_conciliacion
      claveUnica,       // AE clave_unica
      false,            // AF existe_en_ventas
      JSON.stringify({ order_id: orderId, order_number: orderNumber, item: item }) // AG json_raw
    ]);
  }
  return rows;
}

// =================================================================== 
// ── Helper: claves únicas existentes ───────────────────────────
// Clave interna: sh_order_number|sku_maestro (cols D=4, I=9)
// =================================================================== 
function sh_getExistingKeys_(sh) {
  const set     = new Set();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return set;

  const orderNums = sh.getRange(2, 4, lastRow - 1, 1).getValues();
  const skus      = sh.getRange(2, 9, lastRow - 1, 1).getValues();

  for (var i = 0; i < orderNums.length; i++) {
    const on  = String(orderNums[i][0] || '').trim();
    const sku = String(skus[i][0]      || '').trim();
    if (on && sku) set.add(on + '|' + sku);
  }
  return set;
}

// =================================================================== 
// ── Crear hoja Ventas SH si no existe ──────────────────────────
// =================================================================== 
function sh_crearHojaVentasSH_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh   = ss.getSheetByName(SH_SHEET_VENTAS);
  if (sh) return sh;

  sh = ss.insertSheet(SH_SHEET_VENTAS);
  sh.getRange(1, 1, 1, SH_HEADERS.length).setValues([SH_HEADERS]);
  sh.setFrozenRows(1);

  // Formato fecha cols A(1) y G(7)
  sh.getRange(2, 1, sh.getMaxRows() - 1, 1).setNumberFormat('dd/mm/yyyy hh:mm:ss');
  sh.getRange(2, 7, sh.getMaxRows() - 1, 1).setNumberFormat('dd/mm/yyyy hh:mm:ss');

  // Formato número cols L(12),M(13),N(14),O(15),P(16),R(18),S(19)
  [12, 13, 14, 15, 16, 18, 19].forEach(function(c) {
    sh.getRange(2, c, sh.getMaxRows() - 1, 1).setNumberFormat('#,##0');
  });

  return sh;
}

// ===================================================================
// FIN DE VENTAS SH — Ingesta de Órdenes Shopify
// ===================================================================

// ===================================================================
// VENTAS SH — Etapa 2: Conciliación
// Cruza hoja "Ventas SH" contra hoja "Ventas" (fuente de verdad).
// Clave: sh_order_number + sku_maestro (2 niveles)
// Campos: monto_total, cargo_gateway, cargo_envio, monto_neto_parcial
// ===================================================================

function sh_conciliar() {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const shSh     = ss.getSheetByName(SH_SHEET_VENTAS);
  const shVentas = ss.getSheetByName('Ventas');

  if (!shSh)     throw new Error('No existe hoja "Ventas SH".');
  if (!shVentas) throw new Error('No existe hoja "Ventas".');

  SpreadsheetApp.getActive().toast('Conciliación SH iniciada...', 'EHI', 5);

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

    const precioVenta = sh_parseNumber_(hVentas[r][colPrecioV]);
    if (precioVenta < 0) continue;  // notas de crédito — se omiten

    const sku   = String(hVentas[r][colSku]).trim();
    const datos = {
      precioVenta: precioVenta,
      cargosVenta: sh_parseNumber_(hVentas[r][colCargosV]),
      cargoEnvio:  sh_parseNumber_(hVentas[r][colCargoEnvio]),
      recaudado:   sh_parseNumber_(hVentas[r][colRecaudado])
    };

    mapaVentasSimple[numVenta]                = datos;
    mapaVentasCompuesto[numVenta + '|' + sku] = datos;
  }

  // ── 2. Leer hoja Ventas SH ────────────────────────────────────
  const hSh   = shSh.getDataRange().getValues();
  const mapSh = {};
  hSh[0].forEach(function(h, i) { if (h) mapSh[String(h).trim()] = i; });

  const COL_ORDER_NUMBER = mapSh['sh_order_number'];  // D=col 4
  const COL_SKU_MAESTRO  = mapSh['sku_maestro'];       // I=col 9
  const COL_MONTO_TOTAL  = mapSh['monto_total'];       // M=col 13
  const COL_CARGO_GW     = mapSh['cargo_gateway'];     // N=col 14
  const COL_CARGO_ENVIO  = mapSh['cargo_envio'];       // Q=col 17
  const COL_NETO_PARC    = mapSh['monto_neto_parcial'];// S=col 19
  const COL_ESTADO_CONC  = mapSh['estado_conciliacion'];
  const COL_FECHA_CONC   = mapSh['fecha_conciliacion'];
  const COL_MENSAJE_CONC = mapSh['mensaje_conciliacion'];

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
  for (var i = 1; i < hSh.length; i++) {
    const estadoConc = String(hSh[i][COL_ESTADO_CONC]).trim();
    if (ESTADOS_OMITIR.has(estadoConc)) continue;

    procesadas++;
    const orderNumber = String(hSh[i][COL_ORDER_NUMBER]).trim();
    const skuMaestro  = String(hSh[i][COL_SKU_MAESTRO]).trim();

    // Búsqueda 2 niveles: compuesto → simple
    const claveComp = orderNumber + '|' + skuMaestro;
    const ventaRow  = mapaVentasCompuesto[claveComp] || mapaVentasSimple[orderNumber];

    // ── No encontrada ─────────────────────────────────────────
    if (!ventaRow) {
      shSh.getRange(i + 1, COL_ESTADO_CONC  + 1).setValue('NO ENCONTRADA EN VENTAS');
      shSh.getRange(i + 1, COL_FECHA_CONC   + 1).setValue(ahora);
      shSh.getRange(i + 1, COL_MENSAJE_CONC + 1).setValue('# Venta ' + orderNumber + ' no existe en hoja Ventas');
      noEncontradas++;
      continue;
    }

    // ── Comparar 4 campos con tolerancia ±$1 ─────────────────
    // Opción B aprobada: cargo_gateway vs Cargos porVenta
    //                    monto_neto_parcial vs Recaudado
    const montoTotal   = sh_parseNumber_(hSh[i][COL_MONTO_TOTAL]);
    const cargoGateway = sh_parseNumber_(hSh[i][COL_CARGO_GW]);
    const cargoEnvio   = sh_parseNumber_(hSh[i][COL_CARGO_ENVIO]);
    const netoParcial  = sh_parseNumber_(hSh[i][COL_NETO_PARC]);

    const diffs = [];
    if (Math.abs(montoTotal   - ventaRow.precioVenta) > TOLERANCIA)
      diffs.push('monto_total: '     + montoTotal   + ' vs ' + ventaRow.precioVenta);
    if (Math.abs(cargoGateway - ventaRow.cargosVenta) > TOLERANCIA)
      diffs.push('cargo_gateway: '   + cargoGateway + ' vs ' + ventaRow.cargosVenta);
    if (Math.abs(cargoEnvio   - ventaRow.cargoEnvio)  > TOLERANCIA)
      diffs.push('cargo_envio: '     + cargoEnvio   + ' vs ' + ventaRow.cargoEnvio);
    if (Math.abs(netoParcial  - ventaRow.recaudado)   > TOLERANCIA)
      diffs.push('monto_neto_parcial: ' + netoParcial + ' vs ' + ventaRow.recaudado);

    let nuevoEstado, mensaje;

    if (!diffs.length) {
      nuevoEstado = 'CONCILIADA';
      mensaje     = 'OK';
      conciliadas++;
    } else {
      // Solo monto_neto_parcial difiere → cargo_shopify fin de mes no confirmado
      const soloNeto = diffs.every(function(d) { return d.indexOf('monto_neto_parcial') === 0; });
      if (soloNeto) {
        nuevoEstado = 'DIFERENCIA_CONOCIDA';
        mensaje     = 'Diferencia por cargo_shopify fin de mes (no confirmado). ' + diffs.join(' | ');
        diferenciasConocidas++;
      } else {
        nuevoEstado = 'DIFERENCIA';
        mensaje     = diffs.join(' | ');
        diferencias++;
      }
    }

    shSh.getRange(i + 1, COL_ESTADO_CONC  + 1).setValue(nuevoEstado);
    shSh.getRange(i + 1, COL_FECHA_CONC   + 1).setValue(ahora);
    shSh.getRange(i + 1, COL_MENSAJE_CONC + 1).setValue(mensaje);
  }

  const resumen = {
    procesadas: procesadas, conciliadas: conciliadas,
    diferencias: diferencias, diferenciasConocidas: diferenciasConocidas,
    noEncontradas: noEncontradas
  };
  Logger.log('[SH-CONC] Resumen: ' + JSON.stringify(resumen));
  SpreadsheetApp.getActive().toast(
    'SH conciliación: ' + procesadas + ' procesadas | ' +
    conciliadas + ' ok | ' + diferencias + ' diff | ' +
    diferenciasConocidas + ' diff_conocida | ' + noEncontradas + ' no encontradas',
    'EHI', 10
  );
  return resumen;
}

// =================================================================== 
// ── Helper: parsea número formato chileno o número directo ──────
// =================================================================== 
function sh_parseNumber_(val) {
  if (val === null || val === undefined || val === '') return 0;
  if (typeof val === 'number') return Math.round(val);
  return parseInt(String(val).replace(/\./g, '').replace(/,/g, ''), 10) || 0;
}

// ===================================================================
// FIN VENTAS SH — Etapa 2: Conciliación
// ===================================================================