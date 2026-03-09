// Obtiene la configuración Shopify
function getShopifyConfig() {
  const props = PropertiesService.getScriptProperties();
  return {
    storeDomain: props.getProperty('SHOPIFY_STORE_DOMAIN'), // p.ej. quickdropshipping.myshopify.com
    token: props.getProperty('SHOPIFY_ADMIN_TOKEN'),
    apiVersion: '2024-10', // puedes ajustarlo si Shopify cambia
  };
}

// Invoca a los parametros de Shopify
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

/************ 1) CONSULTAR PRODUCTOS → SHEET ************/
// Obtiene los headers
function getHeaderIndex_(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const idx = {};
  headers.forEach(function(name, i) {
    idx[name] = i;
  });
  return idx;
}


// Hoja donde trabajar
const SHEET_NAME = 'Prods. Shopify';  // cambia al nombre que uses

// Orden de columnas requerido
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

// Trae TODOS los productos (hasta 250 por llamada) y los deja en la hoja
// Trae TODOS los productos y los deja en la hoja
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

// Paginación sencilla (hasta que no haya "link: rel=next")
// Paginación sencilla (hasta que no haya "link: rel=next")
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


/************ 2) ACTUALIZAR PRECIOS Y STOCK DESDE LA HOJA ************/

// Actualiza Variant Price, Compare At Price y Stock para las filas con "Actualizar" = TRUE
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

// Obtiene el primer location_id y lo guarda para reutilizarlo
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

// Función para actualizar solo STOCK
function updateShopifyStockOnly() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME); // 'Prods. Shopify'
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

// Función para actualizar solo PRECIO PRINCIPAL
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

// Actualiza sólo el estado del producto
function updateShopifyStatusOnly() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(SHEET_NAME); // "Prods. Shopify"
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
  const sheet = ss.getSheetByName(SHEET_NAME); // 'Prods. Shopify'
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

/**
 * Actualiza el título principal de los productos en Shopify
 * basándose en la columna "Title" de la Google Sheet.
 */
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

// ========================================================================================
// UTILITARIOS
// ========================================================================================

// Muestra el ID de Localicación de mi Shotify
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

// Prueba la actualización de SEOs
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