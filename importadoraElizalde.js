// ===================================================================
// importadoraElizalde.gs
// Ajustado EHI - registro de venta sin fila en blanco en Cta. Banco
// ===================================================================


// ===================================================================
// Función: Registra Venta
// ===================================================================
function registrarVenta() {
  var libro = SpreadsheetApp.getActiveSpreadsheet();
  var hojaVta = libro.getSheetByName('Ventas');

  hojaVta.getRange('A1').activate();
  hojaVta.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.DOWN).activate();
  var celdaActiva = hojaVta.getCurrentCell();
  var vecDatos = hojaVta.getRange(celdaActiva.getRow(), celdaActiva.getColumn(), 1, 24).getValues();

  var hojaCtaBco = libro.getSheetByName('Cta. Banco');

  // buscar última fila con dato en Q desde fila 5
  var lastRow = hojaCtaBco.getLastRow();
  var qVals = hojaCtaBco.getRange(5, 17, Math.max(lastRow - 4, 1), 1).getValues();
  var ultimaFilaConDatoQ = 4;

  for (var i = qVals.length - 1; i >= 0; i--) {
    if (String(qVals[i][0]).trim() !== "") {
      ultimaFilaConDatoQ = i + 5;
      break;
    }
  }

  var filaNext = ultimaFilaConDatoQ + 1;

  hojaCtaBco.getRange(filaNext,1).setFormula('=if(B'+filaNext+'<>"";DATE(YEAR(B'+filaNext+');MONTH(B'+filaNext+');1);"")');
  hojaCtaBco.getRange(filaNext,2).setValue(vecDatos[0][0]);   // fecha
  hojaCtaBco.getRange(filaNext,3).setValue(vecDatos[0][2]);   // #Venta
  hojaCtaBco.getRange(filaNext,4).setValue(vecDatos[0][3]);   // sku
  hojaCtaBco.getRange(filaNext,7).setValue(vecDatos[0][14]);  // canal
  hojaCtaBco.getRange(filaNext,8).setValue(vecDatos[0][12]);  // cantidad
  hojaCtaBco.getRange(filaNext,9).setFormula('=Ventas!V'+celdaActiva.getRow()); // monto recaudado

  SpreadsheetApp.flush();   // clave

  calculaSaldo();

  hojaCtaBco.getRange('Q'+filaNext).activate();

  actualizarTodosStock(vecDatos[0][14], vecDatos[0][3], vecDatos[0][13]);
}

// ===================================================================
// Función: Consulta todos los stocks multicanal
// Orden:
// 1) Shopify
// 2) Paris
// 3) Falabella
// 4) Mercado Libre
// ===================================================================
function consultaTodosStock() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const pasos = [
    {
      canal: 'Shopify',
      fnName: 'syncShopifyProductsToSheet'
    },
    {
      canal: 'Paris',
      fnName: 'paris_fetchProducts_mapSkus'
    },
    {
      canal: 'Falabella',
      fnName: 'consultaProdsFS'
    },
    {
      canal: 'Mercado Libre',
      fnName: 'consultaProdsML'
    }
  ];

  const resultados = [];

  ss.toast('Iniciando consulta de stocks multicanal...', 'EHI', 5);
  Logger.log('[EHI][STOCKS] Inicio consultaTodosStock');

  for (let i = 0; i < pasos.length; i++) {
    const paso = pasos[i];

    try {
      if (typeof this[paso.fnName] !== 'function') {
        throw new Error('No existe función: ' + paso.fnName);
      }

      ss.toast('Consultando stock: ' + paso.canal + '...', 'EHI', 5);
      Logger.log('[EHI][STOCKS] Inicio ' + paso.canal + ' -> ' + paso.fnName);

      this[paso.fnName]();

      SpreadsheetApp.flush();

      resultados.push({
        canal: paso.canal,
        estado: 'OK'
      });

      Logger.log('[EHI][STOCKS] OK ' + paso.canal);
      ss.toast('Finalizó consulta: ' + paso.canal, 'EHI', 4);

      Utilities.sleep(1500);

    } catch (e) {
      const msg = e && e.message ? e.message : String(e);

      resultados.push({
        canal: paso.canal,
        estado: 'ERROR',
        detalle: msg
      });

      Logger.log('[EHI][STOCKS] ERROR ' + paso.canal + ': ' + msg);

      ss.toast('Error consultando ' + paso.canal + '. Revisa logs.', 'EHI', 8);

      // Continúa con el siguiente marketplace
      Utilities.sleep(1500);
    }
  }

  const ok = resultados.filter(r => r.estado === 'OK').length;
  const err = resultados.filter(r => r.estado === 'ERROR').length;

  Logger.log('[EHI][STOCKS] Resumen consultaTodosStock: ' + JSON.stringify(resultados));

  ss.toast(
    'Consulta stocks finalizada: ' + ok + ' OK / ' + err + ' error(es).',
    'EHI',
    8
  );

  return resultados;
}

// ===================================================================
// Función: calcula el saldo banco
// ===================================================================
function calculaSaldo() {
  limpiaFiltro("Cta. Banco");

  var libro = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = libro.getSheetByName("Cta. Banco");
  var lastRow = hoja.getLastRow();

  if (lastRow < 6) {
    hoja.getRange(2, 9).setValue(0);   // I2
    hoja.getRange(2, 16).setValue(0);  // P2
    hoja.getRange(2, 17).setValue(0);  // Q2
    creaFiltro("Cta. Banco", 3);
    return;
  }

  var saldoInicial =
    Number(hoja.getRange(4, 9).getValue()) +
    Number(hoja.getRange(5, 9).getValue());

  var numFilasDatos = lastRow - 5;

  var ingresos = hoja.getRange(6, 9, numFilasDatos, 1).getValues();   // I
  var egresos  = hoja.getRange(6, 16, numFilasDatos, 1).getValues();  // P

  var ultimaFilaConMovimiento = -1;

  // Buscar última fila real considerando ingresos y egresos
  for (var i = numFilasDatos - 1; i >= 0; i--) {
    var rawIngreso = ingresos[i][0];
    var rawEgreso = egresos[i][0];

    var hayIngreso = rawIngreso !== "" && rawIngreso !== null;
    var hayEgreso = rawEgreso !== "" && rawEgreso !== null;

    if (hayIngreso || hayEgreso) {
      ultimaFilaConMovimiento = i;
      break;
    }
  }

  var ttlIngresos = 0;
  var ttlEgresos = 0;
  var saldo = saldoInicial;
  var vecSaldo = [];

  if (ultimaFilaConMovimiento >= 0) {
    for (var j = 0; j <= ultimaFilaConMovimiento; j++) {
      var rawIngresoFila = ingresos[j][0];
      var rawEgresoFila = egresos[j][0];

      var ingreso = Number(rawIngresoFila) || 0;
      var egreso = Number(rawEgresoFila) || 0;

      ttlIngresos += ingreso;
      ttlEgresos += egreso;

      saldo = saldo + ingreso - egreso;

      // Nunca dejar blanco dentro del tramo válido
      vecSaldo.push([saldo, ""]);
    }
  }

  hoja.getRange(2, 9).setValue(ttlIngresos);  // I2
  hoja.getRange(2, 16).setValue(ttlEgresos);  // P2
  hoja.getRange(2, 17).setValue(saldo);       // Q2

  // Limpiar todo el tramo de salida previo
  hoja.getRange(6, 17, numFilasDatos, 2).clearContent();

  // Escribir solo hasta la última fila con movimiento real
  if (vecSaldo.length > 0) {
    hoja.getRange(6, 17, vecSaldo.length, 2).setValues(vecSaldo);
  }

  creaFiltro("Cta. Banco", 3);
}


// ===================================================================
// Limpia filtro
// ===================================================================
function limpiaFiltro(nombreHoja) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(nombreHoja);
  if (!sheet) return;

  const filter = sheet.getFilter();
  if (filter) filter.remove();
}


// ===================================================================
// Crea filtro
// ===================================================================
function creaFiltro(nombreHoja, numFila) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(nombreHoja);
  if (!sheet) return;

  const filter = sheet.getFilter();
  if (!filter) {
    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();
    if (lastRow <= 0 || lastColumn <= 0) return;

    const range = sheet.getRange(numFila, 1, lastRow - numFila + 1, lastColumn);
    range.createFilter();
  }
}


// ===================================================================
// Respaldo
// ===================================================================
function respaldoImportadoraElizalde() {
  const originalSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const originalSpreadsheetName = originalSpreadsheet.getName();
  const originalSpreadsheetId = originalSpreadsheet.getId();
  const originalFile = DriveApp.getFileById(originalSpreadsheetId);
  const parentFolders = originalFile.getParents();
  const parentFolder = parentFolders.hasNext() ? parentFolders.next() : DriveApp.getRootFolder();

  const timestamp = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    'dd.MM.yyyy HH:mm:ss'
  );

  const backupSpreadsheet = SpreadsheetApp.create(
    'Backup - ' + originalSpreadsheetName + ' - ' + timestamp
  );

  const sheets = originalSpreadsheet.getSheets();

  sheets.forEach(function(sheet) {
    const backupSheet = backupSpreadsheet.insertSheet(sheet.getName());
    const range = sheet.getDataRange();
    const values = range.getValues();

    if (values.length > 0 && values[0].length > 0) {
      backupSheet.getRange(1, 1, values.length, values[0].length).setValues(values);

      const formats = range.getNumberFormats();
      backupSheet.getRange(1, 1, formats.length, formats[0].length).setNumberFormats(formats);

      const backgrounds = range.getBackgrounds();
      backupSheet.getRange(1, 1, backgrounds.length, backgrounds[0].length).setBackgrounds(backgrounds);

      const fontStyles = range.getFontFamilies();
      backupSheet.getRange(1, 1, fontStyles.length, fontStyles[0].length).setFontFamilies(fontStyles);

      const alignments = range.getHorizontalAlignments();
      backupSheet.getRange(1, 1, alignments.length, alignments[0].length).setHorizontalAlignments(alignments);
    }
  });

  const initialSheet = backupSpreadsheet.getSheets()[0];
  if (initialSheet.getLastRow() === 0 && initialSheet.getLastColumn() === 0) {
    backupSpreadsheet.deleteSheet(initialSheet);
  }

  const backupFile = DriveApp.getFileById(backupSpreadsheet.getId());
  backupFile.moveTo(parentFolder);
}


// ===================================================================
// Helpers
// ===================================================================
function getLastDataRowFromRow_(sheet, col, startRow) {
  var lastRow = sheet.getLastRow();
  if (lastRow < startRow) return startRow - 1;

  var values = sheet.getRange(startRow, col, lastRow - startRow + 1, 1).getValues();
  for (var i = values.length - 1; i >= 0; i--) {
    if (String(values[i][0] || '').trim() !== '') {
      return startRow + i;
    }
  }
  return startRow - 1;
}

function getLastDataRow_(sheet, col) {
  return getLastDataRowFromRow_(sheet, col, 1);
}

