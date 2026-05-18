function onEdit(e) {
  var sheet = e.source.getActiveSheet();
  var range = e.range;
  if (sheet.getName()==="Cta. Banco") {
    if (range.getColumn() === 9 || range.getColumn() === 16) { //Edita la columna I: Monto Venta o P: Monto Gasto
      calculaSaldo();
    } else if (range.getColumn() === 12 && range.getRow() >= 6){ //Edita la columna L: Tipo
      var value = range.getValue();
      var validationRange;      
      if (value === "Variable") {
        validationRange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Parámetros").getRange("I2:I8");
      } else if (value === "Fijo") {
        validationRange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Parámetros").getRange("H2:H7");
      } else if (value === "Inversiones") {
        validationRange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Parámetros").getRange("J2:J6");
      }
      // Establece la validación de datos en la columna M de la misma fila
      if (validationRange) {
        var rule = SpreadsheetApp.newDataValidation()
          .requireValueInRange(validationRange)
          .setAllowInvalid(false)
          .build();
        sheet.getRange(range.getRow(), 13).setDataValidation(rule);
      } else {
        // Elimina la validación si no hay selección válida
        sheet.getRange(range.getRow(), 13).clearDataValidations();
      }
    }
  }
}

/* Crea la validación de datos para todas las filas de la hoja Cta. Banco, columna M: tipo (13) */
function crearValidacionDeDatos() {
  var hojaBanco = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Cta. Banco");
  var hojaParametros = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Parámetros");
  var lastRow = hojaBanco.getLastRow();
  for (var i = 6; i <= lastRow; i++) {
    var celdaL = hojaBanco.getRange("L" + i).getValue();
    var rangoValidacion = null;
    if (celdaL === "Variable") {
      rangoValidacion = hojaParametros.getRange("I2:I8");
    } else if (celdaL === "Fijo") {
      rangoValidacion = hojaParametros.getRange("H2:H7");
    } else if (celdaL === "Inversiones") {
      rangoValidacion = hojaParametros.getRange("J2:J6");
    }    
    if (rangoValidacion) {
      var reglaValidacion = SpreadsheetApp.newDataValidation()
        .requireValueInRange(rangoValidacion)
        .setAllowInvalid(false)
        .build();
      hojaBanco.getRange("M" + i).setDataValidation(reglaValidacion);
    } else {
      hojaBanco.getRange("M" + i).clearDataValidations().clearContent();
    }
  }
}
