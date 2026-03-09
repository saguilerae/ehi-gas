/**
 * Función principal del Proyecto EHI para activar el modo vacaciones.
 * Se activa desde el menú "Importadora Elizalde".
 */
function lead_time_logic() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shParam = ss.getSheetByName("Parámetros");
  
  // 1. Buscamos la última fila real en B (como ya validamos)
  const dataB = shParam.getRange("B:B").getValues();
  let ultimaFilaReal = 0;
  for (let i = dataB.length - 1; i >= 0; i--) {
    if (dataB[i][0] !== "" && dataB[i][0] !== null) { ultimaFilaReal = i + 1; break; }
  }

  const rowData = shParam.getRange(ultimaFilaReal, 1, 1, 6).getValues()[0];
  const indicador = rowData[3]; // Columna D (ON/OFF)
  const estado    = rowData[4]; // Columna E (Pendiente/Procesado)
  const dias      = rowData[5]; // Columna F (Días Lead Time)

  const isON = (String(indicador).toUpperCase().trim() === "ON" || indicador === true);
  const isPendiente = (String(estado).toUpperCase().trim() === "PENDIENTE");

  if (isON && isPendiente) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert("⚠️ ATENCIÓN", "¿Deseas iniciar el barrido masivo de Lead Time (" + dias + " días) en todos los canales?", ui.ButtonSet.YES_NO);
    
    if (response == ui.Button.YES) {
      try {
        // Ejecución del Barrido en Paris (Usando tu estrategia de SKU/SKU_MKP)
        EHI_Barrido_Paris(dias);
        
        // Aquí irán los siguientes barridos (FS y ML)
        Logger.log("Sincronización de Paris completada.");

        // 2. Marcamos como PROCESADO para evitar re-ejecución
        shParam.getRange(ultimaFilaReal, 5).setValue("Procesado");
        ui.alert("✅ Proceso completado. Paris actualizado y parámetros cerrados.");
        
      } catch (e) {
        ui.alert("❌ Error durante el barrido: " + e.message);
      }
    }
  } else {
    Logger.log("Condiciones no cumplidas: ON=" + isON + ", Pendiente=" + isPendiente);
  }
}

/**
 * Ejecuta el barrido masivo de Lead Time para Paris.
 * Basado en la estrategia de SKU_MKP y bloques (chunks) de Importadora Elizalde.
 */
function EHI_Barrido_Paris(dias) {
  Logger.log("🚀 Iniciando Barrido Final en Paris (PM)...");

  const warehouse = paris_getWarehouse_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shPM = ss.getSheetByName("Prods. PM");
  const dataPM = shPM.getDataRange().getValues();
  
  // Usamos el stockMap para mantener el stock real de la tienda
  const stockMap = paris_fetchAllStockMap_(); 
  
  let updates = [];

  for (let i = 1; i < dataPM.length; i++) {
    const skuMkp = String(dataPM[i][0]).trim();    // Columna A (Solo para cruzar datos)
    const skuSeller = String(dataPM[i][1]).trim(); // Columna B (El que manda)
    
    // Si tenemos el SKU_SELLER y existe en el mapa de stock de Paris
    if (skuSeller && stockMap.hasOwnProperty(skuMkp)) {
      updates.push({
        "sku_seller": skuSeller,      
        "quantity": stockMap[skuMkp], 
        "lead_time": Number(dias),    
        "warehouse": warehouse
      });
    }
  }

  if (updates.length > 0) {
    Logger.log("Enviando " + updates.length + " SKUs a Paris...");
    paris_postStockInChunks_(updates); 
    Logger.log("✅ Barrido masivo de Paris completado exitosamente.");
  }
}
