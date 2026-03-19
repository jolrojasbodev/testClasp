// main.gs

/**
 * Procesa la hoja activa para sumar los valores de cada fila
 * y colocar el resultado en una nueva columna llamada "SUMATORIA".
 */
function realizarSumatoria() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  
  if (data.length < 1) return; // Hoja vacía

  const headers = data[0];
  let sumatoriaColIndex = headers.indexOf("SUMATORIA");

  // Si no existe la columna, la creamos al final
  if (sumatoriaColIndex === -1) {
    sumatoriaColIndex = headers.length;
    sheet.getRange(1, sumatoriaColIndex + 1).setValue("SUMATORIA")
         .setBackground("#f3f3f3")
         .setFontWeight("bold");
  }

  // Preparamos los resultados (empezamos desde la fila 2 para saltar el encabezado)
  const resultados = [];
  
  for (let i = 1; i < data.length; i++) {
    let filaSuma = 0;
    for (let j = 0; j < data[i].length; j++) {
      // Sumamos solo si es un número y no es la propia columna de sumatoria
      if (typeof data[i][j] === 'number' && j !== sumatoriaColIndex) {
        filaSuma += data[i][j];
      }
    }
    resultados.push([filaSuma]);
  }

  // Escribimos todos los resultados de una sola vez por eficiencia
  if (resultados.length > 0) {
    sheet.getRange(2, sumatoriaColIndex + 1, resultados.length, 1).setValues(resultados);
  }
  
  SpreadsheetApp.getUi().alert("¡Sumatoria completada con éxito!");
}

/**
 * Crea un menú personalizado al abrir la hoja de cálculo.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Funciones Personalizadas')
    .addItem('Calcular Sumatoria', 'realizarSumatoria')
    .addToUi();
}