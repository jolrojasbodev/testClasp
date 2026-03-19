/**
 * Archivo: main.ts
 * PROYECTO: Automatización Energuaviare - Fase 1
 * Nota: Este archivo es .ts para aprovechar el autocompletado de VS Code.
 * Al hacer 'clasp push', se convertirá automáticamente en .gs en Google.
 */

function onOpen(): void {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🚀 Gestión Energuaviare')
    .addItem('F1: Sumar Selección', 'sumarSeleccion')
    .addToUi();
}

/**
 * Suma los valores numéricos del rango seleccionado.
 */
function sumarSeleccion(): void {
  const range = SpreadsheetApp.getActiveRange();
  
  if (!range) {
    SpreadsheetApp.getUi().alert('Error', 'No hay un rango seleccionado.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const values = range.getValues() as any[][];
  let sumaTotal = 0;
  let celdasContadas = 0;

  values.forEach(row => {
    row.forEach(cell => {
      if (typeof cell === 'number' && !isNaN(cell)) {
        sumaTotal += cell;
        celdasContadas++;
      }
    });
  });

  const ui = SpreadsheetApp.getUi();
  if (celdasContadas === 0) {
    ui.alert('Aviso', 'No se encontraron valores numéricos en la selección.', ui.ButtonSet.OK);
  } else {
    ui.alert(
      'Resultado de la Suma', 
      `Se sumaron ${celdasContadas} celdas.\n\nTotal: ${sumaTotal.toFixed(2)}`, 
      ui.ButtonSet.OK
    );
  }
}