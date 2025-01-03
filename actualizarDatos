/* Funcion Actualizar Datos */
function actualizarDatos() {
  let spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('BORRADOR'), true);
  /* Separamos el GTIN del nombre del producto */
  spreadsheet.getRange('B:B').splitTextToColumns(']');
  spreadsheet.getRange('B:B').splitTextToColumns('[');
  spreadsheet.getRange('D:D').clear();
  spreadsheet.getRange('E6').copyTo(spreadsheet.getRange('D6:D'), SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);

  /* Copiamos los datos de la Hoja BORRADOR */
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('BD'), true);
  spreadsheet.getRange('BORRADOR!A6:B').copyTo(spreadsheet.getRange('B2'), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('BORRADOR!D6:D').copyTo(spreadsheet.getRange('D2'), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  spreadsheet.getRange('D:D').setNumberFormat('General');
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('FORMULARIO'), false)
};
