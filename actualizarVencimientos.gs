/* Sheets VENCIMIENTOS */

/*Función Ordenar y Actualizar Vencimientos del Formulario */
function actualizarVencimientos() {
  let spreadsheet = SpreadsheetApp.getActive();
  let sheetBaseForm = SpreadsheetApp.getActive().getSheetByName('BASE FORM')

  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('BASE FORM'), true);
  spreadsheet.getRange('BASE FORM!A3:D').clearContent();
  spreadsheet.getActiveSheet().insertRowsAfter(spreadsheet.getActiveSheet().getMaxRows(), 1000);

  /* Limpiamos el filtrado de la Hoja FORMULARIO */
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('FORMULARIO'), true);
  let limpiarFiltroFormulario = SpreadsheetApp.newFilterCriteria()
    .build();
  spreadsheet.getSheetByName('FORMULARIO').getFilter().setColumnFilterCriteria(5, limpiarFiltroFormulario);

  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('FORMULARIO'), true);
  spreadsheet.getRange('B2:C2').copyTo(spreadsheet.getRange('B2:C'), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('D:D').setHorizontalAlignment('left');
  spreadsheet.getRange('E:E').setNumberFormat('mmm"-"yyyy').setHorizontalAlignment('center');
  spreadsheet.getRange('B:B').activate();

  /* Hacemos copia respaldo en la misma tabla */
  spreadsheet.getRange('FORMULARIO!B2:D').copyTo(spreadsheet.getRange('I2'), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);

  /* limpiamos el filtro de BASE FORM */
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('BASE FORM'), true);
  let filtroBaseForm = SpreadsheetApp.newFilterCriteria()
    .build();
  spreadsheet.getActiveSheet().getFilter().setColumnFilterCriteria(4, filtroBaseForm);

  spreadsheet.getRange('FORMULARIO!B2:C2').copyTo(spreadsheet.getRange('FORMULARIO!B2:C'), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);

  let meses = spreadsheet.getRange('BASE FORM!C1:D1').getValue();//tomamos los valores de la formula =MAYUSC(CONCATENAR(TEXTO(HOY();"mmm");" - ";TEXTO(HOY()+30;"mmm");" - ";TEXTO(HOY()+60;"mmm")))
  let mesVencimiento = meses.split(' - ');//separamos los datos para tener 3 valores especificos

  /* Lo siguientes sirve para tomar el dato del año vigente y poder formular el dato del siguiente año  */
  let fecha = spreadsheet.getRange('BORRADOR!A2').getValue().split(' - ');// seleccionamos el dato de la fecha y la separamos
  let seleccionarFecha = fecha[0].split('/')// ejecutamos otra separacion en el primer dato del array
  let separarYear = parseFloat(seleccionarFecha[2])// parseamos el tercer dato del array
  let yearVigente = separarYear // constituimos la variable del año vigente
  let nextYear = separarYear + 1 // constituimos la variable del siguiente año

  /* Creamos una funcion que nos va a servir para filtrar las fechas que tenemos en rangoVencimientos, para al final convertirlas en un array, unificando formatos de fechas obtenido en el getValue de la variable meses */
  function filtrarFechas(mes1, mes2, mes3) {

    // Creamos un array para contener todos los datos de la columna VENCIMIENTOS
    let rangoVencimientos = spreadsheet.getRange('FORMULARIO!E3:E').getValues()

    // Inicializamos un nuevo array para almacenar los elementos aplanados
    let vencimientoTodos = [];

    // Iteramos a través de los arrays internos y agregamos sus elementos al array aplanado
    for (i = 0; i < rangoVencimientos.length; i++) {
      for (j = 0; j < rangoVencimientos[i].length; j++) {
        vencimientoTodos.push(rangoVencimientos[i][j]);
      }
    }
    // Creamos un arreglo auxiliar para verificar duplicados
    let fechasSinDuplicados = vencimientoTodos.reduce(function (resultado, fecha) {
      let isoString = fecha.toISOString();
      if (!resultado.includes(isoString)) {
        resultado.push(isoString);
      }
      return resultado;
    }, []);
    // Convertimos el arreglo de cadenas ISO nuevamente en objetos Date
    fechasSinDuplicados = fechasSinDuplicados.map(function (isoString) {
      return new Date(isoString);
    });

    // Arreglo para almacenar las fechas convertidas
    let fechasConvertidas = [];

    // Función para convertir una fecha individual
    function convertirFecha(fecha) {
      let meses = ["ENE", "FEB", "MAR", "ABR", "MAY", "JUN", "JUL", "AGO", "SEPT", "OCT", "NOV", "DIC"];
      let mesAbreviado = meses[fecha.getMonth()];
      let ano = fecha.getFullYear();
      return mesAbreviado + '-' + ano;
    }
    // Iterar a través de todas las fechas originales y convertirlas
    for (i = 0; i < fechasSinDuplicados.length; i++) {
      let fechaConvertida = convertirFecha(fechasSinDuplicados[i]);
      fechasConvertidas.push(fechaConvertida);
    }
    /* Con esta funcion eliminamos los datos del array, array que vamos a ocultar despues en el filtro */
    fechasConvertidas = fechasConvertidas.filter(function (fecha) {
      return fecha !== mes1 && fecha !== mes2 && fecha !== mes3;
    });
    /* Filtro que nos sirve para ocultar el array fechasConvertidas, asi dejando visibles las que eliminamos anteriormente */
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('FORMULARIO'), true);
    let ocultarVencimientos = SpreadsheetApp.newFilterCriteria().setHiddenValues(fechasConvertidas).build();
    /* Iniciamos con ambos modos (por spreasheet o por sheet) por si falla uno, se aplica desde el otro metodo */
    spreadsheet.getActiveSheet().getFilter().setColumnFilterCriteria(5, ocultarVencimientos);

    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('BASE FORM'), true);
    spreadsheet.getRange('\'FORMULARIO\'!B1:E').copyTo(spreadsheet.getRange('A2'), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    spreadsheet.getRange('D:D').setNumberFormat('mmm"-"yyyy');
    spreadsheet.getSheetByName('FORMULARIO').getFilter().setColumnFilterCriteria(5, limpiarFiltroFormulario);
  }

  /* Establecemos distintos Casos de los meses proximos a vencer */
  if (mesVencimiento[0] == 'NOV') {
    /* Caso 1: mes1 y mes2 van a ser del año vigente, mientras que mes3 corresponde al siguiente año */
    let mes1 = mesVencimiento[0] + '-' + yearVigente;
    let mes2 = mesVencimiento[1] + '-' + yearVigente;
    let mes3 = mesVencimiento[2] + '-' + nextYear;
    // ejecutamos la funcion
    filtrarFechas(mes1, mes2, mes3);
  } else if (mesVencimiento[0] == 'DIC') {
    /* Caso 2: mes1 es Diciembre, por ende mes2 y mes3 son del siguiente año */
    let mes1 = mesVencimiento[0] + '-' + yearVigente;
    let mes2 = mesVencimiento[1] + '-' + nextYear;
    let mes3 = mesVencimiento[2] + '-' + nextYear;
    // ejecutamos la funcion
    filtrarFechas(mes1, mes2, mes3);
  } else {
    /* Caso 3, es que todas las variables de 'mes' corresponden al año vigente */
    let mes1 = mesVencimiento[0] + '-' + yearVigente;
    let mes2 = mesVencimiento[1] + '-' + yearVigente;
    let mes3 = mesVencimiento[2] + '-' + yearVigente;
    // ejecutamos la funcion
    filtrarFechas(mes1, mes2, mes3);
  }

  /* Removemos duplicados */
  let vencimientos = spreadsheet.getRange('BASE FORM!A3:D');
  vencimientos.removeDuplicates([1, 2, 3]).activate();
  /* Eliminamos filas vacias */
  spreadsheet.getRange('BASE FORM!A2').activate();
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.DOWN).activate();
  let ultimaFila = (spreadsheet.getCurrentCell().getRowIndex()) + 1;
  let cantidadFilasEliminar = vencimientos.getLastRow() - ultimaFila;
  console.log('Ultima Fila: ' + ultimaFila, 'Rango a Eliminar: ' + cantidadFilasEliminar)

  /* AQUI ELIMINAMOS LAS FILAS, 
  ANTE ALGUN ERROR RELACIONADO,
  COMENTAR LA SIGUIENTE LINEA DE CODIGO */
  sheetBaseForm.deleteRows(ultimaFila, (cantidadFilasEliminar + 1))

  /* Ultimos detalles para la planilla */
  spreadsheet.getRange('A:A').setHorizontalAlignment('center');
  spreadsheet.getRange('B:B').setHorizontalAlignment('left');
  spreadsheet.getRange('C2:C').setHorizontalAlignment('left');
  spreadsheet.getRange('D:D').setHorizontalAlignment('center');

  /* Ultimos detalles para finalizar la planilla */
  let menorAMayor = spreadsheet.getRange('BASE FORM!A2:D');
  menorAMayor.getFilter().sort(1, true);
  spreadsheet.getRange('A2:D').setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);// y por ultimo insertamos bordes a todas las filas seleccionadas

  Utilities.sleep(500)

  createPDF(ssId, sheetBaseForm, pdfName)
};
