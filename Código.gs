function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Formularios de evaluación')
      .addItem('Listar Enlaces de Formularios', 'listFormsInSubfolders')
      .addItem('Generar Vista de Tarjetas', 'generateCardView')
      .addToUi();
}

function listFormsInSubfolders() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var fileId = spreadsheet.getId();
  var file = DriveApp.getFileById(fileId);
  var parentFolder = file.getParents().next(); // Obtiene la primera carpeta que contiene el archivo

  // Función para obtener el encabezado del formulario
  function getFormHeader(formId) {
    var form = FormApp.openById(formId);
    var lines = form.getDescription().split(/\r?\n|\r|\n/g);
    return lines;
  }

  // Función para verificar si el formulario acepta respuestas
  function isAcceptingResponses(formId) {
    var form = FormApp.openById(formId);
    return form.isAcceptingResponses();
  }

  //Regla para convertir celda en checkbox
  var rule = SpreadsheetApp.newDataValidation().requireCheckbox().build();

  // Función para listar formularios en una carpeta dada
  function listFormsInFolder(folder, sheet) {
    sheet.clear(); // Borra el contenido actual

    // Limpia las reglas de validación de toda la columna "Acepta Respuestas"
    sheet.getRange('E:E').clearDataValidations();

    sheet.appendRow(['Nombre del Formulario', 'Enlace del Formulario', 'Asignatura', 'Profesor', 'Acepta Respuestas','ID']);
    
    var files = folder.getFiles();

    var row = 2; // La fila donde empiezan los datos
    while (files.hasNext()) {
      var file = files.next();
      if (file.getMimeType() === 'application/vnd.google-apps.form') {
        var fileName = file.getName();
        var fileUrl = file.getUrl();
        var formId = file.getId();
        var formHeader = getFormHeader(formId); // Obtiene el encabezado del formulario
        var asignatura = formHeader[5] || ''; // Manejo de índices fuera de rango
        var profesor = formHeader[6] || ''; // Manejo de índices fuera de rango
        var acceptingResponses = isAcceptingResponses(formId); // Verifica si el formulario acepta respuestas
        sheet.appendRow([fileName, fileUrl, asignatura, profesor, acceptingResponses,formId]);
        var cell = sheet.getRange(row, 5);
        cell.setDataValidation(rule);
        row++;
      }
    }

    // Ajusta el ancho de las columnas al contenido
    for (var col = 1; col <= 5; col++) {
      if (col < 3) { 
        sheet.setColumnWidth(col, 200); 
      } else {
        sheet.autoResizeColumn(col);
      }
    }
    //Oculta el id del formulario
    sheet.hideColumns(6);

    // Ordena los datos por la columna "Asignatura" (columna C)
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      var range = sheet.getRange('A2:F' + lastRow); // Rango que incluye los datos
      range.sort([{column: 3, ascending: true}]); // Ordena por la columna C (Asignatura)
    }
  }

  // Limpia las hojas existentes antes de empezar, excepto la activa
  var sheets = spreadsheet.getSheets();
  var activeSheet = spreadsheet.getActiveSheet();
  for (var i = sheets.length - 1; i >= 0; i--) {
    if (sheets[i].getSheetId() !== activeSheet.getSheetId()) {
      spreadsheet.deleteSheet(sheets[i]);
    }
  }

  var subfoldersArray = [];
  var subfolders = parentFolder.getFolders();
  while (subfolders.hasNext()) {
    var subfolder = subfolders.next();
    subfoldersArray.push(subfolder);
  }

  // Ordena las subcarpetas alfabéticamente por nombre
  subfoldersArray.sort(function(a, b) {
    return a.getName().localeCompare(b.getName());
  });

  if (subfoldersArray.length > 0) {
    var firstSubfolder = subfoldersArray[0];
    var firstSubfolderName = firstSubfolder.getName();
    activeSheet.setName(firstSubfolderName); // Renombra la hoja activa con el nombre de la primera subcarpeta
    listFormsInFolder(firstSubfolder, activeSheet); // Lista los formularios en la primera subcarpeta en la hoja activa

    for (var i = 1; i < subfoldersArray.length; i++) {
      var subfolder = subfoldersArray[i];
      var subfolderName = subfolder.getName();
      var sheetName = subfolderName.length > 99 ? subfolderName.substring(0, 99) : subfolderName; // Google Sheets permite nombres de hasta 100 caracteres
      var existingSheet = spreadsheet.getSheetByName(sheetName);
      if (existingSheet) {
        spreadsheet.deleteSheet(existingSheet); // Elimina la hoja existente si ya existe
      }
      var newSheet = spreadsheet.insertSheet(sheetName);
      listFormsInFolder(subfolder, newSheet); // Llama recursivamente para explorar subcarpetas
    }
  }
}



// Función que maneja los cambios en la hoja de cálculo
function onEdit(e) {
  var sheet = e.range.getSheet();
  var range = e.range;
  var row = range.getRow();
  var col = range.getColumn();

  // Verifica si el cambio está en la columna "Acepta Respuestas" (columna 5)
  if (col === 5 && row > 1) {
    var acceptingResponses = range.getValue();
    var formId = sheet.getRange(row, 6).getValue(); // Obtiene el ID del formulario de la columna 6
    var form = FormApp.openById(formId);
    
    // Actualiza el estado de aceptación de respuestas del formulario
    form.setAcceptingResponses(acceptingResponses);
  }
}


function generateCardView() {
  var htmlOutput = HtmlService.createTemplateFromFile('Index');
  htmlOutput.data = getData(); // Pasa los datos al template HTML
  var html = htmlOutput.evaluate()
      .setTitle('Vista de Tarjetas de Formularios')
      .setWidth(800)
      .setHeight(600);

  SpreadsheetApp.getUi().showModalDialog(html, 'Vista de Tarjetas');
}

function getData() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = spreadsheet.getSheets();
  var data = {};

  sheets.forEach(sheet => {
    var sheetName = sheet.getName();
    var values = sheet.getDataRange().getValues();
    data[sheetName] = values.slice(1).map(row => ({
      nombre: row[0],
      enlace: row[1],
      asignatura: row[2],
      profesor: row[3]
    }));
  });

  return data;
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function doGet() {
  var htmlOutput = HtmlService.createTemplateFromFile('Index');
  htmlOutput.data = getData(); // Pasa los datos al template HTML
  return htmlOutput.evaluate()
      .setTitle('Vista de Tarjetas de Formularios')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

}
