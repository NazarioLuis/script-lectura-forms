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

  // Función para listar formularios en una carpeta dada
  function listFormsInFolder(folder, sheet) {
    sheet.clear(); // Borra el contenido actual
    sheet.appendRow(['Nombre del Formulario', 'Enlace del Formulario', 'Asignatura', 'Profesor']);

    var files = folder.getFiles();
    while (files.hasNext()) {
      var file = files.next();
      if (file.getMimeType() === 'application/vnd.google-apps.form') {
        var fileName = file.getName();
        var fileUrl = file.getUrl();
        var formId = file.getId();
        var formHeader = getFormHeader(formId); // Obtiene el encabezado del formulario
        var asignatura = formHeader[5] || ''; // Manejo de índices fuera de rango
        var profesor = formHeader[6] || ''; // Manejo de índices fuera de rango
        sheet.appendRow([fileName, fileUrl, asignatura, profesor]);
      }
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

  var subfolders = parentFolder.getFolders();
  if (subfolders.hasNext()) {
    var firstSubfolder = subfolders.next();
    var firstSubfolderName = firstSubfolder.getName();
    activeSheet.setName(firstSubfolderName); // Renombra la hoja activa con el nombre de la primera subcarpeta
    listFormsInFolder(firstSubfolder, activeSheet); // Lista los formularios en la primera subcarpeta en la hoja activa

    while (subfolders.hasNext()) {
      var subfolder = subfolders.next();
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
