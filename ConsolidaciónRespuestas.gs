function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Formularios de evaluación')
      .addItem('Consolidar datos', 'extractFormResponses')
      .addToUi();
}

function extractFormResponses() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var fileId = spreadsheet.getId();
  var file = DriveApp.getFileById(fileId);
  var parentFolder = file.getParents().next(); // Obtiene la primera carpeta que contiene el archivo

  // Obtiene o crea la subcarpeta "RESPUESTAS" en la carpeta principal
  var responsesFolder = getOrCreateSubfolder(parentFolder, 'RESPUESTAS');

  // Obtiene el archivo 'template' en la carpeta "RESPUESTAS"
  var templateFile = getTemplateFile(responsesFolder);

  // Función para obtener los formularios de una carpeta dada
  function getFormsInFolder(folder) {
    var forms = [];
    var files = folder.getFiles();
    while (files.hasNext()) {
      var file = files.next();
      if (file.getMimeType() === 'application/vnd.google-apps.form') {
        forms.push(file);
      }
    }
    return forms;
  }

  // Función para obtener el encabezado del formulario
  function getFormHeader(formId) {
    var form = FormApp.openById(formId);
    var lines = form.getDescription().split(/\r?\n|\r|\n/g);
    return lines;
  }

  // Solo dejar el valor numérico
  function extractNumericValue(value) {
    var match = String(value).match(/\d+/);
    return match ? parseFloat(match[0]) : '';
  }

  // Función para extraer las respuestas de un formulario
  function getFormResponses(formId) {
    var form = FormApp.openById(formId);
    var responses = form.getResponses();
    var responsesData = [];

    responses.forEach(function(response) {
      var itemResponses = response.getItemResponses();
      var responseRow = [];

      itemResponses.forEach(function(itemResponse) {
        var responseValue = itemResponse.getResponse();
        var itemType = itemResponse.getItem().getType();
        
        if (itemType === FormApp.ItemType.GRID) {
          // Si la pregunta es una cuadrícula, divide las respuestas en columnas
          var gridResponses = responseValue; // Asume que ya es un array
          var numericResponses = gridResponses.map(function(value) {
            return extractNumericValue(value);
          });
          responseRow = responseRow.concat(numericResponses);
        } else if (itemType === FormApp.ItemType.MULTIPLE_CHOICE) {
          // Para MULTIPLE_CHOICE, extrae solo el valor numérico
          var numericResponse = extractNumericValue(responseValue);
          responseRow.push(numericResponse);
        } else {
          if (Array.isArray(responseValue)) {
            // Si la respuesta es una lista o un array, conviértelo a texto separado por comas
            responseValue = responseValue.join(", ");
          } else if (typeof responseValue === 'object') {
            // Si la respuesta es un objeto complejo, conviértelo a texto
            responseValue = String(responseValue);
          }
          responseRow.push(responseValue);
        }
      });

      responsesData.push(responseRow);
    });

    return responsesData;
  }

  // Función para obtener los encabezados de un formulario
  function getFormHeaders(formId) {
    var form = FormApp.openById(formId);
    var items = form.getItems();
    var headers = ['Curso', 'Materia', 'Docente', 'Fecha de Envío'];
    
    items.forEach(function(item) {
      var title = item.getTitle().trim(); // Aplica trim a las cabeceras
      var itemType = item.getType();
      if (itemType === FormApp.ItemType.GRID) {
        // Para cuadrículas, añade una columna para cada opción en cada fila
        var gridItem = item.asGridItem();
        var rows = gridItem.getRows();
        
        rows.forEach(function(row) {
          var headerTitle = title + " - " + row;
          headers.push(headerTitle.trim()); // Aplica trim a las cabeceras
        });
      } else if (itemType !== FormApp.ItemType.PAGE_BREAK) {
        // Excluye PAGE_BREAK y maneja otros tipos
        headers.push(title);
      }
    });
    
    return headers;
  }

  // Limpia la hoja existente antes de empezar
  var sheet = spreadsheet.getSheetByName('Respuestas');
  if (sheet) {
    spreadsheet.deleteSheet(sheet); // Elimina la hoja existente si ya existe
  }
  sheet = spreadsheet.insertSheet('Respuestas');

  var firstFormProcessed = false;

  // Función para escribir las respuestas en una hoja de cálculo
  function writeResponsesToSheet(sheet, subfolderName, forms) {
    var asignaturaResponses = {};

    forms.forEach(function(file) {
      var formId = file.getId();
      var formHeader = getFormHeader(formId); // Obtiene el encabezado del formulario
      var asignatura = (formHeader[5] || '').replace(/Asignatura:/i, "").trim(); // Manejo de índices fuera de rango
      var profesor = (formHeader[6] || '').replace(/Docente:/i, "").trim(); // Manejo de índices fuera de rango
        
      var responsesData = getFormResponses(formId);
      
      if (!firstFormProcessed) {
        // Obtiene los encabezados del primer formulario y los escribe en la hoja
        var headers = getFormHeaders(formId);
        sheet.appendRow(headers);
        firstFormProcessed = true;
      }

      responsesData.forEach(function(response) {
        // Añade subcarpeta, formulario y respuestas
        sheet.appendRow([subfolderName, asignatura, profesor, new Date()].concat(response));
        
        if (!asignaturaResponses[asignatura]) {
          asignaturaResponses[asignatura] = [];
        }
        asignaturaResponses[asignatura].push([subfolderName, asignatura, profesor, new Date()].concat(response));
      });
    });

    return asignaturaResponses;
  }

  // Función para obtener o crear una subcarpeta
  function getOrCreateSubfolder(parentFolder, subfolderName) {
    var subfolderIterator = parentFolder.getFoldersByName(subfolderName);
    
    if (subfolderIterator.hasNext()) {
      // Si la carpeta ya existe, la recupera
      return subfolderIterator.next();
    } else {
      // Si la carpeta no existe, la crea
      return parentFolder.createFolder(subfolderName);
    }
  }

  // Función para obtener o crear una subcarpeta
  function createSubfolder(parentFolder, subfolderName) {
    // Borra la subcarpeta si existe y no es la carpeta 'RESPUESTAS'
    var subfolderIterator = parentFolder.getFoldersByName(subfolderName);
    if (subfolderIterator.hasNext()) {
      var subfolder = subfolderIterator.next();
      
      if (subfolderName !== 'RESPUESTAS') {
        Logger.log(subfolderName);
        deleteAllFilesInFolder(subfolder);
        Drive.Files.remove(subfolder.getId()); // Elimina permanentemente la carpeta existente
      }
    }

    // Crea una nueva subcarpeta si no existe
    return parentFolder.createFolder(subfolderName);
  }

   // Función para eliminar todos los archivos y subcarpetas en una carpeta
  function deleteAllFilesInFolder(folder) {
    var files = folder.getFiles();
    while (files.hasNext()) {
      var file = files.next();
      file.setTrashed(true); // Mueve el archivo a la papelera
      Drive.Files.remove(file.getId()); // Elimina permanentemente el archivo
    }
    
    var subfolders = folder.getFolders();
    while (subfolders.hasNext()) {
      var subfolder = subfolders.next();
      deleteAllFilesInFolder(subfolder); // Recursivamente borra subcarpetas
      subfolder.setTrashed(true); // Mueve la carpeta a la papelera
      Drive.Files.remove(subfolder.getId()); // Elimina permanentemente la subcarpeta
    }
  }


  // Función para obtener el archivo 'template' en la carpeta "RESPUESTAS"
  function getTemplateFile(folder) {
    var files = folder.getFilesByName('template');
    if (files.hasNext()) {
      return files.next();
    } else {
      throw new Error('No se encontró el archivo template en la carpeta RESPUESTAS');
    }
  }

  // Función para crear una nueva planilla a partir del template
  function createNewSpreadsheetFromTemplate(templateFile, newName, parentFolder) {
    var newFile = templateFile.makeCopy(newName, parentFolder);
    return SpreadsheetApp.openById(newFile.getId());
  }

  function refreshData(spreadsheet,curso,asignatura,anho){
    var filtrosSheet = spreadsheet.getSheetByName('Filtros');
    var informeSheet = spreadsheet.getSheetByName('Informe');
    filtrosSheet.getRange('curso').setValue(curso); // Curso
    SpreadsheetApp.flush();
    filtrosSheet.getRange('asignatura').setValue(asignatura); // Asignatura
    filtrosSheet.getRange('anho').setValue(anho); // Año
    refreshFormulas(filtrosSheet,2);
    refreshFormulas(informeSheet,3);
    refreshFormulas(informeSheet,4);
    var sugerenciasRange = informeSheet.getRange("sugerencias");
    sugerenciasRange.setValue(filtrosSheet.getRange("sugerencias_formula").getValue())
  }

  function refreshFormulas(sheet, column) {
    var lastRow = sheet.getLastRow();

    // Obtener el rango de la columna
    var range = sheet.getRange(1, column, lastRow);
    var formulas = range.getFormulas(); // Obtiene todas las fórmulas

    for (var i = 0; i < formulas.length; i++) {
      if (formulas[i][0] !== '') {
        // Solo reescribir si hay una fórmula
        range.getCell(i + 1, 1).setFormula(formulas[i][0]);
      }
    }
  }

  // Obtiene todas las subcarpetas en la carpeta principal y ordénalas alfabéticamente
  var subfolders = [];
  var folderIterator = parentFolder.getFolders();
  while (folderIterator.hasNext()) {
    subfolders.push(folderIterator.next());
  }
  subfolders.sort(function(a, b) {
    return a.getName().localeCompare(b.getName());
  });

  subfolders = subfolders.filter(function(folder) {
    return folder.getName().toLowerCase().includes('curso');
  });
  var firstAsignaturaProcessed = true;
  subfolders.forEach(function(subfolder) {
    var subfolderName = subfolder.getName();
    var forms = getFormsInFolder(subfolder);

    // Crear subcarpeta de curso dentro de "RESPUESTAS" si no existe
    var courseResponseFolder = createSubfolder(responsesFolder, subfolderName);
    
    var asignaturaResponses = writeResponsesToSheet(sheet, subfolderName, forms);

    for (var asignatura in asignaturaResponses) {
      var newSpreadsheet = createNewSpreadsheetFromTemplate(templateFile, asignatura, courseResponseFolder);
      // Limpia la hoja existente antes de empezar
      var newSheet = newSpreadsheet.getSheetByName('Respuestas');
      if (newSheet) {
        newSpreadsheet.deleteSheet(newSheet); // Elimina la hoja existente si ya existe
      }
      newSheet = newSpreadsheet.insertSheet('Respuestas');

      // Copiar las cabeceras desde la hoja principal
      var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      newSheet.appendRow(headers);
      
      var responses = asignaturaResponses[asignatura];
      responses.forEach(function(response) {
        newSheet.appendRow(response);
      });
      if(firstAsignaturaProcessed)
        refreshData(spreadsheet,subfolderName,asignatura,new Date(responses[0][3]).getFullYear());
      refreshData(newSpreadsheet,subfolderName,asignatura,new Date(responses[0][3]).getFullYear());
      firstAsignaturaProcessed = false;
    }
  });
}

function onEdit(e) {
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var filtrosSheet = activeSpreadsheet.getSheetByName('Filtros');
  var informeSheet = activeSpreadsheet.getSheetByName('Informe');

  var editedRange = e.range;
  // Obtiene todos los rangos nombrados en la hoja activa
  var namedRanges = filtrosSheet.getNamedRanges();
  var rangeName = '';

  // Busca si el rango editado está dentro de alguno de los rangos nombrados
  namedRanges.forEach(function(namedRange) {
    var range = namedRange.getRange();
    if (isRangeIntersect(range, editedRange)) {
      rangeName = namedRange.getName();
    }
  });

  // Imprime el nombre del rango editado en el registro
  if (rangeName == "asignatura") {
    var sugerenciasRange = informeSheet.getRange("sugerencias");
    sugerenciasRange.setValue(filtrosSheet.getRange("sugerencias_formula").getValue())
    SpreadsheetApp.flush();
  } 
  
}

// Función para verificar si dos rangos se intersectan
function isRangeIntersect(range1, range2) {
  var start1 = range1.getRow();
  var end1 = start1 + range1.getNumRows() - 1;
  var start2 = range2.getRow();
  var end2 = start2 + range2.getNumRows() - 1;

  var colStart1 = range1.getColumn();
  var colEnd1 = colStart1 + range1.getNumColumns() - 1;
  var colStart2 = range2.getColumn();
  var colEnd2 = colStart2 + range2.getNumColumns() - 1;

  return !(end1 < start2 || end2 < start1 || colEnd1 < colStart2 || colEnd2 < colStart1);
}
