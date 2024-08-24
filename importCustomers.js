/**
 * @OnlyCurrentDoc
 * Este script gestiona la importación de datos de Clientes desde archivos XML en Google Drive.
 * Incluye funciones para obtener archivos, procesar XML y manejar la interfaz de usuario.
 */

const CUST_FOLDER_ID   = '1jFAEy1VhshwOkM2p9IHa0LwSEUEIEKGF2jm-nQxIV-c'; // ID de la carpeta de clientes en Drive
const CUST_SHEET_NAME  = 'Clientes'; // Nombre de la hoja de clientes

/**
 * Importa datos de clientes desde archivos XML en Google Drive.
 */
function importCustomers() {
  const ui    = SpreadsheetApp.getUi();
  const files = getFilesFromFolderCust(CUST_FOLDER_ID, '.xml');

  if (files.length === 0) {
    ui.alert('Error', 'No se encontraron archivos XML en la carpeta especificada.', ui.ButtonSet.OK);
    return;
  }

  const latestFile = getLatestFileCust(files);
  const response   = ui.alert(
    'Importar Datos de Clientes',
    `El último archivo encontrado es:\n\n"${latestFile.getName()}" (modificado el ${latestFile.getLastUpdated().toLocaleString()}).\n\n¿Deseas importar este archivo?`,
    ui.ButtonSet.YES_NO
  );

  if (response === ui.Button.NO) {
    const selectedFile = getFileFromUserCust(ui, files);
    if (!selectedFile) return; // El usuario canceló la selección
    prepareAndImportXMLFileCust(selectedFile, ui);
  } else {
    prepareAndImportXMLFileCust(latestFile, ui);
  }
}

/**
 * Prepara e importa un archivo XML.
 * @param {File} file - El archivo XML a importar.
 * @param {Object} ui - La interfaz de usuario para mostrar alertas.
 */
function prepareAndImportXMLFileCust(file, ui) {
  try {
    const xmlContent = file.getBlob().getDataAsString('ISO-8859-1');
    const xml        = XmlService.parse(xmlContent);
    const root       = xml.getRootElement();
    const entries    = root.getChildren('DATO');

    // Verificar si el XML tiene 31 campos (nuevo número de campos)
    const firstEntry = entries[0];
    if (firstEntry && firstEntry.getChildren().length !== 31) {
      ui.alert('Error', 'Formato de archivo incorrecto. El XML debe tener 31 campos por registro.', ui.ButtonSet.OK);
      return;
    }

    const response = ui.alert(
      'Confirmar importación',
      `Se encontraron ${entries.length} registros en "${file.getName()}".\n\n¿Deseas continuar?`, 
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) return;

    const startTime     = new Date();
    const importedData  = importXMLFileDataCust(entries);
    
    // Verificar duplicados en la columna A (Campo 1)
    const duplicates = findDuplicatesCust(importedData.map(row => row[0])); // Campo 1
    if (duplicates.length > 0) {
        highlightDuplicatesCust(sheetCust, duplicates); // Resaltar duplicados en la columna A
    }

    const endTime       = new Date();
    const durationSeconds     = (endTime - startTime) / 1000;
    const durationMinutes     = Math.floor(durationSeconds / 60);
    const remainingSeconds    = durationSeconds % 60;

    ui.alert(
      'Importación Completada', 
      `Se importaron ${importedData.length} registros, con ${duplicates.length} registros duplicados.\n\n` +
      `Duración: ${durationMinutes} minutos y ${remainingSeconds.toFixed(2)} segundos.`, 
      ui.ButtonSet.OK
    );

    
    return importedData.length;
  } catch (error) {
    ui.alert('Error en la Importación', `Ha ocurrido un error: ${error.message}`, ui.ButtonSet.OK); 
    console.error(error);
    logErrorCust(error); // Registrar el error
    return 0;
  }
}

/**
 * Importa los datos del archivo XML a la hoja de cálculo.
 * @param {Array} entries - Las entradas del XML.
 * @returns {Array} - Los datos importados.
 */
function importXMLFileDataCust(entries) {
  const sheetCust = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CUST_SHEET_NAME) || 
                    SpreadsheetApp.getActiveSpreadsheet().insertSheet(CUST_SHEET_NAME);
  const lastRow = sheetCust.getLastRow();
  if (lastRow > 1) {
    sheetCust.getRange('A2:Q' + lastRow).clearContent(); // Limpiar columnas A a Q
  }
  const data = entries.map(entry => [
    parseInt(entry.getChildText('CodCliente')), // Campo 1: CodCliente (convertido a entero)
    entry.getChildText('RazonSocialdelCliente'), // Campo 2: RazonSocialdelCliente
    entry.getChildText('TipoDoc'), // Campo 3: TipoDoc
    entry.getChildText('NroDocumento'), // Campo 4: NroDocumento
    entry.getChildText('Direccion'), // Campo 5: Direccion
    entry.getChildText('CodPostal'), // Campo 6: CodPostal
    entry.getChildText('Localidad'), // Campo 7: Localidad
    entry.getChildText('Zona'), // Campo 8: Zona
    entry.getChildText('Provincia'), // Campo 9: Provincia
    entry.getChildText('Pais'), // Campo 10: Pais
    entry.getChildText('TipodeCliente'), // Campo 11: TipodeCliente
    entry.getChildText('CategoriaCliente'), // Campo 12: CategoriaCliente
    entry.getChildText('SubCategoriaCliente'), // Campo 13: SubCategoriaCliente
    parseInt(entry.getChildText('CodVendedor')), // Campo 14: CodVendedor (convertido a entero)
    entry.getChildText('Vendedor'), // Campo 15: Vendedor
    entry.getChildText('ListadePrecios'), // Campo 16: ListadePrecios
    entry.getChildText('CondiciondeVentaPredeterminada'), // Campo 17: CondiciondeVentaPredeterminada (ahora como string)
    entry.getChildText('SF_FechadeActualizacion'), // Campo 18: SF_FechadeActualizacion (ahora como string)
    convertToBoolean(entry.getChildText('ControlaCredito')), // Campo 19: ControlaCredito (convertido a booleano)
    entry.getChildText('SF_CreditoMaximo'), // Campo 20: SF_CreditoMaximo (sin transformación)
    entry.getChildText('SF_PenddeFacturar'), // Campo 21: SF_PenddeFacturar (sin transformación)
    entry.getChildText('SF_ChequesenCartera'), // Campo 22: SF_ChequesenCartera (sin transformación)
    entry.getChildText('SF_ChequesRechazados'), // Campo 23: SF_ChequesRechazados (sin transformación)
    entry.getChildText('SF_CreditoaVencer'), // Campo 24: SF_CreditoaVencer (sin transformación)
    entry.getChildText('SF_CreditoVencido'), // Campo 25: SF_CreditoVencido (sin transformación)
    convertToBoolean(entry.getChildText('SF_Moroso')), // Campo 26: SF_Moroso (convertido a booleano)
    convertToBoolean(entry.getChildText('SF_Engestionjudicial')), // Campo 27: SF_Engestionjudicial (convertido a booleano)
    convertToBoolean(entry.getChildText('SF_Incobrable')), // Campo 28: SF_Incobrable (convertido a booleano)
    entry.getChildText('FechaUltimaCompra'), // Campo 29: FechaUltimaCompra
    convertToDateTime(entry.getChildText('FechaUltModificacion')), // Campo 30: FechaUltModificacion (convertido a DateTime)
    convertToBoolean(entry.getChildText('Habilitado')) // Campo 31: Habilitado (convertido a booleano)
  ]);

  if (data.length > 0) {
    sheetCust.getRange(2, 1, data.length, data[0].length).setValues(data);
    
    // Eliminar filas excedentes
    const newLastRow = sheetCust.getLastRow();
    if (newLastRow < sheetCust.getMaxRows()) {
      sheetCust.deleteRows(newLastRow + 1, sheetCust.getMaxRows() - newLastRow);
    }
    
    return data;
  } else {
    SpreadsheetApp.getUi().alert('No hay datos para importar.');
    return [];
  }
}

/**
 * Convierte una cadena de fecha y hora en formato "dd/MM/yyyy HH:mm:ss a.m./p.m." a un objeto Date.
 * @param {string} dateTimeStr - La cadena de fecha y hora a convertir.
 * @returns {Date} - El objeto Date resultante.
 */
function convertToDateTime(dateTimeStr) {
    if (!dateTimeStr || typeof dateTimeStr !== 'string') {
        return null; // Manejar el caso donde la cadena es undefined o no es un string
    }
    
    const parts = dateTimeStr.split(' ');
    if (parts.length !== 3) {
        return null; // Manejar el caso donde el formato no es el esperado
    }

    const [datePart, timePart, period] = parts;
    const [day, month, year] = datePart.split('/').map(Number);
    const [hours, minutes, seconds] = timePart.split(':').map(Number);

    let adjustedHours = hours;
    if (period.toLowerCase() === 'p.m.' && hours < 12) {
        adjustedHours += 12; // Convertir a formato 24 horas
    } else if (period.toLowerCase() === 'a.m.' && hours === 12) {
        adjustedHours = 0; // Ajustar 12 a.m. a 0 horas
    }

    return new Date(year, month - 1, day, adjustedHours, minutes, seconds);
}
/**
 * Obtiene los archivos de una carpeta específica en Google Drive.
 * @param {string} folderId - El ID de la carpeta.
 * @param {string} extension - La extensión de los archivos a buscar.
 * @returns {Array} - Una lista de archivos que coinciden con la extensión.
 */
function getFilesFromFolderCust(folderId, extension) {
  try {
    const folder       = DriveApp.getFolderById(folderId);
    const filesIterator = folder.getFiles();
    const files        = [];

    while (filesIterator.hasNext()) {
      const file = filesIterator.next();
      if (file.getName().endsWith(extension)) {
        files.push(file);
      }
    }

    return files;
  } catch (error) {
    console.error('Error accessing folder:', error);
    throw new Error('No se pudo acceder a la carpeta especificada.');
  }
}

/**
 * Obtiene el archivo más reciente de una lista de archivos.
 * @param {Array} files - La lista de archivos.
 * @returns {File} - El archivo más reciente.
 */
function getLatestFileCust(files) {
  return files.reduce((a, b) => a.getLastUpdated() > b.getLastUpdated() ? a : b);
}

/**
 * Solicita al usuario que ingrese la URL de un archivo XML en Google Drive.
 * @param {Object} ui - La interfaz de usuario para mostrar alertas.
 * @param {Array} files - La lista de archivos disponibles.
 * @returns {File|null} - El archivo seleccionado o null si se cancela.
 */
function getFileFromUserCust(ui, files) {
  let fileUrl = ui.prompt(
    'Buscar archivo XML',
    'Ingresa la URL del archivo XML en Google Drive (link de compartir archivo de Drive):',
    ui.ButtonSet.OK_CANCEL
  ).getResponseText();

  while (true) {
    if (fileUrl) {
      try {
        const fileIdMatch = fileUrl.match(/\/d\/([^\/]+)/);
        if (fileIdMatch) {
          const fileId = fileIdMatch[1];
          return DriveApp.getFileById(fileId);
        } else {
          ui.alert('Error', 'La URL no es válida. Asegúrate de que sea un enlace para compartir de Google Drive.', ui.ButtonSet.OK);
        }
      } catch (error) {
        ui.alert('Error', 'Ocurrió un error al obtener el archivo. Verifica la URL e inténtalo de nuevo.', ui.ButtonSet.OK);
      }
    } else {
      ui.alert('Importación cancelada', 'El usuario canceló la selección del archivo.', ui.ButtonSet.OK);
      return null;
    }

    fileUrl = ui.prompt(
      'Buscar archivo XML',
      'Ingresa la URL del archivo XML en Google Drive (link de compartir archivo de Drive):',
      ui.ButtonSet.OK_CANCEL
    ).getResponseText();
  }
}

/**
 * Registra un error en la hoja de cálculo "Error Log".
 * @param {Error} error - El error que se ha producido.
 */
function logErrorCust(error) {
  const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Error Log');
  if (!logSheet) return; // No crear una nueva hoja si no existe
  logSheet.appendRow([new Date(), error.message]);
}

/**
 * Encuentra duplicados en un array.
 * @param {Array} values - Los valores a verificar.
 * @returns {Array} - Los valores duplicados.
 */
function findDuplicatesCust(values) {
    const seen = new Set();
    const duplicates = new Set();
    values.forEach(value => {
        if (seen.has(value)) {
            duplicates.add(value);
        } else {
            seen.add(value);
        }
    });
    return Array.from(duplicates);
}

/**
 * Resalta las filas duplicadas en la hoja.
 * @param {Object} sheet - La hoja de cálculo.
 * @param {Array} duplicates - Los valores duplicados.
 */
function highlightDuplicatesCust(sheet, duplicates) {
    const range = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
    const values = range.getValues();
    
    values.forEach((row, index) => {
        if (duplicates.includes(row[0])) { // Campo 1 en la columna A
            range.getCell(index + 1, 1).setBackground('yellow'); // Resaltar en amarillo
        }
    });
}
/**
 * Convierte una cadena "Verdadero" o "Falso" a un valor booleano para Google Sheets.
 * @param {string} value - El valor a convertir.
 * @returns {boolean} - El valor booleano correspondiente.
 */
function convertToBoolean(value) {
  return value.toLowerCase() === 'verdadero';
}
