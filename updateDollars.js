/**
 * @fileoverview Script para actualizar y gestionar cotizaciones del dólar en Google Sheets.
 * @author [Tu Nombre]
 */

/**
 * Clase de configuración que contiene constantes y configuraciones globales.
 */
var Config = {
  SHEET_NAME: 'Dolar',
  API_CURRENT: 'https://dolarapi.com/v1/dolares',
  API_HISTORICAL: 'https://api.argentinadatos.com/v1/cotizaciones/dolares/',
  COLUMN_NAMES: {
    DATE: 'Fecha',
    OFICIAL_COMPRA: 'Oficial Compra',
    OFICIAL_VENTA: 'Oficial Venta',
    BLUE_COMPRA: 'Blue Compra',
    BLUE_VENTA: 'Blue Venta',
    MEP_COMPRA: 'MEP Compra',
    MEP_VENTA: 'MEP Venta',
    CRIPTO_COMPRA: 'Cripto Compra',
    CRIPTO_VENTA: 'Cripto Venta',
    MAYORISTA_COMPRA: 'Mayorista Compra',
    MAYORISTA_VENTA: 'Mayorista Venta',
    BLUE_OFICIAL_BREACH: 'Brecha Blue/Oficial',
    BLUE_MEP_BREACH: 'Brecha Blue/MEP',
    LAST_UPDATE: 'Última Actualización',
    MODIFICATION_DATE: 'Fecha de Modificación'
  }
};

/**
 * Objeto para manejar el logging de la aplicación.
 */
var CustomLogger = {
  /**
   * Registra un mensaje de log.
   * @param {string} message - El mensaje a registrar.
   * @param {string} [level='INFO'] - El nivel de log (por defecto 'INFO').
   */
  log: function(message, level) {
    level = level || 'INFO';
    console.log('[' + level + '] ' + new Date().toISOString() + ': ' + message);
  },

  /**
   * Registra un mensaje de error.
   * @param {string} message - El mensaje de error a registrar.
   */
  error: function(message) {
    this.log(message, 'ERROR');
  }
};

/**
 * Obtiene la interfaz de usuario de la hoja de cálculo.
 * @returns {GoogleAppsScript.Base.Ui} La interfaz de usuario.
 */
function getUI() {
  return SpreadsheetApp.getUi();
}

/**
 * Obtiene la hoja de cálculo principal.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet|null} La hoja de cálculo o null si no se encuentra.
 */
function getSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(Config.SHEET_NAME);
  if (!sheet) {
    CustomLogger.error('No se encontró la hoja: ' + Config.SHEET_NAME);
    return null;
  }
  return sheet;
}

/**
 * Obtiene los datos de la hoja de cálculo.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - La hoja de cálculo.
 * @param {number} [startRow=1] - Fila inicial (opcional, por defecto 1).
 * @param {number} [startColumn=1] - Columna inicial (opcional, por defecto 1).
 * @returns {Array<Array<any>>} Los datos de la hoja.
 * @throws {Error} Si no se pueden obtener los datos de la hoja.
 */
function getSheetData(sheet, startRow, startColumn) {
  startRow = startRow || 1;
  startColumn = startColumn || 1;
  try {
    var lastRow = sheet.getLastRow();
    var lastColumn = sheet.getLastColumn();
    if (lastRow < startRow || lastColumn < startColumn) {
      throw new Error('No hay datos en el rango especificado');
    }
    return sheet.getRange(startRow, startColumn, lastRow - startRow + 1, lastColumn - startColumn + 1).getValues();
  } catch (error) {
    throw new Error('Error al obtener datos de la hoja: ' + error.message);
  }
}

/**
 * Maneja los errores de la aplicación de manera uniforme.
 * @param {Error} error - El error a manejar.
 * @param {string} functionName - El nombre de la función donde ocurrió el error.
 * @param {GoogleAppsScript.Base.Ui} [ui] - La interfaz de usuario (opcional).
 */
function handleError(error, functionName, ui) {
  var errorMessage = 'Error en ' + functionName + ': ' + error.message;
  CustomLogger.error(errorMessage);
  
  if (ui) {
    ui.alert('Error', 'Ocurrió un error durante la operación. Por favor, intente nuevamente.\n\nDetalles: ' + errorMessage, ui.ButtonSet.OK);
  }
}

/**
 * Actualiza los valores del dólar en la hoja de cálculo.
 * Esta función es el punto de entrada principal para la actualización de datos.
 */
function updateDollarsUntilToday() {
  var ui = getUI();
  try {
    CustomLogger.log('Iniciando actualización de datos del dólar');
    var sheet = getSheet();
    if (!sheet) return;

    var data = getSheetData(sheet);
    var lastRow = data.length;
    var lastDate = new Date(data[lastRow - 1][0]);
    var today = new Date();
    today.setHours(0, 0, 0, 0);

    if (isToday(lastDate)) {
      handleTodayUpdate(ui, sheet, lastRow);
    } else if (isYesterday(lastDate)) {
      handleYesterdayUpdate(ui, sheet, lastRow);
    } else {
      handleHistoricalUpdate(ui, sheet, lastDate, today);
    }

    CustomLogger.log('Actualización de datos del dólar completada');
  } catch (error) {
    handleError(error, 'updateDollarsUntilToday', ui);
  }
}

/**
 * Verifica si una fecha es la fecha actual.
 * @param {Date} date - Fecha a verificar.
 * @returns {boolean} True si la fecha es la fecha actual, False en caso contrario.
 */
function isToday(date) {
  var today = new Date();
  return date.toDateString() === today.toDateString();
}

/**
 * Verifica si una fecha es la fecha de ayer.
 * @param {Date} date - Fecha a verificar.
 * @returns {boolean} True si la fecha es la fecha de ayer, False en caso contrario.
 */
function isYesterday(date) {
  var yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  return date.toDateString() === yesterday.toDateString();
}

/**
 * Maneja la actualización de datos para el día de hoy.
 * Esta función utiliza la API de datos actuales para obtener la última cotización del día.
 * 
 * @param {GoogleAppsScript.Base.Ui} ui - Interfaz de usuario de la hoja de cálculo.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Hoja de cálculo.
 * @param {number} lastRow - Número de la última fila con datos.
 * 
 * @description
 * - Utiliza la API de datos actuales (Config.API_CURRENT).
 * - Solicita confirmación al usuario antes de actualizar.
 * - Si se confirma, obtiene los datos más recientes y actualiza la fila correspondiente a hoy.
 * - Es útil para actualizar los valores que pueden cambiar durante el día.
 */
function handleTodayUpdate(ui, sheet, lastRow) {
  var response = ui.alert(
    'Confirmar actualización',
    'La tabla tiene valores actualizados para hoy. ¿Desea actualizar los valores de hoy con los últimos datos disponibles?',
    ui.ButtonSet.YES_NO
  );

  if (response === ui.Button.YES) {
    var today = new Date();
    var apiResponse = fetchDataFromApis(today, 'current');
    if (apiResponse) {
      var values = processCurrentApiData(apiResponse.data);
      updateDolarData(sheet, lastRow, values, today);
      ui.alert('Datos actualizados correctamente con la última información disponible.');
    } else {
      ui.alert('No se pudieron obtener los datos actualizados. Por favor, intente nuevamente más tarde.');
    }
  }
}

/**
 * Maneja la actualización de datos para el día de ayer.
 * Esta función utiliza la API de datos históricos para obtener la cotización del día anterior.
 * 
 * @param {GoogleAppsScript.Base.Ui} ui - Interfaz de usuario de la hoja de cálculo.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Hoja de cálculo.
 * @param {number} lastRow - Número de la última fila con datos.
 * 
 * @description
 * - Utiliza la API de datos históricos (Config.API_HISTORICAL).
 * - Solicita confirmación al usuario antes de agregar los datos de hoy.
 * - Si se confirma, obtiene los datos de ayer y agrega una nueva fila con estos datos.
 * - Es útil para agregar el primer registro del día actual, basado en los datos consolidados de ayer.
 */
function handleYesterdayUpdate(ui, sheet, lastRow) {
  var response = ui.alert(
    'Confirmar actualización',
    'La tabla tiene valores hasta ayer. ¿Desea agregar los valores de hoy basados en los datos de cierre de ayer?',
    ui.ButtonSet.YES_NO
  );

  if (response === ui.Button.YES) {
    var today = new Date();
    var yesterday = new Date(today);
    yesterday.setDate(yesterday.getDate() - 1);
    var apiResponse = fetchDataFromApis(yesterday, 'historical');
    if (apiResponse) {
      var values = processHistoricalApiData(apiResponse.data);
      updateDolarData(sheet, lastRow + 1, values, today);
      ui.alert('Se han agregado los datos de hoy basados en el cierre de ayer.');
    } else {
      ui.alert('No se pudieron obtener los datos de ayer. Por favor, intente nuevamente más tarde.');
    }
  }
}

/**
 * Maneja la actualización de datos históricos.
 * @param {GoogleAppsScript.Base.Ui} ui - Interfaz de usuario de la hoja de cálculo.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Hoja de cálculo.
 * @param {Date} lastDate - Última fecha en la hoja de cálculo.
 * @param {Date} today - Fecha actual.
 */
function handleHistoricalUpdate(ui, sheet, lastDate, today) {
  var response = ui.alert(
    'Confirmar actualización',
    'La tabla tiene valores hasta ' + formatDate(lastDate, 'verbose') + '. ¿Desea actualizar los valores hasta hoy?',
    ui.ButtonSet.YES_NO
  );

  if (response === ui.Button.YES) {
    var yesterday = new Date(today);
    yesterday.setDate(yesterday.getDate() - 1);
    
    // Actualizar datos históricos hasta ayer
    if (lastDate < yesterday) {
      var historicResult = updateHistoricalValues(sheet, lastDate, yesterday);
      if (historicResult.error) {
        ui.alert('Error al actualizar datos históricos: ' + historicResult.error);
        return;
      }
    }
    
    // Agregar datos de hoy
    var apiResponse = fetchDataFromApis(today, 'current');
    if (apiResponse) {
      var values = processCurrentApiData(apiResponse.data);
      updateDolarData(sheet, sheet.getLastRow() + 1, values, today);
      ui.alert('Datos actualizados correctamente hasta hoy.');
    } else {
      ui.alert('No se pudieron obtener los datos de hoy. Se han actualizado los datos históricos hasta ayer.');
    }
  }
}

/**
 * Obtiene datos de la API para una fecha específica.
 * @param {Date} date - Fecha para la cual obtener datos.
 * @param {string} apiType - Tipo de API a utilizar ('current' o 'historical').
 * @returns {Object|null} Datos de la API y tipo de API, o null si hay un error.
 * @throws {Error} Si no se pueden obtener los datos de la API.
 */
function fetchDataFromApis(date, apiType) {
  CustomLogger.log('Obteniendo datos de la API para ' + formatDate(date));
  try {
    var url;

    if (apiType === 'current') {
      url = Config.API_CURRENT;
    } else {
      var formattedDate = formatDate(date, 'api');
      url = Config.API_HISTORICAL + formattedDate;
    }

    var response = getCachedOrFetch(url);
    if (!response) {
      CustomLogger.log('No se obtuvieron datos de la API para ' + formatDate(date));
      return null;
    }
    
    // Asegurarse de que los datos estén en formato JSON
    var data = typeof response === 'string' ? JSON.parse(response) : response;
    
    CustomLogger.log('Datos de la API obtenidos para ' + formatDate(date));
    return { data: data, apiType: apiType };
  } catch (error) {
    CustomLogger.error('Error al obtener datos para ' + formatDate(date) + ': ' + error.message);
    return null;
  }
}

/**
 * Procesa los datos de la API para datos actuales.
 * @param {Object|Array} data - Datos de la API.
 * @returns {Object} Valores procesados.
 */
function processCurrentApiData(data) {
  // Si data es un string, intentamos parsearlo como JSON
  if (typeof data === 'string') {
    try {
      data = JSON.parse(data);
    } catch (e) {
      CustomLogger.error('Error al parsear datos de la API: ' + e.message);
      throw e;
    }
  }

  // Si data es un objeto, lo convertimos en un array
  if (!Array.isArray(data)) {
    data = Object.values(data);
  }

  var findValue = function(casa) {
    var item = data.find(function(d) { return d.casa === casa; });
    return item ? { compra: parseFloat(item.compra), venta: parseFloat(item.venta) } : { compra: 0, venta: 0 };
  };

  var oficial = findValue('oficial');
  var blue = findValue('blue');
  var mep = findValue('bolsa');
  var cripto = findValue('cripto');
  var mayorista = findValue('mayorista');

  var values = {
    oficialCompra: oficial.compra,
    oficialVenta: oficial.venta,
    blueCompra: blue.compra,
    blueVenta: blue.venta,
    mepCompra: mep.compra,
    mepVenta: mep.venta,
    criptoCompra: cripto.compra,
    criptoVenta: cripto.venta,
    mayoristaCompra: mayorista.compra,
    mayoristaVenta: mayorista.venta,
    fechaActualizacion: data[0] ? data[0].fechaActualizacion : new Date().toISOString()
  };

  values.blueOficialBreach = ((values.blueVenta - values.oficialVenta) / values.oficialVenta) * 100;
  values.blueMepBreach = ((values.blueVenta - values.mepVenta) / values.mepVenta) * 100;

  return values;
}

/**
 * Procesa los datos de la API para datos históricos.
 * @param {Object} data - Datos de la API.
 * @returns {Object} Valores procesados.
 */
function processHistoricalApiData(data) {
  CustomLogger.log('Procesando datos históricos de la API');
  try {
    var oficialData = data.find(d => d.casa === 'oficial');
    var blueData = data.find(d => d.casa === 'blue');
    var bolsaData = data.find(d => d.casa === 'bolsa');
    var mayoristaData = data.find(d => d.casa === 'mayorista');
    var criptoData = data.find(d => d.casa === 'cripto');

    var values = {
      oficialCompra: parseFloat(oficialData?.compra) || 0,
      oficialVenta: parseFloat(oficialData?.venta) || 0,
      blueCompra: parseFloat(blueData?.compra) || 0,
      blueVenta: parseFloat(blueData?.venta) || 0,
      mepCompra: parseFloat(bolsaData?.compra) || 0,
      mepVenta: parseFloat(bolsaData?.venta) || 0,
      criptoCompra: parseFloat(criptoData?.compra) || 0,
      criptoVenta: parseFloat(criptoData?.venta) || 0,
      mayoristaCompra: parseFloat(mayoristaData?.compra) || 0,
      mayoristaVenta: parseFloat(mayoristaData?.venta) || 0,
      fechaActualizacion: data[0]?.fecha || new Date().toISOString()
    };

    values.blueOficialBreach = values.oficialVenta !== 0 ? ((values.blueVenta - values.oficialVenta) / values.oficialVenta) : 0;
    values.blueMepBreach = values.mepVenta !== 0 ? ((values.blueVenta - values.mepVenta) / values.mepVenta) : 0;

    CustomLogger.log('Datos históricos procesados correctamente');
    return values;
  } catch (error) {
    CustomLogger.error('Error al procesar datos históricos: ' + error.message);
    throw error;
  }
}

/**
 * Actualiza los valores históricos de una hoja, optimizado para escritura en lotes.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Hoja de cálculo a actualizar.
 * @param {Date} startDate - Fecha de inicio del rango.
 * @param {Date} endDate - Fecha de fin del rango.
 * @returns {Object} Resultado de la operación.
 */
function updateHistoricalValues(sheet, startDate, endDate) {
  CustomLogger.log('Actualizando valores históricos desde ' + formatDate(startDate) + ' hasta ' + formatDate(endDate));
  var startTime = new Date().getTime();

  try {
    validateDate(startDate, new Date('2015-01-01'), new Date());
    validateDate(endDate, startDate, new Date());

    var existingData = getSheetData(sheet);
    var existingDates = new Map(existingData.map(row => [formatDate(new Date(row[0]), 'sheet'), row]));

    var rowsToUpdate = [];
    var rowsToAdd = [];
    var currentDate = new Date(startDate);
    while (currentDate <= endDate) {
      var apiResponse = fetchDataFromApis(currentDate, 'historical');
      if (apiResponse) {
        var rowData = prepareHistoricalRowData(currentDate, apiResponse);
        if (rowData) {
          var dateKey = formatDate(currentDate, 'sheet');
          if (existingDates.has(dateKey)) {
            rowsToUpdate.push({ row: existingData.findIndex(row => formatDate(new Date(row[0]), 'sheet') === dateKey) + 1, data: rowData });
          } else {
            rowsToAdd.push(rowData);
          }
        } else {
          CustomLogger.log('Datos incompletos para la fecha ' + formatDate(currentDate));
        }
      }
      currentDate.setDate(currentDate.getDate() + 1);
    }

    // Actualizar filas existentes
    rowsToUpdate.forEach(update => {
      if (update.row > 0 && update.row <= sheet.getLastRow()) {
        sheet.getRange(update.row, 1, 1, update.data.length).setValues([update.data]);
      } else {
        CustomLogger.log('Fila inválida para actualizar: ' + update.row + '. Agregando como nueva fila.');
        rowsToAdd.push(update.data);
      }
    });

    // Agregar nuevas filas
    if (rowsToAdd.length > 0) {
      var startRow = sheet.getLastRow() + 1;
      var columnOrder = Object.values(Config.COLUMN_NAMES);
      sheet.getRange(startRow, 1, rowsToAdd.length, columnOrder.length).setValues(rowsToAdd);
    }

    sortSheetByDate(sheet);

    var endTime = new Date().getTime();
    var duration = ((endTime - startTime) / 1000).toFixed(2);
    return {
      success: true,
      message: 'Se actualizó el tipo de cambio entre las fechas ' + formatDate(startDate, 'verbose') + ' y ' + formatDate(endDate, 'verbose') + '.\n\nSe actualizaron ' + rowsToUpdate.length + ' filas existentes y se agregaron ' + rowsToAdd.length + ' nuevas filas.\n\nEl proceso tomó ' + duration + ' segundos.'
    };
  } catch (error) {
    CustomLogger.error('Error en updateHistoricalValues: ' + error.message);
    return { error: error.message };
  }
}

function prepareHistoricalRowData(date, apiResponse) {
  CustomLogger.log('Preparando fila de datos históricos para ' + formatDate(date));
  try {
    if (!apiResponse || !apiResponse.data) {
      CustomLogger.log('Datos de API no válidos para ' + formatDate(date));
      return null;
    }

    var data = apiResponse.data;
    var values = processHistoricalApiData(data);

    // Función auxiliar para formatear números
    function formatNumber(value, decimals = 2) {
      return value.toFixed(decimals).replace('.', ',');
    }

    // Función auxiliar para formatear porcentajes
    function formatPercentage(value) {
      return (value * 100).toFixed(2).replace('.', ',') + '%';
    }

    var currentDateTime = new Date();

    var rowData = [
      formatDate(date, 'dayColumn'),
      formatNumber(values.oficialCompra),
      formatNumber(values.oficialVenta),
      formatNumber(values.blueCompra),
      formatNumber(values.blueVenta),
      formatNumber(values.mepCompra),
      formatNumber(values.mepVenta),
      formatNumber(values.criptoCompra),
      formatNumber(values.criptoVenta),
      formatNumber(values.mayoristaCompra),
      formatNumber(values.mayoristaVenta),
      formatPercentage(values.blueOficialBreach),
      formatPercentage(values.blueMepBreach),
      formatDate(currentDateTime, 'dateTime'), // Fecha y hora de los datos
      formatDate(currentDateTime, 'dateTime')  // Fecha y hora de modificación del registro
    ];

    CustomLogger.log('Fila de datos históricos preparada para ' + formatDate(date));
    return rowData;
  } catch (error) {
    CustomLogger.error('Error al preparar fila de datos históricos para ' + formatDate(date) + ': ' + error.message);
    return null;
  }
}

/**
 * Elimina las filas existentes en un rango de fechas.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Hoja de cálculo.
 * @param {Date} startDate - Fecha de inicio del rango.
 * @param {Date} endDate - Fecha de fin del rango.
 */
function deleteExistingRows(sheet, startDate, endDate) {
  var data = sheet.getDataRange().getValues();
  var rowsToDelete = [];

  for (var i = data.length - 1; i > 0; i--) {
    var rowDate = new Date(data[i][0]);
    if (rowDate >= startDate && rowDate <= endDate) {
      rowsToDelete.push(i + 1);
    }
  }

  for (var i = rowsToDelete.length - 1; i >= 0; i--) {
    sheet.deleteRow(rowsToDelete[i]);
  }
}

/**
 * Obtiene y prepara los datos históricos.
 * @param {Date} startDate - Fecha de inicio del rango.
 * @param {Date} endDate - Fecha de fin del rango.
 * @returns {Array} Filas de datos preparadas para insertar.
 */
function fetchAndPrepareHistoricalData(startDate, endDate) {
  var rowsToAdd = [];
  var currentDate = new Date(startDate);
  while (currentDate <= endDate) {
    var apiResponse = fetchDataFromApis(currentDate);
    if (apiResponse) {
      var rowData = prepareRowData(currentDate, apiResponse);
      if (rowData) {
        rowsToAdd.push(rowData);
      } else {
        CustomLogger.log('Datos incompletos para la fecha ' + formatDate(currentDate));
      }
    }
    currentDate.setDate(currentDate.getDate() + 1);
  }
  return rowsToAdd;
}

/**
 * Inserta los datos históricos en la hoja.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Hoja de cálculo.
 * @param {Array} rowsToAdd - Filas de datos a insertar.
 */
function insertHistoricalData(sheet, rowsToAdd) {
  if (rowsToAdd.length > 0) {
    var startRow = sheet.getLastRow() + 1;
    var columnOrder = Object.values(Config.COLUMN_NAMES);
    sheet.getRange(startRow, 1, rowsToAdd.length, columnOrder.length).setValues(rowsToAdd);
  }
}

/**
 * Genera el resultado de la actualización.
 * @param {Date} startDate - Fecha de inicio del rango.
 * @param {Date} endDate - Fecha de fin del rango.
 * @param {number} rowsAdded - Número de filas agregadas.
 * @param {number} startTime - Tiempo de inicio de la operación.
 * @returns {string} Mensaje con el resultado de la actualización.
 */
function generateUpdateResult(startDate, endDate, rowsAdded, startTime) {
  var endTime = new Date().getTime();
  var duration = ((endTime - startTime) / 1000).toFixed(2);
  return 'Se actualizó el tipo de cambio entre las fechas ' + formatDate(startDate, 'verbose') + ' y ' + formatDate(endDate, 'verbose') + '.\n\nSe agregaron en total ' + rowsAdded + ' días con todas las cotizaciones.\n\nEl proceso tomó ' + duration + ' segundos.';
}

/**
 * Obtiene datos de la caché o de la API si no están en caché.
 * @param {string} url - URL de la API.
 * @returns {Object|null} Datos de la API o null si hay un error.
 */
function getCachedOrFetch(url, cacheTime = 3600) { // cacheTime en segundos
  const cache = CacheService.getScriptCache();
  const cached = cache.get(url);
  if (cached != null) {
    return JSON.parse(cached);
  }
  
  const response = UrlFetchApp.fetch(url);
  const data = response.getContentText();
  cache.put(url, JSON.stringify(data), cacheTime);
  return data;
}

/**
 * Valida que una fecha esté dentro de un rango permitido.
 * @param {Date} date - La fecha a validar.
 * @param {Date} minDate - La fecha mínima permitida.
 * @param {Date} maxDate - La fecha máxima permitida.
 * @throws {Error} Si la fecha no es válida o está fuera del rango.
 */
function validateDate(date, minDate, maxDate) {
  if (!(date instanceof Date) || isNaN(date)) {
    throw new Error('Fecha inválida');
  }
  if (date < minDate || date > maxDate) {
    throw new Error('La fecha debe estar entre ' + formatDate(minDate) + ' y ' + formatDate(maxDate));
  }
}

/**
 * Formatea una fecha según el formato especificado.
 * @param {Date} date - Fecha a formatear.
 * @param {string} format - Formato deseado ('verbose', 'sheet', 'api', 'dateTime', 'verboseLong', 'dayColumn').
 * @returns {string} - Fecha formateada.
 */
function formatDate(date, format) {
  format = format || 'short';
  var daysOfWeek = ['Domingo', 'Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado'];
  var months = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'];
  
  // Ajustar la fecha a la zona horaria de Argentina (GMT-3)
  var argentinaDate = new Date(date.getTime() - (3 * 60 * 60 * 1000));
  
  var day = ('0' + argentinaDate.getUTCDate()).slice(-2);
  var month = ('0' + (argentinaDate.getUTCMonth() + 1)).slice(-2);
  var year = argentinaDate.getUTCFullYear();
  var shortYear = year.toString().slice(-2);
  var hours = ('0' + argentinaDate.getUTCHours()).slice(-2);
  var minutes = ('0' + argentinaDate.getUTCMinutes()).slice(-2);
  var seconds = ('0' + argentinaDate.getUTCSeconds()).slice(-2);

  switch (format) {
    case 'verbose':
      return daysOfWeek[argentinaDate.getUTCDay()] + ' ' + day + ' de ' + months[argentinaDate.getUTCMonth()] + ' de ' + year;
    case 'sheet':
      return year + '-' + month + '-' + day;
    case 'api':
      return year + '/' + month + '/' + day;
    case 'dateTime':
      return `${day}/${month}/${shortYear} ${hours}:${minutes}:${seconds}`;
    case 'verboseLong':
      return daysOfWeek[argentinaDate.getUTCDay()] + ' ' + day + ' de ' + months[argentinaDate.getUTCMonth()] + ' ' + hours + ':' + minutes + ':' + seconds;
    case 'dayColumn':
      return daysOfWeek[argentinaDate.getUTCDay()] + ' ' + day + '/' + month + '/' + year;
    case 'fullLong':
      return daysOfWeek[argentinaDate.getUTCDay()] + ' ' + day + '-' + month + '-' + shortYear;
    case 'time':
      return `${hours}:${minutes}:${seconds}`;
    case 'fullLongWithTime':
      return `${daysOfWeek[argentinaDate.getUTCDay()]} ${('0' + argentinaDate.getUTCDate()).slice(-2)}-${('0' + (argentinaDate.getUTCMonth() + 1)).slice(-2)}-${argentinaDate.getUTCFullYear().toString().slice(-2)} ${('0' + argentinaDate.getUTCHours()).slice(-2)}:${('0' + argentinaDate.getUTCMinutes()).slice(-2)}:${('0' + argentinaDate.getUTCSeconds()).slice(-2)}`;
    default:
      return day + '/' + month + '/' + year;
  }
}

/**
 * Ordena la hoja por fecha.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Hoja de cálculo.
 */
function sortSheetByDate(sheet) {
  CustomLogger.log('Ordenando hoja por fecha');
  try {
    var dataRange = sheet.getDataRange();
    var numRows = dataRange.getNumRows();
    var numCols = dataRange.getNumColumns();
    
    if (numRows > 1) {
      sheet.getRange(2, 1, numRows - 1, numCols).sort({column: 1, ascending: true});
    }

    CustomLogger.log('Hoja ordenada por fecha');
  } catch (error) {
    CustomLogger.error('Error en sortSheetByDate: ' + error.message);
    throw error;
  }
}

/**
 * Elimina filas duplicadas basadas en la fecha.
 */
function removeDuplicateDates() {
  CustomLogger.log('Eliminando filas duplicadas');
  try {
    var sheet = getSheet();
    if (!sheet) return;

    var data = sheet.getDataRange().getValues();
    var uniqueDates = new Set();
    var rowsToDelete = [];

    for (var i = data.length - 1; i > 0; i--) {
      var dateStr = formatDate(new Date(data[i][0]), 'sheet');
      if (uniqueDates.has(dateStr)) {
        rowsToDelete.push(i + 1);
      } else {
        uniqueDates.add(dateStr);
      }
    }

    for (var i = rowsToDelete.length - 1; i >= 0; i--) {
      sheet.deleteRow(rowsToDelete[i]);
    }

    getUI().alert('Se eliminaron ' + rowsToDelete.length + ' filas duplicadas.');
    CustomLogger.log('Se eliminaron ' + rowsToDelete.length + ' filas duplicadas');
  } catch (error) {
    CustomLogger.error('Error en removeDuplicateDates: ' + error.message);
    getUI().alert('Ocurrió un error al eliminar filas duplicadas. Por favor, intente nuevamente.');
  }
}

/**
 * Solicita al usuario un rango de fechas y actualiza/agrega datos de tipo de cambio.
 */
function updateDollarsByRange() {
  var ui = SpreadsheetApp.getUi();
  
  // Solicitar fecha de inicio
  var startDateResult = ui.prompt(
    'Fecha desde',
    'Ingrese la fecha de inicio (dd/mm o dd/mm/yy):',
    ui.ButtonSet.OK_CANCEL);

  if (startDateResult.getSelectedButton() != ui.Button.OK) {
    return; // El usuario canceló o cerró el diálogo
  }

  var startDateStr = startDateResult.getResponseText();

  // Solicitar fecha de fin
  var endDateResult = ui.prompt(
    'Fecha hasta',
    'Ingrese la fecha de fin (dd/mm o dd/mm/yy):',
    ui.ButtonSet.OK_CANCEL);

  if (endDateResult.getSelectedButton() != ui.Button.OK) {
    return; // El usuario canceló o cerró el diálogo
  }

  var endDateStr = endDateResult.getResponseText();

  // Procesar las fechas
  try {
    var startDate = parseCustomDate(startDateStr);
    var endDate = parseCustomDate(endDateStr);

    // Validar las fechas
    validateDateRange(startDate, endDate);

    // Llamar a la función que actualiza el rango de fechas
    updateHistoricalRange(startDate, endDate);
  } catch (error) {
    ui.alert('Error', 'Hubo un problema al procesar las fechas: ' + error.message, ui.ButtonSet.OK);
  }
}

/**
 * Convierte una fecha en formato dd/mm o dd/mm/yy a un objeto Date.
 * @param {string} dateStr - Fecha en formato dd/mm o dd/mm/yy.
 * @returns {Date} Objeto Date correspondiente.
 */
function parseCustomDate(dateStr) {
  var parts = dateStr.split(/[\/\-]/); // Acepta tanto "/" como "-"
  if (parts.length < 2 || parts.length > 3) {
    throw new Error('Formato de fecha inválido. Use dd/mm o dd/mm/yy.');
  }

  var day = parseInt(parts[0], 10);
  var month = parseInt(parts[1], 10) - 1; // Los meses en JavaScript van de 0 a 11
  var year = new Date().getFullYear(); // Por defecto, usa el año actual

  if (parts.length === 3) {
    year = parseInt(parts[2], 10);
    // Si el año tiene dos dígitos, asumimos que es del siglo XXI
    if (year < 100) {
      year += 2000;
    }
  }
  
  var date = new Date(year, month, day);
  
  // Validar que la fecha sea válida
  if (isNaN(date.getTime())) {
    throw new Error('Fecha inválida');
  }
  
  return date;
}

function validateDateRange(startDate, endDate) {
  var minDate = new Date('2015-01-01');
  var maxDate = new Date();

  if (startDate > endDate) {
    throw new Error('La fecha de inicio debe ser anterior o igual a la fecha de fin.');
  }

  validateDate(startDate, minDate, maxDate);
  validateDate(endDate, minDate, maxDate);
}

function updateHistoricalRange(startDate, endDate) {
  var ui = SpreadsheetApp.getUi();
  try {
    CustomLogger.log('Actualizando datos históricos desde ' + formatDate(startDate) + ' hasta ' + formatDate(endDate));
    var sheet = getSheet();
    if (!sheet) return;

    var result = updateHistoricalValues(sheet, startDate, endDate);
    if (result.success) {
      ui.alert('Actualización Completada', result.message, ui.ButtonSet.OK);
    } else {
      ui.alert('Error en la Actualización', result.error, ui.ButtonSet.OK);
    }
  } catch (error) {
    handleError(error, 'updateHistoricalRange', ui);
  }
}

/**
 * Verifica si una fecha es válida y está dentro de un rango permitido.
 * @param {Date} date Fecha a validar.
 * @param {Date} minDate Fecha mínima permitida.
 * @param {Date} maxDate Fecha máxima permitida.
 * @returns {boolean} True si la fecha es válida y está dentro del rango, False en caso contrario.
 */
function isValidDate(date, minDate, maxDate) {
  CustomLogger.log('Validando fecha ' + formatDate(date) + ' entre ' + formatDate(minDate) + ' y ' + formatDate(maxDate));
  try {
    return date && date >= minDate && date <= maxDate;
  } catch (error) {
    CustomLogger.error('Error en isValidDate: ' + error.message);
    throw error;
  }
}

/**
 * Convierte una fecha en formato dd/mm o dd/mm/yy a un objeto Date.
 * @param {string} dateStr - Fecha en formato dd/mm o dd/mm/yy.
 * @returns {Date} Objeto Date correspondiente.
 */
function parseCustomDate(dateStr) {
  var parts = dateStr.split(/[\/\-]/); // Acepta tanto "/" como "-"
  if (parts.length < 2 || parts.length > 3) {
    throw new Error('Formato de fecha inválido. Use dd/mm o dd/mm/yy.');
  }

  var day = parseInt(parts[0], 10);
  var month = parseInt(parts[1], 10) - 1; // Los meses en JavaScript van de 0 a 11
  var year = new Date().getFullYear(); // Por defecto, usa el año actual

  if (parts.length === 3) {
    year = parseInt(parts[2], 10);
    // Si el año tiene dos dígitos, asumimos que es del siglo XXI
    if (year < 100) {
      year += 2000;
    }
  }
  
  var date = new Date(year, month, day);
  
  // Validar que la fecha sea válida
  if (isNaN(date.getTime())) {
    throw new Error('Fecha inválida');
  }
  
  return date;
}

// Función para obtener los datos más recientes de la API
function fetchLatestExchangeRates() {
  const url = "https://dolarapi.com/v1/dolares";
  try {
    const response = UrlFetchApp.fetch(url);
    const data = JSON.parse(response.getContentText());
    
    // Extraer la fecha de actualización
    const fechaActualizacion = data[0].fechaActualizacion; // Asumimos que todos los tipos tienen la misma fecha de actualización
    
    // Convertir la fecha a la zona horaria de Argentina
    const fechaArgentina = new Date(fechaActualizacion);
    fechaArgentina.setHours(fechaArgentina.getHours() - 3); // Ajuste para GMT-3
    
    return {
      data: data,
      fechaActualizacion: fechaArgentina
    };
  } catch (error) {
    CustomLogger.error('Error al obtener datos de la API: ' + error.message);
    throw error;
  }
}

// Función para formatear fechas
function formatDate(date, format) {
  format = format || 'short';
  var daysOfWeek = ['Domingo', 'Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado'];
  var months = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'];
  
  // Ajustar la fecha a la zona horaria de Argentina (GMT-3)
  var argentinaDate = new Date(date.getTime() - (3 * 60 * 60 * 1000));
  
  var day = ('0' + argentinaDate.getUTCDate()).slice(-2);
  var month = ('0' + (argentinaDate.getUTCMonth() + 1)).slice(-2);
  var year = argentinaDate.getUTCFullYear();
  var shortYear = year.toString().slice(-2);
  var hours = ('0' + argentinaDate.getUTCHours()).slice(-2);
  var minutes = ('0' + argentinaDate.getUTCMinutes()).slice(-2);
  var seconds = ('0' + argentinaDate.getUTCSeconds()).slice(-2);

  switch (format) {
    case 'verbose':
      return daysOfWeek[argentinaDate.getUTCDay()] + ' ' + day + ' de ' + months[argentinaDate.getUTCMonth()] + ' de ' + year;
    case 'sheet':
      return year + '-' + month + '-' + day;
    case 'api':
      return year + '/' + month + '/' + day;
    case 'dateTime':
      return `${day}/${month}/${shortYear} ${hours}:${minutes}:${seconds}`;
    case 'verboseLong':
      return daysOfWeek[argentinaDate.getUTCDay()] + ' ' + day + ' de ' + months[argentinaDate.getUTCMonth()] + ' ' + hours + ':' + minutes + ':' + seconds;
    case 'dayColumn':
      return daysOfWeek[argentinaDate.getUTCDay()] + ' ' + day + '/' + month + '/' + year;
    case 'fullLong':
      return daysOfWeek[argentinaDate.getUTCDay()] + ' ' + day + '-' + month + '-' + shortYear;
    case 'time':
      return `${hours}:${minutes}:${seconds}`;
    case 'fullLongWithTime':
      return `${daysOfWeek[argentinaDate.getUTCDay()]} ${('0' + argentinaDate.getUTCDate()).slice(-2)}-${('0' + (argentinaDate.getUTCMonth() + 1)).slice(-2)}-${argentinaDate.getUTCFullYear().toString().slice(-2)} ${('0' + argentinaDate.getUTCHours()).slice(-2)}:${('0' + argentinaDate.getUTCMinutes()).slice(-2)}:${('0' + argentinaDate.getUTCSeconds()).slice(-2)}`;
    default:
      return day + '/' + month + '/' + year;
  }
}

// Función para generar el mensaje formateado
function generateFormattedMessage(date, values, type) {
  function formatNumber(value) {
    return parseFloat(value).toLocaleString('es-AR', {minimumFractionDigits: 2, maximumFractionDigits: 2});
  }

  function formatPercentage(value) {
    return (parseFloat(value) * 100).toLocaleString('es-AR', {minimumFractionDigits: 2, maximumFractionDigits: 2}) + '%';
  }

  var lastUpdateDate = type === 'current' ? new Date(values.fechaActualizacion) : new Date();
  var isValidDate = !isNaN(lastUpdateDate.getTime());

  var styles = `
    <style>
      body { font-family: 'Roboto', sans-serif; color: #333; line-height: 1.6; }
      h2 { color: #1a73e8; text-align: center; margin-bottom: 15px; }
      .date { text-align: center; font-size: 16px; font-weight: bold; margin-bottom: 20px; font-style: italic; }
      table { width: 100%; border-collapse: collapse; margin-bottom: 20px; }
      th, td { padding: 10px; text-align: right; border-bottom: 1px solid #ddd; font-weight: bold; }
      th { background-color: #e6f2ff; color: #1a73e8; text-align: center; }
      td:first-child { text-align: center; }
      .breach-table { margin-bottom: 20px; background-color: #E7E8EE; }
      .breach-table td { border-bottom: none; text-align: center; }
      .breach-value { color: #B02B2B; }
      .update { font-size: 14px; text-align: center; font-weight: bold; }
    </style>
  `;

  var lastUpdateContent = isValidDate ? `
    <p class="update">Últimos datos disponibles:</p>
    <p class="update">${formatDate(lastUpdateDate, 'fullLongWithTime')}</p>
  ` : '<p class="update">Fecha de actualización no disponible</p>';

  var content = `
    ${styles}
    <h2>Cotizaciones</h2>
    <p class="date">${formatDate(date, 'fullLong')}</p>
    <table>
      <tr><th>Dolar</th><th>Compra</th><th>Venta</th></tr>
      <tr><td>Oficial</td><td>${formatNumber(values.oficialCompra)}</td><td>${formatNumber(values.oficialVenta)}</td></tr>
      <tr><td>Blue</td><td>${formatNumber(values.blueCompra)}</td><td>${formatNumber(values.blueVenta)}</td></tr>
      <tr><td>MEP</td><td>${formatNumber(values.mepCompra)}</td><td>${formatNumber(values.mepVenta)}</td></tr>
      <tr><td>Cripto</td><td>${formatNumber(values.criptoCompra)}</td><td>${formatNumber(values.criptoVenta)}</td></tr>
      <tr><td>Mayorista</td><td>${formatNumber(values.mayoristaCompra)}</td><td>${formatNumber(values.mayoristaVenta)}</td></tr>
    </table>
    <table class="breach-table">
      <tr><td>Brecha Blue/Oficial</td><td class="breach-value">${formatPercentage(values.blueOficialBreach)}</td></tr>
      <tr><td>Brecha Blue/MEP</td><td class="breach-value">${formatPercentage(values.blueMepBreach)}</td></tr>
    </table>
    ${lastUpdateContent}
  `;

  return content;
}

function consultLatestExchangeRate() {
  var ui = getUI();
  try {
    CustomLogger.log('Consultando últimas cotizaciones');
    var apiResponse = fetchDataFromApis(new Date(), 'current');
    if (!apiResponse || !apiResponse.data) {
      throw new Error('No se pudieron obtener datos de la API');
    }

    var values = processCurrentApiData(apiResponse.data);
    
    // Sumar 3 horas a la fecha de actualización
    var updatedDate = new Date(values.fechaActualizacion);
    updatedDate.setHours(updatedDate.getHours() + 3);
    values.fechaActualizacion = updatedDate;

    // Corregir los valores de las brechas
    values.blueOficialBreach /= 100;
    values.blueMepBreach /= 100;

    var message = generateFormattedMessage(new Date(), values, 'current');

    ui.showModalDialog(
      HtmlService.createHtmlOutput(message)
        .setWidth(400)
        .setHeight(600),
      'Consulta de Datos Exitosa'
    );
    CustomLogger.log('Últimas cotizaciones consultadas');
  } catch (error) {
    handleError(error, 'consultLatestExchangeRate', ui);
  }
}

function consultExchangeRateByDate(dateStr) {
  var ui = getUI();
  try {
    CustomLogger.log('Consultando cotizaciones de una fecha específica');
    if (!dateStr) {
      showDateInputDialog();
      return;
    }

    var inputDate = parseCustomDate(dateStr);
    var minDate = new Date('2015-01-01');
    var maxDate = new Date();

    validateDate(inputDate, minDate, maxDate);

    var apiResponse = fetchDataFromApis(inputDate, 'historical');
    if (!apiResponse || !apiResponse.data) {
      throw new Error('No se pudieron obtener datos históricos de la API');
    }

    var values = processHistoricalApiData(apiResponse.data);
    var message = generateFormattedMessage(inputDate, values, 'historical');

    ui.showModalDialog(
      HtmlService.createHtmlOutput(message)
        .setWidth(400)
        .setHeight(600),
      'Consulta de Datos Exitosa'
    );
    CustomLogger.log('Cotizaciones para ' + formatDate(inputDate) + ' consultadas');
  } catch (error) {
    handleError(error, 'consultExchangeRateByDate', ui);
  }
}

function showDateInputDialog() {
  var ui = SpreadsheetApp.getUi();
  var result = ui.prompt(
    'Ingresar Fecha',
    'Ingrese la fecha (dd/mm o dd/mm/yy):',
    ui.ButtonSet.OK_CANCEL);

  var button = result.getSelectedButton();
  var text = result.getResponseText();
  
  if (button == ui.Button.OK) {
    processDateInput(text);
  } else if (button == ui.Button.CANCEL) {
    // El usuario canceló, no hacemos nada
  } else if (button == ui.Button.CLOSE) {
    // El usuario cerró el diálogo, no hacemos nada
  }
}

function processDateInput(dateStr) {
  consultExchangeRateByDate(dateStr);
}

/**
 * Formatea un número para su visualización.
 * @param {number} value - El número a formatear.
 * @returns {string} El número formateado.
 */
function formatNumber(value) {
  return value.toFixed(2).replace('.', ',');
}

/**
 * Prepara una fila de datos para insertar en la hoja.
 * @param {Date} date Fecha de los datos.
 * @param {Object} apiResponse Respuesta de la API que incluye los datos y el tipo de API.
 * @returns {Array} Fila de datos preparada para insertar.
 */
function prepareRowData(date, apiResponse) {
  CustomLogger.log('Preparando fila de datos para ' + formatDate(date));
  try {
    if (!apiResponse || !apiResponse.data) {
      CustomLogger.log('Datos de API no válidos para ' + formatDate(date));
      return null;
    }

    var data = apiResponse.data;
    var apiType = apiResponse.apiType;
    var values = apiType === 'current' ? processCurrentApiData(data) : processHistoricalApiData(data);

    // Función auxiliar para formatear números
    function formatNumber(value, decimals = 2) {
      return value.toFixed(decimals).replace('.', ',');
    }

    // Función auxiliar para formatear porcentajes
    function formatPercentage(value) {
      return (value * 100).toFixed(2).replace('.', ',') + '%';
    }

    var rowData = [
      formatDate(date, 'dayColumn'),
      formatNumber(values.oficialCompra),
      formatNumber(values.oficialVenta),
      formatNumber(values.blueCompra),
      formatNumber(values.blueVenta),
      formatNumber(values.mepCompra),
      formatNumber(values.mepVenta),
      formatNumber(values.criptoCompra),
      formatNumber(values.criptoVenta),
      formatNumber(values.mayoristaCompra),
      formatNumber(values.mayoristaVenta),
      formatPercentage(values.blueOficialBreach),
      formatPercentage(values.blueMepBreach),
      formatDate(new Date(values.fechaActualizacion), 'dateTime'),
      formatDate(new Date(), 'dateTime') // Fecha y hora de modificación del registro
    ];

    CustomLogger.log('Fila de datos preparada para ' + formatDate(date));
    return rowData;
  } catch (error) {
    CustomLogger.error('Error al preparar fila de datos para ' + formatDate(date) + ': ' + error.message);
    return null;
  }
}

/**
 * Actualiza los datos del dólar en la hoja de cálculo.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - La hoja de cálculo.
 * @param {number} row - El número de fila a actualizar.
 * @param {Object} values - Los valores a insertar.
 * @param {Date} date - La fecha de los datos.
 */
function updateDolarData(sheet, row, values, date) {
  CustomLogger.log('Actualizando datos del dólar para ' + formatDate(date));
  try {
    var columnOrder = Object.values(Config.COLUMN_NAMES);
    var fechaActualizacion = new Date(values.fechaActualizacion);
    
    // Sumamos 3 horas a la fecha de actualización
    fechaActualizacion.setHours(fechaActualizacion.getHours() + 3);
    
    var rowData = [
      formatDate(date, 'dayColumn'),
      values.oficialCompra,
      values.oficialVenta,
      values.blueCompra,
      values.blueVenta,
      values.mepCompra,
      values.mepVenta,
      values.criptoCompra,
      values.criptoVenta,
      values.mayoristaCompra,
      values.mayoristaVenta,
      values.blueOficialBreach.toFixed(1).replace('.', ',') + '%',
      values.blueMepBreach.toFixed(1).replace('.', ',') + '%',
      formatDate(fechaActualizacion, 'dateTime'), // Fecha y hora de los datos de la API + 3 horas
      formatDate(new Date(), 'dateTime') // Fecha y hora de modificación del registro
    ];
    sheet.getRange(row, 1, 1, columnOrder.length).setValues([rowData]);
    CustomLogger.log('Datos del dólar actualizados para ' + formatDate(date));
  } catch (error) {
    CustomLogger.error('Error al actualizar datos del dólar para ' + formatDate(date) + ': ' + error.message);
    throw error;
  }
}

/**
 * Actualiza los valores del dólar en la hoja de cálculo.
 * Esta función es el punto de entrada principal para la actualización de datos.
 */
function updateDollarsUntilToday() {
  var ui = getUI();
  try {
    CustomLogger.log('Iniciando actualización de datos del dólar');
    var sheet = getSheet();
    if (!sheet) return;

    var data = getSheetData(sheet);
    var lastRow = data.length;
    var lastDate = new Date(data[lastRow - 1][0]);
    var today = new Date();
    today.setHours(0, 0, 0, 0);

    if (isToday(lastDate)) {
      handleTodayUpdate(ui, sheet, lastRow);
    } else if (isYesterday(lastDate)) {
      handleYesterdayUpdate(ui, sheet, lastRow);
    } else {
      handleHistoricalUpdate(ui, sheet, lastDate, today);
    }

    CustomLogger.log('Actualización de datos del dólar completada');
  } catch (error) {
    handleError(error, 'updateDollarsUntilToday', ui);
  }
}

/**
 * Maneja la actualización de datos para el día de hoy.
 * Esta función utiliza la API de datos actuales para obtener la última cotización del día.
 * 
 * @param {GoogleAppsScript.Base.Ui} ui - Interfaz de usuario de la hoja de cálculo.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Hoja de cálculo.
 * @param {number} lastRow - Número de la última fila con datos.
 * 
 * @description
 * - Utiliza la API de datos actuales (Config.API_CURRENT).
 * - Solicita confirmación al usuario antes de actualizar.
 * - Si se confirma, obtiene los datos más recientes y actualiza la fila correspondiente a hoy.
 * - Es útil para actualizar los valores que pueden cambiar durante el día.
 */
function handleTodayUpdate(ui, sheet, lastRow) {
  var response = ui.alert(
    'Confirmar actualización',
    'La tabla tiene valores actualizados para hoy. ¿Desea actualizar los valores de hoy con los últimos datos disponibles?',
    ui.ButtonSet.YES_NO
  );

  if (response === ui.Button.YES) {
    var today = new Date();
    var apiResponse = fetchDataFromApis(today, 'current');
    if (apiResponse) {
      var values = processCurrentApiData(apiResponse.data);
      updateDolarData(sheet, lastRow, values, today);
      ui.alert('Datos actualizados correctamente con la última información disponible.');
    } else {
      ui.alert('No se pudieron obtener los datos actualizados. Por favor, intente nuevamente más tarde.');
    }
  }
}

/**
 * Maneja la actualización de datos para el día de ayer.
 * Esta función utiliza la API de datos históricos para obtener la cotización del día anterior.
 * 
 * @param {GoogleAppsScript.Base.Ui} ui - Interfaz de usuario de la hoja de cálculo.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Hoja de cálculo.
 * @param {number} lastRow - Número de la última fila con datos.
 * 
 * @description
 * - Utiliza la API de datos históricos (Config.API_HISTORICAL).
 * - Solicita confirmación al usuario antes de agregar los datos de hoy.
 * - Si se confirma, obtiene los datos de ayer y agrega una nueva fila con estos datos.
 * - Es útil para agregar el primer registro del día actual, basado en los datos consolidados de ayer.
 */
function handleYesterdayUpdate(ui, sheet, lastRow) {
  var response = ui.alert(
    'Confirmar actualización',
    'La tabla tiene valores hasta ayer. ¿Desea agregar los valores de hoy basados en los datos de cierre de ayer?',
    ui.ButtonSet.YES_NO
  );

  if (response === ui.Button.YES) {
    var today = new Date();
    var yesterday = new Date(today);
    yesterday.setDate(yesterday.getDate() - 1);
    var apiResponse = fetchDataFromApis(yesterday, 'historical');
    if (apiResponse) {
      var values = processHistoricalApiData(apiResponse.data);
      updateDolarData(sheet, lastRow + 1, values, today);
      ui.alert('Se han agregado los datos de hoy basados en el cierre de ayer.');
    } else {
      ui.alert('No se pudieron obtener los datos de ayer. Por favor, intente nuevamente más tarde.');
    }
  }
}

/**
 * Maneja la actualización histórica de datos.
 * @param {GoogleAppsScript.Base.Ui} ui - La interfaz de usuario.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - La hoja de cálculo.
 * @param {Date} lastDate - La última fecha en la hoja.
 * @param {Date} today - La fecha de hoy.
 */
function handleHistoricalUpdate(ui, sheet, lastDate, today) {
  var response = ui.alert(
    'Confirmar actualización',
    'La tabla tiene valores hasta ' + formatDate(lastDate, 'verbose') + '. ¿Desea actualizar los valores hasta hoy?',
    ui.ButtonSet.YES_NO
  );

  if (response === ui.Button.YES) {
    var yesterday = new Date(today);
    yesterday.setDate(yesterday.getDate() - 1);
    
    // Actualizar datos históricos hasta ayer
    if (lastDate < yesterday) {
      var historicResult = updateHistoricalValues(sheet, lastDate, yesterday);
      if (historicResult.error) {
        ui.alert('Error al actualizar datos históricos: ' + historicResult.error);
        return;
      }
    }
    
    // Agregar datos de hoy
    var apiResponse = fetchDataFromApis(today, 'current');
    if (apiResponse) {
      var values = processCurrentApiData(apiResponse.data);
      updateDolarData(sheet, sheet.getLastRow() + 1, values, today);
      ui.alert('Datos actualizados correctamente hasta hoy.');
    } else {
      ui.alert('No se pudieron obtener los datos de hoy. Se han actualizado los datos históricos hasta ayer.');
    }
  }
}

/**
 * Obtiene datos de la API para una fecha específica.
 * @param {Date} date - Fecha para la cual obtener datos.
 * @param {string} apiType - Tipo de API a utilizar ('current' o 'historical').
 * @returns {Object|null} Datos de la API y tipo de API, o null si hay un error.
 * @throws {Error} Si no se pueden obtener los datos de la API.
 */
function fetchDataFromApis(date, apiType) {
  CustomLogger.log('Obteniendo datos de la API para ' + formatDate(date));
  try {
    var url;

    if (apiType === 'current') {
      url = Config.API_CURRENT;
    } else {
      var formattedDate = formatDate(date, 'api');
      url = Config.API_HISTORICAL + formattedDate;
    }

    var response = getCachedOrFetch(url);
    if (!response) {
      CustomLogger.log('No se obtuvieron datos de la API para ' + formatDate(date));
      return null;
    }
    
    // Asegurarse de que los datos estén en formato JSON
    var data = typeof response === 'string' ? JSON.parse(response) : response;
    
    CustomLogger.log('Datos de la API obtenidos para ' + formatDate(date));
    return { data: data, apiType: apiType };
  } catch (error) {
    CustomLogger.error('Error al obtener datos para ' + formatDate(date) + ': ' + error.message);
    return null;
  }
}

/**
 * Actualiza los valores históricos de una hoja, optimizado para escritura en lotes.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Hoja de cálculo a actualizar.
 * @param {Date} startDate - Fecha de inicio del rango.
 * @param {Date} endDate - Fecha de fin del rango.
 * @returns {Object} Resultado de la operación.
 */
function updateHistoricalValues(sheet, startDate, endDate) {
  CustomLogger.log('Actualizando valores históricos desde ' + formatDate(startDate) + ' hasta ' + formatDate(endDate));
  var startTime = new Date().getTime();

  try {
    validateDate(startDate, new Date('2015-01-01'), new Date());
    validateDate(endDate, startDate, new Date());

    var existingData = getSheetData(sheet);
    var existingDates = new Map(existingData.map(row => [formatDate(new Date(row[0]), 'sheet'), row]));

    var rowsToUpdate = [];
    var rowsToAdd = [];
    var currentDate = new Date(startDate);
    while (currentDate <= endDate) {
      var apiResponse = fetchDataFromApis(currentDate, 'historical');
      if (apiResponse) {
        var rowData = prepareHistoricalRowData(currentDate, apiResponse);
        if (rowData) {
          var dateKey = formatDate(currentDate, 'sheet');
          if (existingDates.has(dateKey)) {
            rowsToUpdate.push({ row: existingData.findIndex(row => formatDate(new Date(row[0]), 'sheet') === dateKey) + 1, data: rowData });
          } else {
            rowsToAdd.push(rowData);
          }
        } else {
          CustomLogger.log('Datos incompletos para la fecha ' + formatDate(currentDate));
        }
      }
      currentDate.setDate(currentDate.getDate() + 1);
    }

    // Actualizar filas existentes
    rowsToUpdate.forEach(update => {
      if (update.row > 0 && update.row <= sheet.getLastRow()) {
        sheet.getRange(update.row, 1, 1, update.data.length).setValues([update.data]);
      } else {
        CustomLogger.log('Fila inválida para actualizar: ' + update.row + '. Agregando como nueva fila.');
        rowsToAdd.push(update.data);
      }
    });

    // Agregar nuevas filas
    if (rowsToAdd.length > 0) {
      var startRow = sheet.getLastRow() + 1;
      var columnOrder = Object.values(Config.COLUMN_NAMES);
      sheet.getRange(startRow, 1, rowsToAdd.length, columnOrder.length).setValues(rowsToAdd);
    }

    sortSheetByDate(sheet);

    var endTime = new Date().getTime();
    var duration = ((endTime - startTime) / 1000).toFixed(2);
    return {
      success: true,
      message: 'Se actualizó el tipo de cambio entre las fechas ' + formatDate(startDate, 'verbose') + ' y ' + formatDate(endDate, 'verbose') + '.\n\nSe actualizaron ' + rowsToUpdate.length + ' filas existentes y se agregaron ' + rowsToAdd.length + ' nuevas filas.\n\nEl proceso tomó ' + duration + ' segundos.'
    };
  } catch (error) {
    CustomLogger.error('Error en updateHistoricalValues: ' + error.message);
    return { error: error.message };
  }
}

/**
 * Procesa los datos de la API para datos históricos.
 * @param {Object} data - Datos de la API.
 * @returns {Object} Valores procesados.
 */
function processHistoricalApiData(data) {
  CustomLogger.log('Procesando datos históricos de la API');
  try {
    var oficialData = data.find(d => d.casa === 'oficial');
    var blueData = data.find(d => d.casa === 'blue');
    var bolsaData = data.find(d => d.casa === 'bolsa');
    var mayoristaData = data.find(d => d.casa === 'mayorista');
    var criptoData = data.find(d => d.casa === 'cripto');

    var values = {
      oficialCompra: parseFloat(oficialData?.compra) || 0,
      oficialVenta: parseFloat(oficialData?.venta) || 0,
      blueCompra: parseFloat(blueData?.compra) || 0,
      blueVenta: parseFloat(blueData?.venta) || 0,
      mepCompra: parseFloat(bolsaData?.compra) || 0,
      mepVenta: parseFloat(bolsaData?.venta) || 0,
      criptoCompra: parseFloat(criptoData?.compra) || 0,
      criptoVenta: parseFloat(criptoData?.venta) || 0,
      mayoristaCompra: parseFloat(mayoristaData?.compra) || 0,
      mayoristaVenta: parseFloat(mayoristaData?.venta) || 0,
      fechaActualizacion: data[0]?.fecha || new Date().toISOString()
    };

    values.blueOficialBreach = values.oficialVenta !== 0 ? ((values.blueVenta - values.oficialVenta) / values.oficialVenta) : 0;
    values.blueMepBreach = values.mepVenta !== 0 ? ((values.blueVenta - values.mepVenta) / values.mepVenta) : 0;

    CustomLogger.log('Datos históricos procesados correctamente');
    return values;
  } catch (error) {
    CustomLogger.error('Error al procesar datos históricos: ' + error.message);
    throw error;
  }
}

// Funciones a desarrollar:

/**
 * Muestra un resumen mensual de las cotizaciones.
 * @todo Implementar la lógica para calcular y mostrar un resumen mensual de cotizaciones.
 */
function monthlyExchangeRateSummary() {
  showDevelopmentMessage('Resumen mensual');
}

/**
 * Exporta los datos a diferentes formatos.
 * @todo Implementar la lógica para exportar datos a formatos como CSV o JSON.
 */
function exportData() {
  showDevelopmentMessage('Exportar datos');
}

/**
 * Permite ajustar las fuentes de datos.
 * @todo Implementar la lógica para configurar las URLs de las APIs o seleccionar entre diferentes fuentes de datos.
 */
function adjustDataSources() {
  showDevelopmentMessage('Ajustar fuentes de datos');
}

/**
 * Configura notificaciones basadas en criterios específicos.
 * @todo Implementar la lógica para configurar alertas o notificaciones basadas en ciertos criterios.
 */
function configureNotifications() {
  showDevelopmentMessage('Configurar notificaciones');
}

/**
 * Muestra una guía de uso detallada de la aplicación.
 * @todo Implementar una guía de uso completa con instrucciones para cada función.
 */
function showUserGuide() {
  showDevelopmentMessage('Guía de uso');
}

/**
 * Muestra información sobre la aplicación.
 * @todo Implementar una pantalla de "Acerca de" con detalles de la aplicación.
 */
function showAbout() {
  showDevelopmentMessage('Acerca de');
}

/**
 * Muestra un mensaje indicando que la función está en desarrollo.
 * @param {string} functionName - El nombre de la función en desarrollo.
 */
function showDevelopmentMessage(functionName) {
  var ui = SpreadsheetApp.getUi();
  ui.alert('Función en Desarrollo', 'La función "' + functionName + '" está actualmente en desarrollo.', ui.ButtonSet.OK);
}
/** Fin. */