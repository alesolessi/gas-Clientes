/**
 * Crea el menú personalizado en la hoja de cálculo.
 * Este menú permite al usuario acceder a diversas funciones relacionadas con las cotizaciones de dólares.
 */
function onOpen() {
    Logger.log("Ejecutando onOpen"); // Para depuración
    var ui = SpreadsheetApp.getUi();
    
    // Menú principal "Cotizaciones Dolar"
    ui.createMenu('Cotizaciones Dolar')
      .addSubMenu(ui.createMenu('Actualizar Cotizaciones')
        .addItem('Actualizar desde Última Fila Hasta Hoy', 'updateDollarsUntilToday') // Actualiza las cotizaciones desde la última fila hasta hoy
        .addItem('Actualizar un Rango de Fechas Histórico', 'updateDollarsByRange')) // Actualiza cotizaciones en un rango de fechas
      .addSubMenu(ui.createMenu('Consultar Cotizaciones')
        .addItem('Últimas Cotizaciones', 'consultLatestExchangeRate') // Consulta las últimas cotizaciones
        .addItem('Cotización Histórica por Fecha', 'consultExchangeRateByDate') // Consulta la cotización histórica por fecha
        .addItem('Resumen Mensual', 'monthlyExchangeRateSummary')) // Muestra un resumen mensual de las cotizaciones
      .addToUi();
    
    // Menú de Importación de Datos
    ui.createMenu('Importación de Datos') // Menú separado para importación de datos
      .addItem('Importar Clientes con Situación Financiera', 'importCustomers') // Ítem de menú para importar clientes
      .addToUi();
    
    // Llamar a la función para posicionar el cursor
    positionToLastRow(); // Posiciona el cursor en la última fila de la hoja
}

/**
 * Posiciona el cursor en la última fila de la hoja 'Base Ventas'.
 * Si la hoja no existe o no hay filas, no se realiza ninguna acción.
 */
function positionToLastRow() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Clientes');
    if (sheet) {
        var lastRow = sheet.getLastRow();
        if (lastRow > 0) {
            sheet.setActiveRange(sheet.getRange(lastRow, 1)); // Posiciona el cursor en la última fila, columna A
        }
    }
}