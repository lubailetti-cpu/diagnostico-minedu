/**
 * Diagnóstico de carga operativa — MINEDU
 *
 * Este script recibe las respuestas del formulario HTML y las guarda
 * en la hoja de cálculo "Respuestas" del documento donde está pegado.
 *
 * Cómo instalarlo (instrucciones detalladas en el tutorial):
 *  1. Crear una nueva hoja de cálculo en Google Drive.
 *  2. Menú: Extensiones > Apps Script.
 *  3. Pegar este código completo en el editor (borrar lo que venga por defecto).
 *  4. Guardar (Ctrl+S).
 *  5. Menú: Implementar > Nueva implementación > Tipo: Aplicación web.
 *  6. Configurar:
 *       - Ejecutar como: Tú mismo
 *       - Quién tiene acceso: Cualquier persona
 *  7. Copiar la URL que termine en /exec y enviársela a Lucía.
 *
 * La primera vez Google va a pedir autorización para que el script
 * pueda escribir en la hoja. Es normal; aceptar.
 */

// Si querés, podés cambiar este nombre de hoja. Por defecto se usa "Respuestas".
const SHEET_NAME = 'Respuestas';

/**
 * Recibe POST del formulario y agrega una fila con las respuestas.
 */
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);

    // Si la hoja no existe, la crea con encabezados.
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      writeHeaders(sheet);
    }

    // Si está vacía, escribir encabezados primero.
    if (sheet.getLastRow() === 0) {
      writeHeaders(sheet);
    }

    // Construir fila siguiendo el orden de columnas.
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const row = headers.map(h => data[h] !== undefined ? data[h] : '');

    sheet.appendRow(row);

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Para verificar manualmente que el endpoint funciona.
 * Si abrís la URL del Apps Script en el navegador, verás este mensaje.
 */
function doGet() {
  return ContentService
    .createTextOutput('Endpoint del diagnóstico MINEDU está activo. Para enviar datos, usar POST con JSON.')
    .setMimeType(ContentService.MimeType.TEXT);
}

/**
 * Escribe los encabezados de las columnas en la hoja.
 * El orden define el orden de columnas en el Sheet.
 */
function writeHeaders(sheet) {
  const headers = [
    'timestamp',
    'nombre',
    'p1_regimen',
    'p2_unidad',
    'p3_rol',
    'p4_anios',
    'p5_frecuencia',
    // P6: matriz volumen + tiempo por tipo de documento (9 tipos × 2 columnas)
    'p6_cant_0', 'p6_tiempo_0',
    'p6_cant_1', 'p6_tiempo_1',
    'p6_cant_2', 'p6_tiempo_2',
    'p6_cant_3', 'p6_tiempo_3',
    'p6_cant_4', 'p6_tiempo_4',
    'p6_cant_5', 'p6_tiempo_5',
    'p6_cant_6', 'p6_tiempo_6',
    'p6_cant_7', 'p6_tiempo_7',
    'p6_cant_8', 'p6_tiempo_8',
    'p7_dificultad',
    'p8_plantillas',
    'p9_reproceso',
    'p10_rondas',
    'p11_cambios',
    'p12_devolucion',
    'p13_razon_dev',
    'p14_aprobadores',
    'p15_busqueda_norma',
    'p16_busqueda_antec',
    'p17_seguridad',
    'p18_urgentes',
    'p19_plazo',
    'p20_poi',
    'p21_jornada',
    'p22_agotamiento',
    'p23_satis_calidad',
    'p24_satis_tiempo',
    'p25_carga_equitativa',
    'p26_uso_tiempo',
    'p27_nivel_ia',
    'p28_usa_ia',
    'p29_tareas_ia',
    // P30: 6 afirmaciones Likert
    'p30_lik_reduce_tiempo',
    'p30_lik_calidad_borradores',
    'p30_lik_mas_tiempo_estrategico',
    'p30_lik_confiaria',
    'p30_lik_errores_haria_dejar',
    'p30_lik_piloto',
    'p31_preocup',
    'p32_docs_utiles',
    'p33_condicion',
    // P34: matriz de impacto (5 tipos)
    'p34_impacto_0',
    'p34_impacto_1',
    'p34_impacto_2',
    'p34_impacto_3',
    'p34_impacto_4',
    'p35_repetitivas',
    'p36_funcionalidad',
    'p37_ojala'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  // Formato bonito a los encabezados
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#1E3A5F')
    .setFontColor('#ffffff');
  sheet.setFrozenRows(1);
}
