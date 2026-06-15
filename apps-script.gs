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
 */

const SHEET_NAME = 'Respuestas';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      writeHeaders(sheet);
    }
    if (sheet.getLastRow() === 0) {
      writeHeaders(sheet);
    }

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

function doGet() {
  return ContentService
    .createTextOutput('Endpoint del diagnóstico MINEDU está activo. Para enviar datos, usar POST con JSON.')
    .setMimeType(ContentService.MimeType.TEXT);
}

function writeHeaders(sheet) {
  const headers = [
    'timestamp',
    'p0_nivel',
    'p1_regimen',
    'p2_categoria',
    'p2_unidad',
    'p2_dre',
    'p2_ugel',
    'p3_rol',
    'p4_anios',
    // P5: distribución porcentual del tiempo semanal (4 categorías)
    'p6_distrib_sustantivas',
    'p6_distrib_documentos',
    'p6_distrib_administrativas',
    'p6_distrib_externas',
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
    'p12_devolucion',
    'p13_razon_dev',
    'p15_busqueda_norma',
    'p15_vigencia',
    'p18_urgentes',
    'p19_plazo',
    // BCP — afirmaciones (9 + 2 DRE/UGEL) escala semáforo 1-4
    'pbcp_bcp_poi_retrasos',
    'pbcp_bcp_carga_no_distrib',
    'pbcp_bcp_sobre_horas',
    'pbcp_bcp_faltan_manos',
    'pbcp_bcp_satis_calidad',
    'pbcp_bcp_manual_auto',
    'pbcp_bcp_info_personas',
    'pbcp_bcp_datos_dispersos',
    'pbcp_bcp_indicadores',
    'pbcp_bcp_normas_iiee',
    'pbcp_bcp_consultas_repetidas',
    'p26_uso_tiempo',
    'p27_nivel_ia',
    'p28_usa_ia',
    'p29_tareas_ia',
    // Likert plataforma (4)
    'p30_lik_reduce_tiempo',
    'p30_lik_calidad_borradores',
    'p30_lik_mas_tiempo_estrategico',
    'p30_lik_piloto',
    'p31_preocup',
    'p33_condicion',
    // Situaciones específicas y priorización
    'situaciones',
    'situaciones_orden',
    // ===== Específicas DRE / UGEL (vacías si nivel = MINEDU) =====
    'pdu_num_iiee',
    'pdu_visitas_mes',
    'pdu_tiempo_informe',
    'pdu_pre_visita',
    'pdu_consolidar',
    'pdu_traduce_normas',
    'pdu_consultas',
    'pdu_oportunidades'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#1E3A5F')
    .setFontColor('#ffffff');
  sheet.setFrozenRows(1);
}
