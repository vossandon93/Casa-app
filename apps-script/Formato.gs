/**
 * Formato visual de la planilla (colores, anchos, bandas, números).
 *
 * Va en un archivo aparte de Code.gs: en Apps Script es otro archivo del mismo
 * proyecto (botón + junto a "Archivos" → Secuencia de comandos → "Formato").
 * Todo lo de aquí es solo apariencia: no lee ni cambia ningún dato.
 *
 * - darFormato: ejecutar desde el editor para dejar TODA la planilla bonita.
 *   También corre solo una vez al abrir la app con una versión nueva (FORMATO_VERSION).
 * - Los resúmenes se reescriben con cada gasto, así que Code.gs llama a
 *   estiloResumen_ cada vez que escribe uno.
 * - Se guía por los nombres de las hojas y de las columnas: si más adelante se
 *   agrega una hoja o una columna, igual queda con el estilo general.
 */

var FORMATO_VERSION = '2026-11-04 v1';

var ESTILO = {
  fuente: 'Arial',
  tamano: 10,
  texto: '#1F2937',
  textoSuave: '#6B7280',
  encabezado: '#1E3A5F',      // azul petróleo oscuro
  encabezadoTexto: '#FFFFFF',
  banda: '#F3F6FA',           // fila par, muy suave
  blanco: '#FFFFFF',
  borde: '#D5DCE5',
  total: '#E8EEF6',           // columna Total en resúmenes
  titulo: '#1E3A5F'
};

// Colores de pestaña por grupo.
var COLOR_PESTANA = {
  'Gastos': '#1E3A5F',
  'Resumen mensual': '#2E7D5B', 'Resumen anual': '#2E7D5B', 'Resumen por cuenta': '#2E7D5B',
  'Centros': '#9AA5B1', 'Cuentas': '#9AA5B1', 'Personas': '#9AA5B1',
  'Rendiciones': '#B7791F',
  'Azúcar': '#8E5AA8', 'Proteína': '#8E5AA8'
};

// Orden de las pestañas (las que no están aquí quedan al final, como estén).
var ORDEN_PESTANAS = ['Gastos', 'Resumen mensual', 'Resumen anual', 'Resumen por cuenta',
  'Centros', 'Cuentas', 'Personas', 'Rendiciones', 'Azúcar', 'Proteína'];

// El espacio al final separa el monto de la columna siguiente.
var FMT_PESOS = '$#,##0" ";-$#,##0" "';
var FMT_PESOS_CERO = '$#,##0" ";-$#,##0" ";"– "';   // en resúmenes el 0 se ve como "–"

// Estilo por nombre de columna: ancho (px), formato de número y alineación.
var COLUMNAS_ESTILO = {
  'ID': { ancho: 80 },
  'Fecha': { ancho: 95, formato: 'dd/mm/yyyy', alinear: 'center' },
  'Creado': { ancho: 125, formato: 'dd/mm/yyyy hh:mm', alinear: 'center', suave: true },
  'Centro de costo': { ancho: 135 },
  'Cuenta': { ancho: 145 },
  'Monto': { ancho: 105, formato: FMT_PESOS, alinear: 'right' },
  'Total': { ancho: 115, formato: FMT_PESOS, alinear: 'right' },
  'Descripción': { ancho: 280, ajustar: true },
  'Registró': { ancho: 90, alinear: 'center' },
  'Persona': { ancho: 140 },
  'Comprobante': { ancho: 160 },
  'Estado': { ancho: 110, alinear: 'center' },
  'Correo': { ancho: 180 },
  'Evento': { ancho: 160 },
  'Gramos': { ancho: 80, formato: '0', alinear: 'right' },
  'Alimento': { ancho: 200 },
  'Año': { ancho: 60, formato: '0', alinear: 'center' },
  'Mes': { ancho: 55, formato: '00', alinear: 'center' }
};
var ANCHO_POR_DEFECTO = 120;

// ---------------------------------------------------------------------------
// Entrada manual: ejecutar desde el editor.
// ---------------------------------------------------------------------------
function darFormato() {
  var ss = libro_();
  ss.getSheets().forEach(function (hoja) {
    try {
      estiloHoja_(hoja);
    } catch (err) {
      Logger.log('No se pudo dar formato a "' + hoja.getName() + '": ' + err);
    }
  });
  ordenarPestanas_(ss);
  PropertiesService.getScriptProperties().setProperty('FORMATO_VERSION', FORMATO_VERSION);
  Logger.log('Listo: formato aplicado a ' + ss.getSheets().length + ' hojas.');
}

// Corre sola una vez por versión (se llama al abrir la app). Nunca debe romper nada.
function formatoSiFalta_() {
  try {
    if (PropertiesService.getScriptProperties().getProperty('FORMATO_VERSION') !== FORMATO_VERSION) {
      darFormato();
      return;
    }
    // Si la hoja Gastos creció más allá de sus bandas, se le vuelve a dar formato.
    var gastos = libro_().getSheetByName('Gastos');
    var b = gastos && gastos.getBandings()[0];
    if (gastos && (!b || b.getRange().getLastRow() < gastos.getMaxRows())) estiloHoja_(gastos);
  } catch (err) {
    Logger.log('formatoSiFalta_: ' + err);
  }
}

// ---------------------------------------------------------------------------
// Una hoja cualquiera: decide si es resumen (título en A1 + encabezado en fila 2)
// o tabla normal (encabezado en fila 1).
// ---------------------------------------------------------------------------
function estiloHoja_(hoja) {
  if (hoja.isSheetHidden()) return;   // respaldos ocultos: no se tocan
  var nombre = hoja.getName();
  if (COLOR_PESTANA[nombre]) hoja.setTabColor(COLOR_PESTANA[nombre]);
  if (hoja.getLastRow() === 0 || hoja.getLastColumn() === 0) return;
  if (nombre.indexOf('Resumen') === 0) {
    var enc = hoja.getRange(2, 1, 1, hoja.getLastColumn()).getValues()[0];
    var colTotal = enc.indexOf('Total') + 1;
    estiloResumen_(hoja, colTotal > 0 ? colTotal : 2);
    return;
  }
  estiloTabla_(hoja, 1);
}

// Tabla con encabezado en filaEnc y datos debajo.
function estiloTabla_(hoja, filaEnc) {
  var nCols = hoja.getLastColumn();
  var maxFilas = hoja.getMaxRows();
  var enc = hoja.getRange(filaEnc, 1, 1, nCols).getValues()[0].map(function (v) { return String(v).trim(); });
  var nDatos = Math.max(maxFilas - filaEnc, 1);

  // Base: fuente, color y alineación vertical en toda la tabla.
  hoja.getRange(filaEnc, 1, nDatos + 1, nCols)
    .setFontFamily(ESTILO.fuente).setFontSize(ESTILO.tamano).setFontColor(ESTILO.texto)
    .setVerticalAlignment('middle');
  hoja.setHiddenGridlines(true);
  hoja.setFrozenRows(filaEnc);

  // Encabezado.
  hoja.getRange(filaEnc, 1, 1, nCols)
    .setBackground(ESTILO.encabezado).setFontColor(ESTILO.encabezadoTexto)
    .setFontWeight('bold').setHorizontalAlignment('center').setWrap(true);
  hoja.setRowHeight(filaEnc, 30);

  // Columnas: ancho, formato y alineación según su nombre.
  enc.forEach(function (nombre, i) {
    var cfg = COLUMNAS_ESTILO[nombre] || {};
    var col = i + 1;
    hoja.setColumnWidth(col, cfg.ancho || ANCHO_POR_DEFECTO);
    var datos = hoja.getRange(filaEnc + 1, col, nDatos, 1);
    if (cfg.formato) datos.setNumberFormat(cfg.formato);
    datos.setHorizontalAlignment(cfg.alinear || 'left');
    hoja.getRange(filaEnc, col).setHorizontalAlignment(cfg.alinear || 'left');
    datos.setWrapStrategy(cfg.ajustar ? SpreadsheetApp.WrapStrategy.WRAP : SpreadsheetApp.WrapStrategy.CLIP);
    if (cfg.suave) datos.setFontColor(ESTILO.textoSuave);
  });

  // Filas alternadas (bandas nativas de Sheets: siguen al agregar filas).
  bandas_(hoja, hoja.getRange(filaEnc, 1, nDatos + 1, nCols));

  // Filtro en el encabezado (solo en hojas de datos, si no tiene uno).
  if (['Gastos', 'Rendiciones', 'Proteína'].indexOf(hoja.getName()) !== -1 && !hoja.getFilter()) {
    hoja.getRange(filaEnc, 1, Math.max(hoja.getLastRow() - filaEnc + 1, 2), nCols).createFilter();
  }

  if (hoja.getName() === 'Rendiciones') colorEstado_(hoja, enc, filaEnc, nDatos);
}

// ---------------------------------------------------------------------------
// Resúmenes: A1 = título, fila 2 = encabezado, desde la fila 3 los montos.
// Code.gs lo llama después de reescribir cada resumen.
// ---------------------------------------------------------------------------
function estiloResumen_(hoja, primeraColMonto) {
  var nCols = Math.max(hoja.getLastColumn(), 1);
  var nFilas = Math.max(hoja.getLastRow(), 2);
  var enc = hoja.getRange(2, 1, 1, nCols).getValues()[0].map(function (v) { return String(v).trim(); });

  hoja.getRange(1, 1, hoja.getMaxRows(), nCols)
    .setFontFamily(ESTILO.fuente).setFontSize(ESTILO.tamano).setFontColor(ESTILO.texto)
    .setVerticalAlignment('middle');
  hoja.setHiddenGridlines(true);
  hoja.setFrozenRows(2);
  if (COLOR_PESTANA[hoja.getName()]) hoja.setTabColor(COLOR_PESTANA[hoja.getName()]);

  // Título.
  hoja.getRange('A1').setFontSize(14).setFontWeight('bold').setFontColor(ESTILO.titulo);
  hoja.setRowHeight(1, 36);

  // Encabezado.
  hoja.getRange(2, 1, 1, nCols)
    .setBackground(ESTILO.encabezado).setFontColor(ESTILO.encabezadoTexto)
    .setFontWeight('bold').setHorizontalAlignment('center').setWrap(true);
  hoja.setRowHeight(2, 32);
  if (nCols >= primeraColMonto) hoja.getRange(2, primeraColMonto, 1, nCols - primeraColMonto + 1).setHorizontalAlignment('right');

  // Anchos y formatos.
  enc.forEach(function (nombre, i) {
    var col = i + 1;
    var cfg = COLUMNAS_ESTILO[nombre];
    var esMonto = col >= primeraColMonto;
    hoja.setColumnWidth(col, cfg ? cfg.ancho : (esMonto ? 105 : ANCHO_POR_DEFECTO));
  });
  if (nFilas >= 3) {
    var datos = hoja.getRange(3, 1, nFilas - 2, nCols);
    datos.setHorizontalAlignment('left');
    enc.forEach(function (nombre, i) {
      var cfg = COLUMNAS_ESTILO[nombre];
      if (cfg && cfg.formato && i + 1 < primeraColMonto) {
        hoja.getRange(3, i + 1, nFilas - 2, 1).setNumberFormat(cfg.formato).setHorizontalAlignment(cfg.alinear || 'left');
      }
    });
    var montos = hoja.getRange(3, primeraColMonto, nFilas - 2, nCols - primeraColMonto + 1);
    montos.setNumberFormat(FMT_PESOS_CERO).setHorizontalAlignment('right');
  }

  bandas_(hoja, hoja.getRange(2, 1, nFilas - 1, nCols));

  // Columna Total destacada. Va como formato condicional porque las bandas
  // taparían un color de fondo normal. Las reglas de los resúmenes son solo de aquí.
  var colTotal = enc.indexOf('Total') + 1;
  var reglas = [];
  if (colTotal > 0 && nFilas >= 3) {
    var rangoTotal = hoja.getRange(3, colTotal, nFilas - 2, 1);
    rangoTotal.setFontWeight('bold');
    reglas.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=TRUE')
      .setBackground(ESTILO.total).setRanges([rangoTotal]).build());
  }
  hoja.setConditionalFormatRules(reglas);

  // Línea más marcada cuando cambia el año (Resumen mensual y por cuenta).
  var colAnio = enc.indexOf('Año') + 1;
  if (colAnio > 0 && enc.indexOf('Mes') !== -1 && nFilas >= 4) {
    var anios = hoja.getRange(3, colAnio, nFilas - 2, 1).getValues();
    for (var r = 1; r < anios.length; r++) {
      if (anios[r][0] !== anios[r - 1][0]) {
        hoja.getRange(3 + r, 1, 1, nCols).setBorder(true, null, null, null, null, null,
          ESTILO.encabezado, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      }
    }
  }
}

// ---------------------------------------------------------------------------
// Ayudantes
// ---------------------------------------------------------------------------

// Quita las bandas que haya y aplica filas alternadas blanco / gris azulado.
function bandas_(hoja, rango) {
  hoja.getBandings().forEach(function (b) { b.remove(); });
  rango.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false)
    .setHeaderRowColor(ESTILO.encabezado)
    .setFirstRowColor(ESTILO.blanco)
    .setSecondRowColor(ESTILO.banda);
}

// En Rendiciones: Pendiente en ámbar, Reembolsada en verde.
function colorEstado_(hoja, enc, filaEnc, nDatos) {
  var col = enc.indexOf('Estado') + 1;
  if (col === 0) return;
  var rango = hoja.getRange(filaEnc + 1, col, nDatos, 1);
  var otras = hoja.getConditionalFormatRules().filter(function (r) {
    return !r.getRanges().some(function (x) { return x.getColumn() === col; });
  });
  otras.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Pendiente')
    .setBackground('#FDF3D8').setFontColor('#8A5A00').setRanges([rango]).build());
  otras.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Reembolsada')
    .setBackground('#DFF3E7').setFontColor('#1E6B3F').setRanges([rango]).build());
  hoja.setConditionalFormatRules(otras);
}

function ordenarPestanas_(ss) {
  var pos = 1;
  ORDEN_PESTANAS.forEach(function (nombre) {
    var hoja = ss.getSheetByName(nombre);
    if (!hoja) return;
    ss.setActiveSheet(hoja);
    ss.moveActiveSheet(pos++);
  });
  var gastos = ss.getSheetByName('Gastos');
  if (gastos) ss.setActiveSheet(gastos);
}
