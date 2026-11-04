/**
 * Backend de Gastos Familia (Google Apps Script).
 *
 * Se pega en Extensiones > Apps Script de la planilla de Google.
 * Ver SETUP.md para las instrucciones completas.
 *
 * El PIN se guarda en Configuración del proyecto > Propiedades del script,
 * con la clave PIN. Nunca se escribe en este archivo.
 */

var HOJA_GASTOS = 'Gastos';
var HOJA_CENTROS = 'Centros';
var HOJA_PERSONAS = 'Personas';
var HOJA_CUENTAS = 'Cuentas';
var HOJA_RESUMEN_MENSUAL = 'Resumen mensual';
var HOJA_RESUMEN_ANUAL = 'Resumen anual';
var HOJA_RESUMEN_CUENTAS = 'Resumen por cuenta';
var HOJA_RENDICIONES = 'Rendiciones';
var HOJA_AZUCAR = 'Azúcar';
var HOJA_PROTEINA = 'Proteína';
var HOJA_ENTRENOS = 'Entrenamientos';
var HOJA_DIARIO = 'Diario Vicente';

// Orden lógico que usa el código (f[0]..f[7]). En la hoja las columnas se buscan por
// su encabezado, así que pueden estar en otro orden (ORDEN_HOJA_GASTOS).
var COLUMNAS_GASTOS = ['ID', 'Fecha', 'Centro de costo', 'Monto', 'Descripción', 'Registró', 'Creado', 'Cuenta'];
var ORDEN_HOJA_GASTOS = ['ID', 'Fecha', 'Centro de costo', 'Cuenta', 'Monto', 'Descripción', 'Registró', 'Creado'];
var CENTROS_INICIALES = ['Auto', 'Familia', 'Salud', 'Vicky', 'Fran', 'Vicente', 'Casa'];
var PERSONAS_INICIALES = ['Vicente', 'Fran'];
// Cuentas que se crean solas si existe un centro con ese nombre (sin importar mayúsculas).
var CUENTAS_INICIALES = {
  'Auto': ['Seguro', 'Bencina', 'TAG'],
  'Familia': ['Regalos', 'Eventos / Salidas', 'Supermercado', 'Luna'],
  'Salud': ['Isapre', 'Farmacia'],
  'Vicky': ['Sueldo', 'Traslado', 'Imposiciones'],
  'Casa': ['Fumigación', 'Conectividad', 'Gas', 'Agua', 'Tintorería']
};
// Cuentas que cambiaron de nombre: { centro: { 'nombre viejo': 'nombre nuevo' } }.
var CUENTAS_RENOMBRADAS = { 'Vicky': { 'Contribuciones': 'Imposiciones' } };
// Centros antiguos que se juntan con otro (la cuenta se mantiene): { 'viejo': 'nuevo' }.
var CENTROS_RENOMBRADOS = { 'Nana': 'Vicky' };
var ULTIMOS_GASTOS = 20;
var COLUMNAS_RENDICIONES = ['ID', 'Fecha', 'Monto', 'Descripción', 'Comprobante', 'Estado', 'Correo', 'Registró', 'Creado'];
var COLUMNAS_AZUCAR = ['Fecha', 'Evento', 'Creado'];
// Un registro por persona y día. Las filas viejas (una por comida, sin Persona) son de Fran.
var COLUMNAS_PROTEINA = ['ID', 'Fecha', 'Gramos', 'Nota', 'Creado', 'Persona'];
var CARPETA_RENDICIONES = 'Rendiciones - comprobantes';
var META_PROTEINA_INICIAL = 100;
// Deporte (pestaña de Vicente): entrenamientos y diario de cigarros, comida y peso.
// La proteína de Vicente va en la hoja Proteína (mismo componente que Salud).
var COLUMNAS_ENTRENOS = ['ID', 'Fecha', 'Categorías', 'Minutos', 'Ejercicios', 'Nota', 'Creado'];
var COLUMNAS_DIARIO = ['Fecha', 'Cigarros', 'Comida', 'Etiquetas', 'Peso (kg)', 'Actualizado'];
var CATEGORIAS_DEPORTE = ['Upper', 'Lower', 'Core', 'Cardio', 'Full body', 'Movilidad'];
var METAS_DEPORTE = {
  semana: { prop: 'DEPORTE_META_SEMANA', inicial: 4, min: 1, max: 7 },
  peso: { prop: 'DEPORTE_META_PESO', inicial: 76, min: 30, max: 200 }
};
var ULTIMAS_RENDICIONES = 50;
// Se muestra al abrir la URL /exec: sirve para comprobar que Google tiene el código nuevo.
var VERSION = '2026-11-04 · proteína diaria con nota + pestaña Deporte';

// ---------------------------------------------------------------------------
// Instalación: ejecutar una vez desde el editor (botón "Ejecutar" con setup).
// ---------------------------------------------------------------------------

function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    // Proyecto creado en script.google.com (no dentro de una planilla):
    // se usa la planilla creada antes o se crea una nueva en tu Drive.
    var props = PropertiesService.getScriptProperties();
    var id = props.getProperty('SHEET_ID');
    ss = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.create('Gastos Familia');
    props.setProperty('SHEET_ID', ss.getId());
    Logger.log('Planilla: ' + ss.getUrl());
  }
  ss.setSpreadsheetTimeZone('America/Santiago');

  var gastos = obtenerOCrearHoja_(ss, HOJA_GASTOS);
  if (gastos.getLastRow() === 0) {
    gastos.appendRow(ORDEN_HOJA_GASTOS);
  }
  ordenarColumnasGastos_(gastos);
  formatoGastos_(gastos);
  gastos.hideColumns(1);
  asegurarColumnaCuenta_(gastos);

  var centros = obtenerOCrearHoja_(ss, HOJA_CENTROS);
  if (centros.getLastRow() === 0) {
    centros.appendRow(['Centro de costo']);
    CENTROS_INICIALES.forEach(function (c) { centros.appendRow([c]); });
    centros.getRange(1, 1).setFontWeight('bold');
    centros.setFrozenRows(1);
  }

  var personas = obtenerOCrearHoja_(ss, HOJA_PERSONAS);
  if (personas.getLastRow() === 0) {
    personas.appendRow(['Persona']);
    PERSONAS_INICIALES.forEach(function (p) { personas.appendRow([p]); });
    personas.getRange(1, 1).setFontWeight('bold');
    personas.setFrozenRows(1);
  }

  var cuentas = obtenerOCrearHoja_(ss, HOJA_CUENTAS);
  if (cuentas.getLastRow() === 0) {
    cuentas.appendRow(['Centro de costo', 'Cuenta']);
    cuentas.getRange(1, 1, 1, 2).setFontWeight('bold');
    cuentas.setFrozenRows(1);
    leerLista_(HOJA_CENTROS, ss).forEach(function (centro) {
      Object.keys(CUENTAS_INICIALES).forEach(function (k) {
        if (k.toLowerCase() !== centro.toLowerCase()) return;
        CUENTAS_INICIALES[k].forEach(function (c) { cuentas.appendRow([centro, c]); });
      });
    });
  }

  crearResumenes_(ss);
  aplicarListasGastos_(ss);

  prepararHoja_(ss, HOJA_RENDICIONES, COLUMNAS_RENDICIONES);
  prepararHoja_(ss, HOJA_AZUCAR, COLUMNAS_AZUCAR);
  prepararHoja_(ss, HOJA_PROTEINA, COLUMNAS_PROTEINA);
  prepararHoja_(ss, HOJA_ENTRENOS, COLUMNAS_ENTRENOS);
  prepararHoja_(ss, HOJA_DIARIO, COLUMNAS_DIARIO);

  // Pide (una vez) los permisos de Drive y Gmail que usan las rendiciones.
  carpetaRendiciones_();
  MailApp.getRemainingDailyQuota();

  var hoja1 = ss.getSheetByName('Hoja 1') || ss.getSheetByName('Sheet1');
  if (hoja1 && hoja1.getLastRow() === 0 && ss.getSheets().length > 1) {
    ss.deleteSheet(hoja1);
  }

  if (!PropertiesService.getScriptProperties().getProperty('PIN')) {
    Logger.log('Falta definir el PIN en Configuración del proyecto > Propiedades del script.');
  }
  if (!PropertiesService.getScriptProperties().getProperty('CORREO_RENDICION')) {
    Logger.log('Para que las rendiciones se envíen por correo, define CORREO_RENDICION en Propiedades del script.');
  }
}

// Crea la hoja con sus encabezados y formatos si no existe. Se usa también
// desde la API, así las hojas nuevas aparecen aunque no se vuelva a correr setup.
function prepararHoja_(ss, nombre, columnas) {
  var hoja = ss.getSheetByName(nombre);
  if (hoja && hoja.getLastRow() > 0) return hoja;
  hoja = hoja || ss.insertSheet(nombre);
  hoja.appendRow(columnas);
  hoja.setFrozenRows(1);
  hoja.getRange(1, 1, 1, columnas.length).setFontWeight('bold');
  columnas.forEach(function (c, i) {
    var col = hoja.getRange(1, i + 1, hoja.getMaxRows(), 1);
    if (c === 'Fecha') col.setNumberFormat('dd/mm/yyyy');
    if (c === 'Creado') col.setNumberFormat('dd/mm/yyyy hh:mm');
    if (c === 'Monto') col.setNumberFormat('$#,##0');
  });
  hoja.getRange(1, 1, 1, columnas.length).setNumberFormat('@');
  if (columnas[0] === 'ID') hoja.hideColumns(1);
  return hoja;
}

function hoja_(nombre, columnas) {
  return prepararHoja_(libro_(), nombre, columnas);
}

// ---------------------------------------------------------------------------
// Ordenar centros y cuentas: ejecutar UNA vez desde el editor (reordenarCentrosYCuentas).
//
// 1. Deja las hojas Centros y Cuentas igual a CENTROS_INICIALES / CUENTAS_INICIALES.
// 2. Mueve los gastos antiguos cuyo centro ahora es una cuenta (Luna, Supermercado,
//    Farmacia…) a su centro nuevo, con esa cuenta. Se cambia la MISMA fila: no se
//    copia nada, así que ningún gasto queda en dos centros ni se suma doble.
// 3. Junta centros antiguos (Nana → Vicky) y renombra cuentas (Contribuciones → Imposiciones).
// 4. Antes de tocar nada guarda una copia de la hoja Gastos ("Gastos respaldo …").
// Se puede ejecutar de nuevo sin problema: lo que ya está ordenado no cambia.
// ---------------------------------------------------------------------------
function reordenarCentrosYCuentas() {
  var ss = libro_();
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var gastos = hojaGastos_();
    asegurarColumnaCuenta_(gastos);
    var tz = ss.getSpreadsheetTimeZone();
    var respaldo = gastos.copyTo(ss).setName('Gastos respaldo ' + Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH.mm'));
    respaldo.hideSheet();

    // Centros
    var hojaCentros = obtenerOCrearHoja_(ss, HOJA_CENTROS);
    hojaCentros.clearContents();
    hojaCentros.getRange(1, 1, CENTROS_INICIALES.length + 1, 1)
      .setValues([['Centro de costo']].concat(alfabetico_(CENTROS_INICIALES).map(function (c) { return [c]; })));
    hojaCentros.getRange(1, 1).setFontWeight('bold');

    // Cuentas
    var filasCuentas = [['Centro de costo', 'Cuenta']];
    alfabetico_(Object.keys(CUENTAS_INICIALES)).forEach(function (centro) {
      alfabetico_(CUENTAS_INICIALES[centro]).forEach(function (c) { filasCuentas.push([centro, c]); });
    });
    var hojaCuentas = obtenerOCrearHoja_(ss, HOJA_CUENTAS);
    hojaCuentas.clearContents();
    hojaCuentas.getRange(1, 1, filasCuentas.length, 2).setValues(filasCuentas);
    hojaCuentas.getRange(1, 1, 1, 2).setFontWeight('bold');

    // Cada cuenta → su centro (sin mayúsculas), para mover los centros antiguos.
    var centroDeCuenta = {};
    Object.keys(CUENTAS_INICIALES).forEach(function (centro) {
      CUENTAS_INICIALES[centro].forEach(function (c) { centroDeCuenta[c.toLowerCase()] = { centro: centro, cuenta: c }; });
    });
    var centroOficial = {};
    CENTROS_INICIALES.forEach(function (c) { centroOficial[c.toLowerCase()] = c; });

    var cambios = {}, sinCentro = {}, n = gastos.getLastRow() - 1;
    if (n > 0) {
      var filas = leerGastos_(gastos);
      filas.forEach(function (f) {
        if (!f[0] && !f[3]) return;
        var centro = String(f[2]).trim(), cuenta = String(f[7] || '').trim();
        var antes = centro + (cuenta ? ' / ' + cuenta : '');
        Object.keys(CENTROS_RENOMBRADOS).forEach(function (viejo) {
          if (centro.toLowerCase() === viejo.toLowerCase()) centro = CENTROS_RENOMBRADOS[viejo];
        });
        var nuevo = centroOficial[centro.toLowerCase()];
        if (nuevo) {
          centro = nuevo;
        } else if (centroDeCuenta[centro.toLowerCase()]) {
          var destino = centroDeCuenta[centro.toLowerCase()];
          centro = destino.centro;
          cuenta = cuenta || destino.cuenta;
        } else {
          sinCentro[centro] = (sinCentro[centro] || 0) + 1;
        }
        var ren = CUENTAS_RENOMBRADAS[centro] || {};
        Object.keys(ren).forEach(function (viejo) { if (cuenta.toLowerCase() === viejo.toLowerCase()) cuenta = ren[viejo]; });
        var despues = centro + (cuenta ? ' / ' + cuenta : '');
        if (despues !== antes) {
          f[2] = centro;
          f[7] = cuenta;
          var k = antes + '  →  ' + despues;
          cambios[k] = (cambios[k] || 0) + 1;
        }
      });
      guardarCentroYCuenta_(gastos, filas);
    }
    actualizarResumenes_(ss);

    var lineas = ['Respaldo guardado en la hoja oculta "' + respaldo.getName() + '".'];
    var claves = Object.keys(cambios);
    lineas.push(claves.length ? 'Gastos movidos:' : 'No había gastos que mover.');
    claves.forEach(function (k) { lineas.push('  ' + k + '  (' + cambios[k] + ')'); });
    Object.keys(sinCentro).forEach(function (c) {
      lineas.push('OJO: ' + sinCentro[c] + ' gasto(s) con centro "' + c + '" no calzan con ningún centro ni cuenta: quedaron igual.');
    });
    Logger.log(lineas.join('\n'));
    PropertiesService.getScriptProperties().setProperty('ORDEN_CENTROS', ORDEN_CENTROS_VERSION);
  } finally {
    lock.releaseLock();
  }
}

// Los resúmenes se escriben como valores calculados por el script (sin fórmulas
// QUERY, que fallaban según la configuración regional de la planilla). Se
// recalculan cada vez que se guarda o borra un gasto desde la app. Si editas la
// hoja Gastos a mano, ejecuta actualizarResumenes desde el editor.
function crearResumenes_(ss) {
  actualizarResumenes_(ss);
}

function actualizarResumenes() {
  actualizarResumenes_(libro_());
}

function actualizarResumenes_(ss) {
  var filas = totalesPorMes_(ss);
  var centros = leerLista_(HOJA_CENTROS, ss);
  filas.forEach(function (f) { if (centros.indexOf(f.centro) === -1) centros.push(f.centro); });
  centros.sort(function (a, b) { return a.localeCompare(b, 'es', { sensitivity: 'base' }); });

  var porMes = {};
  var porAnio = {};
  filas.forEach(function (f) {
    var anio = f.mes.slice(0, 4);
    porMes[f.mes] = porMes[f.mes] || {};
    porMes[f.mes][f.centro] = (porMes[f.mes][f.centro] || 0) + f.total;
    porAnio[anio] = porAnio[anio] || {};
    porAnio[anio][f.centro] = (porAnio[anio][f.centro] || 0) + f.total;
  });

  var mensual = [['Año', 'Mes', 'Total'].concat(centros)];
  Object.keys(porMes).sort().reverse().forEach(function (mes) {
    mensual.push([Number(mes.slice(0, 4)), Number(mes.slice(5, 7))].concat(filaTotales_(porMes[mes], centros)));
  });
  escribirResumen_(obtenerOCrearHoja_(ss, HOJA_RESUMEN_MENSUAL), 'Total por mes y centro de costo', mensual, 3);

  var anual = [['Año', 'Total'].concat(centros)];
  Object.keys(porAnio).sort().reverse().forEach(function (anio) {
    anual.push([Number(anio)].concat(filaTotales_(porAnio[anio], centros)));
  });
  escribirResumen_(obtenerOCrearHoja_(ss, HOJA_RESUMEN_ANUAL), 'Total por año y centro de costo', anual, 2);

  var porCuenta = {};
  filas.forEach(function (f) {
    var clave = f.mes + '|' + f.centro + '|' + f.cuenta;
    porCuenta[clave] = porCuenta[clave] || [Number(f.mes.slice(0, 4)), Number(f.mes.slice(5, 7)), f.centro, f.cuenta || 'Sin cuenta', 0];
    porCuenta[clave][4] += f.total;
  });
  var filasCuenta = Object.keys(porCuenta).map(function (k) { return porCuenta[k]; }).sort(function (a, b) {
    return (b[0] - a[0]) || (b[1] - a[1]) || a[2].localeCompare(b[2], 'es') || (b[4] - a[4]);
  });
  escribirResumen_(obtenerOCrearHoja_(ss, HOJA_RESUMEN_CUENTAS), 'Total por mes, centro de costo y cuenta',
    [['Año', 'Mes', 'Centro de costo', 'Cuenta', 'Total']].concat(filasCuenta), 5);
}

function filaTotales_(montos, centros) {
  var valores = centros.map(function (c) { return montos[c] || 0; });
  var total = valores.reduce(function (a, b) { return a + b; }, 0);
  return [total].concat(valores);
}

function escribirResumen_(hoja, titulo, tabla, primeraColMonto) {
  hoja.clear();
  hoja.getRange('A1').setValue(titulo).setFontWeight('bold');
  hoja.getRange(2, 1, tabla.length, tabla[0].length).setValues(tabla);
  hoja.getRange(2, 1, 1, tabla[0].length).setFontWeight('bold');
  if (tabla.length > 1) {
    hoja.getRange(3, primeraColMonto, tabla.length - 1, tabla[0].length - primeraColMonto + 1).setNumberFormat('$#,##0');
  } else {
    hoja.getRange('A3').setValue('Sin gastos todavía');
  }
  hoja.setFrozenRows(2);
}

// Suma de montos agrupada por mes (yyyy-MM), centro, cuenta y persona.
function totalesPorMes_(ss) {
  ss = ss || libro_();
  var hoja = ss.getSheetByName(HOJA_GASTOS);
  if (!hoja || hoja.getLastRow() < 2) return [];
  var tz = ss.getSpreadsheetTimeZone();
  var grupos = {};
  leerGastos_(hoja).forEach(function (f) {
    var monto = Number(f[3]) || 0;
    if (!(f[1] instanceof Date) || !monto) return;
    var mes = Utilities.formatDate(f[1], tz, 'yyyy-MM');
    var centro = String(f[2]).trim();
    var persona = String(f[5]).trim();
    var cuenta = String(f[7] || '').trim();
    var clave = mes + '|' + centro + '|' + cuenta + '|' + persona;
    grupos[clave] = grupos[clave] || { mes: mes, centro: centro, cuenta: cuenta, persona: persona, total: 0 };
    grupos[clave].total += monto;
  });
  return Object.keys(grupos).map(function (k) { return grupos[k]; });
}

function libro_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) return ss;
  var id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if (!id) throw new Error('Falta la planilla: ejecuta setup() primero');
  return SpreadsheetApp.openById(id);
}

function obtenerOCrearHoja_(ss, nombre) {
  return ss.getSheetByName(nombre) || ss.insertSheet(nombre);
}

// Planillas creadas antes de las cuentas no tienen el título de esa columna.
function asegurarColumnaCuenta_(hoja) {
  if (encabezadosGastos_(hoja).indexOf('Cuenta') !== -1) return;
  var col = Math.max(hoja.getLastColumn(), COLUMNAS_GASTOS.length - 1) + 1;
  hoja.getRange(1, col).setValue('Cuenta').setFontWeight('bold');
}

function encabezadosGastos_(hoja) {
  var ancho = Math.max(hoja.getLastColumn(), 1);
  return hoja.getRange(1, 1, 1, ancho).getValues()[0].map(function (v) { return String(v).trim(); });
}

// Para cada columna lógica (COLUMNAS_GASTOS), su posición (0..) en la hoja.
function mapaGastos_(hoja) {
  var enc = encabezadosGastos_(hoja);
  return COLUMNAS_GASTOS.map(function (nombre, i) {
    var j = enc.indexOf(nombre);
    return j === -1 ? i : j;
  });
}

// Filas de la hoja Gastos en orden lógico, desde la fila `desde` (por defecto todas).
function leerGastos_(hoja, desde, n) {
  desde = desde || 2;
  if (n === undefined) n = hoja.getLastRow() - desde + 1;
  if (n <= 0) return [];
  var mapa = mapaGastos_(hoja);
  var ancho = Math.max.apply(null, mapa) + 1;
  return hoja.getRange(desde, 1, n, ancho).getValues().map(function (f) {
    return mapa.map(function (j) { return f[j]; });
  });
}

function filaFisicaGastos_(hoja, logica) {
  var mapa = mapaGastos_(hoja);
  var fila = [];
  for (var k = 0; k <= Math.max.apply(null, mapa); k++) fila.push('');
  mapa.forEach(function (j, i) { fila[j] = logica[i]; });
  return fila;
}

// Escribe solo las columnas Centro de costo y Cuenta (lo único que cambian las migraciones).
function guardarCentroYCuenta_(hoja, filas) {
  if (!filas.length) return;
  var mapa = mapaGastos_(hoja);
  hoja.getRange(2, mapa[2] + 1, filas.length, 1).setValues(filas.map(function (f) { return [f[2]]; }));
  hoja.getRange(2, mapa[7] + 1, filas.length, 1).setValues(filas.map(function (f) { return [f[7]]; }));
}

// Deja la hoja Gastos con la Cuenta al lado del Centro de costo. Mueve la columna
// completa (con sus datos y formato): no se copia ni se duplica ningún gasto.
function ordenarColumnasGastos_(hoja) {
  asegurarColumnaCuenta_(hoja);
  var enc = encabezadosGastos_(hoja);
  var colCentro = enc.indexOf('Centro de costo') + 1;
  var colCuenta = enc.indexOf('Cuenta') + 1;
  if (colCentro > 0 && colCuenta > colCentro + 1) {
    hoja.moveColumns(hoja.getRange(1, colCuenta), colCentro + 1);
    return true;
  }
  return false;
}

function formatoGastos_(hoja) {
  var enc = encabezadosGastos_(hoja);
  hoja.setFrozenRows(1);
  hoja.getRange(1, 1, 1, enc.length).setFontWeight('bold');
  var formatos = { 'Fecha': 'dd/mm/yyyy', 'Monto': '$#,##0', 'Creado': 'dd/mm/yyyy hh:mm' };
  Object.keys(formatos).forEach(function (nombre) {
    var c = enc.indexOf(nombre) + 1;
    if (c > 0) hoja.getRange(2, c, Math.max(hoja.getMaxRows() - 1, 1), 1).setNumberFormat(formatos[nombre]);
  });
}

// Corre sola una vez (al abrir la app con este código) o a mano con ordenarHojaGastos.
var COLUMNAS_VERSION = '2026-11-03 cuenta al lado';
function ordenarHojaGastosSiFalta_() {
  if (propiedad_('COLUMNAS_VERSION') === COLUMNAS_VERSION) return;
  ordenarHojaGastos();
}

function ordenarHojaGastos() {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var hoja = hojaGastos_();
    var movida = ordenarColumnasGastos_(hoja);
    formatoGastos_(hoja);
    PropertiesService.getScriptProperties().setProperty('COLUMNAS_VERSION', COLUMNAS_VERSION);
    Logger.log(movida ? 'Listo: la columna Cuenta quedó al lado de Centro de costo.' : 'La hoja Gastos ya estaba ordenada.');
  } finally {
    lock.releaseLock();
  }
}

// ---------------------------------------------------------------------------
// Editar la planilla a mano
// - onEdit: al editar la hoja Gastos se recalculan los resúmenes solos; al cambiar
//   las hojas Centros o Cuentas se actualizan las listas desplegables.
// - Listas desplegables en Gastos: Centro de costo (de la hoja Centros) y Cuenta
//   (solo las cuentas del centro de esa fila).
// - revisarGastos: registra las filas con centro, cuenta, fecha o monto inválidos.
// ---------------------------------------------------------------------------
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    var hoja = e.range.getSheet();
    var nombre = hoja.getName();
    var ss = e.source || libro_();
    if (nombre === HOJA_CENTROS || nombre === HOJA_CUENTAS) {
      aplicarListasGastos_(ss);
      return;
    }
    if (nombre !== HOJA_GASTOS) return;
    var desde = Math.max(e.range.getRow(), 2);
    var hasta = e.range.getLastRow();
    if (hasta < 2) return;
    var mapa = mapaGastos_(hoja);
    var colCentro = mapa[2] + 1, colCuenta = mapa[7] + 1;
    if (e.range.getColumn() <= colCentro && e.range.getLastColumn() >= colCentro) {
      // Cambió el centro: la cuenta anterior se borra si no es de este centro.
      var cuentas = leerCuentas_(ss);
      var centros = hoja.getRange(desde, colCentro, hasta - desde + 1, 1).getValues();
      var celdasCuenta = hoja.getRange(desde, colCuenta, hasta - desde + 1, 1);
      var actuales = celdasCuenta.getValues();
      var borradas = 0;
      var nuevas = actuales.map(function (f, i) {
        var cu = String(f[0]).trim();
        if (cu && (cuentas[String(centros[i][0]).trim()] || []).indexOf(cu) === -1) { borradas++; return ['']; }
        return [f[0]];
      });
      aplicarListaCuentas_(hoja, desde, hasta - desde + 1, ss);
      if (borradas) {
        celdasCuenta.setValues(nuevas);
        ss.toast('Se borró la cuenta de ' + borradas + ' fila(s) porque no era de ese centro. Elígela de nuevo.', 'Gastos', 6);
      }
    }
    actualizarResumenes_(ss);
  } catch (err) {
    // Editar a mano nunca debe fallar por esto.
  }
}

function aplicarListasGastos_(ss) {
  ss = ss || libro_();
  var hoja = ss.getSheetByName(HOJA_GASTOS);
  var hojaCentros = ss.getSheetByName(HOJA_CENTROS);
  if (!hoja || !hojaCentros) return;
  asegurarColumnaCuenta_(hoja);
  var mapa = mapaGastos_(hoja);
  var n = hoja.getMaxRows() - 1;
  if (n < 1) return;
  var reglaCentro = SpreadsheetApp.newDataValidation()
    .requireValueInRange(hojaCentros.getRange('A2:A'), true)
    .setAllowInvalid(false)
    .setHelpText('Elige un centro de la hoja Centros (para uno nuevo, agrégalo ahí o desde la app).')
    .build();
  hoja.getRange(2, mapa[2] + 1, n, 1).setDataValidation(reglaCentro);
  aplicarListaCuentas_(hoja, 2, n, ss);
}

// Lista de cuentas de cada fila según su centro (filas sin centro: sin lista).
function aplicarListaCuentas_(hoja, desde, n, ss) {
  if (n < 1) return;
  ss = ss || libro_();
  var mapa = mapaGastos_(hoja);
  var cuentas = leerCuentas_(ss);
  var reglas = {};
  var valores = hoja.getRange(desde, mapa[2] + 1, n, 1).getValues().map(function (f) {
    var centro = String(f[0]).trim();
    if (!centro || !(cuentas[centro] || []).length) return [null];
    if (!reglas[centro]) {
      reglas[centro] = SpreadsheetApp.newDataValidation()
        .requireValueInList(alfabetico_(cuentas[centro]), true)
        .setAllowInvalid(false)
        .setHelpText('Cuentas de ' + centro + ' (o déjala vacía).')
        .build();
    }
    return [reglas[centro]];
  });
  hoja.getRange(desde, mapa[7] + 1, n, 1).setDataValidations(valores);
}

var LISTAS_VERSION = '2026-11-03';
function listasGastosSiFalta_() {
  if (propiedad_('LISTAS_VERSION') === LISTAS_VERSION) return;
  aplicarListasGastos_();
  PropertiesService.getScriptProperties().setProperty('LISTAS_VERSION', LISTAS_VERSION);
}

// Ejecutar desde el editor: muestra en el registro las filas que no cuadran.
function revisarGastos() {
  var problemas = problemasGastos_().map(function (p) {
    return 'Fila ' + p.fila + ': ' + p.motivos.join('; ') + (p.descripcion ? '  (' + p.descripcion + ')' : '');
  });
  Logger.log(problemas.length ? problemas.length + ' fila(s) para revisar:\n' + problemas.join('\n') : 'Todo bien: no hay filas con problemas.');
}

// Lo mismo, como datos para el botón "Revisar planilla" de la app.
function problemasGastos_() {
  var ss = libro_();
  var hoja = hojaGastos_();
  var centros = leerLista_(HOJA_CENTROS, ss);
  var cuentas = leerCuentas_(ss);
  var problemas = [];
  leerGastos_(hoja).forEach(function (f, i) {
    if (!f[0] && !f[3] && !f[2]) return;
    var fila = i + 2, motivos = [];
    var centro = String(f[2]).trim(), cuenta = String(f[7] || '').trim();
    if (centros.indexOf(centro) === -1) motivos.push('centro "' + centro + '" no está en la hoja Centros');
    else if (cuenta && (cuentas[centro] || []).indexOf(cuenta) === -1) motivos.push('cuenta "' + cuenta + '" no es de ' + centro);
    if (!(f[1] instanceof Date)) motivos.push('fecha "' + f[1] + '" no es una fecha');
    if (typeof f[3] !== 'number' || !(f[3] > 0)) motivos.push('monto "' + f[3] + '" no es un número');
    if (motivos.length) {
      problemas.push({
        fila: fila,
        motivos: motivos,
        descripcion: String(f[4] || ''),
        centro: centro,
        cuenta: cuenta,
        monto: typeof f[3] === 'number' ? f[3] : 0,
        fecha: f[1] instanceof Date ? Utilities.formatDate(f[1], ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd') : '',
        persona: String(f[5] || '')
      });
    }
  });
  return problemas;
}

// ---------------------------------------------------------------------------
// API web. Todas las acciones van por POST (text/plain para evitar CORS).
// ---------------------------------------------------------------------------

function doGet() {
  return json_({ ok: true, mensaje: 'Gastos Familia: backend funcionando.', version: VERSION });
}

function doPost(e) {
  try {
    var body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (body.action === 'modo') return json_(modoAcceso_());
    if (body.action === 'loginGoogle') return json_(loginGoogle_(body));
    autorizar_(body);

    switch (body.action) {
      case 'config':
        return json_(config_());
      case 'list':
        return json_({ ok: true, gastos: ultimosGastos_(), totalMes: totalMesActual_() });
      case 'add':
        agregarGasto_(body);
        actualizarResumenes_(libro_());
        return json_({ ok: true, gastos: ultimosGastos_(), totalMes: totalMesActual_() });
      case 'addCentro':
        agregarCentro_(body.nombre);
        return json_(config_());
      case 'addCuenta':
        agregarCuenta_(body.centro, body.nombre);
        return json_(config_());
      case 'delete':
        borrarGasto_(body.id);
        actualizarResumenes_(libro_());
        return json_({ ok: true, gastos: ultimosGastos_(), totalMes: totalMesActual_() });
      case 'revisar':
        return json_({ ok: true, problemas: problemasGastos_(), url: libro_().getUrl() });
      case 'resumen':
        return json_({ ok: true, filas: totalesPorMes_() });
      case 'detalle':
        return json_({ ok: true, gastos: gastosDeCentro_(body.centro, body.periodo) });

      // Rendiciones (pestaña de Vicente)
      case 'rendiciones':
        return json_(Object.assign({ ok: true }, listarRendiciones_()));
      case 'rendicionAdd':
        var correo = agregarRendicion_(body);
        return json_(Object.assign({ ok: true, correo: correo }, listarRendiciones_()));
      case 'rendicionEstado':
        cambiarEstadoRendicion_(body.id, body.estado);
        return json_(Object.assign({ ok: true }, listarRendiciones_()));
      case 'rendicionDelete':
        borrarRendicion_(body.id);
        return json_(Object.assign({ ok: true }, listarRendiciones_()));

      // Salud (pestaña de Fran): días sin azúcar y proteína diaria
      case 'salud':
        return json_(Object.assign({ ok: true }, datosSalud_()));
      case 'azucarEvento':
        agregarEventoAzucar_(body.evento, body.fecha);
        return json_(Object.assign({ ok: true }, datosSalud_()));
      case 'azucarDeshacer':
        deshacerEventoAzucar_();
        return json_(Object.assign({ ok: true }, datosSalud_()));

      // Proteína diaria: el mismo registro para Fran (Salud) y Vicente (Deporte), por persona.
      case 'proteina':
        return json_(Object.assign({ ok: true }, datosProteina_(body.persona)));
      case 'proteinaGuardar':
        guardarProteina_(body);
        return json_(Object.assign({ ok: true }, datosProteina_(body.persona)));
      case 'proteinaBorrar':
        borrarProteinaDia_(body.persona, body.fecha);
        return json_(Object.assign({ ok: true }, datosProteina_(body.persona)));
      case 'proteinaMeta':
        cambiarMetaProteina_(body.persona, body.meta);
        return json_(Object.assign({ ok: true }, datosProteina_(body.persona)));

      // Deporte (pestaña de Vicente): entrenamientos + diario de cigarros, comida y peso
      case 'deporte':
        return json_(Object.assign({ ok: true }, datosDeporte_()));
      case 'entrenoAdd':
        agregarEntreno_(body);
        return json_(Object.assign({ ok: true }, datosDeporte_()));
      case 'entrenoDelete':
        borrarFilaPorId_(hoja_(HOJA_ENTRENOS, COLUMNAS_ENTRENOS), body.id);
        return json_(Object.assign({ ok: true }, datosDeporte_()));
      case 'diarioSumar':
        sumarDiario_(body);
        return json_(Object.assign({ ok: true }, datosDeporte_()));
      case 'diarioGuardar':
        guardarDiario_(body);
        return json_(Object.assign({ ok: true }, datosDeporte_()));
      case 'deporteMeta':
        cambiarMetaDeporte_(body.clave, body.valor);
        return json_(Object.assign({ ok: true }, datosDeporte_()));
      case 'deporteOcultar':
        ocultarEjercicio_(body.id, body.ocultar);
        return json_(Object.assign({ ok: true }, datosDeporte_()));
      default:
        throw new Error('Acción desconocida');
    }
  } catch (err) {
    return json_({ ok: false, error: err.message });
  }
}

// ---------------------------------------------------------------------------
// Acceso
// - Con las propiedades GOOGLE_CLIENT_ID y CORREOS_PERMITIDOS: solo entran esos
//   correos, con su cuenta de Google. El PIN deja de servir.
// - Sin ellas: PIN (como antes).
// En ambos casos, 5 intentos fallidos seguidos bloquean el ingreso por 15 minutos.
// Tras entrar con Google el equipo recibe una sesión (válida 1 año) que se puede
// revocar con cerrarTodasLasSesiones.
// ---------------------------------------------------------------------------
var MAX_FALLOS = 5;
var BLOQUEO_SEGUNDOS = 15 * 60;
var DURACION_SESION_DIAS = 365;

function propiedad_(k) { return PropertiesService.getScriptProperties().getProperty(k) || ''; }

function correosPermitidos_() {
  return propiedad_('CORREOS_PERMITIDOS').split(/[\s,;]+/)
    .map(function (c) { return c.trim().toLowerCase(); })
    .filter(function (c) { return c; });
}

function modoGoogle_() { return !!(propiedad_('GOOGLE_CLIENT_ID') && correosPermitidos_().length); }

function modoAcceso_() {
  return { ok: true, google: modoGoogle_(), clientId: modoGoogle_() ? propiedad_('GOOGLE_CLIENT_ID') : '' };
}

function revisarBloqueo_() {
  var n = Number(CacheService.getScriptCache().get('fallos') || 0);
  if (n >= MAX_FALLOS) throw new Error('Demasiados intentos fallidos. Espera 15 minutos e intenta de nuevo.');
}

function sumarFallo_() {
  var cache = CacheService.getScriptCache();
  cache.put('fallos', String(Number(cache.get('fallos') || 0) + 1), BLOQUEO_SEGUNDOS);
}

function claveSesion_(token) {
  var hash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(token));
  return 'SES_' + Utilities.base64EncodeWebSafe(hash).replace(/=+$/, '');
}

function autorizar_(body) {
  if (modoGoogle_()) {
    var token = String(body.sesion || '');
    var guardada = token ? propiedad_(claveSesion_(token)) : '';
    var s = null;
    try { s = guardada ? JSON.parse(guardada) : null; } catch (e) {}
    var vencida = s && (Date.now() - s.creado) > DURACION_SESION_DIAS * 864e5;
    // Si sacaste un correo de la lista, sus sesiones dejan de servir al tiro.
    if (!s || vencida || correosPermitidos_().indexOf(s.email) === -1) {
      if (s) PropertiesService.getScriptProperties().deleteProperty(claveSesion_(token));
      throw new Error('Sesión no válida: entra con tu cuenta de Google');
    }
    return s.email;
  }
  revisarBloqueo_();
  var esperado = propiedad_('PIN');
  if (!esperado) throw new Error('El PIN no está configurado en el script');
  if (String(body.pin || '') !== String(esperado)) {
    sumarFallo_();
    throw new Error('PIN incorrecto');
  }
  return '';
}

// Recibe el id_token que Google entregó a la página y lo verifica con Google.
function loginGoogle_(body) {
  if (!modoGoogle_()) throw new Error('El ingreso con Google no está activado');
  revisarBloqueo_();
  var idToken = String(body.idToken || '');
  var nonce = String(body.nonce || '');
  if (!idToken || !nonce) throw new Error('Faltan datos de Google');
  var r = UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(idToken), { muteHttpExceptions: true });
  if (r.getResponseCode() !== 200) throw new Error('Google no validó el ingreso. Intenta de nuevo.');
  var t = JSON.parse(r.getContentText());
  var valido = t.aud === propiedad_('GOOGLE_CLIENT_ID') &&
    (t.iss === 'accounts.google.com' || t.iss === 'https://accounts.google.com') &&
    String(t.email_verified) === 'true' &&
    Number(t.exp) * 1000 > Date.now() &&
    t.nonce === nonce;
  if (!valido) throw new Error('Google no validó el ingreso. Intenta de nuevo.');
  var email = String(t.email || '').toLowerCase();
  if (correosPermitidos_().indexOf(email) === -1) {
    sumarFallo_();
    throw new Error('La cuenta ' + email + ' no tiene acceso a esta app.');
  }
  var token = Utilities.getUuid().replace(/-/g, '');
  PropertiesService.getScriptProperties().setProperty(claveSesion_(token), JSON.stringify({ email: email, creado: Date.now() }));
  return { ok: true, sesion: token, email: email };
}

// Ejecutar UNA vez desde el editor después de poner GOOGLE_CLIENT_ID y
// CORREOS_PERMITIDOS: pide el permiso para hablar con Google y revisa la configuración.
function activarAccesoGoogle() {
  UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?id_token=x', { muteHttpExceptions: true });
  var lineas = [];
  lineas.push(propiedad_('GOOGLE_CLIENT_ID') ? '✓ GOOGLE_CLIENT_ID configurado.' : '✗ Falta la propiedad GOOGLE_CLIENT_ID.');
  var correos = correosPermitidos_();
  lineas.push(correos.length ? '✓ Correos con acceso: ' + correos.join(', ') : '✗ Falta la propiedad CORREOS_PERMITIDOS.');
  lineas.push(modoGoogle_() ? 'Listo: desde ahora solo se entra con esas cuentas de Google (publica una nueva versión).' : 'Todavía se entra con PIN.');
  Logger.log(lineas.join('\n'));
}

// Cierra la sesión en todos los equipos (todos deberán volver a entrar con Google).
function cerrarTodasLasSesiones() {
  var props = PropertiesService.getScriptProperties();
  var n = 0;
  Object.keys(props.getProperties()).forEach(function (k) {
    if (k.indexOf('SES_') === 0) { props.deleteProperty(k); n++; }
  });
  Logger.log('Sesiones cerradas: ' + n);
}

// Muestra qué cuentas tienen sesiones abiertas y desde cuándo.
function verSesiones() {
  var props = PropertiesService.getScriptProperties().getProperties();
  var lineas = Object.keys(props).filter(function (k) { return k.indexOf('SES_') === 0; }).map(function (k) {
    var s = JSON.parse(props[k]);
    return s.email + ' · desde ' + new Date(s.creado).toLocaleString();
  });
  Logger.log(lineas.length ? lineas.join('\n') : 'No hay sesiones abiertas.');
}

function leerLista_(nombreHoja, ss) {
  var hoja = (ss || libro_()).getSheetByName(nombreHoja);
  if (!hoja || hoja.getLastRow() < 2) return [];
  return hoja.getRange(2, 1, hoja.getLastRow() - 1, 1).getValues()
    .map(function (fila) { return String(fila[0]).trim(); })
    .filter(function (v) { return v; });
}

function agregarCentro_(nombre) {
  nombre = String(nombre || '').replace(/\s+/g, ' ').trim().slice(0, 40);
  if (!nombre) throw new Error('Escribe un nombre para el centro de costo');
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var existentes = leerLista_(HOJA_CENTROS);
    var repetido = existentes.some(function (c) { return c.toLowerCase() === nombre.toLowerCase(); });
    if (!repetido) libro_().getSheetByName(HOJA_CENTROS).appendRow([nombre]);
    ordenarHoja_(HOJA_CENTROS, 1);
  } finally {
    lock.releaseLock();
  }
  return leerLista_(HOJA_CENTROS);
}

function config_() {
  ordenarHojaGastosSiFalta_();
  reordenarSiFalta_();
  agregarCuentasFaltantes_();
  moverGastosAntiguosSiFalta_();
  listasGastosSiFalta_();
  var cuentas = leerCuentas_();
  Object.keys(cuentas).forEach(function (c) { cuentas[c] = alfabetico_(cuentas[c]); });
  return { ok: true, centros: alfabetico_(leerLista_(HOJA_CENTROS)), personas: leerLista_(HOJA_PERSONAS), cuentas: cuentas };
}

// Centros y cuentas siempre en orden alfabético (sin importar mayúsculas ni tildes).
function alfabetico_(lista) {
  return lista.slice().sort(function (a, b) { return String(a).localeCompare(String(b), 'es', { sensitivity: 'base' }); });
}

// Deja la hoja ordenada A→Z (Centros por nombre; Cuentas por centro y luego cuenta).
function ordenarHoja_(nombreHoja, columnas) {
  var hoja = libro_().getSheetByName(nombreHoja);
  if (!hoja || hoja.getLastRow() < 3) return;
  var orden = [];
  for (var i = 1; i <= columnas; i++) orden.push({ column: i, ascending: true });
  hoja.getRange(2, 1, hoja.getLastRow() - 1, columnas).sort(orden);
}

// La primera vez que la app abre con este código, ordena centros y cuentas sola
// (no hace falta ejecutar reordenarCentrosYCuentas a mano). Solo corre una vez.
// Cuentas nuevas de la lista oficial que aún no están en la hoja Cuentas: se agregan
// solas (sin borrar las que ya hay ni tocar gastos). Cambiar CUENTAS_VERSION al
// agregar cuentas a CUENTAS_INICIALES para que corra de nuevo.
var CUENTAS_VERSION = '2026-11-03 casa';
function agregarCuentasFaltantes_() {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('CUENTAS_VERSION') === CUENTAS_VERSION) return;
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var ss = libro_();
    var hoja = obtenerOCrearHoja_(ss, HOJA_CUENTAS);
    if (hoja.getLastRow() === 0) hoja.appendRow(['Centro de costo', 'Cuenta']);
    var centros = leerLista_(HOJA_CENTROS, ss);
    var actuales = leerCuentas_(ss);
    var nuevas = [];
    Object.keys(CUENTAS_INICIALES).forEach(function (k) {
      var centro = centros.filter(function (c) { return c.toLowerCase() === k.toLowerCase(); })[0];
      if (!centro) return;
      var ya = (actuales[centro] || []).map(function (c) { return c.toLowerCase(); });
      CUENTAS_INICIALES[k].forEach(function (c) {
        if (ya.indexOf(c.toLowerCase()) === -1) nuevas.push([centro, c]);
      });
    });
    if (nuevas.length) {
      hoja.getRange(hoja.getLastRow() + 1, 1, nuevas.length, 2).setValues(nuevas);
      ordenarHoja_(HOJA_CUENTAS, 2);
    }
    props.setProperty('CUENTAS_VERSION', CUENTAS_VERSION);
  } finally {
    lock.releaseLock();
  }
}

// ---------------------------------------------------------------------------
// Gastos antiguos → cuentas nuevas (corre sola una vez por GASTOS_VERSION, o a mano
// con moverGastosAntiguos). Cambia la MISMA fila (nunca copia), así no se suma doble:
// 1. Centro antiguo con nombre de cuenta o sinónimo (Gas, Agua, Internet, Entel…)
//    → su centro nuevo con esa cuenta (ej. "Internet" → Casa / Conectividad).
// 2. Gastos de Casa sin cuenta cuya descripción la deja clara (ej. "Metrogas",
//    "plan Entel") → esa cuenta. Si calza con más de una cuenta, no se toca.
// Antes guarda un respaldo oculto de la hoja Gastos.
// ---------------------------------------------------------------------------
var GASTOS_VERSION = '2026-11-03 casa';
var SINONIMOS_CENTRO = {
  'internet': ['Casa', 'Conectividad'], 'entel': ['Casa', 'Conectividad'], 'plan entel': ['Casa', 'Conectividad'],
  'wifi': ['Casa', 'Conectividad'], 'telefono': ['Casa', 'Conectividad'], 'celular': ['Casa', 'Conectividad'],
  'fumigacion': ['Casa', 'Fumigación'], 'tintoreria': ['Casa', 'Tintorería'], 'lavanderia': ['Casa', 'Tintorería'],
  'gas': ['Casa', 'Gas'], 'agua': ['Casa', 'Agua']
};
var PALABRAS_CUENTA = {
  'Casa': {
    'Gas': /\b(gas|metrogas|lipigas|abastible|gasco|balon|balones)\b/,
    'Agua': /\b(agua|aguas andinas|essbio|esval|smapa)\b/,
    'Conectividad': /\b(internet|entel|wifi|wi-fi|fibra|plan movil|plan celular|vtr|movistar|wom)\b/,
    'Fumigación': /\b(fumigacion|fumigar|fumigador|plagas|desratizacion)\b/,
    'Tintorería': /\b(tintoreria|lavanderia|lavaseco)\b/
  }
};

function sinTildes_(t) {
  return String(t || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/\s+/g, ' ').trim();
}

function moverGastosAntiguosSiFalta_() {
  if (PropertiesService.getScriptProperties().getProperty('GASTOS_VERSION') === GASTOS_VERSION) return;
  moverGastosAntiguos();
}

function moverGastosAntiguos() {
  var ss = libro_();
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var props = PropertiesService.getScriptProperties();
    var gastos = hojaGastos_();
    asegurarColumnaCuenta_(gastos);
    var n = gastos.getLastRow() - 1;
    var cambios = {};
    if (n > 0) {
      var centros = leerLista_(HOJA_CENTROS, ss);
      var oficial = {};
      centros.forEach(function (c) { oficial[sinTildes_(c)] = c; });
      var cuentas = leerCuentas_(ss);
      // Nombre de cuenta (sin tildes) → [centro, cuenta], más los sinónimos.
      var destinoDe = {};
      Object.keys(cuentas).forEach(function (c) {
        cuentas[c].forEach(function (cu) { destinoDe[sinTildes_(cu)] = [c, cu]; });
      });
      Object.keys(SINONIMOS_CENTRO).forEach(function (k) {
        var d = SINONIMOS_CENTRO[k];
        if (oficial[sinTildes_(d[0])]) destinoDe[k] = [oficial[sinTildes_(d[0])], d[1]];
      });

      var filas = leerGastos_(gastos);
      filas.forEach(function (f) {
        if (!f[0] && !f[3]) return;
        var centro = String(f[2]).trim(), cuenta = String(f[7] || '').trim();
        var antes = centro + (cuenta ? ' / ' + cuenta : '');
        var clave = sinTildes_(centro);
        if (!oficial[clave] && destinoDe[clave]) {
          centro = destinoDe[clave][0];
          cuenta = cuenta || destinoDe[clave][1];
        }
        var reglas = PALABRAS_CUENTA[centro];
        if (reglas && !cuenta) {
          var desc = sinTildes_(f[4]);
          var calzan = Object.keys(reglas).filter(function (cu) {
            return reglas[cu].test(desc) && (cuentas[centro] || []).indexOf(cu) !== -1;
          });
          if (calzan.length === 1) cuenta = calzan[0];
        }
        var despues = centro + (cuenta ? ' / ' + cuenta : '');
        if (despues !== antes) {
          f[2] = centro;
          f[7] = cuenta;
          var k = antes + (f[4] ? ' ("' + String(f[4]).slice(0, 30) + '")' : '') + '  →  ' + despues;
          cambios[k] = (cambios[k] || 0) + 1;
        }
      });
      var claves = Object.keys(cambios);
      if (claves.length) {
        var tz = ss.getSpreadsheetTimeZone();
        gastos.copyTo(ss).setName('Gastos respaldo ' + Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH.mm.ss')).hideSheet();
        guardarCentroYCuenta_(gastos, filas);
        actualizarResumenes_(ss);
      }
    }
    var lineas = Object.keys(cambios);
    Logger.log(lineas.length ? 'Gastos movidos:\n' + lineas.map(function (k) { return '  ' + k + '  (' + cambios[k] + ')'; }).join('\n') : 'No había gastos que mover.');
    props.setProperty('GASTOS_VERSION', GASTOS_VERSION);
  } finally {
    lock.releaseLock();
  }
}

var ORDEN_CENTROS_VERSION = '2026-11-03';
function reordenarSiFalta_() {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('ORDEN_CENTROS') === ORDEN_CENTROS_VERSION) return;
  reordenarCentrosYCuentas();
}

// { 'Vicky': ['Sueldo', 'Contribuciones', 'Traslado'], ... } con los centros tal como están en la hoja Centros.
function leerCuentas_(ss) {
  ss = ss || libro_();
  var hoja = ss.getSheetByName(HOJA_CUENTAS);
  var res = {};
  if (!hoja || hoja.getLastRow() < 2) return res;
  var centros = leerLista_(HOJA_CENTROS, ss);
  hoja.getRange(2, 1, hoja.getLastRow() - 1, 2).getValues().forEach(function (f) {
    var escrito = String(f[0]).trim().toLowerCase();
    var cuenta = String(f[1]).replace(/\s+/g, ' ').trim();
    var centro = centros.filter(function (c) { return c.toLowerCase() === escrito; })[0];
    if (!centro || !cuenta) return;
    res[centro] = res[centro] || [];
    var repetida = res[centro].some(function (c) { return c.toLowerCase() === cuenta.toLowerCase(); });
    if (!repetida) res[centro].push(cuenta);
  });
  return res;
}

function agregarCuenta_(centro, nombre) {
  centro = String(centro || '').trim();
  nombre = String(nombre || '').replace(/\s+/g, ' ').trim().slice(0, 40);
  if (leerLista_(HOJA_CENTROS).indexOf(centro) === -1) throw new Error('Centro de costo inválido');
  if (!nombre) throw new Error('Escribe un nombre para la cuenta');
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var ss = libro_();
    var hoja = ss.getSheetByName(HOJA_CUENTAS);
    if (!hoja) {
      hoja = ss.insertSheet(HOJA_CUENTAS);
      hoja.appendRow(['Centro de costo', 'Cuenta']);
      hoja.getRange(1, 1, 1, 2).setFontWeight('bold');
      hoja.setFrozenRows(1);
    }
    var actuales = leerCuentas_(ss)[centro] || [];
    var repetida = actuales.some(function (c) { return c.toLowerCase() === nombre.toLowerCase(); });
    if (!repetida) hoja.appendRow([centro, nombre]);
    ordenarHoja_(HOJA_CUENTAS, 2);
  } finally {
    lock.releaseLock();
  }
  aplicarListasGastos_();
}

function agregarGasto_(body) {
  var centro = String(body.centro || '').trim();
  var persona = String(body.persona || '').trim();
  var descripcion = String(body.descripcion || '').trim().slice(0, 200);
  var monto = Math.round(Number(body.monto));
  var fecha = parsearFecha_(body.fecha);
  var cuenta = String(body.cuenta || '').trim();

  if (leerLista_(HOJA_CENTROS).indexOf(centro) === -1) throw new Error('Centro de costo inválido');
  if (cuenta && (leerCuentas_()[centro] || []).indexOf(cuenta) === -1) throw new Error('Cuenta inválida para ' + centro);
  if (leerLista_(HOJA_PERSONAS).indexOf(persona) === -1) throw new Error('Persona inválida');
  if (!(monto > 0)) throw new Error('Monto inválido');

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var hoja = hojaGastos_();
    asegurarColumnaCuenta_(hoja);
    hoja.appendRow(filaFisicaGastos_(hoja, [Utilities.getUuid(), fecha, centro, monto, descripcion, persona, new Date(), cuenta]));
    aplicarListaCuentas_(hoja, hoja.getLastRow(), 1);
  } finally {
    lock.releaseLock();
  }
}

function borrarGasto_(id) {
  if (!id) throw new Error('Falta el ID');
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var hoja = hojaGastos_();
    var ultima = hoja.getLastRow();
    if (ultima < 2) throw new Error('Gasto no encontrado');
    var ids = hoja.getRange(2, 1, ultima - 1, 1).getValues();
    for (var i = ids.length - 1; i >= 0; i--) {
      if (ids[i][0] === id) {
        hoja.deleteRow(i + 2);
        return;
      }
    }
    throw new Error('Gasto no encontrado');
  } finally {
    lock.releaseLock();
  }
}

function ultimosGastos_() {
  var hoja = hojaGastos_();
  var ultima = hoja.getLastRow();
  if (ultima < 2) return [];
  var n = Math.min(ULTIMOS_GASTOS, ultima - 1);
  var tz = libro_().getSpreadsheetTimeZone();
  return leerGastos_(hoja, ultima - n + 1, n)
    .reverse()
    .map(function (f) {
      return {
        id: f[0],
        fecha: f[1] instanceof Date ? Utilities.formatDate(f[1], tz, 'yyyy-MM-dd') : String(f[1]),
        centro: f[2],
        monto: Number(f[3]) || 0,
        descripcion: f[4],
        persona: f[5],
        cuenta: String(f[7] || '').trim()
      };
    });
}

function totalMesActual_() {
  var hoja = hojaGastos_();
  var ultima = hoja.getLastRow();
  if (ultima < 2) return 0;
  var tz = libro_().getSpreadsheetTimeZone();
  var mesActual = Utilities.formatDate(new Date(), tz, 'yyyy-MM');
  return leerGastos_(hoja).reduce(function (suma, f) {
    if (f[1] instanceof Date && Utilities.formatDate(f[1], tz, 'yyyy-MM') === mesActual) {
      return suma + (Number(f[3]) || 0);
    }
    return suma;
  }, 0);
}

// Gastos de un centro en un mes ('yyyy-MM') o año ('yyyy').
function gastosDeCentro_(centro, periodo) {
  centro = String(centro || '').trim();
  periodo = String(periodo || '');
  if (!/^\d{4}(-\d{2})?$/.test(periodo)) throw new Error('Periodo inválido');
  var hoja = hojaGastos_();
  if (hoja.getLastRow() < 2) return [];
  var tz = libro_().getSpreadsheetTimeZone();
  return leerGastos_(hoja)
    .filter(function (f) {
      return f[1] instanceof Date && String(f[2]).trim() === centro &&
        Utilities.formatDate(f[1], tz, 'yyyy-MM').slice(0, periodo.length) === periodo;
    })
    .map(function (f) {
      return {
        id: f[0],
        fecha: Utilities.formatDate(f[1], tz, 'yyyy-MM-dd'),
        monto: Number(f[3]) || 0,
        descripcion: f[4],
        persona: f[5],
        cuenta: String(f[7] || '').trim()
      };
    });
}

function hojaGastos_() {
  var hoja = libro_().getSheetByName(HOJA_GASTOS);
  if (!hoja) throw new Error('Falta la hoja Gastos: ejecuta setup() primero');
  return hoja;
}

function parsearFecha_(texto) {
  var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(texto || ''));
  if (!m) return new Date();
  // Mediodía para que ningún cambio de zona horaria mueva el día.
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), 12, 0, 0);
}

// ---------------------------------------------------------------------------
// Rendiciones: gasto + foto del comprobante. La foto queda en Drive, la fila en
// la hoja Rendiciones y se manda un correo con la foto adjunta a CORREO_RENDICION.
// ---------------------------------------------------------------------------

function agregarRendicion_(body) {
  var monto = Math.round(Number(body.monto));
  var descripcion = String(body.descripcion || '').trim().slice(0, 200);
  var persona = String(body.persona || '').trim();
  var fecha = parsearFecha_(body.fecha);
  if (!(monto > 0)) throw new Error('Monto inválido');
  if (!descripcion) throw new Error('Escribe una descripción');
  if (leerLista_(HOJA_PERSONAS).indexOf(persona) === -1) throw new Error('Persona inválida');
  if (!body.foto) throw new Error('Falta la foto del comprobante');

  var tz = libro_().getSpreadsheetTimeZone();
  var tipo = /^(image\/(jpeg|png|webp|heic|heif)|application\/pdf)$/.test(body.fotoTipo) ? body.fotoTipo : 'image/jpeg';
  var bytes = Utilities.base64Decode(String(body.foto));
  if (bytes.length > 15 * 1024 * 1024) throw new Error('El archivo es muy grande (máx. 15 MB)');
  var ext = tipo === 'application/pdf' ? 'pdf' : tipo === 'image/jpeg' ? 'jpg' : tipo.split('/')[1];
  var nombre = (Utilities.formatDate(fecha, tz, 'yyyy-MM-dd') + ' ' + pesos_(monto) + ' ' + descripcion.slice(0, 60))
    .replace(/[\\/:*?"<>|]/g, '-') + '.' + ext;
  var blob = Utilities.newBlob(bytes, tipo, nombre);
  var archivo = carpetaRendiciones_().createFile(blob);

  var r = { fecha: fecha, monto: monto, descripcion: descripcion, persona: persona, url: archivo.getUrl() };
  var correo = enviarCorreoRendicion_(r, blob, tz);

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    hoja_(HOJA_RENDICIONES, COLUMNAS_RENDICIONES).appendRow([
      Utilities.getUuid(), fecha, monto, descripcion, r.url, 'Pendiente', correo.texto, persona, new Date()
    ]);
  } finally {
    lock.releaseLock();
  }
  return correo;
}

function enviarCorreoRendicion_(r, blob, tz) {
  var para = String(PropertiesService.getScriptProperties().getProperty('CORREO_RENDICION') || '').trim();
  if (!para) return { enviado: false, texto: 'No enviado: falta CORREO_RENDICION en Propiedades del script' };
  var fechaTxt = Utilities.formatDate(r.fecha, tz, 'dd/MM/yyyy');
  var asunto = 'Rendición ' + fechaTxt + ' · ' + pesos_(r.monto) + ' · ' + r.descripcion;
  var texto = 'Rendición de gasto\n\n' +
    'Fecha: ' + fechaTxt + '\nMonto: ' + pesos_(r.monto) + '\nDescripción: ' + r.descripcion +
    '\nRegistró: ' + r.persona + '\n\nComprobante adjunto y en Drive: ' + r.url;
  var html = '<p><b>Rendición de gasto</b></p><table cellpadding="4">' +
    '<tr><td>Fecha</td><td><b>' + fechaTxt + '</b></td></tr>' +
    '<tr><td>Monto</td><td><b>' + pesos_(r.monto) + '</b></td></tr>' +
    '<tr><td>Descripción</td><td>' + html_(r.descripcion) + '</td></tr>' +
    '<tr><td>Registró</td><td>' + html_(r.persona) + '</td></tr></table>' +
    '<p>Comprobante adjunto. También está en <a href="' + r.url + '">Google Drive</a>.</p>';
  try {
    MailApp.sendEmail({ to: para, subject: asunto, body: texto, htmlBody: html, attachments: [blob] });
    return { enviado: true, para: para, texto: 'Enviado a ' + para + ' el ' + Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm') };
  } catch (err) {
    return { enviado: false, texto: 'No enviado: ' + err.message };
  }
}

function carpetaRendiciones_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('CARPETA_RENDICIONES_ID');
  if (id) {
    try {
      var carpeta = DriveApp.getFolderById(id);
      if (!carpeta.isTrashed()) return carpeta;
    } catch (e) {}
  }
  var nueva = DriveApp.createFolder(CARPETA_RENDICIONES);
  props.setProperty('CARPETA_RENDICIONES_ID', nueva.getId());
  return nueva;
}

function listarRendiciones_() {
  var hoja = hoja_(HOJA_RENDICIONES, COLUMNAS_RENDICIONES);
  if (hoja.getLastRow() < 2) return { rendiciones: [], pendiente: 0 };
  var tz = libro_().getSpreadsheetTimeZone();
  var todas = hoja.getRange(2, 1, hoja.getLastRow() - 1, COLUMNAS_RENDICIONES.length).getValues()
    .filter(function (f) { return f[0]; })
    .map(function (f) {
      return {
        id: f[0],
        fecha: f[1] instanceof Date ? Utilities.formatDate(f[1], tz, 'yyyy-MM-dd') : String(f[1]),
        monto: Number(f[2]) || 0,
        descripcion: f[3],
        foto: f[4],
        estado: f[5] || 'Pendiente',
        correo: f[6],
        persona: f[7]
      };
    });
  var pendiente = todas.reduce(function (s, r) { return r.estado === 'Reembolsada' ? s : s + r.monto; }, 0);
  todas.sort(function (a, b) { return a.fecha < b.fecha ? 1 : a.fecha > b.fecha ? -1 : 0; });
  return { rendiciones: todas.slice(0, ULTIMAS_RENDICIONES), pendiente: pendiente };
}

function cambiarEstadoRendicion_(id, estado) {
  if (['Pendiente', 'Reembolsada'].indexOf(estado) === -1) throw new Error('Estado inválido');
  var hoja = hoja_(HOJA_RENDICIONES, COLUMNAS_RENDICIONES);
  hoja.getRange(filaDeId_(hoja, id), 6).setValue(estado);
}

function borrarRendicion_(id) {
  var hoja = hoja_(HOJA_RENDICIONES, COLUMNAS_RENDICIONES);
  var fila = filaDeId_(hoja, id);
  var url = String(hoja.getRange(fila, 5).getValue());
  var m = /[-\w]{25,}/.exec(url);
  if (m) {
    try { DriveApp.getFileById(m[0]).setTrashed(true); } catch (e) {}
  }
  borrarFilaPorId_(hoja, id);
}

// ---------------------------------------------------------------------------
// Salud: contador de días sin azúcar.
// ---------------------------------------------------------------------------

function datosSalud_() {
  var tz = libro_().getSpreadsheetTimeZone();
  var hojaA = hoja_(HOJA_AZUCAR, COLUMNAS_AZUCAR);
  var eventos = hojaA.getLastRow() < 2 ? [] : hojaA.getRange(2, 1, hojaA.getLastRow() - 1, 2).getValues()
    .filter(function (f) { return f[0] instanceof Date; })
    .map(function (f) { return { fecha: Utilities.formatDate(f[0], tz, 'yyyy-MM-dd'), evento: String(f[1]) }; });

  return { azucar: { eventos: eventos } };
}

function agregarEventoAzucar_(evento, fecha) {
  var nombres = { inicio: 'Inicio', recaida: 'Comí azúcar' };
  if (!nombres[evento]) throw new Error('Evento inválido');
  hoja_(HOJA_AZUCAR, COLUMNAS_AZUCAR).appendRow([parsearFecha_(fecha), nombres[evento], new Date()]);
}

function deshacerEventoAzucar_() {
  var hoja = hoja_(HOJA_AZUCAR, COLUMNAS_AZUCAR);
  if (hoja.getLastRow() < 2) throw new Error('No hay nada que deshacer');
  hoja.deleteRow(hoja.getLastRow());
}

// ---------------------------------------------------------------------------
// Proteína diaria: un total de gramos por día + una nota opcional ("base de huevos,
// comí azúcar, me sentí bien…"). Es por persona; Fran y Vicente usan lo mismo.
// ---------------------------------------------------------------------------

function personaProteina_(persona) {
  persona = String(persona || '').trim();
  if (leerLista_(HOJA_PERSONAS).indexOf(persona) === -1) throw new Error('Persona inválida');
  return persona;
}

// Hoja Proteína al formato nuevo: "Alimento" pasa a "Nota" y se agrega "Persona".
function hojaProteina_() {
  var hoja = hoja_(HOJA_PROTEINA, COLUMNAS_PROTEINA);
  var enc = hoja.getRange(1, 1, 1, COLUMNAS_PROTEINA.length).getValues()[0];
  COLUMNAS_PROTEINA.forEach(function (c, i) {
    if (String(enc[i]) !== c) hoja.getRange(1, i + 1).setValue(c).setFontWeight('bold');
  });
  return hoja;
}

function filasProteina_(hoja, persona) {
  var tz = libro_().getSpreadsheetTimeZone();
  if (hoja.getLastRow() < 2) return [];
  return hoja.getRange(2, 1, hoja.getLastRow() - 1, COLUMNAS_PROTEINA.length).getValues()
    .map(function (f, i) {
      return {
        fila: i + 2,
        fecha: f[1] instanceof Date ? Utilities.formatDate(f[1], tz, 'yyyy-MM-dd') : '',
        gramos: Number(f[2]) || 0,
        nota: String(f[3] || '').trim(),
        persona: String(f[5] || '').trim() || 'Fran'
      };
    })
    .filter(function (f) { return f.fecha && f.persona === persona; });
}

function datosProteina_(persona) {
  persona = personaProteina_(persona);
  // Junta por día (los días antiguos pueden tener varias filas, una por comida).
  var dias = {};
  filasProteina_(hojaProteina_(), persona).forEach(function (f) {
    var d = dias[f.fecha] = dias[f.fecha] || { fecha: f.fecha, gramos: 0, notas: [] };
    d.gramos += f.gramos;
    if (f.nota) d.notas.push(f.nota);
  });
  var registros = Object.keys(dias).sort().reverse().slice(0, 120).map(function (k) {
    var d = dias[k];
    return { fecha: d.fecha, gramos: d.gramos, nota: d.notas.join(' · ') };
  });
  return { persona: persona, meta: metaProteina_(persona), registros: registros };
}

function guardarProteina_(body) {
  var persona = personaProteina_(body.persona);
  var gramos = Math.round(Number(body.gramos));
  var nota = String(body.nota || '').replace(/[ \t]+/g, ' ').trim().slice(0, 500);
  if (!(gramos >= 0 && gramos <= 1000)) throw new Error('Gramos inválidos');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(body.fecha || ''))) throw new Error('Fecha inválida');
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var hoja = hojaProteina_();
    borrarFilasProteina_(hoja, persona, body.fecha);
    hoja.appendRow([Utilities.getUuid(), parsearFecha_(body.fecha), gramos, nota, new Date(), persona]);
  } finally {
    lock.releaseLock();
  }
}

function borrarProteinaDia_(persona, fecha) {
  persona = personaProteina_(persona);
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    borrarFilasProteina_(hojaProteina_(), persona, String(fecha || ''));
  } finally {
    lock.releaseLock();
  }
}

function borrarFilasProteina_(hoja, persona, fecha) {
  filasProteina_(hoja, persona)
    .filter(function (f) { return f.fecha === fecha; })
    .reverse()
    .forEach(function (f) { hoja.deleteRow(f.fila); });
}

function metaProteina_(persona) {
  var props = PropertiesService.getScriptProperties();
  var meta = Number(props.getProperty('PROTEINA_META_' + persona));
  // La meta antigua (sin persona) era la de Fran.
  if (!meta && persona === 'Fran') meta = Number(props.getProperty('PROTEINA_META'));
  return meta || META_PROTEINA_INICIAL;
}

function cambiarMetaProteina_(persona, meta) {
  persona = personaProteina_(persona);
  meta = Math.round(Number(meta));
  if (!(meta >= 10 && meta <= 500)) throw new Error('Meta inválida');
  PropertiesService.getScriptProperties().setProperty('PROTEINA_META_' + persona, String(meta));
}

// ---------------------------------------------------------------------------
// Deporte (Vicente): registro de entrenamientos y diario.
// - Entrenamientos: una fila por sesión. Los ejercicios van en una celda, uno por
//   línea, como "Nombre — detalle" (ej. "Dominadas — 4x6").
// - Diario Vicente: una fila por día con cigarros, comida y peso.
// ---------------------------------------------------------------------------
var DIAS_ENTRENOS = 400;
var DIAS_DIARIO = 60;

function datosDeporte_() {
  var tz = libro_().getSpreadsheetTimeZone();
  var desdeE = Utilities.formatDate(new Date(Date.now() - DIAS_ENTRENOS * 864e5), tz, 'yyyy-MM-dd');
  var desdeD = Utilities.formatDate(new Date(Date.now() - DIAS_DIARIO * 864e5), tz, 'yyyy-MM-dd');

  var hojaE = hoja_(HOJA_ENTRENOS, COLUMNAS_ENTRENOS);
  var sesiones = hojaE.getLastRow() < 2 ? [] : hojaE.getRange(2, 1, hojaE.getLastRow() - 1, 6).getValues()
    .filter(function (f) { return f[0] && f[1] instanceof Date; })
    .map(function (f) {
      return {
        id: f[0],
        fecha: Utilities.formatDate(f[1], tz, 'yyyy-MM-dd'),
        categorias: String(f[2]).split(',').map(function (c) { return c.trim(); }).filter(function (c) { return c; }),
        minutos: Number(f[3]) || 0,
        ejercicios: String(f[4]).split('\n').map(function (l) {
          var i = l.indexOf(' — ');
          return i === -1 ? { nombre: l.trim(), detalle: '' } : { nombre: l.slice(0, i).trim(), detalle: l.slice(i + 3).trim() };
        }).filter(function (e) { return e.nombre; }),
        nota: String(f[5] || '')
      };
    })
    .filter(function (s) { return s.fecha >= desdeE; });
  sesiones.sort(function (a, b) { return a.fecha < b.fecha ? 1 : a.fecha > b.fecha ? -1 : 0; });

  var hojaD = hoja_(HOJA_DIARIO, COLUMNAS_DIARIO);
  var filas = hojaD.getLastRow() < 2 ? [] : hojaD.getRange(2, 1, hojaD.getLastRow() - 1, 5).getValues()
    .filter(function (f) { return f[0] instanceof Date; })
    .map(function (f) {
      return {
        fecha: Utilities.formatDate(f[0], tz, 'yyyy-MM-dd'),
        cigarros: Number(f[1]) || 0,
        comida: String(f[2] || ''),
        etiquetas: String(f[3] || '').split(',').map(function (c) { return c.trim(); }).filter(function (c) { return c; }),
        peso: f[4] === '' || f[4] === null ? null : Number(f[4]) || null
      };
    });
  filas.sort(function (a, b) { return a.fecha < b.fecha ? -1 : a.fecha > b.fecha ? 1 : 0; });
  var pesos = filas.filter(function (f) { return f.peso; }).map(function (f) { return { fecha: f.fecha, peso: f.peso }; });

  var props = PropertiesService.getScriptProperties();
  var metas = {};
  Object.keys(METAS_DEPORTE).forEach(function (k) {
    metas[k] = Number(props.getProperty(METAS_DEPORTE[k].prop)) || METAS_DEPORTE[k].inicial;
  });
  var ocultos = [];
  try { ocultos = JSON.parse(props.getProperty('DEPORTE_OCULTOS') || '[]'); } catch (e) {}

  return {
    sesiones: sesiones,
    diario: filas.filter(function (f) { return f.fecha >= desdeD; }),
    pesos: pesos.slice(-30),
    metas: metas,
    ocultos: ocultos
  };
}

function agregarEntreno_(body) {
  var categorias = (body.categorias || []).filter(function (c) { return CATEGORIAS_DEPORTE.indexOf(c) !== -1; });
  if (!categorias.length) throw new Error('Elige al menos una categoría');
  var minutos = Math.round(Number(body.minutos) || 0);
  if (minutos < 0 || minutos > 600) throw new Error('Minutos inválidos');
  var limpiar = function (t, n) { return String(t || '').replace(/\s+/g, ' ').trim().replace(/^[=+\-@]+/, '').slice(0, n); };
  var ejercicios = (body.ejercicios || []).slice(0, 40).map(function (e) {
    var nombre = limpiar(e && e.nombre, 60).replace(/ — /g, ' - ');
    var detalle = limpiar(e && e.detalle, 60);
    return nombre ? nombre + (detalle ? ' — ' + detalle : '') : '';
  }).filter(function (l) { return l; });
  var nota = limpiar(body.nota, 300);
  hoja_(HOJA_ENTRENOS, COLUMNAS_ENTRENOS).appendRow([
    Utilities.getUuid(), parsearFecha_(body.fecha), categorias.join(', '), minutos || '', ejercicios.join('\n'), nota, new Date()
  ]);
}

// Devuelve el número de fila del día (o 0 si no existe).
function filaDiario_(hoja, fecha, tz) {
  var buscada = Utilities.formatDate(fecha, tz, 'yyyy-MM-dd');
  var ultima = hoja.getLastRow();
  if (ultima < 2) return 0;
  var fechas = hoja.getRange(2, 1, ultima - 1, 1).getValues();
  for (var i = fechas.length - 1; i >= 0; i--) {
    if (fechas[i][0] instanceof Date && Utilities.formatDate(fechas[i][0], tz, 'yyyy-MM-dd') === buscada) return i + 2;
  }
  return 0;
}

// Cambia los campos de un día (crea la fila si no existe). cambios: { columna (1-based): fn(valorActual) }.
function editarDiario_(fechaTexto, cambios) {
  var tz = libro_().getSpreadsheetTimeZone();
  var fecha = parsearFecha_(fechaTexto);
  var hoja = hoja_(HOJA_DIARIO, COLUMNAS_DIARIO);
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var fila = filaDiario_(hoja, fecha, tz);
    var valores = fila ? hoja.getRange(fila, 1, 1, COLUMNAS_DIARIO.length).getValues()[0] : [fecha, '', '', '', '', ''];
    Object.keys(cambios).forEach(function (col) { valores[col - 1] = cambios[col](valores[col - 1]); });
    valores[COLUMNAS_DIARIO.length - 1] = new Date();
    if (fila) hoja.getRange(fila, 1, 1, valores.length).setValues([valores]);
    else hoja.appendRow(valores);
  } finally {
    lock.releaseLock();
  }
}

var CAMPOS_DIARIO = { cigarros: { col: 2, max: 200 } };

function sumarDiario_(body) {
  var c = CAMPOS_DIARIO[body.campo];
  var n = Math.round(Number(body.n));
  if (!c) throw new Error('Campo inválido');
  if (!n || Math.abs(n) > c.max) throw new Error('Cantidad inválida');
  var cambios = {};
  cambios[c.col] = function (v) { return Math.min(c.max, Math.max(0, (Number(v) || 0) + n)); };
  editarDiario_(body.fecha, cambios);
}

// Solo cambia los campos que vienen: cigarros, comida, etiquetas, peso.
function guardarDiario_(body) {
  var cambios = {};
  Object.keys(CAMPOS_DIARIO).forEach(function (k) {
    if (!(k in body)) return;
    var n = Math.round(Number(body[k]));
    if (!(n >= 0 && n <= CAMPOS_DIARIO[k].max)) throw new Error('Cantidad inválida');
    cambios[CAMPOS_DIARIO[k].col] = function () { return n; };
  });
  if ('comida' in body) {
    var comida = String(body.comida || '').replace(/\s+/g, ' ').trim().replace(/^[=+\-@]+/, '').slice(0, 300);
    cambios[3] = function () { return comida; };
  }
  if ('etiquetas' in body) {
    var etiquetas = (body.etiquetas || []).map(function (t) { return String(t).replace(/,/g, ' ').trim().slice(0, 30); })
      .filter(function (t) { return t; }).slice(0, 15).join(', ');
    cambios[4] = function () { return etiquetas; };
  }
  if ('peso' in body) {
    var peso = body.peso === null || body.peso === '' ? '' : Math.round(Number(String(body.peso).replace(',', '.')) * 10) / 10;
    if (peso !== '' && !(peso >= 30 && peso <= 250)) throw new Error('Peso inválido');
    cambios[5] = function () { return peso; };
  }
  if (!Object.keys(cambios).length) throw new Error('No hay nada que guardar');
  editarDiario_(body.fecha, cambios);
}

function cambiarMetaDeporte_(clave, valor) {
  var m = METAS_DEPORTE[clave];
  if (!m) throw new Error('Meta inválida');
  var v = Math.round(Number(String(valor).replace(',', '.')) * 10) / 10;
  if (!(v >= m.min && v <= m.max)) throw new Error('Meta inválida');
  PropertiesService.getScriptProperties().setProperty(m.prop, String(v));
}

// Ejercicios que Vicente no quiere que se le propongan.
function ocultarEjercicio_(id, ocultar) {
  id = String(id || '').replace(/[^\w-]/g, '').slice(0, 40);
  if (!id) throw new Error('Falta el ejercicio');
  var props = PropertiesService.getScriptProperties();
  var lista = [];
  try { lista = JSON.parse(props.getProperty('DEPORTE_OCULTOS') || '[]'); } catch (e) {}
  lista = lista.filter(function (x) { return x !== id; });
  if (ocultar) lista.push(id);
  props.setProperty('DEPORTE_OCULTOS', JSON.stringify(lista.slice(-200)));
}

// ---------------------------------------------------------------------------
// Utilidades
// ---------------------------------------------------------------------------

function filaDeId_(hoja, id) {
  if (!id) throw new Error('Falta el ID');
  var ultima = hoja.getLastRow();
  if (ultima >= 2) {
    var ids = hoja.getRange(2, 1, ultima - 1, 1).getValues();
    for (var i = ids.length - 1; i >= 0; i--) {
      if (ids[i][0] === id) return i + 2;
    }
  }
  throw new Error('Registro no encontrado');
}

function borrarFilaPorId_(hoja, id) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    hoja.deleteRow(filaDeId_(hoja, id));
  } finally {
    lock.releaseLock();
  }
}

function pesos_(n) {
  return '$' + String(Math.round(n)).replace(/\B(?=(\d{3})+(?!\d))/g, '.');
}

function html_(t) {
  return String(t).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
