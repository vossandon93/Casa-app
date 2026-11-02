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

var COLUMNAS_GASTOS = ['ID', 'Fecha', 'Centro de costo', 'Monto', 'Descripción', 'Registró', 'Creado', 'Cuenta'];
var CENTROS_INICIALES = ['Farmacia', 'Luna', 'Supermercado', 'Fran'];
var PERSONAS_INICIALES = ['Vicente', 'Fran'];
// Cuentas que se crean solas si existe un centro con ese nombre (sin importar mayúsculas).
var CUENTAS_INICIALES = { 'Vicky': ['Sueldo', 'Contribuciones', 'Traslado'], 'Nana': ['Sueldo', 'Contribuciones', 'Traslado'] };
var ULTIMOS_GASTOS = 20;
// Se muestra al abrir la URL /exec: sirve para comprobar que Google tiene el código nuevo.
var VERSION = '2026-11-02 · cuentas dentro de cada centro';

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
    gastos.appendRow(COLUMNAS_GASTOS);
  }
  gastos.setFrozenRows(1);
  gastos.getRange(1, 1, 1, COLUMNAS_GASTOS.length).setFontWeight('bold');
  gastos.getRange('B:B').setNumberFormat('dd/mm/yyyy');
  gastos.getRange('D:D').setNumberFormat('$#,##0');
  gastos.getRange('G:G').setNumberFormat('dd/mm/yyyy hh:mm');
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

  var hoja1 = ss.getSheetByName('Hoja 1') || ss.getSheetByName('Sheet1');
  if (hoja1 && hoja1.getLastRow() === 0 && ss.getSheets().length > 1) {
    ss.deleteSheet(hoja1);
  }

  if (!PropertiesService.getScriptProperties().getProperty('PIN')) {
    Logger.log('Falta definir el PIN en Configuración del proyecto > Propiedades del script.');
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
  hoja.getRange(2, 1, hoja.getLastRow() - 1, COLUMNAS_GASTOS.length).getValues().forEach(function (f) {
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
  var col = COLUMNAS_GASTOS.length;
  var celda = hoja.getRange(1, col);
  if (!String(celda.getValue() || '').trim()) celda.setValue(COLUMNAS_GASTOS[col - 1]).setFontWeight('bold');
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
    validarPin_(body.pin);

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
      case 'resumen':
        return json_({ ok: true, filas: totalesPorMes_() });
      case 'detalle':
        return json_({ ok: true, gastos: gastosDeCentro_(body.centro, body.periodo) });
      default:
        throw new Error('Acción desconocida');
    }
  } catch (err) {
    return json_({ ok: false, error: err.message });
  }
}

function validarPin_(pin) {
  var esperado = PropertiesService.getScriptProperties().getProperty('PIN');
  if (!esperado) throw new Error('El PIN no está configurado en el script');
  if (String(pin || '') !== String(esperado)) throw new Error('PIN incorrecto');
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
  } finally {
    lock.releaseLock();
  }
  return leerLista_(HOJA_CENTROS);
}

function config_() {
  return { ok: true, centros: leerLista_(HOJA_CENTROS), personas: leerLista_(HOJA_PERSONAS), cuentas: leerCuentas_() };
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
  } finally {
    lock.releaseLock();
  }
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
    hoja.appendRow([Utilities.getUuid(), fecha, centro, monto, descripcion, persona, new Date(), cuenta]);
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
  return hoja.getRange(ultima - n + 1, 1, n, COLUMNAS_GASTOS.length).getValues()
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
  return hoja.getRange(2, 2, ultima - 1, 3).getValues().reduce(function (suma, f) {
    if (f[0] instanceof Date && Utilities.formatDate(f[0], tz, 'yyyy-MM') === mesActual) {
      return suma + (Number(f[2]) || 0);
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
  return hoja.getRange(2, 1, hoja.getLastRow() - 1, COLUMNAS_GASTOS.length).getValues()
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

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
