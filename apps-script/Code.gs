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

var COLUMNAS_GASTOS = ['ID', 'Fecha', 'Centro de costo', 'Monto', 'Descripción', 'Registró', 'Creado', 'Cuenta'];
var CENTROS_INICIALES = ['Auto', 'Familia', 'Salud', 'Vicky', 'Fran', 'Vicente', 'Casa'];
var PERSONAS_INICIALES = ['Vicente', 'Fran'];
// Cuentas que se crean solas si existe un centro con ese nombre (sin importar mayúsculas).
var CUENTAS_INICIALES = {
  'Auto': ['Seguro', 'Bencina', 'TAG'],
  'Familia': ['Regalos', 'Eventos / Salidas', 'Supermercado', 'Luna'],
  'Salud': ['Isapre', 'Farmacia'],
  'Vicky': ['Sueldo', 'Traslado', 'Imposiciones']
};
// Cuentas que cambiaron de nombre: { centro: { 'nombre viejo': 'nombre nuevo' } }.
var CUENTAS_RENOMBRADAS = { 'Vicky': { 'Contribuciones': 'Imposiciones' } };
var ULTIMOS_GASTOS = 20;
var COLUMNAS_RENDICIONES = ['ID', 'Fecha', 'Monto', 'Descripción', 'Comprobante', 'Estado', 'Correo', 'Registró', 'Creado'];
var COLUMNAS_AZUCAR = ['Fecha', 'Evento', 'Creado'];
var COLUMNAS_PROTEINA = ['ID', 'Fecha', 'Gramos', 'Alimento', 'Creado'];
var CARPETA_RENDICIONES = 'Rendiciones - comprobantes';
var META_PROTEINA_INICIAL = 100;
var ULTIMAS_RENDICIONES = 50;
// Se muestra al abrir la URL /exec: sirve para comprobar que Google tiene el código nuevo.
var VERSION = '2026-11-02 · cuentas + rendiciones y salud';

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

  prepararHoja_(ss, HOJA_RENDICIONES, COLUMNAS_RENDICIONES);
  prepararHoja_(ss, HOJA_AZUCAR, COLUMNAS_AZUCAR);
  prepararHoja_(ss, HOJA_PROTEINA, COLUMNAS_PROTEINA);

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
// 3. Renombra cuentas (Vicky / Contribuciones → Imposiciones).
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
      .setValues([['Centro de costo']].concat(CENTROS_INICIALES.map(function (c) { return [c]; })));
    hojaCentros.getRange(1, 1).setFontWeight('bold');

    // Cuentas
    var filasCuentas = [['Centro de costo', 'Cuenta']];
    Object.keys(CUENTAS_INICIALES).forEach(function (centro) {
      CUENTAS_INICIALES[centro].forEach(function (c) { filasCuentas.push([centro, c]); });
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
      var rango = gastos.getRange(2, 1, n, COLUMNAS_GASTOS.length);
      var filas = rango.getValues();
      filas.forEach(function (f) {
        if (!f[0] && !f[3]) return;
        var centro = String(f[2]).trim(), cuenta = String(f[7] || '').trim();
        var antes = centro + (cuenta ? ' / ' + cuenta : '');
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
      rango.setValues(filas);
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
      case 'proteinaAdd':
        agregarProteina_(body);
        return json_(Object.assign({ ok: true }, datosSalud_()));
      case 'proteinaDelete':
        borrarFilaPorId_(hoja_(HOJA_PROTEINA, COLUMNAS_PROTEINA), body.id);
        return json_(Object.assign({ ok: true }, datosSalud_()));
      case 'proteinaMeta':
        cambiarMetaProteina_(body.meta);
        return json_(Object.assign({ ok: true }, datosSalud_()));
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
// Salud: contador de días sin azúcar y proteína diaria.
// ---------------------------------------------------------------------------

function datosSalud_() {
  var tz = libro_().getSpreadsheetTimeZone();
  var hojaA = hoja_(HOJA_AZUCAR, COLUMNAS_AZUCAR);
  var eventos = hojaA.getLastRow() < 2 ? [] : hojaA.getRange(2, 1, hojaA.getLastRow() - 1, 2).getValues()
    .filter(function (f) { return f[0] instanceof Date; })
    .map(function (f) { return { fecha: Utilities.formatDate(f[0], tz, 'yyyy-MM-dd'), evento: String(f[1]) }; });

  var hojaP = hoja_(HOJA_PROTEINA, COLUMNAS_PROTEINA);
  var filas = hojaP.getLastRow() < 2 ? [] : hojaP.getRange(2, 1, hojaP.getLastRow() - 1, 4).getValues()
    .filter(function (f) { return f[0] && f[1] instanceof Date; })
    .map(function (f) {
      return { id: f[0], fecha: Utilities.formatDate(f[1], tz, 'yyyy-MM-dd'), gramos: Number(f[2]) || 0, alimento: String(f[3]) };
    });
  var desde = new Date();
  desde.setDate(desde.getDate() - 35);
  desde = Utilities.formatDate(desde, tz, 'yyyy-MM-dd');

  // Alimentos más usados (con sus últimos gramos) para agregarlos de un toque.
  var usados = {};
  filas.slice(-300).forEach(function (f) {
    var k = f.alimento.toLowerCase();
    if (!k) return;
    usados[k] = { alimento: f.alimento, gramos: f.gramos, veces: ((usados[k] || {}).veces || 0) + 1 };
  });
  var frecuentes = Object.keys(usados).map(function (k) { return usados[k]; })
    .sort(function (a, b) { return b.veces - a.veces; }).slice(0, 8);

  return {
    azucar: { eventos: eventos },
    proteina: {
      meta: Number(PropertiesService.getScriptProperties().getProperty('PROTEINA_META')) || META_PROTEINA_INICIAL,
      registros: filas.filter(function (f) { return f.fecha >= desde; }),
      frecuentes: frecuentes
    }
  };
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

function agregarProteina_(body) {
  var gramos = Math.round(Number(body.gramos));
  var alimento = String(body.alimento || '').replace(/\s+/g, ' ').trim().slice(0, 60);
  if (!(gramos > 0 && gramos <= 500)) throw new Error('Gramos inválidos');
  if (!alimento) throw new Error('Escribe qué comiste');
  hoja_(HOJA_PROTEINA, COLUMNAS_PROTEINA).appendRow([Utilities.getUuid(), parsearFecha_(body.fecha), gramos, alimento, new Date()]);
}

function cambiarMetaProteina_(meta) {
  meta = Math.round(Number(meta));
  if (!(meta >= 10 && meta <= 500)) throw new Error('Meta inválida');
  PropertiesService.getScriptProperties().setProperty('PROTEINA_META', String(meta));
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
