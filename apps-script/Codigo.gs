/* =====================================================================
   Codigo.gs  —  Servidor del sistema de ventas (Google Apps Script)

   Se pega dentro de la hoja de cálculo: Extensiones > Apps Script.
   Recibe las peticiones de la página, valida los datos, evita
   duplicados y escribe en las pestañas VENTAS, DETALLE_VENTAS y
   PRODUCTOS.

   La clave de acceso NO va escrita aquí: se guarda en
   Configuración del proyecto > Propiedades de la secuencia de comandos
   con el nombre CLAVE_ACCESO (ver instrucciones.txt).
   ===================================================================== */

var ZONA = 'America/Costa_Rica';

var HOJA_VENTAS = 'VENTAS';
var HOJA_DETALLE = 'DETALLE_VENTAS';
var HOJA_PRODUCTOS = 'PRODUCTOS';

var ENCABEZADOS = {
  VENTAS: ['ID de venta', 'Fecha', 'Hora', 'Nombre del cliente', 'Cantidad total de unidades', 'Total de la venta', 'Estado del registro'],
  DETALLE_VENTAS: ['ID de venta', 'Fecha', 'Nombre del producto', 'Precio unitario', 'Cantidad', 'Subtotal'],
  PRODUCTOS: ['ID del producto', 'Nombre', 'Precio', 'Estado']
};

var PRODUCTOS_INICIALES = [
  ['P001', 'Hamburguesa', 1500, 'ACTIVO'],
  ['P002', 'Papas doraditas', 1000, 'ACTIVO'],
  ['P003', 'Coca-Cola', 800, 'ACTIVO'],
  ['P004', 'Fresco natural', 600, 'ACTIVO'],
  ['P005', 'Perro caliente', 1200, 'ACTIVO']
];

// Estados de una venta en la columna "Estado del registro"
var EN_PROCESO = 'EN_PROCESO';   // se empezó a escribir, todavía no cuenta
var CONFIRMADA = 'CONFIRMADA';   // venta completa: es la única que se cuenta
var DESCARTADA = 'DESCARTADA';   // intento incompleto que el usuario descartó

// Deben coincidir con logica.js
var PRECIO_MAX = 1000000;
var CANTIDAD_MAX = 999;
var LINEAS_MAX = 50;
var CLIENTE_MAX = 60;
var NOMBRE_PRODUCTO_MAX = 40;

/* ------------------------------------------------------------------ */
/* Puntos de entrada                                                  */
/* ------------------------------------------------------------------ */

/** Al abrir la URL en el navegador solo se muestra un aviso. */
function doGet() {
  return ContentService.createTextOutput('El servidor del sistema de ventas está activo.');
}

/** Todas las operaciones llegan por POST con un JSON { accion, clave, ... }. */
function doPost(e) {
  var respuesta;
  try {
    var datos = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    verificarClave(datos.clave);
    prepararHojas();
    switch (datos.accion) {
      case 'probar':           respuesta = { ok: true, mensaje: 'Conexión correcta.' }; break;
      case 'obtenerProductos': respuesta = { ok: true, productos: leerProductos() }; break;
      case 'guardarProductos': respuesta = guardarProductos(datos.productos); break;
      case 'registrarVenta':   respuesta = registrarVenta(datos.venta); break;
      case 'descartarIntento': respuesta = descartarIntento(datos.id); break;
      case 'ventasPorFecha':   respuesta = ventasPorFecha(datos.fecha); break;
      default: throw errorCon('ACCION', 'Acción desconocida.');
    }
  } catch (err) {
    respuesta = { ok: false, codigo: err.codigo || 'INTERNO', error: err.codigo ? err.message : 'Error interno del servidor: ' + err.message };
  }
  return ContentService.createTextOutput(JSON.stringify(respuesta))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ------------------------------------------------------------------ */
/* Utilidades                                                         */
/* ------------------------------------------------------------------ */

function errorCon(codigo, mensaje) {
  var e = new Error(mensaje);
  e.codigo = codigo;
  return e;
}

function verificarClave(clave) {
  var guardada = PropertiesService.getScriptProperties().getProperty('CLAVE_ACCESO');
  if (!guardada) throw errorCon('SIN_CLAVE', 'El servidor no tiene CLAVE_ACCESO configurada. Revise instrucciones.txt.');
  if (typeof clave !== 'string' || clave !== guardada) throw errorCon('CLAVE', 'La clave de acceso no es correcta.');
}

function esEnteroPositivo(v, max) {
  return typeof v === 'number' && Math.floor(v) === v && v > 0 && v <= max;
}

/** Evita que un texto empiece con =, +, - o @ y Sheets lo tome como fórmula. */
function textoSeguro(texto, largo) {
  var t = String(texto == null ? '' : texto).replace(/[\r\n\t]+/g, ' ').trim().slice(0, largo);
  return /^[=+\-@]/.test(t) ? "'" + t : t;
}

function hoja(nombre) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombre);
}

/** Crea las pestañas y encabezados si todavía no existen. */
function prepararHojas() {
  var libro = SpreadsheetApp.getActiveSpreadsheet();
  [HOJA_VENTAS, HOJA_DETALLE, HOJA_PRODUCTOS].forEach(function (nombre) {
    var h = libro.getSheetByName(nombre);
    if (!h) h = libro.insertSheet(nombre);
    if (h.getLastRow() === 0) {
      var enc = ENCABEZADOS[nombre];
      h.getRange(1, 1, 1, enc.length).setValues([enc]).setFontWeight('bold');
      h.setFrozenRows(1);
      // Fecha y hora como texto para que no cambien según el idioma de la hoja
      if (nombre === HOJA_VENTAS) {
        h.getRange('B:C').setNumberFormat('@');
        h.getRange('F:F').setNumberFormat('"₡"#,##0');
      }
      if (nombre === HOJA_DETALLE) {
        h.getRange('B:B').setNumberFormat('@');
        h.getRange('D:D').setNumberFormat('"₡"#,##0');
        h.getRange('F:F').setNumberFormat('"₡"#,##0');
      }
      if (nombre === HOJA_PRODUCTOS) {
        h.getRange('C:C').setNumberFormat('"₡"#,##0');
        h.getRange(2, 1, PRODUCTOS_INICIALES.length, 4).setValues(PRODUCTOS_INICIALES);
      }
    }
  });
}

/** Función para ejecutar a mano una vez desde el editor (crea las pestañas). */
function inicializar() {
  prepararHojas();
  Logger.log('Pestañas listas: VENTAS, DETALLE_VENTAS y PRODUCTOS.');
}

/** Todas las filas de datos (sin encabezado). */
function filas(h) {
  var n = h.getLastRow() - 1;
  if (n < 1) return [];
  return h.getRange(2, 1, n, h.getLastColumn()).getValues();
}

/** Número de fila donde la columna A es exactamente "id" (o 0). */
function buscarFila(h, id) {
  var celda = h.getRange('A:A').createTextFinder(id).matchEntireCell(true).findNext();
  return celda ? celda.getRow() : 0;
}

function borrarDetalles(id) {
  var h = hoja(HOJA_DETALLE);
  var celdas = h.getRange('A:A').createTextFinder(id).matchEntireCell(true).findAll();
  var numeros = celdas.map(function (c) { return c.getRow(); }).sort(function (a, b) { return b - a; });
  numeros.forEach(function (n) { h.deleteRow(n); });  // de abajo hacia arriba
}

function comoFecha(v) {
  return v instanceof Date ? Utilities.formatDate(v, ZONA, 'yyyy-MM-dd') : String(v);
}

function comoHora(v) {
  return v instanceof Date ? Utilities.formatDate(v, ZONA, 'HH:mm:ss') : String(v);
}

function conBloqueo(funcion) {
  var bloqueo = LockService.getScriptLock();
  try {
    bloqueo.waitLock(25000);
  } catch (e) {
    throw errorCon('OCUPADO', 'El servidor está ocupado. La venta NO se guardó; intente de nuevo.');
  }
  try {
    return funcion();
  } finally {
    bloqueo.releaseLock();
  }
}

/* ------------------------------------------------------------------ */
/* Productos                                                          */
/* ------------------------------------------------------------------ */

function leerProductos() {
  return filas(hoja(HOJA_PRODUCTOS))
    .filter(function (f) { return String(f[0]).trim() !== ''; })
    .map(function (f) {
      return {
        id: String(f[0]).trim(),
        nombre: String(f[1]),
        precio: Number(f[2]),
        activo: String(f[3]).trim().toUpperCase() !== 'INACTIVO'
      };
    });
}

function guardarProductos(lista) {
  if (!Array.isArray(lista) || lista.length === 0) throw errorCon('VALIDACION', 'No se recibieron productos.');
  lista.forEach(function (p) {
    if (!p || typeof p.id !== 'string') throw errorCon('VALIDACION', 'Producto sin identificador.');
    var nombre = String(p.nombre || '').trim();
    if (!nombre || nombre.length > NOMBRE_PRODUCTO_MAX) throw errorCon('VALIDACION', 'El nombre del producto ' + p.id + ' no es válido.');
    if (!esEnteroPositivo(p.precio, PRECIO_MAX)) throw errorCon('VALIDACION', 'El precio de ' + nombre + ' no es válido.');
  });
  return conBloqueo(function () {
    var h = hoja(HOJA_PRODUCTOS);
    lista.forEach(function (p) {
      var fila = buscarFila(h, p.id);
      if (!fila) throw errorCon('VALIDACION', 'El producto ' + p.id + ' no existe en la pestaña PRODUCTOS.');
      h.getRange(fila, 2, 1, 3).setValues([[textoSeguro(p.nombre, NOMBRE_PRODUCTO_MAX), p.precio, p.activo === false ? 'INACTIVO' : 'ACTIVO']]);
    });
    SpreadsheetApp.flush();
    return { ok: true, productos: leerProductos() };
  });
}

/* ------------------------------------------------------------------ */
/* Ventas                                                             */
/* ------------------------------------------------------------------ */

/** Revisa la venta recibida y recalcula todo. No confía en el navegador. */
function validarVenta(v) {
  if (!v || typeof v !== 'object') throw errorCon('VALIDACION', 'No se recibió la venta.');
  if (typeof v.id !== 'string' || !/^V-\d{8}-\d{6}-[A-Z0-9]{6}$/.test(v.id)) throw errorCon('VALIDACION', 'Identificador de venta inválido.');

  var momento = new Date(v.fechaHora);
  if (isNaN(momento.getTime())) throw errorCon('VALIDACION', 'Fecha de la venta inválida.');
  var ahora = Date.now();
  if (momento.getTime() > ahora + 10 * 60 * 1000) throw errorCon('VALIDACION', 'La fecha de la venta está en el futuro. Revise la fecha y hora de la tablet.');
  if (momento.getTime() < ahora - 30 * 24 * 60 * 60 * 1000) throw errorCon('VALIDACION', 'La venta tiene más de 30 días. Regístrela manualmente en la hoja.');

  if (!Array.isArray(v.items) || v.items.length === 0) throw errorCon('VALIDACION', 'La venta no tiene productos.');
  if (v.items.length > LINEAS_MAX) throw errorCon('VALIDACION', 'La venta tiene demasiados productos distintos.');

  var existentes = {};
  leerProductos().forEach(function (p) { existentes[p.id] = true; });

  var vistos = {};
  var total = 0;
  var unidades = 0;
  var items = v.items.map(function (it) {
    if (!it || typeof it.productoId !== 'string' || !existentes[it.productoId]) throw errorCon('VALIDACION', 'Producto desconocido en la venta.');
    if (vistos[it.productoId]) throw errorCon('VALIDACION', 'Un producto aparece repetido en la venta.');
    vistos[it.productoId] = true;
    var nombre = String(it.nombre || '').trim();
    if (!nombre) throw errorCon('VALIDACION', 'Producto sin nombre.');
    if (!esEnteroPositivo(it.precio, PRECIO_MAX)) throw errorCon('VALIDACION', 'Precio inválido para ' + nombre + '.');
    if (!esEnteroPositivo(it.cantidad, CANTIDAD_MAX)) throw errorCon('VALIDACION', 'Cantidad inválida para ' + nombre + '.');
    var subtotal = it.precio * it.cantidad;
    total += subtotal;
    unidades += it.cantidad;
    return { nombre: textoSeguro(nombre, NOMBRE_PRODUCTO_MAX), precio: it.precio, cantidad: it.cantidad, subtotal: subtotal };
  });

  // El total que calculó el navegador debe coincidir con el del servidor
  if (v.total !== total) throw errorCon('VALIDACION', 'El total enviado (' + v.total + ') no coincide con el calculado (' + total + ').');
  if (v.unidades !== unidades) throw errorCon('VALIDACION', 'La cantidad de unidades no coincide.');

  return {
    id: v.id,
    fecha: Utilities.formatDate(momento, ZONA, 'yyyy-MM-dd'),
    hora: Utilities.formatDate(momento, ZONA, 'HH:mm:ss'),
    cliente: textoSeguro(v.cliente, CLIENTE_MAX),
    unidades: unidades,
    total: total,
    items: items
  };
}

/**
 * Guarda una venta una sola vez (idempotente).
 * 1. Escribe la fila en VENTAS con estado EN_PROCESO.
 * 2. Escribe las filas en DETALLE_VENTAS.
 * 3. Cambia el estado a CONFIRMADA.
 * Si algo falla a mitad, la venta queda EN_PROCESO y no se cuenta.
 * Un reintento con el mismo ID borra el detalle incompleto y lo rehace.
 * Si el ID ya está CONFIRMADA, no se vuelve a escribir.
 */
function registrarVenta(ventaRecibida) {
  var v = validarVenta(ventaRecibida);
  return conBloqueo(function () {
    if (CacheService.getScriptCache().get('descartada_' + v.id)) {
      throw errorCon('DESCARTADA', 'Esta venta fue descartada y no se guardará.');
    }
    var hv = hoja(HOJA_VENTAS);
    var fila = buscarFila(hv, v.id);

    if (fila) {
      var existente = hv.getRange(fila, 1, 1, 7).getValues()[0];
      var estado = String(existente[6]);
      if (estado === CONFIRMADA) {
        if (Number(existente[5]) !== v.total) {
          throw errorCon('CONFLICTO', 'Ya existe una venta con este ID y otro total. Revise el historial.');
        }
        return { ok: true, duplicada: true, id: v.id, mensaje: 'La venta ya estaba guardada; no se duplicó.' };
      }
      if (estado === DESCARTADA) throw errorCon('DESCARTADA', 'Esta venta fue descartada y no se guardará.');
      borrarDetalles(v.id);                       // limpiar un intento incompleto
    } else {
      fila = hv.getLastRow() + 1;
    }

    hv.getRange(fila, 1, 1, 7).setValues([[v.id, v.fecha, v.hora, v.cliente, v.unidades, v.total, EN_PROCESO]]);

    var hd = hoja(HOJA_DETALLE);
    var detalle = v.items.map(function (it) {
      return [v.id, v.fecha, it.nombre, it.precio, it.cantidad, it.subtotal];
    });
    hd.getRange(hd.getLastRow() + 1, 1, detalle.length, 6).setValues(detalle);

    hv.getRange(fila, 7).setValue(CONFIRMADA);
    SpreadsheetApp.flush();
    return { ok: true, duplicada: false, id: v.id, mensaje: 'Venta guardada.' };
  });
}

/**
 * Se usa cuando la página no supo si una venta llegó (se cortó la conexión)
 * y el usuario quiere corregirla. Si ya estaba CONFIRMADA lo avisa; si no,
 * la marca como descartada para que un envío atrasado no la guarde después.
 */
function descartarIntento(id) {
  if (typeof id !== 'string' || !/^V-\d{8}-\d{6}-[A-Z0-9]{6}$/.test(id)) throw errorCon('VALIDACION', 'Identificador inválido.');
  return conBloqueo(function () {
    var hv = hoja(HOJA_VENTAS);
    var fila = buscarFila(hv, id);
    if (fila && String(hv.getRange(fila, 7).getValue()) === CONFIRMADA) {
      return { ok: true, yaConfirmada: true, mensaje: 'Esta venta ya estaba guardada. No se puede descartar.' };
    }
    if (fila) {
      borrarDetalles(id);
      hv.getRange(fila, 7).setValue(DESCARTADA);
    }
    CacheService.getScriptCache().put('descartada_' + id, '1', 21600); // 6 horas
    SpreadsheetApp.flush();
    return { ok: true, yaConfirmada: false, mensaje: 'Intento descartado.' };
  });
}

/** Ventas CONFIRMADAS de un día, cada una con su detalle. */
function ventasPorFecha(fecha) {
  if (typeof fecha !== 'string' || !/^\d{4}-\d{2}-\d{2}$/.test(fecha)) throw errorCon('VALIDACION', 'Fecha inválida.');
  var ventas = [];
  var porId = {};
  filas(hoja(HOJA_VENTAS)).forEach(function (f) {
    if (comoFecha(f[1]) !== fecha || String(f[6]) !== CONFIRMADA) return;
    var id = String(f[0]);
    if (porId[id]) return;                          // nunca contar dos veces
    var v = { id: id, fecha: fecha, hora: comoHora(f[2]), cliente: String(f[3]), unidades: Number(f[4]), total: Number(f[5]), detalles: [] };
    porId[id] = v;
    ventas.push(v);
  });
  filas(hoja(HOJA_DETALLE)).forEach(function (f) {
    var v = porId[String(f[0])];
    if (v) v.detalles.push({ nombre: String(f[2]), precio: Number(f[3]), cantidad: Number(f[4]), subtotal: Number(f[5]) });
  });
  ventas.sort(function (a, b) { return a.hora < b.hora ? -1 : a.hora > b.hora ? 1 : 0; });
  return { ok: true, fecha: fecha, ventas: ventas };
}
