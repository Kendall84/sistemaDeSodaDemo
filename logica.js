/* =====================================================================
   logica.js
   Funciones "puras" del sistema: cálculos del carrito, validaciones,
   formato de colones, identificadores y resumen del día.
   No tocan la pantalla ni la red, por eso se pueden probar solas
   (ver carpeta pruebas/).
   ===================================================================== */
(function (global) {
  'use strict';

  // Límites de seguridad. Deben coincidir con los de Codigo.gs.
  var LIMITES = {
    PRECIO_MAX: 1000000,     // ₡1.000.000 por unidad
    CANTIDAD_MAX: 999,       // unidades de un mismo producto en una venta
    LINEAS_MAX: 50,          // productos distintos en una venta
    CLIENTE_MAX: 60,         // caracteres del nombre del cliente
    NOMBRE_PRODUCTO_MAX: 40
  };

  var ZONA_HORARIA = 'America/Costa_Rica';

  /** Entero positivo dentro de un máximo (sin decimales, sin texto). */
  function esEnteroPositivo(valor, maximo) {
    return typeof valor === 'number' &&
      Number.isInteger(valor) &&
      valor > 0 &&
      valor <= maximo;
  }

  /** ₡1.500 — colones sin decimales y con punto como separador de miles. */
  function formatearColones(monto) {
    var n = Math.round(Number(monto) || 0);
    var signo = n < 0 ? '-' : '';
    var texto = String(Math.abs(n)).replace(/\B(?=(\d{3})+(?!\d))/g, '.');
    return signo + '₡' + texto;
  }

  /** Convierte lo que escribe el usuario ("1.500", "₡1500") en entero, o null. */
  function leerPrecio(texto) {
    var limpio = String(texto == null ? '' : texto).replace(/[₡\s.,]/g, '');
    if (!/^\d+$/.test(limpio)) return null;
    var n = parseInt(limpio, 10);
    return esEnteroPositivo(n, LIMITES.PRECIO_MAX) ? n : null;
  }

  /* -------------------------- Carrito -------------------------- */
  // El carrito es una lista de líneas:
  // { productoId, nombre, precio, cantidad }
  // El nombre y el precio se copian en el momento de agregar el producto,
  // así un cambio de precio posterior no altera la venta.

  function agregarAlCarrito(carrito, producto) {
    var copia = carrito.map(function (l) { return Object.assign({}, l); });
    var linea = copia.find(function (l) { return l.productoId === producto.id; });
    if (linea) {
      if (linea.cantidad < LIMITES.CANTIDAD_MAX) linea.cantidad += 1;
    } else {
      copia.push({
        productoId: producto.id,
        nombre: producto.nombre,
        precio: producto.precio,
        cantidad: 1
      });
    }
    return copia;
  }

  /** Suma (o resta) a una línea. Si llega a 0 la línea se elimina. */
  function cambiarCantidad(carrito, productoId, delta) {
    return carrito
      .map(function (l) {
        if (l.productoId !== productoId) return Object.assign({}, l);
        var nueva = Math.min(l.cantidad + delta, LIMITES.CANTIDAD_MAX);
        return Object.assign({}, l, { cantidad: nueva });
      })
      .filter(function (l) { return l.cantidad > 0; });
  }

  function quitarDelCarrito(carrito, productoId) {
    return carrito.filter(function (l) { return l.productoId !== productoId; });
  }

  function subtotalLinea(linea) {
    return linea.precio * linea.cantidad;
  }

  function totalCarrito(carrito) {
    return carrito.reduce(function (suma, l) { return suma + subtotalLinea(l); }, 0);
  }

  function unidadesCarrito(carrito) {
    return carrito.reduce(function (suma, l) { return suma + l.cantidad; }, 0);
  }

  /** Devuelve una lista de errores (vacía si el carrito es válido). */
  function validarCarrito(carrito) {
    var errores = [];
    if (!Array.isArray(carrito) || carrito.length === 0) {
      return ['La venta no tiene productos. Toque al menos un producto.'];
    }
    if (carrito.length > LIMITES.LINEAS_MAX) {
      errores.push('La venta tiene demasiados productos distintos.');
    }
    var vistos = {};
    carrito.forEach(function (l) {
      var nombre = l && l.nombre ? l.nombre : 'Producto';
      if (!l || typeof l.productoId !== 'string' || !l.productoId) {
        errores.push('Hay un producto sin identificador.');
        return;
      }
      if (vistos[l.productoId]) errores.push(nombre + ' aparece repetido.');
      vistos[l.productoId] = true;
      if (!esEnteroPositivo(l.precio, LIMITES.PRECIO_MAX)) {
        errores.push('El precio de ' + nombre + ' no es válido.');
      }
      if (!esEnteroPositivo(l.cantidad, LIMITES.CANTIDAD_MAX)) {
        errores.push('La cantidad de ' + nombre + ' no es válida.');
      }
    });
    if (errores.length === 0) {
      var total = totalCarrito(carrito);
      if (!Number.isSafeInteger(total) || total <= 0) {
        errores.push('El total de la venta no es válido.');
      }
    }
    return errores;
  }

  /* --------------------- Fechas e identificadores --------------------- */

  /** Partes de fecha/hora en Costa Rica, sin importar la zona de la tablet. */
  function partesFechaCR(fecha) {
    var f = new Intl.DateTimeFormat('en-CA', {
      timeZone: ZONA_HORARIA, year: 'numeric', month: '2-digit', day: '2-digit',
      hour: '2-digit', minute: '2-digit', second: '2-digit', hourCycle: 'h23'
    }).formatToParts(fecha || new Date());
    var p = {};
    f.forEach(function (x) { p[x.type] = x.value; });
    return p;
  }

  /** "2026-10-25" (fecha de hoy en Costa Rica). */
  function fechaHoyCR(fecha) {
    var p = partesFechaCR(fecha);
    return p.year + '-' + p.month + '-' + p.day;
  }

  /** "25/10/2026" a partir de "2026-10-25". */
  function fechaLegible(iso) {
    var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso || '');
    return m ? m[3] + '/' + m[2] + '/' + m[1] : (iso || '');
  }

  function textoAleatorio(largo) {
    var letras = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
    var bytes = new Uint8Array(largo);
    var c = global.crypto || (typeof require === 'function' ? require('crypto').webcrypto : null);
    c.getRandomValues(bytes);
    var s = '';
    for (var i = 0; i < largo; i++) s += letras[bytes[i] % letras.length];
    return s;
  }

  /** Identificador único y legible: V-20261025-143510-K7M2QX */
  function generarIdVenta(fecha) {
    var p = partesFechaCR(fecha);
    return 'V-' + p.year + p.month + p.day + '-' + p.hour + p.minute + p.second +
      '-' + textoAleatorio(6);
  }

  /** Arma el objeto que se envía al servidor. */
  function construirVenta(carrito, cliente, fecha) {
    var momento = fecha || new Date();
    return {
      id: generarIdVenta(momento),
      fechaHora: momento.toISOString(),
      cliente: String(cliente || '').trim().slice(0, LIMITES.CLIENTE_MAX),
      items: carrito.map(function (l) {
        return { productoId: l.productoId, nombre: l.nombre, precio: l.precio, cantidad: l.cantidad };
      }),
      unidades: unidadesCarrito(carrito),
      total: totalCarrito(carrito)
    };
  }

  /* ------------------------- Resumen del día ------------------------- */

  /**
   * ventas: lista que devuelve el servidor (solo ventas CONFIRMADAS), cada una
   * con { id, total, unidades, detalles: [{ nombre, precio, cantidad, subtotal }] }
   */
  function calcularResumen(ventas) {
    var lista = Array.isArray(ventas) ? ventas : [];
    var vistos = {};
    var unicas = lista.filter(function (v) {   // nunca contar dos veces el mismo ID
      if (vistos[v.id]) return false;
      vistos[v.id] = true;
      return true;
    });
    var porProducto = {};
    var totalDetalle = 0;
    unicas.forEach(function (v) {
      (v.detalles || []).forEach(function (d) {
        if (!porProducto[d.nombre]) porProducto[d.nombre] = { nombre: d.nombre, unidades: 0, total: 0 };
        porProducto[d.nombre].unidades += d.cantidad;
        porProducto[d.nombre].total += d.subtotal;
        totalDetalle += d.subtotal;
      });
    });
    var totalVendido = unicas.reduce(function (s, v) { return s + v.total; }, 0);
    var cantidadVentas = unicas.length;
    return {
      totalVendido: totalVendido,
      cantidadVentas: cantidadVentas,
      unidades: unicas.reduce(function (s, v) { return s + v.unidades; }, 0),
      promedio: cantidadVentas ? Math.round(totalVendido / cantidadVentas) : 0,
      productos: Object.keys(porProducto).map(function (k) { return porProducto[k]; })
        .sort(function (a, b) { return b.total - a.total; }),
      // true si VENTAS y DETALLE_VENTAS suman lo mismo
      coincide: totalDetalle === totalVendido
    };
  }

  var Logica = {
    LIMITES: LIMITES,
    ZONA_HORARIA: ZONA_HORARIA,
    esEnteroPositivo: esEnteroPositivo,
    formatearColones: formatearColones,
    leerPrecio: leerPrecio,
    agregarAlCarrito: agregarAlCarrito,
    cambiarCantidad: cambiarCantidad,
    quitarDelCarrito: quitarDelCarrito,
    subtotalLinea: subtotalLinea,
    totalCarrito: totalCarrito,
    unidadesCarrito: unidadesCarrito,
    validarCarrito: validarCarrito,
    fechaHoyCR: fechaHoyCR,
    fechaLegible: fechaLegible,
    generarIdVenta: generarIdVenta,
    construirVenta: construirVenta,
    calcularResumen: calcularResumen
  };

  if (typeof module !== 'undefined' && module.exports) module.exports = Logica;
  else global.Logica = Logica;
})(typeof window !== 'undefined' ? window : globalThis);
