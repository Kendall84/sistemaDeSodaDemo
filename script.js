/* =====================================================================
   script.js
   Pantalla del sistema: secciones, carrito, envío de ventas a Google
   Sheets, ventas pendientes, historial, resumen, configuración y
   exportación a Excel. Los cálculos están en logica.js.
   ===================================================================== */
(function () {
  'use strict';

  var L = window.Logica;

  /* ------------------------------------------------------------------ */
  /* Almacenamiento local (solo en este dispositivo)                    */
  /* ------------------------------------------------------------------ */
  var CLAVES = {
    config: 'sv_config',          // { url, clave }
    productos: 'sv_productos',    // copia de la pestaña PRODUCTOS
    carrito: 'sv_carrito',        // venta que se está armando
    cliente: 'sv_cliente',
    intento: 'sv_intento',        // venta enviada sin respuesta confirmada
    pendientes: 'sv_pendientes'   // ventas no confirmadas, con su ID original
  };

  function leer(clave, porDefecto) {
    try {
      var t = localStorage.getItem(clave);
      return t ? JSON.parse(t) : porDefecto;
    } catch (e) { return porDefecto; }
  }

  function guardar(clave, valor) {
    try { localStorage.setItem(clave, JSON.stringify(valor)); return true; }
    catch (e) {
      alert('No se pudo guardar en este dispositivo (almacenamiento lleno o bloqueado). ' +
        'No recargue la página hasta resolverlo.');
      return false;
    }
  }

  var PRODUCTOS_INICIALES = [
    { id: 'P001', nombre: 'Hamburguesa', precio: 1500, activo: true },
    { id: 'P002', nombre: 'Papas doraditas', precio: 1000, activo: true },
    { id: 'P003', nombre: 'Coca-Cola', precio: 800, activo: true },
    { id: 'P004', nombre: 'Fresco natural', precio: 600, activo: true },
    { id: 'P005', nombre: 'Perro caliente', precio: 1200, activo: true }
  ];

  var estado = {
    config: leer(CLAVES.config, { url: '', clave: '' }),
    productos: leer(CLAVES.productos, PRODUCTOS_INICIALES),
    carrito: leer(CLAVES.carrito, []),
    intento: leer(CLAVES.intento, null),       // { venta, estado: 'enviando' | 'incierta' }
    pendientes: leer(CLAVES.pendientes, []),   // [{ venta, error, rechazada }]
    historial: [],
    enviandoPendientes: false
  };

  /* ------------------------------------------------------------------ */
  /* Utilidades de pantalla                                             */
  /* ------------------------------------------------------------------ */
  function $(id) { return document.getElementById(id); }

  /** Crea un elemento con texto seguro (nunca usa innerHTML con datos). */
  function el(etiqueta, atributos, hijos) {
    var n = document.createElement(etiqueta);
    Object.keys(atributos || {}).forEach(function (k) {
      if (k === 'texto') n.textContent = atributos[k];
      else if (k === 'clase') n.className = atributos[k];
      else if (k.indexOf('on') === 0) n.addEventListener(k.slice(2), atributos[k]);
      else n.setAttribute(k, atributos[k]);
    });
    (hijos || []).forEach(function (h) { if (h) n.appendChild(h); });
    return n;
  }

  function mostrarMensaje(id, tipo, texto) {
    var m = $(id);
    m.className = 'mensaje mensaje-' + tipo;
    m.textContent = '';
    texto.split('\n').forEach(function (linea, i) {
      if (i) m.appendChild(document.createElement('br'));
      m.appendChild(document.createTextNode(linea));
    });
  }

  function ocultarMensaje(id) { $(id).className = 'mensaje oculto'; }

  /* ------------------------------------------------------------------ */
  /* Comunicación con Google Apps Script                                */
  /* ------------------------------------------------------------------ */

  // El servidor respondió que NO guardó (dato inválido, clave, ocupado...)
  function ErrorServidor(mensaje, codigo) { this.message = mensaje; this.codigo = codigo; }
  // No se llegó a enviar nada (falta configuración)
  function ErrorConfig(mensaje) { this.message = mensaje; }
  // No hubo respuesta clara: NO se sabe si el servidor guardó
  function ErrorRed(mensaje) { this.message = mensaje; }

  var TIEMPO_ESPERA_MS = 30000;

  function llamar(accion, datos) {
    var cfg = estado.config;
    if (!cfg.url || !cfg.clave) {
      return Promise.reject(new ErrorConfig('Falta configurar la URL de Apps Script y la clave en Configuración.'));
    }
    if (!navigator.onLine) {
      return Promise.reject(new ErrorRed('El dispositivo no tiene conexión a Internet.'));
    }
    var cuerpo = Object.assign({ accion: accion, clave: cfg.clave }, datos || {});
    var control = new AbortController();
    var temporizador = setTimeout(function () { control.abort(); }, TIEMPO_ESPERA_MS);

    // text/plain evita la verificación previa (CORS) que Apps Script no admite
    return fetch(cfg.url, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify(cuerpo),
      redirect: 'follow',
      signal: control.signal
    })
      .catch(function () {
        throw new ErrorRed('No hubo respuesta del servidor (sin Internet, tiempo agotado o URL incorrecta).');
      })
      .then(function (resp) {
        return resp.json().catch(function () {
          throw new ErrorRed('El servidor respondió algo inesperado (código ' + resp.status + '). Revise la URL y el despliegue.');
        });
      })
      .then(function (json) {
        if (!json || json.ok !== true) {
          throw new ErrorServidor((json && json.error) || 'El servidor rechazó la operación.', json && json.codigo);
        }
        return json;
      })
      .finally(function () { clearTimeout(temporizador); });
  }

  /* ------------------------------------------------------------------ */
  /* Navegación                                                         */
  /* ------------------------------------------------------------------ */
  function irA(seccion) {
    document.querySelectorAll('.menu-btn').forEach(function (b) {
      b.classList.toggle('activo', b.dataset.seccion === seccion);
    });
    document.querySelectorAll('.seccion').forEach(function (s) {
      s.classList.toggle('activa', s.id === 'seccion-' + seccion);
    });
    document.body.dataset.seccion = seccion;
    if (seccion === 'config') pintarConfig();
    window.scrollTo(0, 0);
  }

  /* ------------------------------------------------------------------ */
  /* Registrar venta                                                    */
  /* ------------------------------------------------------------------ */
  function carritoBloqueado() { return estado.intento !== null; }

  function cambiarCarrito(nuevo) {
    if (carritoBloqueado()) return;
    estado.carrito = nuevo;
    guardar(CLAVES.carrito, estado.carrito);
    ocultarMensaje('mensaje-venta');
    pintarVenta();
  }

  function pintarProductos() {
    var cont = $('productos');
    cont.textContent = '';
    var activos = estado.productos.filter(function (p) { return p.activo; });
    if (!activos.length) {
      cont.appendChild(el('p', { clase: 'vacio', texto: 'No hay productos activos. Actívelos en Configuración.' }));
      return;
    }
    activos.forEach(function (p) {
      var linea = estado.carrito.find(function (l) { return l.productoId === p.id; });
      var boton = el('button', {
        clase: 'producto' + (linea ? ' seleccionado' : ''), type: 'button',
        'aria-label': 'Agregar ' + p.nombre,
        onclick: function () { cambiarCarrito(L.agregarAlCarrito(estado.carrito, p)); }
      }, [
        el('span', { clase: 'producto-nombre', texto: p.nombre }),
        el('span', { clase: 'producto-precio', texto: L.formatearColones(p.precio) }),
        el('span', { clase: 'producto-cantidad', texto: linea ? '× ' + linea.cantidad : '+ Agregar' })
      ]);
      boton.disabled = carritoBloqueado();
      cont.appendChild(boton);
    });
  }

  function pintarCarrito() {
    var bloqueado = carritoBloqueado();
    var cuerpo = $('carrito-lineas');
    cuerpo.textContent = '';
    estado.carrito.forEach(function (l) {
      function boton(texto, etiqueta, accion, clase) {
        var b = el('button', { clase: 'mini ' + (clase || ''), type: 'button', 'aria-label': etiqueta, texto: texto, onclick: accion });
        b.disabled = bloqueado;
        return b;
      }
      cuerpo.appendChild(el('tr', {}, [
        el('td', {}, [
          document.createTextNode(l.nombre),
          el('small', { clase: 'precio-unidad', texto: L.formatearColones(l.precio) + ' c/u' })
        ]),
        el('td', { clase: 'num col-precio', texto: L.formatearColones(l.precio) }),
        el('td', { clase: 'centro' }, [el('div', { clase: 'contador' }, [
          boton('−', 'Quitar uno', function () { cambiarCarrito(L.cambiarCantidad(estado.carrito, l.productoId, -1)); }),
          el('span', { texto: String(l.cantidad) }),
          boton('+', 'Agregar uno', function () { cambiarCarrito(L.cambiarCantidad(estado.carrito, l.productoId, 1)); })
        ])]),
        el('td', { clase: 'num', texto: L.formatearColones(L.subtotalLinea(l)) }),
        el('td', { clase: 'centro' }, [
          boton('✕', 'Eliminar ' + l.nombre, function () { cambiarCarrito(L.quitarDelCarrito(estado.carrito, l.productoId)); }, 'mini-quitar')
        ])
      ]));
    });
    var hay = estado.carrito.length > 0;
    $('carrito-tabla').classList.toggle('oculto', !hay);
    $('carrito-vacio').classList.toggle('oculto', hay);
    var total = L.formatearColones(L.totalCarrito(estado.carrito));
    $('carrito-total').textContent = total;
    $('barra-total-monto').textContent = total;
    document.body.classList.toggle('con-carrito', hay);

    $('cliente').disabled = bloqueado;
    var enviando = estado.intento && estado.intento.estado === 'enviando';
    var registrar = $('btn-registrar');
    registrar.disabled = bloqueado;
    registrar.textContent = enviando ? 'ENVIANDO…' : 'REGISTRAR VENTA';
    registrar.classList.toggle('oculto', !!(estado.intento && estado.intento.estado === 'incierta'));
    $('acciones-incierta').classList.toggle('oculto', !(estado.intento && estado.intento.estado === 'incierta'));
  }

  function pintarVenta() {
    pintarProductos();
    pintarCarrito();
  }

  function registrarVenta() {
    if (carritoBloqueado()) return;                 // protección contra doble toque
    var errores = L.validarCarrito(estado.carrito);
    if (errores.length) {
      mostrarMensaje('mensaje-venta', 'error', errores.join('\n'));
      return;
    }
    var venta = L.construirVenta(estado.carrito, $('cliente').value);
    estado.intento = { venta: venta, estado: 'enviando' };
    // Se guarda ANTES de enviar: si la página se cierra, la venta no se pierde
    guardar(CLAVES.intento, estado.intento);
    enviarIntento();
  }

  function enviarIntento() {
    var intento = estado.intento;
    if (!intento) return;
    intento.estado = 'enviando';
    guardar(CLAVES.intento, intento);
    mostrarMensaje('mensaje-venta', 'info', 'Enviando la venta a Google Sheets… No cierre la página.');
    pintarVenta();

    llamar('registrarVenta', { venta: intento.venta })
      .then(function (r) {
        // Confirmado por el servidor: recién ahora se limpia el carrito
        limpiarVentaActual();
        mostrarMensaje('mensaje-venta', 'exito',
          (r.duplicada ? 'La venta ya estaba guardada (no se duplicó).' : '✔ Venta registrada correctamente.') +
          '\nID: ' + intento.venta.id + ' · Total: ' + L.formatearColones(intento.venta.total));
      })
      .catch(function (e) {
        if (e instanceof ErrorServidor || e instanceof ErrorConfig) {
          // El servidor dijo claramente que NO guardó: se libera el carrito para corregir
          estado.intento = null;
          guardar(CLAVES.intento, null);
          mostrarMensaje('mensaje-venta', 'error',
            'La venta NO se guardó. ' + e.message + '\nLos productos siguen en el carrito; corrija y vuelva a presionar REGISTRAR VENTA.');
        } else {
          intento.estado = 'incierta';
          guardar(CLAVES.intento, intento);
          mostrarMensaje('mensaje-venta', 'alerta',
            'La venta NO ha sido confirmada. ' + e.message +
            '\nNo la cuente como registrada todavía. Puede reintentar (no se duplicará) o dejarla pendiente y atender al siguiente cliente.');
        }
        pintarVenta();
      });
  }

  function limpiarVentaActual() {
    estado.intento = null;
    estado.carrito = [];
    $('cliente').value = '';
    guardar(CLAVES.intento, null);
    guardar(CLAVES.carrito, []);
    guardar(CLAVES.cliente, '');
    pintarVenta();
  }

  function dejarPendiente() {
    var intento = estado.intento;
    if (!intento) return;
    estado.pendientes.push({ venta: intento.venta, error: 'Sin confirmar', rechazada: false });
    if (!guardar(CLAVES.pendientes, estado.pendientes)) { estado.pendientes.pop(); return; }
    limpiarVentaActual();
    mostrarMensaje('mensaje-venta', 'alerta',
      'La venta ' + intento.venta.id + ' quedó PENDIENTE (todavía no confirmada). Se enviará automáticamente cuando haya Internet.');
    pintarPendientes();
    enviarPendientes();
  }

  function corregirProductos() {
    var intento = estado.intento;
    if (!intento) return;
    mostrarMensaje('mensaje-venta', 'info', 'Consultando con el servidor si la venta llegó…');
    llamar('descartarIntento', { id: intento.venta.id })
      .then(function (r) {
        if (r.yaConfirmada) {
          limpiarVentaActual();
          mostrarMensaje('mensaje-venta', 'exito',
            'La venta ' + intento.venta.id + ' sí había quedado guardada (' + L.formatearColones(intento.venta.total) + '). Ya no se puede modificar.');
        } else {
          estado.intento = null;
          guardar(CLAVES.intento, null);
          mostrarMensaje('mensaje-venta', 'info', 'Confirmado: esa venta no se guardó. Ya puede corregir los productos.');
          pintarVenta();
        }
      })
      .catch(function (e) {
        mostrarMensaje('mensaje-venta', 'alerta',
          'No se pudo consultar al servidor: ' + e.message + '\nPara corregir se necesita conexión. Puede reintentar o dejarla pendiente.');
      });
  }

  /* ------------------------------------------------------------------ */
  /* Ventas pendientes                                                  */
  /* ------------------------------------------------------------------ */
  function pintarPendientes() {
    var n = estado.pendientes.length;
    var boton = $('btn-pendientes');
    boton.classList.toggle('oculto', n === 0);
    boton.textContent = '⚠ ' + n + (n === 1 ? ' venta pendiente' : ' ventas pendientes');

    var lista = $('pendientes-lista');
    lista.textContent = '';
    estado.pendientes.forEach(function (p, i) {
      var v = p.venta;
      var item = el('li', { clase: p.rechazada ? 'rechazada' : '' }, [
        el('strong', { texto: v.id }),
        el('span', { texto: L.fechaLegible(L.fechaHoyCR(new Date(v.fechaHora))) + ' · ' + L.formatearColones(v.total) + (v.cliente ? ' · ' + v.cliente : '') }),
        el('small', { texto: p.rechazada ? 'Rechazada por el servidor: ' + p.error : 'Estado: ' + p.error })
      ]);
      if (p.rechazada) {
        // Solo se puede quitar una venta que el servidor rechazó (seguro que no se guardó)
        item.appendChild(el('button', {
          clase: 'btn btn-texto', type: 'button', texto: 'Quitar de la lista',
          onclick: function () {
            if (confirm('El servidor rechazó la venta ' + v.id + ', así que NO está guardada.\n¿Quitarla de la lista? Si fue una venta real, regístrela de nuevo.')) {
              estado.pendientes.splice(i, 1);
              guardar(CLAVES.pendientes, estado.pendientes);
              pintarPendientes();
            }
          }
        }));
      }
      lista.appendChild(item);
    });
    if (n === 0) lista.appendChild(el('li', { texto: 'No hay ventas pendientes.' }));
  }

  /** Reintenta cada pendiente con su ID original (el servidor evita duplicados). */
  function enviarPendientes(manual) {
    if (estado.enviandoPendientes) return Promise.resolve();
    var porEnviar = estado.pendientes.filter(function (p) { return !p.rechazada; });
    if (!porEnviar.length) {
      if (manual) mostrarMensaje('pendientes-mensaje', 'info', 'No hay ventas para enviar.');
      return Promise.resolve();
    }
    estado.enviandoPendientes = true;
    if (manual) mostrarMensaje('pendientes-mensaje', 'info', 'Enviando…');
    var confirmadas = 0;
    var sinConexion = null;

    var cadena = Promise.resolve();
    porEnviar.forEach(function (p) {
      cadena = cadena.then(function () {
        if (sinConexion) return;
        return llamar('registrarVenta', { venta: p.venta })
          .then(function () {
            estado.pendientes = estado.pendientes.filter(function (x) { return x.venta.id !== p.venta.id; });
            confirmadas++;
          })
          .catch(function (e) {
            if (e instanceof ErrorServidor && e.codigo !== 'OCUPADO') { p.rechazada = true; p.error = e.message; }
            else { p.error = e.message; sinConexion = e; }
          })
          .then(function () { guardar(CLAVES.pendientes, estado.pendientes); });
      });
    });

    return cadena.then(function () {
      estado.enviandoPendientes = false;
      pintarPendientes();
      var quedan = estado.pendientes.length;
      var texto = confirmadas + ' venta(s) confirmada(s). ' +
        (quedan ? quedan + ' siguen pendientes.' + (sinConexion ? ' ' + sinConexion.message : '') : 'No quedan pendientes.');
      if (manual || confirmadas) mostrarMensaje('pendientes-mensaje', quedan ? 'alerta' : 'exito', texto);
      if (confirmadas && !estado.intento) mostrarMensaje('mensaje-venta', quedan ? 'alerta' : 'exito', 'Ventas pendientes: ' + texto);
    });
  }

  /* ------------------------------------------------------------------ */
  /* Historial                                                          */
  /* ------------------------------------------------------------------ */
  function consultarDia(fecha) {
    return llamar('ventasPorFecha', { fecha: fecha }).then(function (r) { return r.ventas; });
  }

  function mensajeConsulta(e) {
    if (e instanceof ErrorConfig) return e.message;
    return 'No se pudieron obtener las ventas. ' + e.message;
  }

  function cargarHistorial() {
    var fecha = $('historial-fecha').value;
    if (!fecha) return mostrarMensaje('historial-mensaje', 'error', 'Seleccione una fecha.');
    mostrarMensaje('historial-mensaje', 'info', 'Consultando Google Sheets…');
    var cuerpo = $('historial-lineas');
    consultarDia(fecha).then(function (ventas) {
      estado.historial = ventas;
      cuerpo.textContent = '';
      ventas.forEach(function (v) {
        cuerpo.appendChild(el('tr', {}, [
          el('td', { clase: 'id', texto: v.id }),
          el('td', { texto: v.hora }),
          el('td', { texto: v.cliente || '—' }),
          el('td', { clase: 'num', texto: String(v.unidades) }),
          el('td', { clase: 'num', texto: L.formatearColones(v.total) }),
          el('td', {}, [el('button', { clase: 'btn btn-chico', type: 'button', texto: 'Ver detalle', onclick: function () { verDetalle(v); } })])
        ]));
      });
      var total = ventas.reduce(function (s, v) { return s + v.total; }, 0);
      mostrarMensaje('historial-mensaje', 'info', ventas.length
        ? ventas.length + ' venta(s) el ' + L.fechaLegible(fecha) + ' · Total ' + L.formatearColones(total)
        : 'No hay ventas registradas el ' + L.fechaLegible(fecha) + '.');
    }).catch(function (e) {
      cuerpo.textContent = '';
      mostrarMensaje('historial-mensaje', 'error', mensajeConsulta(e));
    });
  }

  function verDetalle(v) {
    $('detalle-titulo').textContent = 'Venta ' + v.id;
    $('detalle-info').textContent = L.fechaLegible(v.fecha) + ' ' + v.hora + (v.cliente ? ' · Cliente: ' + v.cliente : '');
    var cuerpo = $('detalle-lineas');
    cuerpo.textContent = '';
    v.detalles.forEach(function (d) {
      cuerpo.appendChild(el('tr', {}, [
        el('td', { texto: d.nombre }),
        el('td', { clase: 'num', texto: L.formatearColones(d.precio) }),
        el('td', { clase: 'num', texto: String(d.cantidad) }),
        el('td', { clase: 'num', texto: L.formatearColones(d.subtotal) })
      ]));
    });
    $('detalle-total').textContent = L.formatearColones(v.total);
    $('dialogo-detalle').showModal();
  }

  /* ------------------------------------------------------------------ */
  /* Resumen del día                                                    */
  /* ------------------------------------------------------------------ */
  function cargarResumen() {
    var fecha = $('resumen-fecha').value;
    if (!fecha) return mostrarMensaje('resumen-mensaje', 'error', 'Seleccione una fecha.');
    mostrarMensaje('resumen-mensaje', 'info', 'Consultando Google Sheets…');
    consultarDia(fecha).then(function (ventas) {
      var r = L.calcularResumen(ventas);
      $('r-total').textContent = L.formatearColones(r.totalVendido);
      $('r-ventas').textContent = String(r.cantidadVentas);
      $('r-unidades').textContent = String(r.unidades);
      $('r-promedio').textContent = L.formatearColones(r.promedio);
      var cuerpo = $('resumen-productos');
      cuerpo.textContent = '';
      r.productos.forEach(function (p) {
        cuerpo.appendChild(el('tr', {}, [
          el('td', { texto: p.nombre }),
          el('td', { clase: 'num', texto: String(p.unidades) }),
          el('td', { clase: 'num', texto: L.formatearColones(p.total) })
        ]));
      });
      if (!r.coincide) {
        mostrarMensaje('resumen-mensaje', 'alerta', 'Atención: el total de VENTAS no coincide con la suma de DETALLE_VENTAS. Revise la hoja.');
      } else {
        mostrarMensaje('resumen-mensaje', 'info', 'Resumen del ' + L.fechaLegible(fecha) + ' (solo ventas confirmadas).' +
          (estado.pendientes.length ? ' Hay ' + estado.pendientes.length + ' venta(s) pendientes que todavía no se cuentan.' : ''));
      }
    }).catch(function (e) {
      mostrarMensaje('resumen-mensaje', 'error', mensajeConsulta(e));
    });
  }

  /* ------------------------------------------------------------------ */
  /* Exportar a Excel                                                   */
  /* ------------------------------------------------------------------ */
  var URL_SHEETJS = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';

  function cargarSheetJS() {
    if (window.XLSX) return Promise.resolve(window.XLSX);
    return new Promise(function (ok, falla) {
      var s = document.createElement('script');
      s.src = URL_SHEETJS;
      s.onload = function () { window.XLSX ? ok(window.XLSX) : falla(new Error('sin XLSX')); };
      s.onerror = function () { falla(new Error('No se pudo cargar la librería de Excel')); };
      document.head.appendChild(s);
    });
  }

  function filasExportacion(ventas) {
    var resumen = [['Fecha', 'Hora', 'ID de venta', 'Cliente', 'Unidades', 'Total de la venta']];
    var detalle = [['Fecha', 'Hora', 'ID de venta', 'Cliente', 'Producto', 'Cantidad', 'Precio unitario', 'Subtotal', 'Total de la venta']];
    ventas.forEach(function (v) {
      resumen.push([v.fecha, v.hora, v.id, v.cliente, v.unidades, v.total]);
      v.detalles.forEach(function (d) {
        detalle.push([v.fecha, v.hora, v.id, v.cliente, d.nombre, d.cantidad, d.precio, d.subtotal, v.total]);
      });
    });
    var total = ventas.reduce(function (s, v) { return s + v.total; }, 0);
    resumen.push(['', '', '', 'TOTAL DEL DÍA', ventas.reduce(function (s, v) { return s + v.unidades; }, 0), total]);
    return { resumen: resumen, detalle: detalle };
  }

  function descargarCSV(filas, nombre) {
    var texto = filas.map(function (f) {
      return f.map(function (c) { return '"' + String(c).replace(/"/g, '""') + '"'; }).join(';');
    }).join('\r\n');
    var blob = new Blob(['﻿' + texto], { type: 'text/csv;charset=utf-8' });
    var a = el('a', { href: URL.createObjectURL(blob), download: nombre });
    document.body.appendChild(a);
    a.click();
    a.remove();
  }

  function exportarExcel() {
    var fecha = $('historial-fecha').value;
    if (!fecha) return mostrarMensaje('historial-mensaje', 'error', 'Seleccione una fecha para exportar.');
    var boton = $('btn-exportar');
    boton.disabled = true;
    mostrarMensaje('historial-mensaje', 'info', 'Preparando el archivo de Excel…');
    consultarDia(fecha)
      .then(function (ventas) {
        if (!ventas.length) throw new ErrorServidor('No hay ventas registradas el ' + L.fechaLegible(fecha) + '.');
        var filas = filasExportacion(ventas);
        return cargarSheetJS().then(function (XLSX) {
          var libro = XLSX.utils.book_new();
          var h1 = XLSX.utils.aoa_to_sheet(filas.resumen);
          var h2 = XLSX.utils.aoa_to_sheet(filas.detalle);
          h1['!cols'] = [{ wch: 11 }, { wch: 9 }, { wch: 26 }, { wch: 20 }, { wch: 9 }, { wch: 16 }];
          h2['!cols'] = [{ wch: 11 }, { wch: 9 }, { wch: 26 }, { wch: 20 }, { wch: 20 }, { wch: 9 }, { wch: 14 }, { wch: 12 }, { wch: 16 }];
          XLSX.utils.book_append_sheet(libro, h1, 'Ventas');
          XLSX.utils.book_append_sheet(libro, h2, 'Detalle');
          XLSX.writeFile(libro, 'ventas_' + fecha + '.xlsx');
          mostrarMensaje('historial-mensaje', 'exito', 'Archivo ventas_' + fecha + '.xlsx descargado (' + ventas.length + ' ventas).');
        }, function () {
          // Respaldo si no se pudo cargar la librería: CSV que Excel también abre
          descargarCSV(filas.detalle, 'ventas_' + fecha + '.csv');
          mostrarMensaje('historial-mensaje', 'alerta', 'No se pudo crear el .xlsx; se descargó ventas_' + fecha + '.csv, que Excel también abre.');
        });
      })
      .catch(function (e) { mostrarMensaje('historial-mensaje', 'error', mensajeConsulta(e)); })
      .finally(function () { boton.disabled = false; });
  }

  /* ------------------------------------------------------------------ */
  /* Configuración                                                      */
  /* ------------------------------------------------------------------ */
  function pintarConfig() {
    $('cfg-url').value = estado.config.url || '';
    $('cfg-clave').value = estado.config.clave || '';
    var cont = $('config-productos');
    cont.textContent = '';
    estado.productos.forEach(function (p) { cont.appendChild(filaProducto(p)); });
  }

  /** Fila editable de un producto. Sin id = producto nuevo todavía sin guardar. */
  function filaProducto(p) {
    var nuevo = !p.id;
    var activo = el('input', { type: 'checkbox', 'data-campo': 'activo' });
    activo.checked = p.activo;
    var fila = el('div', { clase: 'config-fila' + (nuevo ? ' config-fila-nueva' : ''), 'data-id': p.id || '' }, [
      el('label', { clase: 'campo' }, [el('span', { texto: nuevo ? 'Nombre (producto nuevo)' : 'Nombre' }), el('input', { type: 'text', maxlength: '40', value: p.nombre, 'data-campo': 'nombre', placeholder: 'Ej.: Gallo pinto' })]),
      el('label', { clase: 'campo' }, [el('span', { texto: 'Precio (₡)' }), el('input', { type: 'text', inputmode: 'numeric', value: p.precio ? String(p.precio) : '', 'data-campo': 'precio', placeholder: 'Ej.: 1500' })]),
      el('label', { clase: 'campo-check' }, [activo, el('span', { texto: 'Activo' })])
    ]);
    if (nuevo) {
      // Un producto nuevo se puede quitar antes de guardarlo; después solo se desactiva
      fila.appendChild(el('button', {
        clase: 'btn btn-texto', type: 'button', texto: 'Quitar',
        onclick: function () { fila.remove(); }
      }));
    }
    return fila;
  }

  function agregarProducto() {
    var fila = filaProducto({ id: '', nombre: '', precio: 0, activo: true });
    $('config-productos').appendChild(fila);
    ocultarMensaje('productos-mensaje');
    fila.querySelector('[data-campo="nombre"]').focus();
  }

  function leerFormularioProductos() {
    var errores = [];
    var nombres = {};
    var lista = Array.prototype.map.call(document.querySelectorAll('.config-fila'), function (fila) {
      var nombre = fila.querySelector('[data-campo="nombre"]').value.trim();
      var precio = L.leerPrecio(fila.querySelector('[data-campo="precio"]').value);
      if (!nombre) errores.push('Hay un producto sin nombre.');
      else if (nombres[nombre.toLowerCase()]) errores.push('Hay dos productos llamados "' + nombre + '".');
      nombres[nombre.toLowerCase()] = true;
      if (precio === null) errores.push('El precio de "' + (nombre || 'producto nuevo') + '" no es válido (use solo números enteros, mayor que 0).');
      return { id: fila.dataset.id || null, nombre: nombre, precio: precio, activo: fila.querySelector('[data-campo="activo"]').checked };
    });
    return { lista: lista, errores: errores };
  }

  /** Sin Google Sheets conectado, los productos nuevos reciben su ID aquí mismo. */
  function asignarIdsLocales(lista) {
    var siguiente = lista.reduce(function (max, p) {
      var m = /^P(\d+)$/.exec(p.id || '');
      return m ? Math.max(max, parseInt(m[1], 10)) : max;
    }, 0) + 1;
    return lista.map(function (p) {
      return p.id ? p : Object.assign({}, p, { id: 'P' + ('00' + siguiente++).slice(-3) });
    });
  }

  function aplicarProductos(lista) {
    estado.productos = lista;
    guardar(CLAVES.productos, lista);
    // Si hay una venta armándose (no enviada), se actualizan sus precios
    if (!carritoBloqueado()) {
      estado.carrito = estado.carrito
        .filter(function (l) { return lista.some(function (p) { return p.id === l.productoId && p.activo; }); })
        .map(function (l) {
          var p = lista.find(function (x) { return x.id === l.productoId; });
          return Object.assign({}, l, { nombre: p.nombre, precio: p.precio });
        });
      guardar(CLAVES.carrito, estado.carrito);
    }
    pintarVenta();
  }

  function guardarProductos() {
    var datos = leerFormularioProductos();
    if (datos.errores.length) return mostrarMensaje('productos-mensaje', 'error', datos.errores.join('\n'));
    if (!estado.config.url) {
      aplicarProductos(asignarIdsLocales(datos.lista));
      pintarConfig();
      return mostrarMensaje('productos-mensaje', 'alerta', 'Guardado solo en este dispositivo, porque Google Sheets todavía no está conectado.');
    }
    var boton = $('btn-guardar-productos');
    boton.disabled = true;
    mostrarMensaje('productos-mensaje', 'info', 'Guardando en Google Sheets…');
    llamar('guardarProductos', { productos: datos.lista })
      .then(function (r) {
        aplicarProductos(r.productos);
        pintarConfig();
        mostrarMensaje('productos-mensaje', 'exito', '✔ Productos guardados en Google Sheets.' +
          (r.creados ? ' Se agregó ' + r.creados + ' producto(s) nuevo(s); ya aparecen en Registrar venta.' : ''));
      })
      .catch(function (e) {
        var texto = 'Los cambios NO se guardaron. ' + e.message;
        // Codigo.gs anterior a esta versión no sabe crear productos
        if (/sin identificador/i.test(e.message)) {
          texto += '\nPara agregar productos hay que actualizar Codigo.gs en Apps Script y publicar una "Nueva versión" (ver instrucciones.txt, parte 11).';
        }
        mostrarMensaje('productos-mensaje', 'error', texto);
      })
      .finally(function () { boton.disabled = false; });
  }

  function cargarProductosDelServidor() {
    if (!estado.config.url) return Promise.resolve();
    return llamar('obtenerProductos').then(function (r) {
      if (Array.isArray(r.productos) && r.productos.length) aplicarProductos(r.productos);
    }).catch(function () { /* se usan los productos guardados en el dispositivo */ });
  }

  function guardarConexion() {
    var url = $('cfg-url').value.trim();
    var clave = $('cfg-clave').value;
    if (!/^https:\/\/script\.google\.com\/macros\/s\/[\w-]+\/exec$/.test(url)) {
      return mostrarMensaje('conexion-mensaje', 'error', 'La URL debe tener la forma https://script.google.com/macros/s/…/exec');
    }
    if (!clave) return mostrarMensaje('conexion-mensaje', 'error', 'Escriba la clave de acceso.');
    estado.config = { url: url, clave: clave };
    guardar(CLAVES.config, estado.config);
    pintarAvisoConfig();
    mostrarMensaje('conexion-mensaje', 'info', 'Probando conexión…');
    llamar('probar')
      .then(function () {
        mostrarMensaje('conexion-mensaje', 'exito', '✔ Conexión correcta con Google Sheets.');
        return cargarProductosDelServidor().then(pintarConfig).then(function () { enviarPendientes(); });
      })
      .catch(function (e) {
        mostrarMensaje('conexion-mensaje', 'error', 'La conexión falló: ' + e.message);
      });
  }

  function pintarAvisoConfig() {
    $('aviso-config').classList.toggle('oculto', !!(estado.config.url && estado.config.clave));
  }

  /* ------------------------------------------------------------------ */
  /* Inicio                                                             */
  /* ------------------------------------------------------------------ */
  function iniciar() {
    document.querySelectorAll('.menu-btn').forEach(function (b) {
      b.addEventListener('click', function () { irA(b.dataset.seccion); });
    });
    $('btn-registrar').addEventListener('click', registrarVenta);
    $('btn-reintentar').addEventListener('click', enviarIntento);
    $('btn-a-pendientes').addEventListener('click', dejarPendiente);
    $('btn-corregir').addEventListener('click', corregirProductos);
    $('btn-historial').addEventListener('click', cargarHistorial);
    $('btn-resumen').addEventListener('click', cargarResumen);
    $('btn-exportar').addEventListener('click', exportarExcel);
    $('btn-guardar-productos').addEventListener('click', guardarProductos);
    $('btn-agregar-producto').addEventListener('click', agregarProducto);
    $('btn-guardar-conexion').addEventListener('click', guardarConexion);
    $('btn-pendientes').addEventListener('click', function () { ocultarMensaje('pendientes-mensaje'); $('dialogo-pendientes').showModal(); });
    $('btn-enviar-pendientes').addEventListener('click', function () { enviarPendientes(true); });
    $('ir-carrito').addEventListener('click', function (ev) { ev.preventDefault(); document.querySelector('.carrito').scrollIntoView(); });
    $('cliente').value = leer(CLAVES.cliente, '');
    $('cliente').addEventListener('input', function () { guardar(CLAVES.cliente, $('cliente').value); });

    var hoy = L.fechaHoyCR();
    $('historial-fecha').value = hoy;
    $('resumen-fecha').value = hoy;
    document.body.dataset.seccion = 'venta';

    pintarAvisoConfig();
    pintarPendientes();
    pintarVenta();

    // Si la página se cerró mientras se enviaba una venta, no se sabe si llegó
    if (estado.intento) {
      estado.intento.estado = 'incierta';
      guardar(CLAVES.intento, estado.intento);
      pintarVenta();
      mostrarMensaje('mensaje-venta', 'alerta',
        'La venta ' + estado.intento.venta.id + ' se estaba enviando cuando se cerró la página y NO está confirmada.\n' +
        'Presione "Reintentar envío" (no se duplicará) o déjela pendiente.');
    }

    cargarProductosDelServidor().then(function () { enviarPendientes(); });
    window.addEventListener('online', function () { enviarPendientes(); });
    setInterval(function () { if (estado.pendientes.length) enviarPendientes(); }, 60000);

    // Permite abrir la página sin Internet si ya se había abierto antes
    if ('serviceWorker' in navigator && (location.protocol === 'https:' || location.hostname === 'localhost')) {
      navigator.serviceWorker.register('sw.js').catch(function () {});
    }
  }

  document.addEventListener('DOMContentLoaded', iniciar);
})();
