/* Pruebas de los cálculos y validaciones (logica.js).
   En el navegador: abra pruebas/pruebas.html.
   Con Node.js (opcional): node pruebas/pruebas.js */
(function () {
  'use strict';
  var L = typeof module !== 'undefined' ? require('../logica.js') : window.Logica;
  var resultados = [];

  function prueba(nombre, fn) {
    try { fn(); resultados.push({ nombre: nombre, ok: true }); }
    catch (e) { resultados.push({ nombre: nombre, ok: false, error: e.message }); }
  }
  function igual(real, esperado, que) {
    if (real !== esperado) throw new Error((que || 'valor') + ': se esperaba ' + JSON.stringify(esperado) + ' y se obtuvo ' + JSON.stringify(real));
  }

  var P = {
    hamburguesa: { id: 'P001', nombre: 'Hamburguesa', precio: 1500 },
    papas: { id: 'P002', nombre: 'Papas doraditas', precio: 1000 },
    coca: { id: 'P003', nombre: 'Coca-Cola', precio: 800 }
  };
  function armar(lista) {
    return lista.reduce(function (c, p) { return L.agregarAlCarrito(c, p); }, []);
  }

  prueba('PRUEBA 1: 1 hamburguesa + 1 Coca-Cola = ₡2.300', function () {
    var c = armar([P.hamburguesa, P.coca]);
    igual(L.totalCarrito(c), 2300, 'total');
    igual(L.formatearColones(L.totalCarrito(c)), '₡2.300', 'formato');
    igual(L.validarCarrito(c).length, 0, 'errores');
  });

  prueba('PRUEBA 2: 2 hamburguesas + 1 papas + 2 Coca-Colas = ₡5.600', function () {
    var c = armar([P.hamburguesa, P.hamburguesa, P.papas, P.coca, P.coca]);
    igual(c.length, 3, 'líneas');
    igual(L.subtotalLinea(c[0]), 3000, 'subtotal hamburguesas');
    igual(L.subtotalLinea(c[1]), 1000, 'subtotal papas');
    igual(L.subtotalLinea(c[2]), 1600, 'subtotal Coca-Colas');
    igual(L.totalCarrito(c), 5600, 'total');
    igual(L.unidadesCarrito(c), 5, 'unidades');
    igual(L.formatearColones(5600), '₡5.600', 'formato');
  });

  prueba('PRUEBA 3: venta sin productos es rechazada', function () {
    igual(L.validarCarrito([]).length > 0, true, 'rechazo');
  });

  prueba('PRUEBA 4: cantidades y precios inválidos son rechazados', function () {
    var base = { productoId: 'P001', nombre: 'Hamburguesa', precio: 1500, cantidad: 1 };
    [
      { cantidad: 0 }, { cantidad: -1 }, { cantidad: 1.5 }, { cantidad: '2' }, { cantidad: 1000 },
      { precio: 0 }, { precio: -500 }, { precio: 1500.5 }, { precio: NaN }, { precio: '1500' }, { precio: 2000000 }
    ].forEach(function (cambio) {
      var linea = Object.assign({}, base, cambio);
      igual(L.validarCarrito([linea]).length > 0, true, 'rechazo de ' + JSON.stringify(cambio));
    });
    igual(L.validarCarrito([base, base]).length > 0, true, 'producto repetido');
  });

  prueba('Bajar cantidad a 0 elimina la línea (nunca queda 0 ni negativo)', function () {
    var c = armar([P.coca]);
    c = L.cambiarCantidad(c, 'P003', -1);
    igual(c.length, 0, 'líneas');
    c = L.cambiarCantidad(armar([P.coca]), 'P003', -5);
    igual(c.length, 0, 'líneas tras restar de más');
  });

  prueba('Quitar producto y actualizar total', function () {
    var c = armar([P.hamburguesa, P.coca]);
    c = L.quitarDelCarrito(c, 'P001');
    igual(L.totalCarrito(c), 800, 'total');
  });

  prueba('PRUEBA 7 (parte local): cambiar el precio no altera líneas ya agregadas', function () {
    var c = armar([P.hamburguesa]);
    var nuevo = Object.assign({}, P.hamburguesa, { precio: 2000 });
    igual(c[0].precio, 1500, 'precio guardado en la línea');
    var c2 = armar([nuevo]);
    igual(c2[0].precio, 2000, 'nueva venta con nuevo precio');
  });

  prueba('Lectura de precios escritos por el usuario', function () {
    igual(L.leerPrecio('1.500'), 1500); igual(L.leerPrecio('₡ 2000'), 2000);
    igual(L.leerPrecio('abc'), null); igual(L.leerPrecio('-5'), null); igual(L.leerPrecio('0'), null);
    igual(L.leerPrecio(''), null);
  });

  prueba('Identificadores únicos con formato V-AAAAMMDD-HHMMSS-XXXXXX', function () {
    var vistos = {};
    for (var i = 0; i < 5000; i++) {
      var id = L.generarIdVenta(new Date('2026-10-25T20:00:00Z'));
      if (!/^V-20261025-140000-[A-Z0-9]{6}$/.test(id)) throw new Error('formato ' + id);
      if (vistos[id]) throw new Error('repetido ' + id);
      vistos[id] = true;
    }
  });

  prueba('La venta enviada lleva total, unidades y fecha correctos', function () {
    var c = armar([P.hamburguesa, P.hamburguesa, P.papas, P.coca, P.coca]);
    var v = L.construirVenta(c, '  Don Carlos  ', new Date('2026-10-25T20:00:00Z'));
    igual(v.total, 5600); igual(v.unidades, 5); igual(v.cliente, 'Don Carlos');
    igual(v.items.length, 3); igual(v.fechaHora, '2026-10-25T20:00:00.000Z');
    igual(L.fechaHoyCR(new Date('2026-10-26T03:00:00Z')), '2026-10-25', 'fecha en Costa Rica (UTC-6)');
  });

  prueba('PRUEBA 9 (parte local): resumen del día y sin duplicados', function () {
    var ventas = [
      { id: 'A', total: 5600, unidades: 5, detalles: [
        { nombre: 'Hamburguesa', precio: 1500, cantidad: 2, subtotal: 3000 },
        { nombre: 'Papas doraditas', precio: 1000, cantidad: 1, subtotal: 1000 },
        { nombre: 'Coca-Cola', precio: 800, cantidad: 2, subtotal: 1600 }] },
      { id: 'B', total: 2300, unidades: 2, detalles: [
        { nombre: 'Hamburguesa', precio: 1500, cantidad: 1, subtotal: 1500 },
        { nombre: 'Coca-Cola', precio: 800, cantidad: 1, subtotal: 800 }] }
    ];
    var r = L.calcularResumen(ventas.concat([ventas[0]])); // la venta A repetida
    igual(r.totalVendido, 7900, 'total'); igual(r.cantidadVentas, 2, 'ventas');
    igual(r.unidades, 7, 'unidades'); igual(r.promedio, 3950, 'promedio');
    igual(r.productos[0].nombre, 'Hamburguesa'); igual(r.productos[0].total, 4500);
    igual(r.coincide, true, 'coincide con detalle');
  });

  /* ---- Mostrar resultados ---- */
  var fallidas = resultados.filter(function (r) { return !r.ok; });
  if (typeof module !== 'undefined') {
    resultados.forEach(function (r) { console.log((r.ok ? '✔ ' : '✘ ') + r.nombre + (r.ok ? '' : '  → ' + r.error)); });
    console.log('\n' + (resultados.length - fallidas.length) + ' de ' + resultados.length + ' pruebas pasaron.');
    if (fallidas.length) process.exitCode = 1;
  } else {
    var lista = document.getElementById('resultados');
    resultados.forEach(function (r) {
      var li = document.createElement('li');
      li.className = r.ok ? 'ok' : 'falla';
      li.textContent = (r.ok ? '✔ ' : '✘ ') + r.nombre + (r.ok ? '' : '  → ' + r.error);
      lista.appendChild(li);
    });
    document.getElementById('total').textContent =
      (resultados.length - fallidas.length) + ' de ' + resultados.length + ' pruebas pasaron.';
  }
})();
