/* =====================================================================
   Usuarios.gs  —  Inicio de sesión (usuario y contraseña)

   - Las contraseñas NUNCA se guardan: se guarda una "huella" (hash)
     PBKDF2-HMAC-SHA256 con sal aleatoria propia de cada usuario.
   - Tras 5 intentos fallidos el usuario queda bloqueado 15 minutos.
   - Cada inicio de sesión recibe un token aleatorio. En el servidor
     solo se guarda la huella del token, con su fecha de vencimiento.
   - Los usuarios se crean desde la hoja de cálculo: menú
     "Sistema de ventas" > "Crear usuario" (ver instrucciones.txt).
   ===================================================================== */

var HOJA_USUARIOS = 'USUARIOS';
var ENCABEZADO_USUARIOS = ['Usuario', 'Rol', 'Estado', 'Huella de contraseña', 'Sal', 'Intentos fallidos', 'Bloqueado hasta'];

var ROL_ADMIN = 'ADMIN';     // vende, consulta y cambia productos
var ROL_CAJERO = 'CAJERO';   // vende y consulta

var ITERACIONES_HASH = 2000;
var INTENTOS_MAX = 5;
var BLOQUEO_MINUTOS = 15;
var FALLOS_GLOBALES_MAX = 30;          // en 10 minutos, sumando todos los usuarios
var SESION_HORAS = 12;
var CONTRASENA_MIN = 8;

/* ------------------------- Huellas (hash) ------------------------- */

function aBytes(texto) {
  return Utilities.newBlob(texto).getBytes();
}

function bytesAHex(bytes) {
  return bytes.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}

/** PBKDF2-HMAC-SHA256 (un bloque de 32 bytes). */
function pbkdf2(contrasena, salHex, iteraciones) {
  var clave = aBytes(contrasena);
  var sal = aBytes(salHex).concat([0, 0, 0, 1]);
  var u = Utilities.computeHmacSha256Signature(sal, clave);
  var t = u.slice();
  for (var i = 1; i < iteraciones; i++) {
    u = Utilities.computeHmacSha256Signature(u, clave);
    for (var j = 0; j < t.length; j++) t[j] = t[j] ^ u[j];
  }
  return bytesAHex(t);
}

function sha256Hex(texto) {
  return bytesAHex(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, texto, Utilities.Charset.UTF_8));
}

function aleatorioHex() {
  return sha256Hex(Utilities.getUuid() + Utilities.getUuid() + Date.now());
}

/** Compara sin cortar en la primera diferencia. */
function igualesSeguro(a, b) {
  if (a.length !== b.length) return false;
  var dif = 0;
  for (var i = 0; i < a.length; i++) dif |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return dif === 0;
}

/* --------------------------- Usuarios --------------------------- */

function hojaUsuarios() {
  var libro = SpreadsheetApp.getActiveSpreadsheet();
  var h = libro.getSheetByName(HOJA_USUARIOS);
  if (!h) h = libro.insertSheet(HOJA_USUARIOS);
  if (h.getLastRow() === 0) {
    h.getRange(1, 1, 1, ENCABEZADO_USUARIOS.length).setValues([ENCABEZADO_USUARIOS]).setFontWeight('bold');
    h.setFrozenRows(1);
    h.getRange('A:A').setNumberFormat('@');
  }
  return h;
}

function normalizarUsuario(u) {
  return String(u == null ? '' : u).trim().toLowerCase();
}

function validarNombreUsuario(u) {
  if (!/^[a-z0-9._-]{3,30}$/.test(u)) throw errorCon('VALIDACION', 'El usuario debe tener de 3 a 30 letras minúsculas, números, punto, guion o guion bajo (sin espacios ni tildes).');
}

function validarContrasena(c) {
  if (typeof c !== 'string' || c.length < CONTRASENA_MIN || c.length > 100) throw errorCon('VALIDACION', 'La contraseña debe tener al menos ' + CONTRASENA_MIN + ' caracteres.');
}

/** { fila, usuario, rol, estado, huella, sal, intentos, bloqueadoHasta } o null */
function buscarUsuario(usuario) {
  var h = hojaUsuarios();
  var fila = buscarFila(h, usuario);
  if (!fila) return null;
  var f = h.getRange(fila, 1, 1, 7).getValues()[0];
  return {
    fila: fila, usuario: String(f[0]), rol: String(f[1]).toUpperCase() === ROL_ADMIN ? ROL_ADMIN : ROL_CAJERO,
    estado: String(f[2]).toUpperCase(), huella: String(f[3]), sal: String(f[4]),
    intentos: Number(f[5]) || 0, bloqueadoHasta: Number(f[6]) || 0
  };
}

/** Crea un usuario o le cambia la contraseña/rol si ya existe. */
function guardarUsuario(usuario, contrasena, rol) {
  usuario = normalizarUsuario(usuario);
  validarNombreUsuario(usuario);
  validarContrasena(contrasena);
  rol = String(rol || '').trim().toUpperCase() === ROL_ADMIN ? ROL_ADMIN : ROL_CAJERO;
  return conBloqueo(function () {
    var h = hojaUsuarios();
    var sal = aleatorioHex().slice(0, 32);
    var fila = [usuario, rol, 'ACTIVO', pbkdf2(contrasena, sal, ITERACIONES_HASH), sal, 0, 0];
    var existente = buscarUsuario(usuario);
    h.getRange(existente ? existente.fila : h.getLastRow() + 1, 1, 1, 7).setValues([fila]);
    cerrarSesionesDe(usuario);   // un cambio de contraseña saca las sesiones abiertas
    SpreadsheetApp.flush();
    return { usuario: usuario, rol: rol, nuevo: !existente };
  });
}

/* --------------------------- Sesiones --------------------------- */

function propiedades() { return PropertiesService.getScriptProperties(); }

function iniciarSesion(usuarioTexto, contrasena) {
  var generico = errorCon('LOGIN', 'Usuario o contraseña incorrectos.');
  var cache = CacheService.getScriptCache();
  var fallos = Number(cache.get('fallos_globales')) || 0;
  if (fallos >= FALLOS_GLOBALES_MAX) throw errorCon('LIMITE', 'Demasiados intentos fallidos. Espere 10 minutos.');
  var usuario = normalizarUsuario(usuarioTexto);
  if (!usuario || typeof contrasena !== 'string' || contrasena.length > 100) throw generico;

  return conBloqueo(function () {
    var u = buscarUsuario(usuario);
    var ahora = Date.now();
    if (u && u.bloqueadoHasta > ahora) {
      throw errorCon('BLOQUEADO', 'Usuario bloqueado por intentos fallidos. Intente de nuevo en ' + Math.ceil((u.bloqueadoHasta - ahora) / 60000) + ' minuto(s).');
    }
    var correcta = u && u.estado === 'ACTIVO' && u.huella && igualesSeguro(pbkdf2(contrasena, u.sal, ITERACIONES_HASH), u.huella);
    var h = hojaUsuarios();
    if (!correcta) {
      cache.put('fallos_globales', String(fallos + 1), 600);
      if (u) {
        var intentos = u.intentos + 1;
        var hasta = intentos >= INTENTOS_MAX ? ahora + BLOQUEO_MINUTOS * 60000 : 0;
        h.getRange(u.fila, 6, 1, 2).setValues([[intentos >= INTENTOS_MAX ? 0 : intentos, hasta]]);
      }
      throw generico;   // mismo mensaje exista o no el usuario
    }
    if (u.intentos || u.bloqueadoHasta) h.getRange(u.fila, 6, 1, 2).setValues([[0, 0]]);

    limpiarSesionesVencidas();
    var token = aleatorioHex() + aleatorioHex();
    propiedades().setProperty('S_' + sha256Hex(token), JSON.stringify({ u: u.usuario, exp: ahora + SESION_HORAS * 3600000 }));
    return { ok: true, token: token, usuario: u.usuario, rol: u.rol, venceEn: SESION_HORAS * 3600000 };
  });
}

/** Devuelve { usuario, rol } o lanza SESION. Renueva la sesión si está por vencer. */
function verificarSesion(token) {
  var vencida = errorCon('SESION', 'La sesión venció o no es válida. Inicie sesión de nuevo.');
  if (typeof token !== 'string' || !/^[0-9a-f]{128}$/.test(token)) throw vencida;
  var clave = 'S_' + sha256Hex(token);
  var guardado = propiedades().getProperty(clave);
  if (!guardado) throw vencida;
  var s = JSON.parse(guardado);
  var ahora = Date.now();
  if (s.exp < ahora) { propiedades().deleteProperty(clave); throw vencida; }
  var u = buscarUsuario(s.u);
  if (!u || u.estado !== 'ACTIVO') { propiedades().deleteProperty(clave); throw vencida; }
  if (s.exp - ahora < (SESION_HORAS / 2) * 3600000) {
    s.exp = ahora + SESION_HORAS * 3600000;
    propiedades().setProperty(clave, JSON.stringify(s));
  }
  return { usuario: u.usuario, rol: u.rol };
}

function cerrarSesion(token) {
  if (typeof token === 'string' && /^[0-9a-f]{128}$/.test(token)) propiedades().deleteProperty('S_' + sha256Hex(token));
  return { ok: true };
}

function cerrarSesionesDe(usuario) {
  var todas = propiedades().getProperties();
  Object.keys(todas).forEach(function (k) {
    if (k.indexOf('S_') !== 0) return;
    try { if (!usuario || JSON.parse(todas[k]).u === usuario) propiedades().deleteProperty(k); }
    catch (e) { propiedades().deleteProperty(k); }
  });
}

function limpiarSesionesVencidas() {
  var todas = propiedades().getProperties();
  var ahora = Date.now();
  Object.keys(todas).forEach(function (k) {
    if (k.indexOf('S_') !== 0) return;
    try { if (JSON.parse(todas[k]).exp < ahora) propiedades().deleteProperty(k); }
    catch (e) { propiedades().deleteProperty(k); }
  });
}

/* ------------------ Menú dentro de Google Sheets ------------------ */
/* Solo lo ve quien abre la hoja (el dueño). No se usa desde la página. */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Sistema de ventas')
    .addItem('Crear usuario o cambiar contraseña', 'menuCrearUsuario')
    .addItem('Desbloquear usuario', 'menuDesbloquearUsuario')
    .addItem('Cerrar todas las sesiones', 'menuCerrarSesiones')
    .addToUi();
}

function preguntar(ui, titulo, texto) {
  var r = ui.prompt(titulo, texto, ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return null;
  return r.getResponseText();
}

function menuCrearUsuario() {
  var ui = SpreadsheetApp.getUi();
  prepararHojas();
  var usuario = preguntar(ui, 'Usuario', 'Nombre de usuario (minúsculas, sin espacios). Ej.: kendall');
  if (usuario === null) return;
  var contrasena = preguntar(ui, 'Contraseña', 'Contraseña de al menos ' + CONTRASENA_MIN + ' caracteres.\n(Se muestra al escribirla: asegúrese de que nadie esté mirando.)');
  if (contrasena === null) return;
  var rol = preguntar(ui, 'Rol', 'Escriba ADMIN (puede cambiar productos) o CAJERO (solo vende y consulta).');
  if (rol === null) return;
  try {
    var r = guardarUsuario(usuario, contrasena, rol);
    ui.alert((r.nuevo ? 'Usuario creado: ' : 'Contraseña actualizada: ') + r.usuario + ' (' + r.rol + ').');
  } catch (e) {
    ui.alert('No se guardó: ' + e.message);
  }
}

function menuDesbloquearUsuario() {
  var ui = SpreadsheetApp.getUi();
  var usuario = preguntar(ui, 'Desbloquear', 'Usuario a desbloquear:');
  if (usuario === null) return;
  var u = buscarUsuario(normalizarUsuario(usuario));
  if (!u) return ui.alert('No existe ese usuario.');
  hojaUsuarios().getRange(u.fila, 6, 1, 2).setValues([[0, 0]]);
  ui.alert('Usuario desbloqueado: ' + u.usuario);
}

function menuCerrarSesiones() {
  cerrarSesionesDe(null);
  SpreadsheetApp.getUi().alert('Se cerraron todas las sesiones. Todos deberán iniciar sesión otra vez.');
}
