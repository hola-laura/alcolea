// Conexión de Presupuestos Alcolea con Google Drive.
// Pegar en la hoja "Presupuestos Alcolea": Extensiones > Apps Script.
const CARPETA = 'PON_AQUI_LA_CARPETA';
const CLAVE = 'PON_AQUI_LA_CLAVE';
const ARCHIVO_DATOS = 'alcolea-datos.json';
const CARPETA_PDF = 'Presupuestos en PDF';

function responder(datos) {
  return ContentService.createTextOutput(JSON.stringify(datos)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  if (e.parameter.clave !== CLAVE) return responder({ ok: false, error: 'clave' });
  const archivo = buscar(DriveApp.getFolderById(CARPETA), ARCHIVO_DATOS);
  return responder({ ok: true, estado: archivo ? JSON.parse(archivo.getBlob().getDataAsString()) : null });
}

function doPost(e) {
  const pedido = JSON.parse(e.postData.contents);
  if (pedido.clave !== CLAVE) return responder({ ok: false, error: 'clave' });
  const cerrojo = LockService.getScriptLock();
  cerrojo.waitLock(20000);
  try {
    if (pedido.accion === 'guardar') return responder(guardar(pedido.estado));
    if (pedido.accion === 'pdf') return responder(subirPdf(pedido.nombre, pedido.base64));
    return responder({ ok: false, error: 'accion' });
  } finally {
    cerrojo.releaseLock();
  }
}

function buscar(carpeta, nombre) {
  const archivos = carpeta.getFilesByName(nombre);
  return archivos.hasNext() ? archivos.next() : null;
}

function guardar(estado) {
  const carpeta = DriveApp.getFolderById(CARPETA);
  const texto = JSON.stringify(estado);
  const archivo = buscar(carpeta, ARCHIVO_DATOS);
  if (archivo) archivo.setContent(texto);
  else carpeta.createFile(ARCHIVO_DATOS, texto, 'application/json');

  const presupuestos = (estado.quotes || []).filter(function (q) { return !q.ejemplo; });
  const filas = presupuestos.map(function (q) {
    const base = q.lineas.reduce(function (s, l) { return s + (Number(l.cantidad) || 0) * (Number(l.precio) || 0); }, 0);
    const iva = base * (Number(q.iva) || 0) / 100;
    return [q.num, q.fecha, q.cliente.nombre, q.cliente.telefono, q.cliente.direccion, q.titulo, base, q.iva, iva, base + iva, q.estado, q.pdfUrl || ''];
  });
  const lineas = [];
  presupuestos.forEach(function (q) {
    q.lineas.forEach(function (l) {
      lineas.push([q.num, q.cliente.nombre, l.nombre, l.cantidad, l.unidad, l.precio, (Number(l.cantidad) || 0) * (Number(l.precio) || 0)]);
    });
  });
  const precios = (estado.catalogo || []).map(function (i) { return [i.cat, i.nombre, i.unidad, i.precio]; });
  escribir('Presupuestos', ['Número', 'Fecha', 'Cliente', 'Teléfono', 'Dirección', 'Trabajo', 'Sin IVA', 'IVA %', 'IVA', 'Total', 'Estado', 'PDF'], filas);
  escribir('Líneas', ['Presupuesto', 'Cliente', 'Concepto', 'Cantidad', 'Unidad', 'Precio', 'Importe'], lineas);
  escribir('Precios', ['Categoría', 'Concepto', 'Unidad', 'Precio'], precios);
  return { ok: true };
}

function escribir(nombre, cabecera, filas) {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = libro.getSheetByName(nombre) || libro.insertSheet(nombre);
  hoja.clearContents();
  const todo = [cabecera].concat(filas);
  hoja.getRange(1, 1, todo.length, cabecera.length).setValues(todo);
  hoja.getRange(1, 1, 1, cabecera.length).setFontWeight('bold');
  hoja.setFrozenRows(1);
}

function subirPdf(nombre, base64) {
  const carpeta = DriveApp.getFolderById(CARPETA);
  const carpetas = carpeta.getFoldersByName(CARPETA_PDF);
  const destino = carpetas.hasNext() ? carpetas.next() : carpeta.createFolder(CARPETA_PDF);
  const anterior = buscar(destino, nombre);
  if (anterior) anterior.setTrashed(true);
  const blob = Utilities.newBlob(Utilities.base64Decode(base64), 'application/pdf', nombre);
  const archivo = destino.createFile(blob);
  archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return { ok: true, url: archivo.getUrl() };
}
