// --- CONFIGURACIÓN E INICIO ---
function getSS() {
  const ssId = PropertiesService.getScriptProperties().getProperty('SS_ID');
  if (!ssId) return null;
  try { return SpreadsheetApp.openById(ssId); } catch(e) { return null; }
}

function doGet(e) {
  if (e && e.parameter && e.parameter.api === 'inventario') {
    const respuesta = apiObtenerInventarioMovil();
    return ContentService.createTextOutput(JSON.stringify(respuesta))
      .setMimeType(ContentService.MimeType.JSON);
  }
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Inventario DyC')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// --- API INVENTARIO ---
function apiObtenerInventarioMovil() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetInv = ss.getSheetByName('Inventario');
    if (!sheetInv) throw new Error("La hoja 'Inventario' no existe.");
    const datos = sheetInv.getDataRange().getValues();
    const marcasUnicas = new Set();
    const productos = [];
    for (let i = 1; i < datos.length; i++) {
      const row = datos[i];
      if (!row[0]) continue;
      const marca = row[2] ? String(row[2]).trim() : "Generico";
      marcasUnicas.add(marca);
      productos.push({
        id: row[0],
        desc: row[1],
        marca: marca,
        ref: row[3],
        ubica: row[4],
        precio: row[6],
        stock: row[7]
      });
    }
    return { exito: true, productos: productos, marcas: Array.from(marcasUnicas).sort() };
  } catch (error) {
    return { exito: false, error: error.message };
  }
}

// --- MOVIMIENTOS ---
function registrarMovimiento(datos) {
  const ss = getSS();
  const hInv = ss.getSheetByName('Inventario');
  const invData = hInv.getDataRange().getValues();
  const hMov = ss.getSheetByName('Movimientos');
  for (let i = 1; i < invData.length; i++) {
    if (invData[i][0].toString() === datos.id.toString()) {
      let stockActual = Number(invData[i][7]);
      let stockNuevo = (datos.tipo === 'Entrada') ? stockActual + Number(datos.cant) : stockActual - Number(datos.cant);
      if (stockNuevo < 0) return "❌ Error: Stock insuficiente.";
      hInv.getRange(i + 1, 8).setValue(stockNuevo);
      if (hMov) {
        hMov.appendRow([new Date(), datos.id, datos.tipo, datos.cant, stockActual, stockNuevo, datos.usuario || 'Admin', datos.motivo || 'Manual']);
      }
      return "✅ Movimiento registrado.";
    }
  }
  return "❌ ID de producto no encontrado.";
}

// --- PROCESAR VENTA CON PDF ---
function procesarVenta(datos) {
  const ss = getSS();
  const hVen = ss.getSheetByName('Ventas');
  const hInv = ss.getSheetByName('Inventario');
  const hPlantilla = ss.getSheetByName('PlantillaFactura');
  if (!ss || !hVen || !hInv || !hPlantilla) throw new Error("procesarVenta: faltan hojas requeridas.");

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);

  try {
    let ultimaFila = hVen.getLastRow();
    let nuevoId = (ultimaFila > 1) ? Number(hVen.getRange(ultimaFila, 1).getValue()) + 1 : 1;

    let totalCosto = 0;
    const invData = hInv.getDataRange().getValues();

    datos.productos.forEach(item => {
      let costoUnitario = 0;
      for (let i = 1; i < invData.length; i++) {
        if (String(invData[i][0]) === String(item.id)) {
          costoUnitario = Number(invData[i][5] || 0);
          break;
        }
      }
      totalCosto += (costoUnitario * item.cant);
      const movRes = registrarMovimiento({ id: item.id, tipo: 'Salida', cant: item.cant, usuario: 'Caja', motivo: `Factura #${nuevoId}` });
      if (typeof movRes === 'string' && movRes.startsWith('❌')) throw new Error(movRes);
    });

    hVen.appendRow([nuevoId, new Date(), datos.cliente.Nombre, datos.total, totalCosto, datos.total - totalCosto, JSON.stringify(datos.productos)]);

    hPlantilla.getRange("C5").setValue(datos.cliente.Nombre);
    hPlantilla.getRange("F4").setValue(nuevoId);
    hPlantilla.getRange("F5").setValue(new Date());
    hPlantilla.getRange("F25").setValue(datos.total);
    hPlantilla.getRange("B10:F24").clearContent();
    let filas = datos.productos.map(p => [p.desc, "", p.cant, p.precio, (p.cant * p.precio)]);
    if (filas.length > 0) hPlantilla.getRange(10, 2, filas.length, 5).setValues(filas);

    SpreadsheetApp.flush();

    const token = ScriptApp.getOAuthToken();
    const gid = hPlantilla.getSheetId();
    const url = `https://docs.google.com/spreadsheets/d/${ss.getId()}/export?format=pdf&gid=${gid}&size=A4&gridlines=false&portrait=true&fitw=true`;
    const res = UrlFetchApp.fetch(url, { headers: { 'Authorization': 'Bearer ' + token } });

    return { ssId: ss.getId(), gid: gid, numFactura: nuevoId, pdfBase64: Utilities.base64Encode(res.getBlob().getBytes()) };

  } finally {
    lock.releaseLock();
  }
}

// --- LIMPIAR PLANTILLA ---
function limpiarPlantilla() {
  const hPlantilla = getSS().getSheetByName('PlantillaFactura');
  if (!hPlantilla) return;
  hPlantilla.getRange("B10:F25").clearContent();
  hPlantilla.getRange("C5,F4,F5").clearContent();
}
