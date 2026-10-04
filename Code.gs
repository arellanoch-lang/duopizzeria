/**
 * PANEL DE PEDIDOS — Dúo Pizzería
 * -------------------------------------
 * Este script hace de "servidor" compartido para todos los dispositivos:
 * - Recibe los pedidos desde la web del cliente y los guarda en la hoja "Pedidos"
 * - Guarda la configuración del restaurante (teléfono, horario, pausa, carta,
 *   pizza destacada, contraseña de admin) en la hoja "Configuracion", para que sea la MISMA
 *   en todos los móviles/tablets/ordenadores que usen la web o el panel
 * - Reparte el número de pedido de forma centralizada, para que nunca se
 *   dupliquen aunque varios clientes pidan al mismo tiempo desde dispositivos distintos
 * - Protege los datos con dos claves:
 *     · Clave del panel (personal): ver pedidos y cambiar su estado.
 *     · Clave de Configuración (dueño): todo lo anterior, más pausa y configuración.
 *   Sin la clave no se entregan pedidos (nombres, teléfonos, direcciones) ni se cambia nada.
 *   La configuración pública que lee la web de clientes nunca incluye las claves.
 *
 * CÓMO INSTALARLO (una sola vez):
 * 1. Crea una Google Sheet nueva (sheets.new)
 * 2. Ve a Extensiones > Apps Script
 * 3. Borra todo el contenido de "Código.gs" y pega este archivo entero
 * 4. Guarda (icono de disco o Ctrl+S)
 * 5. Pulsa "Implementar" (Deploy) > "Nueva implementación"
 * 6. Tipo: "Aplicación web"
 * 7. Configuración:
 *      - Ejecutar como: Yo (tu cuenta)
 *      - Quién tiene acceso: Cualquier usuario
 * 8. Pulsa "Implementar" y autoriza los permisos que te pida Google
 * 9. Copia la URL que te da ("URL de la aplicación web")
 * 10. Pega esa URL en la constante SHEET_SCRIPT_URL al principio del
 *     archivo index.html (y también en panel-pedidos.html al conectar)
 *
 * Cada vez que cambies este código, tendrás que hacer una nueva
 * implementación (Implementar > Gestionar implementaciones > lápiz de editar > Nueva versión)
 * para que los cambios se apliquen a la URL ya existente.
 */

const SHEET_NAME = "Pedidos";
const HEADERS = ["Fecha y hora", "Nº Pedido", "Tipo", "Nombre", "Teléfono", "Dirección", "Código Postal", "Pedido", "Total", "Estado", "Hora Listo", "Hora Entregado"];

const CONFIG_SHEET_NAME = "Configuracion";
const ORDER_COUNTER_KEY = "nextOrderNumber";
const FIRST_ORDER_NUMBER = 1001;
// Clave de Configuración si todavía no se ha guardado ninguna en la hoja.
// Cámbiala desde el panel (Configuración) en cuanto lo instales.
const DEFAULT_ADMIN_PASSWORD = "duo2026";

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getConfigSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG_SHEET_NAME);
    sheet.appendRow(["Clave", "Valor"]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function upsertConfigRow(sheet, key, value) {
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  sheet.appendRow([key, value]);
}

function readConfigMap() {
  const sheet = getConfigSheet();
  const rows = sheet.getDataRange().getValues();
  const map = {};
  for (let i = 1; i < rows.length; i++) {
    map[rows[i][0]] = rows[i][1];
  }
  return map;
}

function getPasswords() {
  const map = readConfigMap();
  return {
    admin: map.adminPassword !== undefined ? String(map.adminPassword) : DEFAULT_ADMIN_PASSWORD,
    staff: map.staffPassword !== undefined ? String(map.staffPassword) : ""
  };
}

// Devuelve "admin", "staff" o null según la clave recibida.
// Si la clave de Configuración está vacía, no se pide clave para nada (como antes).
function roleFor(key) {
  const pw = getPasswords();
  const k = String(key === undefined || key === null ? "" : key);
  if (pw.admin === "" || k === pw.admin) return "admin";
  if (pw.staff !== "" && k === pw.staff) return "staff";
  return null;
}

// includeSecrets = true solo para el dueño ya identificado (pantalla de Configuración).
function loadConfigObject(includeSecrets) {
  const map = readConfigMap();

  const result = {};
  if (map.phone !== undefined) result.phone = String(map.phone);
  if (map.whatsappEnabled !== undefined) result.whatsappEnabled = String(map.whatsappEnabled) === "true";
  if (map.testMode !== undefined) result.testMode = String(map.testMode) === "true";
  if (includeSecrets) {
    const pw = getPasswords();
    result.adminPassword = pw.admin;
    result.staffPassword = pw.staff;
  }
  if (map.pauseActive !== undefined || map.pauseMessage !== undefined) {
    result.pause = {
      active: String(map.pauseActive) === "true",
      message: map.pauseMessage || ""
    };
  }
  if (map.schedule !== undefined) {
    try { result.schedule = JSON.parse(map.schedule); } catch (e) { /* ignorar si está corrupto */ }
  }
  if (map.menu !== undefined) {
    try { result.menu = JSON.parse(map.menu); } catch (e) { /* ignorar si está corrupto */ }
  }
  if (map.featuredId !== undefined || map.featuredLabel !== undefined) {
    result.featured = {
      id: String(map.featuredId || ""),
      label: String(map.featuredLabel || "")
    };
  }
  return result;
}

function saveConfigObject(cfg) {
  const sheet = getConfigSheet();
  if (cfg.phone !== undefined) upsertConfigRow(sheet, "phone", cfg.phone);
  if (cfg.whatsappEnabled !== undefined) upsertConfigRow(sheet, "whatsappEnabled", String(cfg.whatsappEnabled));
  if (cfg.testMode !== undefined) upsertConfigRow(sheet, "testMode", String(cfg.testMode));
  if (cfg.adminPassword !== undefined) upsertConfigRow(sheet, "adminPassword", cfg.adminPassword);
  if (cfg.staffPassword !== undefined) upsertConfigRow(sheet, "staffPassword", cfg.staffPassword);
  if (cfg.pause !== undefined) {
    upsertConfigRow(sheet, "pauseActive", String(cfg.pause.active));
    upsertConfigRow(sheet, "pauseMessage", cfg.pause.message || "");
  }
  if (cfg.schedule !== undefined) upsertConfigRow(sheet, "schedule", JSON.stringify(cfg.schedule));
  if (cfg.menu !== undefined) upsertConfigRow(sheet, "menu", JSON.stringify(cfg.menu));
  if (cfg.featured !== undefined) {
    upsertConfigRow(sheet, "featuredId", cfg.featured.id || "");
    upsertConfigRow(sheet, "featuredLabel", cfg.featured.label || "");
  }
}

function getNextOrderNumber() {
  const sheet = getConfigSheet();
  const rows = sheet.getDataRange().getValues();
  let rowIndex = -1;
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === ORDER_COUNTER_KEY) { rowIndex = i; break; }
  }
  let current;
  if (rowIndex === -1) {
    current = FIRST_ORDER_NUMBER;
    sheet.appendRow([ORDER_COUNTER_KEY, current + 1]);
  } else {
    current = parseInt(rows[rowIndex][1], 10) || FIRST_ORDER_NUMBER;
    sheet.getRange(rowIndex + 1, 2).setValue(current + 1);
  }
  return current;
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    if (data.action === "newOrder") {
      const lock = LockService.getScriptLock();
      lock.waitLock(10000);
      let orderNumber;
      try {
        orderNumber = getNextOrderNumber();
        const sheet = getSheet();
        sheet.appendRow([
          new Date(),
          orderNumber,
          data.type === "delivery" ? "Domicilio" : data.type === "table" ? "Mesa" : "Recoger",
          data.name,
          data.phone,
          data.address || "",
          data.postal || "",
          data.items,
          data.total,
          "Pendiente"
        ]);
      } finally {
        lock.releaseLock();
      }
      return jsonResponse({ success: true, orderNumber: orderNumber });
    }

    // ----- A partir de aquí, todo necesita clave -----
    const role = roleFor(data.key);

    if (data.action === "login") {
      if (!role) return jsonResponse({ success: false, error: "auth" });
      return jsonResponse({ success: true, role: role });
    }

    if (data.action === "listOrders") {
      if (!role) return jsonResponse({ success: false, error: "auth" });
      return jsonResponse({ success: true, orders: listOrders() });
    }

    if (data.action === "adminConfig") {
      if (role !== "admin") return jsonResponse({ success: false, error: "auth" });
      return jsonResponse({ success: true, config: loadConfigObject(true) });
    }

    if (data.action === "updateStatus") {
      if (!role) return jsonResponse({ success: false, error: "auth" });
      const sheet = getSheet();
      const rows = sheet.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][1]) === String(data.orderNumber)) {
          sheet.getRange(i + 1, 10).setValue(data.status);
          const now = new Date();
          if (data.status === "Listo" && !rows[i][10]) {
            sheet.getRange(i + 1, 11).setValue(now);
          }
          if (data.status === "Entregado" && !rows[i][11]) {
            sheet.getRange(i + 1, 12).setValue(now);
          }
          break;
        }
      }
      return jsonResponse({ success: true });
    }

    if (data.action === "saveConfig") {
      if (role !== "admin") return jsonResponse({ success: false, error: "auth" });
      saveConfigObject(data.config || {});
      return jsonResponse({ success: true });
    }

    return jsonResponse({ success: false, error: "Acción no reconocida" });
  } catch (err) {
    return jsonResponse({ success: false, error: err.toString() });
  }
}

function doGet(e) {
  if (e && e.parameter && e.parameter.action === "config") {
    return jsonResponse(loadConfigObject(false));
  }
  // Los pedidos ya no se entregan por GET: el panel los pide con su clave (listOrders).
  return jsonResponse({ success: false, error: "auth" });
}

function listOrders() {
  const sheet = getSheet();
  const rows = sheet.getDataRange().getValues();
  const orders = [];
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row[1]) continue; // fila vacía
    orders.push({
      timestamp: row[0],
      orderNumber: row[1],
      type: row[2],
      name: row[3],
      phone: row[4],
      address: row[5],
      postal: row[6],
      items: row[7],
      total: row[8],
      status: row[9],
      readyAt: row[10] || null,
      deliveredAt: row[11] || null
    });
  }
  return orders;
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
