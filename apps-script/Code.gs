/**
 * Desayunos Diana — puente entre la app y la planilla de Google.
 *
 * CÓMO INSTALAR (una sola vez):
 *  1. Crea una planilla nueva en Google Sheets (ej: "Desayunos Diana - Datos").
 *  2. Menú Extensiones > Apps Script. Borra lo que aparece y pega TODO este archivo.
 *  3. Cambia la CLAVE de abajo por una que solo sepan ustedes.
 *  4. Guarda (ícono del disquete). Luego: Implementar > Nueva implementación >
 *     tipo "Aplicación web" · Ejecutar como: "Yo" · Quién tiene acceso: "Cualquier usuario".
 *  5. Autoriza los permisos y copia la URL que termina en /exec. Esa URL va en la app.
 *
 * Si después cambias este código: Implementar > Administrar implementaciones >
 * lápiz > Versión: "Nueva versión" > Implementar (así la URL sigue siendo la misma).
 */

const CLAVE = 'cambia-esta-clave';

const H_PEDIDOS = ['id', 'fecha', 'num', 'cliente', 'detalle', 'total', 'pagado', 'pendiente',
  'estado', 'tomado_por', 'nota', 'actualizado', 'datos'];
const H_CARTA = ['id', 'nombre', 'precio', 'categoria'];

// ---------- Entrada ----------
function doGet(e) {
  return run_(function () {
    const p = (e && e.parameter) || {};
    checkKey_(p.key);
    const since = Number(p.since || 0);
    const now = Date.now();
    const orders = readOrders_().filter(function (r) { return r.ts > since; }).map(function (r) { return r.o; });
    // La carta se envía siempre (es corta), así también sirven los cambios hechos a mano en la hoja
    return { ok: true, now: now, orders: orders, menu: readMenu_() };
  });
}

function doPost(e) {
  return run_(function () {
    const b = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    checkKey_(b.key);
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      switch (b.action) {
        case 'saveOrder': return { ok: true, order: saveOrder_(b.order) };
        case 'setDone': return { ok: true, order: patch_(b.id, function (o) { o.done = !!b.done; if (o.done) o.doneAt = b.doneAt || Date.now(); else delete o.doneAt; }) };
        case 'addPayment': return { ok: true, order: patch_(b.id, function (o) {
          o.payments = o.payments || [];
          if (!o.payments.some(function (x) { return x.id === b.payment.id; })) o.payments.push(b.payment);
        }) };
        case 'delPayment': return { ok: true, order: patch_(b.id, function (o) {
          o.payments = (o.payments || []).filter(function (x) { return x.id !== b.payId; });
        }) };
        case 'delOrder': return { ok: true, order: patch_(b.id, function (o) { o.deleted = true; }) };
        case 'setMenu': writeMenu_(b.menu); return { ok: true };
        case 'import': return { ok: true, added: import_(b.orders || [], b.menu) };
        default: throw new Error('Acción desconocida');
      }
    } finally {
      lock.releaseLock();
    }
  });
}

// ---------- Pedidos ----------
function sheet_(name, headers) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
    if (name === 'Pedidos') {
      sh.getRange('A:B').setNumberFormat('@');
      sh.getRange('J:K').setNumberFormat('@');
    }
  }
  return sh;
}

function readOrders_() {
  const sh = sheet_('Pedidos', H_PEDIDOS);
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  const vals = sh.getRange(2, 1, n, H_PEDIDOS.length).getValues();
  const out = [];
  vals.forEach(function (v, i) {
    if (!v[0]) return;
    try { out.push({ row: i + 2, ts: Number(v[11]) || 0, o: JSON.parse(v[12]) }); } catch (err) { }
  });
  return out;
}

function rowOf_(o, ts) {
  const items = o.items || [];
  const total = items.reduce(function (s, i) { return s + i.price * i.qty; }, 0);
  const paid = (o.payments || []).reduce(function (s, p) { return s + p.amount; }, 0);
  const pend = Math.max(total - paid, 0);
  o.pend = pend;
  const estado = o.deleted ? 'Borrado' : (o.done ? 'Entregado' : 'En preparación');
  const detalle = items.map(function (i) { return i.qty + '× ' + i.name; }).join(', ');
  return [o.id, o.date, o.num, o.client, detalle, total, paid, pend, estado, o.by || '', o.note || '', ts, JSON.stringify(o)];
}

function writeRow_(sh, row, o) {
  const ts = Date.now();
  const values = rowOf_(o, ts);
  if (row) sh.getRange(row, 1, 1, values.length).setValues([values]);
  else sh.appendRow(values);
  return o;
}

function nextNum_(rows, date) {
  return rows.reduce(function (m, r) {
    return (r.o.date === date && !r.o.deleted) ? Math.max(m, Number(r.o.num) || 0) : m;
  }, 0) + 1;
}

function saveOrder_(o) {
  if (!o || !o.id || !o.date) throw new Error('Pedido inválido');
  const sh = sheet_('Pedidos', H_PEDIDOS);
  const rows = readOrders_();
  const found = rows.filter(function (r) { return r.o.id === o.id; })[0];
  if (!found) {
    // Pedido nuevo: el número lo asigna la planilla para que nunca se repita
    o.num = nextNum_(rows, o.date);
    o.payments = o.payments || [];
    o.created = o.created || Date.now();
    return writeRow_(sh, 0, o);
  }
  // Pedido existente: solo se cambian los datos que se editan en el formulario
  const cur = found.o;
  cur.client = o.client;
  cur.items = o.items;
  cur.note = o.note;
  return writeRow_(sh, found.row, cur);
}

function patch_(id, fn) {
  const sh = sheet_('Pedidos', H_PEDIDOS);
  const found = readOrders_().filter(function (r) { return r.o.id === id; })[0];
  if (!found) return null;
  fn(found.o);
  return writeRow_(sh, found.row, found.o);
}

function import_(orders, menu) {
  const sh = sheet_('Pedidos', H_PEDIDOS);
  const have = {};
  readOrders_().forEach(function (r) { have[r.o.id] = 1; });
  const ts = Date.now();
  const rows = [];
  orders.forEach(function (o) {
    if (!o || !o.id || !o.date || have[o.id]) return;
    have[o.id] = 1;
    o.payments = o.payments || [];
    rows.push(rowOf_(o, ts));
  });
  if (rows.length) sh.getRange(sh.getLastRow() + 1, 1, rows.length, H_PEDIDOS.length).setValues(rows);
  if (menu && menu.items && menu.items.length && !readMenu_().items.length) writeMenu_(menu);
  return rows.length;
}

// ---------- Carta ----------
function readMenu_() {
  const sh = sheet_('Carta', H_CARTA);
  const n = sh.getLastRow() - 1;
  if (n < 1) return { items: [] };
  const items = sh.getRange(2, 1, n, 4).getValues()
    .map(function (v, i) { return v[1] ? v : null; })
    .map(function (v, i) { return v && { id: String(v[0] || ('fila' + (i + 2))), name: String(v[1]), price: Number(v[2]) || 0, cat: String(v[3] || '') }; })
    .filter(function (p) { return p; });
  return { items: items };
}

function writeMenu_(menu) {
  const sh = sheet_('Carta', H_CARTA);
  const n = sh.getLastRow() - 1;
  if (n > 0) sh.getRange(2, 1, n, 4).clearContent();
  const items = (menu && menu.items) || [];
  if (items.length) sh.getRange(2, 1, items.length, 4).setValues(items.map(function (p) { return [p.id, p.name, p.price, p.cat || '']; }));
}

// ---------- Utilidades ----------
function checkKey_(k) {
  if (String(k || '') !== CLAVE) throw new Error('Clave incorrecta');
}

function run_(fn) {
  let out;
  try { out = fn(); } catch (err) { out = { ok: false, error: String(err && err.message || err) }; }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}
