/**
 * Test de Alma: cobro con Mercado Pago y envío del resultado por email.
 *
 * Qué hace:
 * 1. La web envía las respuestas aquí (accion "crear"). Se guarda una fila en la hoja
 *    "Ventas" y se crea un pago en Mercado Pago. La web redirige al comprador a pagar.
 * 2. Mercado Pago avisa aquí cuando hay un pago (webhook). Se comprueba el pago con la
 *    API de Mercado Pago, se marca la fila como pagada y se envía el PDF por email.
 *    Si es el Mensaje personal, además te llega un email con el WhatsApp del comprador.
 * 3. Al volver de Mercado Pago, la web pide aquí el texto completo (doGet, accion "resultado").
 *    Solo se entrega si el pago está aprobado: los textos no están en la web.
 *
 * Configuración (Proyecto > Configuración del proyecto > Propiedades de la secuencia de comandos):
 *   MP_ACCESS_TOKEN  Access token de producción de Mercado Pago (Tu negocio > Credenciales)
 *   SITIO_URL        Dirección de la web, por ejemplo https://diematarin.github.io/test-de-alma/
 *   AVISOS_EMAIL     Email donde quieres recibir los pedidos del Mensaje personal
 * El token nunca se escribe en el código ni en la web.
 */

const PRECIOS = {
  basico: { titulo: 'Test de Alma: Mensaje del alma', ars: 7500 },
  personalizado: { titulo: 'Test de Alma: Mensaje personal', ars: 29000 },
};
const HOJA = 'Ventas';
const COLUMNAS = ['id', 'fecha', 'producto', 'nombre', 'email', 'telefono', 'arquetipo_id', 'perdida', 'respuestas', 'estado', 'pago_id', 'enviado'];

function prop(k) {
  return PropertiesService.getScriptProperties().getProperty(k);
}

function hoja() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(HOJA);
  if (!sh) {
    sh = ss.insertSheet(HOJA);
    sh.appendRow(COLUMNAS);
  }
  return sh;
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData.contents || '{}'); } catch (err) {}
  if (body.accion === 'crear') return crearPago(body);
  // Aviso de Mercado Pago: llega como ?type=payment&data.id=123 o en el cuerpo
  const tipo = (e.parameter && (e.parameter.type || e.parameter.topic)) || body.type;
  const id = (e.parameter && (e.parameter['data.id'] || e.parameter.id)) || (body.data && body.data.id);
  if (tipo === 'payment' && id) procesarPago(String(id));
  return json({ ok: true });
}

function crearPago(b) {
  const prod = PRECIOS[b.producto];
  const arq = ARQUETIPOS.find(a => a.id === Number(b.arquetipo_id));
  if (!prod || !arq || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(b.email || '')) return json({ error: 'Datos incompletos' });
  const id = Utilities.getUuid();
  hoja().appendRow([id, new Date(), b.producto, b.nombre, b.email, b.telefono || '', arq.id, b.perdida ? 'si' : 'no',
    JSON.stringify(b.respuestas || {}), 'pendiente', '', '']);

  const sitio = prop('SITIO_URL');
  const pref = {
    items: [{ title: prod.titulo, quantity: 1, unit_price: prod.ars, currency_id: 'ARS' }],
    payer: { email: b.email, name: b.nombre },
    external_reference: id,
    back_urls: { success: sitio + '?pago=ok&ref=' + id, pending: sitio + '?pago=pendiente&ref=' + id, failure: sitio + '?pago=error&ref=' + id },
    auto_return: 'approved',
    notification_url: ScriptApp.getService().getUrl(),
  };
  const r = UrlFetchApp.fetch('https://api.mercadopago.com/checkout/preferences', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + prop('MP_ACCESS_TOKEN') },
    payload: JSON.stringify(pref),
  });
  const res = JSON.parse(r.getContentText());
  if (!res.init_point) return json({ error: 'Mercado Pago no creó el pago', detalle: res.message || '' });
  return json({ url: res.init_point, ref: id });
}

// La web pide el resultado completo al volver de Mercado Pago.
function doGet(e) {
  const q = e.parameter || {};
  if (q.accion !== 'resultado' || !q.ref) return json({ error: 'Pedido no válido' });
  // Si el aviso de Mercado Pago todavía no llegó, se comprueba el pago ahora.
  if (q.payment_id) procesarPago(String(q.payment_id));
  const datos = hoja().getDataRange().getValues();
  const col = n => COLUMNAS.indexOf(n);
  const fila = datos.find((f, i) => i > 0 && f[col('id')] === q.ref);
  if (!fila) return json({ error: 'No encontramos tu pedido' });
  if (fila[col('estado')] !== 'pagado') return json({ estado: 'pendiente' });
  const arq = ARQUETIPOS.find(a => a.id === Number(fila[col('arquetipo_id')]));
  return json({
    estado: 'pagado',
    arquetipo_id: arq.id,
    texto: arq.texto,
    mensaje_extendido: arq.mensaje_extendido,
    mensaje_duelo: fila[col('perdida')] === 'si' ? MENSAJE_DUELO : '',
  });
}

function procesarPago(pagoId) {
  const r = UrlFetchApp.fetch('https://api.mercadopago.com/v1/payments/' + pagoId, {
    headers: { Authorization: 'Bearer ' + prop('MP_ACCESS_TOKEN') }, muteHttpExceptions: true,
  });
  const p = JSON.parse(r.getContentText());
  if (p.status !== 'approved' || !p.external_reference) return;
  const col = n => COLUMNAS.indexOf(n);

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = hoja();
    const datos = sh.getDataRange().getValues();
    for (let i = 1; i < datos.length; i++) {
      const fila = datos[i];
      if (fila[col('id')] !== p.external_reference) continue;
      if (fila[col('enviado')]) return; // ya enviado: Mercado Pago a veces avisa dos veces
      const esperado = PRECIOS[fila[col('producto')]].ars;
      if (Number(p.transaction_amount) < esperado) return;
      sh.getRange(i + 1, col('estado') + 1).setValue('pagado');
      sh.getRange(i + 1, col('pago_id') + 1).setValue(pagoId);
      enviarResultado(fila, col);
      sh.getRange(i + 1, col('enviado') + 1).setValue(new Date());
      return;
    }
  } finally {
    lock.releaseLock();
  }
}

function enviarResultado(fila, col) {
  const arq = ARQUETIPOS.find(a => a.id === Number(fila[col('arquetipo_id')]));
  const nombre = fila[col('nombre')];
  const perdida = fila[col('perdida')] === 'si';
  const personal = fila[col('producto')] === 'personalizado';
  const html = pdfHtml(nombre, arq, perdida);
  const pdf = Utilities.newBlob(html, 'text/html', 'mensaje.html').getAs('application/pdf').setName('Tu mensaje del alma - ' + arq.nombre + '.pdf');

  let cuerpo = 'Hola ' + nombre + ',\n\nGracias por hacer el Test de Alma. Tu arquetipo es ' + arq.nombre +
    ', y tu carta, ' + arq.carta + '. Te dejamos tu mensaje completo en el PDF adjunto.\n';
  if (personal) cuerpo += '\nEn las próximas 48 horas te escribimos por WhatsApp con un audio grabado solo para vos.\n';
  cuerpo += '\nEste test es una herramienta de reflexión personal y no reemplaza la ayuda de un profesional.\n\nTest de Alma';
  GmailApp.sendEmail(fila[col('email')], 'Tu mensaje del alma está listo', cuerpo, { attachments: [pdf], name: 'Test de Alma' });

  if (personal && prop('AVISOS_EMAIL')) {
    GmailApp.sendEmail(prop('AVISOS_EMAIL'), 'Nuevo Mensaje personal: ' + nombre,
      'Nombre: ' + nombre + '\nEmail: ' + fila[col('email')] + '\nWhatsApp: ' + fila[col('telefono')] +
      '\nArquetipo: ' + arq.nombre + ' (' + arq.carta + ')\nPerdió a alguien: ' + (perdida ? 'sí' : 'no') +
      '\nRespuestas: ' + fila[col('respuestas')] + '\n\nPlazo: 48 horas.');
  }
}

function esc(s) {
  return String(s).replace(/[&<>]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c]));
}

function pdfHtml(nombre, arq, perdida) {
  const parrafos = arq.texto.split(/(?<=[.!?])\s+(?=[A-ZÁÉÍÓÚÑ¿"“])/);
  let grupos = '';
  for (let i = 0; i < parrafos.length; i += 4) grupos += '<p>' + esc(parrafos.slice(i, i + 4).join(' ')) + '</p>';
  return '<html><body style="font-family:Georgia,serif;color:#2a2238;line-height:1.6;padding:24px">' +
    '<p style="letter-spacing:2px;font-size:11px;color:#8a6d1f">TEST DE ALMA · ' + esc(nombre).toUpperCase() + '</p>' +
    '<h1 style="font-size:30px;margin:4px 0">' + esc(arq.nombre) + '</h1>' +
    '<p style="font-style:italic;color:#6b6080">Tu carta: ' + esc(arq.carta) + '. ' + esc(arq.significado) + '</p>' +
    grupos +
    '<h2 style="font-size:20px;color:#8a6d1f;margin-top:28px">El mensaje de ' + esc(arq.carta) + '</h2><p>' + esc(arq.mensaje_extendido) + '</p>' +
    (perdida ? '<h2 style="font-size:20px;color:#8a6d1f;margin-top:28px">Para quien partió</h2><p>' + esc(MENSAJE_DUELO) + '</p>' : '') +
    '<p style="font-size:11px;color:#6b6080;margin-top:32px">Este test es una herramienta de reflexión personal y no reemplaza la ayuda de un profesional.</p>' +
    '</body></html>';
}
