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
 *   WHATSAPP_NUMERO  Tu número para los avisos por WhatsApp, con código de país (ej. 5491122334455)
 *   WHATSAPP_CLAVE   La clave que te da CallMeBot (servicio gratis) al activar los avisos
 *   NTFY_TEMA        Nombre secreto del canal de la app ntfy para recibir avisos en el celular
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
  // Modo prueba: propiedad MODO_PRUEBA = si. Salta Mercado Pago, marca como pagado y envía el email.
  if (prop('MODO_PRUEBA') === 'si') {
    const sh = hoja(), n = sh.getLastRow(), col = k => COLUMNAS.indexOf(k);
    sh.getRange(n, col('estado') + 1).setValue('pagado');
    sh.getRange(n, col('pago_id') + 1).setValue('PRUEBA');
    enviarResultado(sh.getRange(n, 1, 1, COLUMNAS.length).getValues()[0], col);
    sh.getRange(n, col('enviado') + 1).setValue(new Date());
    return json({ url: sitio + '?pago=ok&ref=' + id, ref: id });
  }
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
  const html = pdfHtml(nombre, arq, perdida, fila[col('producto')]);
  const pdf = Utilities.newBlob(html, 'text/html', 'mensaje.html').getAs('application/pdf').setName('Tu mensaje del alma - ' + arq.nombre + '.pdf');
  const adjuntos = [pdf];
  // Mensaje personal: se adjunta el audio del arquetipo (audio/<id>.mp3 en la web), si ya existe.
  if (personal) {
    try {
      const r = UrlFetchApp.fetch(prop('SITIO_URL') + 'audio/' + arq.id + '.mp3', { muteHttpExceptions: true });
      if (r.getResponseCode() === 200) adjuntos.push(r.getBlob().setName('Audio - ' + arq.nombre + '.mp3'));
    } catch (e) {}
  }

  let cuerpo = 'Hola ' + nombre + ',\n\nGracias por hacer el Test de Alma. Tu arquetipo es ' + arq.nombre +
    ', y tu carta, ' + arq.carta + '. Te dejamos tu mensaje completo en el PDF adjunto.\n';
  if (personal) cuerpo += '\nEn las próximas 48 horas te escribimos por WhatsApp con un audio grabado solo para vos.\n';
  cuerpo += '\nEste test es una herramienta de reflexión personal y no reemplaza la ayuda de un profesional.\n\nTest de Alma';
  GmailApp.sendEmail(fila[col('email')], 'Tu mensaje del alma está listo', cuerpo, { attachments: adjuntos, name: 'Test de Alma', htmlBody: emailHtml(nombre, arq, personal) });

  if (personal && prop('AVISOS_EMAIL')) {
    GmailApp.sendEmail(prop('AVISOS_EMAIL'), 'Nuevo Mensaje personal: ' + nombre,
      'Nombre: ' + nombre + '\nEmail: ' + fila[col('email')] + '\nWhatsApp: ' + fila[col('telefono')] +
      '\nArquetipo: ' + arq.nombre + ' (' + arq.carta + ')\nPerdió a alguien: ' + (perdida ? 'sí' : 'no') +
      '\nRespuestas: ' + fila[col('respuestas')] + '\n\nPlazo: 48 horas.');
  }

  avisoWhatsApp(personal
    ? '🌙 Nuevo MENSAJE PERSONAL\n' + nombre + ' - ' + arq.nombre + '\nWhatsApp: ' + fila[col('telefono')] + '\nPerdió a alguien: ' + (perdida ? 'sí' : 'no') + '\nPlazo: 48 h'
    : '✨ Nuevo test común\n' + nombre + ' - ' + arq.nombre);
}

// Aviso al celular: app ntfy (gratis) y/o WhatsApp con CallMeBot (gratis, solo para tu propio número).
function avisoWhatsApp(texto) {
  const tema = prop('NTFY_TEMA');
  if (tema) {
    try {
      UrlFetchApp.fetch('https://ntfy.sh/' + encodeURIComponent(tema), { method: 'post', payload: texto, muteHttpExceptions: true });
    } catch (e) { console.log('Aviso ntfy falló: ' + e); }
  }
  const numero = prop('WHATSAPP_NUMERO'), clave = prop('WHATSAPP_CLAVE');
  if (!numero || !clave) return;
  try {
    UrlFetchApp.fetch('https://api.callmebot.com/whatsapp.php?phone=' + encodeURIComponent(numero) +
      '&text=' + encodeURIComponent(texto) + '&apikey=' + encodeURIComponent(clave), { muteHttpExceptions: true });
  } catch (e) { console.log('Aviso WhatsApp falló: ' + e); }
}

function esc(s) {
  return String(s).replace(/[&<>]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c]));
}

// ---------- PDF ----------
// Las imágenes salen de la web (carpeta img/pdf) y van incrustadas en el PDF.

function slug(s) {
  return s.normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().replace('ermitano', 'ermitanio').replace(/\s+/g, '-');
}

function imagen(nombre) {
  try {
    const r = UrlFetchApp.fetch(prop('SITIO_URL') + 'img/pdf/' + nombre + '.jpg', { muteHttpExceptions: true });
    if (r.getResponseCode() !== 200) return '';
    return 'data:image/jpeg;base64,' + Utilities.base64Encode(r.getBlob().getBytes());
  } catch (e) {
    return '';
  }
}

function oraciones(t) {
  return t.match(/[^.!?]+[.!?]+["”]?\s*/g) || [t];
}

// Separa el texto en: cuerpo, práctica de la semana y frase final.
function partes(texto) {
  let practica = '';
  const i = texto.indexOf('Esta semana');
  const j = i >= 0 ? texto.indexOf('No estás solo', i) : -1;
  if (i >= 0 && j > i) {
    practica = texto.slice(i, j).trim();
    texto = texto.slice(0, i) + texto.slice(j);
  }
  const o = oraciones(texto).map(s => s.trim());
  const cierre = o.slice(-2).join(' ');
  const cuerpo = o.slice(0, -2);
  const parrafos = [];
  for (let k = 0; k < cuerpo.length; k += 4) parrafos.push(cuerpo.slice(k, k + 4).join(' '));
  return { parrafos, practica, cierre };
}

function fechaLarga() {
  const m = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
  const d = new Date();
  return d.getDate() + ' de ' + m[d.getMonth()] + ' de ' + d.getFullYear();
}

function pdfHtml(nombre, arq, perdida, producto, imgs) {
  imgs = imgs || { portada: imagen('portada'), carta: imagen(slug(arq.carta)) };
  const p = partes(arq.texto);
  const violeta = '#4a2a6b', oro = '#b48a2c', tinta = '#2b2433', suave = '#6d6378', papel = '#fbf8f2', lila = '#f2ecf7';
  const sep = '<p style="text-align:center;color:' + oro + ';font-size:13pt;letter-spacing:8px;margin:18pt 0">✦ ✦ ✦</p>';
  const titulo = (eyebrow, h) =>
    '<p style="text-align:center;font-size:8.5pt;letter-spacing:3px;color:' + oro + ';margin:0 0 4pt 0">' + eyebrow + '</p>' +
    '<h2 style="text-align:center;font-size:24pt;font-weight:normal;color:' + violeta + ';margin:0 0 6pt 0">' + h + '</h2>' +
    '<p style="text-align:center;margin:0 0 18pt 0"><span style="display:inline-block;width:60pt;border-top:1px solid ' + oro + '"></span></p>';
  const pagina = '<div style="page-break-before:always"></div>';

  const cuerpo = p.parrafos.map((t, k) =>
    '<p style="margin:0 0 11pt 0;text-align:justify">' +
    (k === 0 ? '<span style="font-size:20pt;color:' + violeta + ';line-height:1">' + esc(t.charAt(0)) + '</span>' + esc(t.slice(1)) : esc(t)) +
    '</p>').join('');

  return '<html><head><meta charset="utf-8"></head>' +
    '<body style="font-family:Georgia,\'Times New Roman\',serif;color:' + tinta + ';font-size:11pt;line-height:1.65;margin:0">' +

    // Portada
    '<table width="100%" cellpadding="0" cellspacing="0" style="border:1.5px solid ' + oro + ';background:' + papel + '"><tr><td style="padding:34pt 30pt;text-align:center;height:690pt;vertical-align:middle">' +
    (imgs.portada ? '<img src="' + imgs.portada + '" style="width:100%;max-width:440pt;border-radius:6pt">' : '') +
    '<p style="font-size:9pt;letter-spacing:4px;color:' + oro + ';margin:34pt 0 6pt 0">TEST DE ALMA</p>' +
    '<p style="font-size:12pt;font-style:italic;color:' + suave + ';margin:0">Un mensaje para</p>' +
    '<h1 style="font-size:30pt;font-weight:normal;color:' + violeta + ';margin:2pt 0 0 0">' + esc(nombre) + '</h1>' +
    sep +
    '<table align="center" cellpadding="0" cellspacing="0"><tr>' +
    (imgs.carta ? '<td style="padding-right:18pt;vertical-align:middle"><img src="' + imgs.carta + '" style="width:110pt;border:1px solid ' + oro + ';border-radius:5pt"></td>' : '') +
    '<td style="text-align:left;vertical-align:middle">' +
    '<p style="font-size:8.5pt;letter-spacing:3px;color:' + oro + ';margin:0">TU ARQUETIPO</p>' +
    '<p style="font-size:22pt;color:' + violeta + ';margin:2pt 0 6pt 0;line-height:1.15">' + esc(arq.nombre) + '</p>' +
    '<p style="font-size:10.5pt;color:' + suave + ';margin:0;font-style:italic">Tu carta: ' + esc(arq.carta) + '</p>' +
    '</td></tr></table>' +
    '<p style="font-size:9pt;color:' + suave + ';margin:40pt 0 0 0">' + fechaLarga() + '</p>' +
    '</td></tr></table>' +

    // Mensaje
    pagina + titulo('TU MENSAJE', esc(arq.nombre)) + cuerpo +
    (p.practica ?
      '<table width="100%" cellpadding="0" cellspacing="0" style="margin:16pt 0"><tr><td style="background:' + lila + ';border-left:3px solid ' + oro + ';padding:14pt 18pt">' +
      '<p style="font-size:8.5pt;letter-spacing:3px;color:' + oro + ';margin:0 0 6pt 0">TU PRÁCTICA DE ESTA SEMANA</p>' +
      '<p style="margin:0">' + esc(p.practica) + '</p></td></tr></table>' : '') +
    '<p style="text-align:center;font-size:15pt;font-style:italic;color:' + violeta + ';margin:22pt 30pt 0 30pt;line-height:1.5">“' + esc(p.cierre) + '”</p>' +

    // Carta
    pagina + titulo('TU CARTA DEL TAROT', esc(arq.carta)) +
    (imgs.carta ? '<p style="text-align:center;margin:0 0 14pt 0"><img src="' + imgs.carta + '" style="width:170pt;border:1.5px solid ' + oro + ';border-radius:8pt"></p>' : '') +
    '<p style="text-align:center;font-style:italic;color:' + suave + ';font-size:12pt;margin:0 0 16pt 0">' + esc(arq.significado) + '</p>' +
    '<p style="text-align:justify;margin:0 30pt">' + esc(arq.mensaje_extendido) + '</p>' +

    // Duelo
    (perdida ? pagina + titulo('PARA QUIEN PARTIÓ', 'El amor no se va') +
      '<p style="text-align:center;font-size:13pt;font-style:italic;line-height:1.8;margin:10pt 40pt">' + esc(MENSAJE_DUELO) + '</p>' + sep : '') +

    // Cierre
    (perdida ? '' : pagina) +
    (producto === 'personalizado' ?
      '<table width="100%" cellpadding="0" cellspacing="0" style="margin:' + (perdida ? '10pt' : '0') + ' 0 18pt 0"><tr><td style="border:1px solid ' + oro + ';padding:14pt 18pt;text-align:center">' +
      '<p style="font-size:8.5pt;letter-spacing:3px;color:' + oro + ';margin:0 0 6pt 0">TU MENSAJE PERSONAL</p>' +
      '<p style="margin:0">En las próximas 48 horas te escribimos por WhatsApp con un audio grabado solo para vos, a partir de tus respuestas.</p></td></tr></table>' : '') +
    '<p style="text-align:center;font-style:italic;color:' + suave + ';margin:' + (perdida ? '10pt' : '120pt') + ' 0 4pt 0">Con cariño,</p>' +
    '<p style="text-align:center;font-size:14pt;color:' + violeta + ';margin:0 0 30pt 0">Test de Alma</p>' +
    '<p style="font-size:8pt;color:' + suave + ';text-align:center;line-height:1.5;margin:0 30pt">Este mensaje es una experiencia de reflexión simbólica y acompañamiento espiritual. No constituye asesoramiento médico, psicológico, legal ni financiero. Si estás atravesando un momento que sentís que te supera, buscar acompañamiento profesional es un acto de amor propio.</p>' +
    '</body></html>';
}

function emailHtml(nombre, arq, personal) {
  const violeta = '#4a2a6b', oro = '#b48a2c', suave = '#6d6378';
  const img = imagen(slug(arq.carta));
  return '<div style="background:#f3eef8;padding:24px 12px;font-family:Georgia,serif;color:#2b2433">' +
    '<table align="center" width="100%" cellpadding="0" cellspacing="0" style="max-width:520px;background:#fbf8f2;border:1px solid ' + oro + ';border-radius:10px;overflow:hidden">' +
    '<tr><td style="background:#1f1830;padding:26px 24px;text-align:center">' +
    '<p style="margin:0;font-size:11px;letter-spacing:4px;color:#e8c46c">TEST DE ALMA</p>' +
    '<p style="margin:8px 0 0;font-size:24px;color:#efe9f6">Tu mensaje está listo</p></td></tr>' +
    '<tr><td style="padding:28px 28px 8px">' +
    '<p style="font-size:20px;color:' + violeta + ';margin:0 0 14px">Hola ' + esc(nombre) + ',</p>' +
    '<table cellpadding="0" cellspacing="0" style="margin:0 0 16px"><tr>' +
    (img ? '<td style="padding-right:14px"><img src="' + img + '" width="70" style="border:1px solid ' + oro + ';border-radius:5px;display:block"></td>' : '') +
    '<td><p style="margin:0;font-size:11px;letter-spacing:3px;color:' + oro + '">TU ARQUETIPO</p>' +
    '<p style="margin:4px 0;font-size:20px;color:' + violeta + '">' + esc(arq.nombre) + '</p>' +
    '<p style="margin:0;font-size:13px;font-style:italic;color:' + suave + '">Tu carta: ' + esc(arq.carta) + '</p></td></tr></table>' +
    '<p style="font-size:15px;line-height:1.6;margin:0 0 12px">Preparamos tu mensaje con cuidado, a partir de lo que estás viviendo y de lo que tu alma necesita escuchar. <b>Lo encontrás adjunto en PDF.</b></p>' +
    '<p style="font-size:15px;line-height:1.6;margin:0 0 12px">Te recomendamos leerlo en un momento tranquilo, sin apuro. Y si algo te resuena, guardalo. Es para vos.</p>' +
    (personal ? '<p style="font-size:15px;line-height:1.6;margin:0 0 12px">También te adjuntamos <b>el audio de tu arquetipo</b>, para escucharlo con calma.</p>' : '') + (personal ? '<p style="font-size:15px;line-height:1.6;margin:0 0 12px;background:#f2ecf7;border-left:3px solid ' + oro + ';padding:10px 12px">En las próximas <b>48 horas</b> te escribimos por WhatsApp con un audio grabado solo para vos.</p>' : '') +
    '<p style="font-size:15px;font-style:italic;color:' + suave + ';margin:20px 0 4px">Con cariño,</p>' +
    '<p style="font-size:16px;color:' + violeta + ';margin:0 0 20px">Test de Alma</p></td></tr>' +
    '<tr><td style="background:#f2ecf7;padding:14px 28px;font-size:11px;line-height:1.5;color:' + suave + '">Este mensaje es una experiencia de reflexión simbólica y acompañamiento espiritual. No constituye asesoramiento médico, psicológico, legal ni financiero. Si estás atravesando un momento que sentís que te supera, buscar acompañamiento profesional es un acto de amor propio.</td></tr>' +
    '</table></div>';
}
