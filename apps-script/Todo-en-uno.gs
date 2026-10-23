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


// Generado desde contenido-test-v2.json. No editar a mano.
const ARQUETIPOS = [
 {
  "id": 1,
  "nombre": "El Buscador",
  "carta": "El Loco",
  "significado": "El comienzo de un viaje. Confianza en lo desconocido.",
  "texto": "Hay algo en vos que no se conforma con lo que ya sabe. Una parte de tu alma que sigue preguntando, que sigue buscando, que no se queda quieta con las respuestas fáciles. Y eso, aunque a veces te canse, es un regalo. No estás perdido. Estás buscando. Y buscar no es lo mismo que estar perdido. El que busca sabe que hay algo más. El que busca intuye que la vida tiene capas que todavía no tocó. Y esa intuición tuya no es casualidad: es un llamado. Lo que sentís no es debilidad. Es amor que sigue fluyendo, incluso cuando ya no encuentra a quién abrazar. Es tu alma pidiéndote que sigas caminando, aunque no veas el final del camino. Tus respuestas dibujan el perfil de alguien que no se contenta con la superficie. Alguien que mira el cielo y se pregunta. Alguien que siente que hay algo más grande, aunque no siempre sepa nombrarlo. Muchas tradiciones espirituales hablan de los buscadores como almas que eligieron venir a esta vida a recordar, no a aprender. A recordar de dónde vienen. A recordar que hay un hilo que las conecta con algo más grande. Y vos, por tus respuestas, parecés ser una de esas almas. Los ángeles que te acompañan no te dan respuestas directas. No te dicen \"esto es así\". Te dan señales. Te ponen personas en el camino. Te hacen encontrar libros, canciones, conversaciones que llegan justo cuando las necesitás. Porque saben que tu alma no necesita que le digan qué hacer: necesita que le recuerden que ya sabe. Hay una imagen que aparece en tus respuestas: la de alguien parado en un cruce de caminos, mirando el horizonte. No está perdido. Está eligiendo. Y elegir, cuando se hace con el alma, es un acto sagrado. No estás perdido. Estás buscando. Y el que busca, ya está en el camino correcto. Esta semana, elegí un momento del día y hacete una sola pregunta: \"¿Qué es lo que mi alma está tratando de recordarme?\" No busques la respuesta con la mente. Dejala caer. Escribila si aparece. Y si no aparece, no te preocupes: la pregunta ya está haciendo su trabajo. No estás solo. Los ángeles que te acompañan no te van a dar un mapa, pero te van a dar señales. Están en las coincidencias, en los sueños, en las personas que aparecen justo cuando las necesitás. Seguí buscando. No estás perdido. Estás en el camino.",
  "mensaje_extendido": "El Loco te invita a dar el primer paso sin saber exactamente a dónde vas. Te recuerda que la fe no es certeza, sino confianza. Y que el camino se hace al andar. Cada vez que te animás a lo desconocido, el universo te abre una puerta. No necesitás tener todo resuelto. Solo necesitás animarte a empezar. Esta carta te dice que el viaje más importante es el que hacés hacia adentro. Y que, aunque parezca que estás solo, siempre hay alguien caminando a tu lado. Confiá. El camino se va a ir abriendo."
 },
 {
  "id": 2,
  "nombre": "El Sanador",
  "carta": "La Emperatriz",
  "significado": "Nutrición, cuidado, abundancia. Dar sin vaciarse.",
  "texto": "Hay algo en vos que te lleva a cuidar a los demás. A escuchar. A sostener. A estar presente incluso cuando vos mismo necesitarías que alguien esté presente para vos. Y eso, aunque es un regalo, también es un peso. No estás solo. Y no tenés que sostener todo. Lo que sentís no es debilidad. Es amor que sigue fluyendo, incluso cuando ya no encuentra a quién abrazar. Pero ese amor también tiene que fluir hacia vos. Tu alma no necesita que salves a nadie más. Necesita que te permitas ser salvado también. Tus respuestas dibujan el perfil de alguien que aprendió a sanar a otros desde su propia herida. Alguien que sabe escuchar porque alguna vez necesitó ser escuchado. Alguien que sabe acompañar porque alguna vez se sintió solo. Los ángeles que te acompañan no te piden que cargues con el dolor del mundo. Te piden que empieces por vos. Porque un sanador que no se sana a sí mismo termina vaciándose. Y vos no viniste a vaciarte. Viniste a fluir. Hay una imagen que aparece en tus respuestas: la de alguien sosteniendo una vela encendida para iluminar el camino de otros. Pero la vela se consume. Y si no se renueva, se apaga. Tu alma te pide que te permitas renovar tu propia luz. No tenés que salvarte a todos. A veces, el acto más sanador es permitirte ser cuidado. Esta semana, elegí un día para no cuidar a nadie. No para ser egoísta, sino para descansar. No ofrezcas ayuda. No resuelvas problemas. No sostengas. Solo estate. Y si aparece la culpa, respirá y decite: \"Esto también es parte de cuidarme.\" No estás solo. Los ángeles que te acompañan no te piden que seas fuerte todo el tiempo. Te piden que te permitas descansar. Cuidate. No estás fallando. Estás volviendo a vos.",
  "mensaje_extendido": "La Emperatriz te recuerda que el cuidado no es debilidad, sino poder. Te invita a nutrirte, a nutrir a otros, a dejar que la vida florezca a tu alrededor. Esta carta te dice que el amor que das no se pierde. Se multiplica. Y que vos también merecés ser nutrido. Permitite recibir. Permitite florecer. El universo te está invitando a crear, a dar a luz algo nuevo, a confiar en que lo que plantás va a crecer. Aunque no veas los frutos todavía, las raíces ya están trabajando."
 },
 {
  "id": 3,
  "nombre": "El Guardián",
  "carta": "El Emperador",
  "significado": "Protección, estructura, liderazgo. Cuidar sin controlar.",
  "texto": "Hay algo en vos que te lleva a proteger. A sostener. A estar firme incluso cuando por dentro todo tiembla. Y eso, aunque es una forma de amor, también es una forma de soledad. No estás solo. Y no tenés que cargar todo. Lo que sentís no es debilidad. Es amor que sigue fluyendo, incluso cuando ya no encuentra a quién abrazar. Pero ese amor también tiene que fluir hacia vos. Tu alma no necesita que seas fuerte todo el tiempo. Necesita que te permitas ser sostenido también. Tus respuestas dibujan el perfil de alguien que aprendió a sostener a otros desde muy temprano. Alguien que sabe proteger porque alguna vez necesitó protección. Alguien que sabe estar firme porque alguna vez se sintió frágil. Los ángeles que te acompañan no te piden que cargues con el peso del mundo. Te piden que empieces por vos. Porque un guardián que no se protege a sí mismo termina agotándose. Y vos no viniste a agotarte. Viniste a vivir. Hay una imagen que aparece en tus respuestas: la de alguien sosteniendo un techo para que otros pasen. Pero el techo pesa. Y si no se apoya en algo, se cae. Tu alma te pide que te permitas apoyarte también. No tenés que sostener todo para ser digno de amor. A veces, el acto más valiente es permitirte descansar. Esta semana, elegí un día para no sostener nada. No para ser irresponsable, sino para descansar. No resuelvas problemas. No te adelantes a lo que puede pasar. Solo estate. Y si aparece la culpa, respirá y decite: \"Esto también es parte de cuidarme.\" No estás solo. Los ángeles que te acompañan no te piden que seas fuerte todo el tiempo. Te piden que te permitas descansar. Descansá. No estás fallando. Estás volviendo a vos.",
  "mensaje_extendido": "El Emperador te recuerda que la protección no es control. Que cuidar no es dominar. Que sostener no es cargar con todo. Esta carta te invita a poner límites, a estructurar tu vida, a construir sobre bases firmes. Pero también te recuerda que el poder real no está en el control, sino en la confianza. Permitite liderar sin cargar. Permitite proteger sin controlar. Y permitite descansar. Porque un guardián que no descansa no puede sostener nada."
 },
 {
  "id": 4,
  "nombre": "El Que Descansa",
  "carta": "El Colgado",
  "significado": "Pausa, rendición, ver desde otra perspectiva.",
  "texto": "Hay algo en vos que pide pausa. Una parte de tu alma que está cansada, no de vivir, sino de sostener. De sostener el ritmo, las expectativas, las responsabilidades, los silencios incómodos. Y ese cansancio no es debilidad: es una señal. No estás solo. Y no tenés que seguir a este ritmo. Lo que sentís no es debilidad. Es amor que sigue fluyendo, incluso cuando ya no encuentra a quién abrazar. Pero ese amor también tiene que fluir hacia vos. Tu alma no necesita más fuerza. Necesita más espacio. Tus respuestas dibujan el perfil de alguien que ha estado corriendo durante mucho tiempo. Alguien que ha sostenido mucho más de lo que ha recibido. Alguien que ha aprendido a aguantar, pero no a descansar. Los ángeles que te acompañan no te piden que sigas. Te piden que pares. Porque un alma agotada no puede escuchar las señales. Y vos necesitás escuchar. Necesitás descansar para volver a ver el camino. Hay una imagen que aparece en tus respuestas: la de un río que ha estado corriendo durante mucho tiempo, arrastrando piedras, ramas, todo lo que encuentra. Y en algún momento, el río necesita ensancharse, bajar la velocidad, dejar que el agua se asiente. No para dejar de ser río. Sino para volver a ver el fondo. Parar no es rendirse. Parar es elegir seguir viviendo. Esta semana, elegí un momento del día para hacer algo que no sirva para nada. No para producir, no para ayudar, no para resolver. Solo para vos. Puede ser mirar por la ventana, tomar un té sin apuro, escuchar una canción con los ojos cerrados. Si aparece la culpa, respirá y decite: \"Esto también es parte de cuidarme.\" No estás solo. Aunque ahora sientas que nadie ve lo que cargás, hay una parte de vos que sí te ve. Y hay algo más grande que vos que también te sostiene, incluso cuando no lo sentís. Descansá. No estás fallando. Estás volviendo a vos.",
  "mensaje_extendido": "El Colgado te invita a pausar. A rendirte. A ver desde otra perspectiva. Esta carta te dice que no todo se resuelve con acción. A veces, la respuesta llega cuando dejás de buscar. Cuando dejás de forzar. Cuando te permitís estar quieto. No estás perdiendo el tiempo. Estás dejando que el tiempo haga su trabajo. Permitite estar en pausa. Permitite no saber. Permitite que el misterio te habite. Y confiá en que, cuando menos lo esperes, la respuesta va a llegar."
 },
 {
  "id": 5,
  "nombre": "El Corazón Invisible",
  "carta": "La Luna",
  "significado": "Misterio, emociones profundas, lo que no se ve.",
  "texto": "Hay algo en vos que siente que no es visto. Que no es escuchado. Que pasa desapercibido incluso cuando está dando todo. Y eso, aunque no lo digas, duele. No estás solo. Y no estás invisible. Lo que sentís no es debilidad. Es amor que sigue fluyendo, incluso cuando ya no encuentra a quién abrazar. Pero ese amor también tiene que fluir hacia vos. Tu alma no necesita que te vean para ser valiosa. Ya lo es. Pero necesita que vos te veas. Tus respuestas dibujan el perfil de alguien que ha aprendido a estar en silencio. Alguien que ha aprendido a no molestar. Alguien que ha aprendido a dar sin pedir. Y ese aprendizaje, aunque fue necesario en algún momento, ahora te pesa. Los ángeles que te acompañan no te piden que grites. Te piden que te escuches. Porque un alma que no se escucha a sí misma termina creyendo que no merece ser escuchada. Y vos merecés ser escuchado. Hay una imagen que aparece en tus respuestas: la de alguien parado en una habitación llena de gente, sintiéndose solo. Pero no está solo. Solo está desconectado de sí mismo. Y el camino de vuelta empieza por vos. No estás invisible. Estás desconectado de vos mismo. Y volver a vos es el primer paso para que los demás te vean. Esta semana, elegí un momento del día para escribir tres cosas que sentís. No para mostrárselas a nadie. Solo para vos. Puede ser una emoción, una necesidad, un deseo. Escribilas y leelas en voz alta. Y si aparece la vergüenza, respirá y decite: \"Esto también es parte de verme.\" No estás solo. Los ángeles que te acompañan te ven, aunque no lo sientas. Los guías que te cuidan te escuchan, aunque no lo sepas. Y los seres amados que te velan te acompañan, aunque no los veas. No estás invisible. Estás volviendo a vos.",
  "mensaje_extendido": "La Luna te invita a mirar lo que no se ve. A escuchar lo que no se dice. A confiar en lo que sentís aunque no puedas explicarlo. Esta carta te dice que hay una sabiduría en la oscuridad. Que las emociones más profundas no son un problema, son una guía. No te apures a salir de la noche. A veces, la noche es el lugar donde el alma descansa. Permitite sentir. Permitite no entender. Permitite que la luna te acompañe. Y confiá en que, aunque no veas el camino, estás siendo guiado."
 },
 {
  "id": 6,
  "nombre": "El Luchador",
  "carta": "La Fuerza",
  "significado": "Valentía suave, resistencia, poder interior.",
  "texto": "Hay algo en vos que no se rinde. Que sigue adelante aunque duela. Que se levanta aunque esté cansado. Y eso, aunque es una forma de fuerza, también es una forma de agotamiento. No estás solo. Y no tenés que luchar todo el tiempo. Lo que sentís no es debilidad. Es amor que sigue fluyendo, incluso cuando ya no encuentra a quién abrazar. Pero ese amor también tiene que fluir hacia vos. Tu alma no necesita que luches más. Necesita que descanses. Tus respuestas dibujan el perfil de alguien que ha aprendido a resistir. Alguien que ha aprendido a aguantar. Alguien que ha aprendido a seguir adelante incluso cuando todo parece perdido. Y esa resistencia, aunque te ha traído hasta acá, también te ha desgastado. Los ángeles que te acompañan no te piden que luches más. Te piden que confíes. Porque un alma que lucha todo el tiempo no puede escuchar las señales. Y vos necesitás escuchar. Hay una imagen que aparece en tus respuestas: la de alguien remando contra la corriente durante mucho tiempo. Pero en algún momento, el río cambia. Y el que sigue remando contra la corriente se agota sin darse cuenta de que ya podría estar fluyendo. No tenés que luchar todo el tiempo. A veces, el acto más valiente es soltar los remos y confiar. Esta semana, elegí un momento del día para no hacer nada. No para ser vago, sino para descansar. No resuelvas problemas. No te adelantes a lo que puede pasar. Solo estate. Y si aparece la culpa, respirá y decite: \"Esto también es parte de confiar.\" No estás solo. Los ángeles que te acompañan no te piden que luches más. Te piden que confíes. Descansá. No estás fallando. Estás volviendo a vos.",
  "mensaje_extendido": "La Fuerza te invita a confiar en tu poder interior. No en la fuerza bruta, sino en la fuerza suave. La que sostiene sin apretar. La que espera sin desesperar. La que resiste sin romperse. Esta carta te dice que ya tenés todo lo que necesitás. Que no necesitás ser más fuerte. Solo necesitás confiar en la fuerza que ya tenés. Permitite ser suave. Permitite ser paciente. Permitite ser valiente. Y confiá en que tu poder interior es más grande de lo que creés."
 },
 {
  "id": 7,
  "nombre": "El Que Suelta",
  "carta": "La Templanza",
  "significado": "Equilibrio, paciencia, sanación lenta.",
  "texto": "Hay algo en vos que carga con un peso que no te corresponde. Una culpa. Un arrepentimiento. Algo que quedó pendiente y que todavía te pesa. Y eso, aunque es humano, también es agotador. No estás solo. Y no tenés que cargar con eso. Lo que sentís no es debilidad. Es amor que sigue fluyendo, incluso cuando ya no encuentra a quién abrazar. Pero ese amor también tiene que fluir hacia vos. Tu alma no necesita que te castigues. Necesita que te perdones. Tus respuestas dibujan el perfil de alguien que ha aprendido a cargar con la culpa. Alguien que se exige más de lo que se permite. Alguien que se juzga más de lo que se acepta. Y esa culpa, aunque parece justa, no te está ayudando. Los ángeles que te acompañan no te piden que te castigues. Te piden que te perdones. Porque un alma que no se perdona no puede avanzar. Y vos necesitás avanzar. Hay una imagen que aparece en tus respuestas: la de alguien cargando una mochila llena de piedras. Cada piedra es una culpa, un arrepentimiento, algo que no se perdonó. Y en algún momento, la mochila pesa tanto que ya no se puede caminar. Pero las piedras no son tuyas. Las cargaste porque creíste que debías. Y podés soltarlas. No tenés que cargar con la culpa para ser digno de amor. A veces, el acto más valiente es perdonarte. Esta semana, elegí un momento tranquilo y escribí una carta de perdón hacia vos mismo. No para justificar lo que hiciste. Para reconocer que hiciste lo mejor que pudiste con lo que tenías. Escribila a mano. Y si al terminarla sentís paz, aunque sea un poco, es señal de que estás soltando. No estás solo. Los ángeles que te acompañan no te piden que te castigues. Te piden que te perdones. Perdonate. No estás fallando. Estás volviendo a vos.",
  "mensaje_extendido": "La Templanza te invita a encontrar el equilibrio. A no apurarte. A confiar en el proceso. Esta carta te dice que la sanación no es lineal. Que hay días mejores y días peores. Y que ambos son parte del camino. No te exijas sanar de un día para el otro. Permitite ir despacio. Permitite mezclar. Permitite encontrar tu propio ritmo. Y confiá en que, aunque parezca que no avanzás, cada paso cuenta. Cada momento de paz cuenta. Cada respiración cuenta."
 },
 {
  "id": 8,
  "nombre": "El Que Espera",
  "carta": "La Estrella",
  "significado": "Esperanza, fe, guía en la oscuridad.",
  "texto": "Hay algo en vos que espera. Una señal. Un cambio. Una respuesta. Algo que te diga que vas por buen camino. Y esa espera, aunque es humana, también es agotadora. No estás solo. Y no tenés que esperar para siempre. Lo que sentís no es debilidad. Es amor que sigue fluyendo, incluso cuando ya no encuentra a quién abrazar. Pero ese amor también tiene que fluir hacia vos. Tu alma no necesita que esperes. Necesita que confíes. Tus respuestas dibujan el perfil de alguien que ha aprendido a esperar. Alguien que ha aprendido a tener paciencia. Alguien que ha aprendido a confiar en que algo va a cambiar. Y esa espera, aunque es una forma de fe, también es una forma de detenerse. Los ángeles que te acompañan no te piden que esperes más. Te piden que confíes. Porque un alma que espera todo el tiempo no puede vivir el presente. Y vos necesitás vivir el presente. Hay una imagen que aparece en tus respuestas: la de alguien parado en una estación, esperando un tren que no llega. Pero el tren no es el único camino. Hay otros. Y a veces, el tren que esperamos no es el que necesitamos. No tenés que esperar para ser feliz. A veces, el acto más valiente es confiar en que lo que buscás ya está en camino. Esta semana, elegí un momento del día para hacer algo que te guste. No para distraerte de la espera, sino para recordarte que la vida también pasa mientras esperás. Puede ser caminar, leer, cocinar, escuchar música. Y si aparece la culpa, respirá y decite: \"Esto también es parte de confiar.\" No estás solo. Los ángeles que te acompañan no te piden que esperes más. Te piden que confíes. Confiá. No estás fallando. Estás volviendo a vos.",
  "mensaje_extendido": "La Estrella te invita a esperar. A confiar en que hay una luz que te guía, incluso en la oscuridad. Esta carta te dice que no estás solo. Que hay algo más grande que vos acompañándote. Que hay un propósito, aunque no lo veas. Permitite esperar. Permitite confiar. Permitite que la estrella te guíe. Y recordá que la espera no es eterna. La vida fluye. Y el reencuentro existe."
 },
 {
  "id": 9,
  "nombre": "El Niño Interior",
  "carta": "El Sol",
  "significado": "Alegría, claridad, plenitud.",
  "texto": "Hay algo en vos que necesita ternura. Un niño interior que todavía espera ser consolado. Un niño interior que todavía espera ser visto. Y eso, aunque no lo digas, es una de las cosas que más te pesa. No estás solo. Y no tenés que consolarte solo. Lo que sentís no es debilidad. Es amor que sigue fluyendo, incluso cuando ya no encuentra a quién abrazar. Pero ese amor también tiene que fluir hacia vos. Tu alma no necesita que seas fuerte. Necesita que te permitas ser cuidado. Tus respuestas dibujan el perfil de alguien que aprendió a crecer rápido. Alguien que aprendió a ser fuerte antes de tiempo. Alguien que aprendió a cuidarse solo porque no había otra opción. Y ese niño interior, aunque creció, todavía sigue ahí. Los ángeles que te acompañan no te piden que sigas siendo fuerte. Te piden que te permitas ser cuidado. Porque un alma que no se permite ser cuidada no puede descansar. Y vos necesitás descansar. Hay una imagen que aparece en tus respuestas: la de alguien abrazando a un niño que llora. Pero el que abraza y el que llora son la misma persona. Y en algún momento, hay que permitir que el niño sea consolado. No tenés que ser fuerte todo el tiempo. A veces, el acto más valiente es permitirte ser cuidado. Esta semana, elegí un momento tranquilo y escribí una carta a tu niño interior. Decile lo que necesitabas escuchar cuando eras chico. Decile que está a salvo. Decile que ya no está solo. Y si al terminarla sentís paz, aunque sea un poco, es señal de que estás sanando. No estás solo. Los ángeles que te acompañan no te piden que seas fuerte. Te piden que te permitas ser cuidado. Permitite ser cuidado. No estás fallando. Estás volviendo a vos.",
  "mensaje_extendido": "El Sol te invita a brillar. A confiar en tu luz. A permitirte ser visto. Esta carta te dice que la alegría no es un pecado. Que la plenitud no es egoísmo. Que tenés derecho a ser feliz. Permitite brillar. Permitite disfrutar. Permitite que la vida te sonría. Y recordá que tu luz no le quita luz a nadie. Al contrario, ilumina a todos los que te rodean."
 },
 {
  "id": 10,
  "nombre": "El Sabio Silencioso",
  "carta": "El Ermitaño",
  "significado": "Introspección, sabiduría, luz interior.",
  "texto": "Hay algo en vos que ya sabe. Una parte de tu alma que intuye el camino, que conoce la respuesta, que sabe qué hacer. Pero no te animás. Y esa desconexión entre lo que sabés y lo que hacés es una de las cosas que más te pesa. No estás solo. Y no tenés que saberlo todo. Lo que sentís no es debilidad. Es amor que sigue fluyendo, incluso cuando ya no encuentra a quién abrazar. Pero ese amor también tiene que fluir hacia vos. Tu alma no necesita más información. Necesita que confíes en lo que ya sabe. Tus respuestas dibujan el perfil de alguien que ha aprendido a escuchar su intuición. Alguien que ha aprendido a confiar en su voz interior. Alguien que ha aprendido a saber sin saber cómo. Y esa sabiduría, aunque es un regalo, también es una responsabilidad. Los ángeles que te acompañan no te piden que sepas más. Te piden que confíes en lo que ya sabés. Porque un alma que sabe pero no se anima a actuar vive en una tensión constante. Y vos necesitás paz. Hay una imagen que aparece en tus respuestas: la de alguien parado frente a una puerta que sabe que tiene que abrir, pero no se anima. La puerta no está cerrada con llave. Solo está cerrada. Y abrirla depende de vos. Ya sabés lo que tenés que hacer. Solo necesitás animarte a hacerlo. Esta semana, elegí una cosa que sabés que tenés que hacer pero no te animás. No tiene que ser grande. Puede ser una conversación, un mensaje, una decisión. Hacela. Y si aparece el miedo, respirá y decite: \"Esto también es parte de confiar.\" No estás solo. Los ángeles que te acompañan no te piden que sepas más. Te piden que confíes en lo que ya sabés. Animate. No estás fallando. Estás volviendo a vos.",
  "mensaje_extendido": "El Ermitaño te invita a mirar hacia adentro. A confiar en tu voz interior. A encontrar tu propia luz. Esta carta te dice que no necesitás que otros te digan qué hacer. Que ya sabés. Que tu intuición es tu mejor guía. Permitite estar solo. Permitite escucharte. Permitite confiar en lo que sentís. Y recordá que la soledad no es castigo. Es un espacio sagrado donde el alma se encuentra a sí misma."
 },
 {
  "id": 11,
  "nombre": "El Que Renace",
  "carta": "La Muerte",
  "significado": "Transformación, renacimiento, dejar atrás.",
  "texto": "Hay algo en vos que está muriendo. No vos, sino una versión de vos. Una versión que ya no te sirve. Una versión que ya no te representa. Y esa muerte, aunque duele, también es un nacimiento. No estás solo. Y no tenés que renacer solo. Lo que sentís no es debilidad. Es amor que sigue fluyendo, incluso cuando ya no encuentra a quién abrazar. Pero ese amor también tiene que fluir hacia vos. Tu alma no necesita que te aferres a lo viejo. Necesita que confíes en lo nuevo. Tus respuestas dibujan el perfil de alguien que está en transición. Alguien que está dejando atrás algo que ya no le sirve. Alguien que está empezando a ser otro. Y esa transición, aunque es incómoda, es necesaria. Los ángeles que te acompañan no te piden que te apures. Te piden que confíes. Porque un alma en transición necesita tiempo para acomodarse. Y vos necesitás tiempo. Hay una imagen que aparece en tus respuestas: la de una oruga que se convierte en mariposa. Pero entre la oruga y la mariposa hay un momento de completa desintegración. La oruga no sabe que va a ser mariposa. Solo sabe que ya no puede ser oruga. Y en ese no saber, está el milagro. No estás perdiendo. Estás renaciendo. Y el renacimiento siempre empieza con una muerte. Esta semana, elegí un momento tranquilo y escribí una lista de cosas que querés dejar atrás. No para castigarte, sino para reconocer que ya no te sirven. Puede ser una creencia, un hábito, una relación, una versión de vos. Escribilas. Y si al terminarla sentís paz, aunque sea un poco, es señal de que estás renaciendo. No estás solo. Los ángeles que te acompañan no te piden que te apures. Te piden que confíes. Renacé. No estás fallando. Estás volviendo a vos.",
  "mensaje_extendido": "La Muerte te invita a soltar. A dejar atrás lo que ya no te sirve. A renacer. Esta carta te dice que no estás perdiendo. Estás transformándote. Que hay algo nuevo naciendo en vos. Aunque duela. Aunque no lo entiendas. Permitite soltar. Permitite morir a lo viejo. Permitite renacer. Y confiá en que lo que viene es mejor que lo que se fue."
 },
 {
  "id": 12,
  "nombre": "El Conectado",
  "carta": "El Mundo",
  "significado": "Plenitud, regreso a casa, conexión con todo.",
  "texto": "Hay algo en vos que confía. Una parte de tu alma que sabe que no está sola, aunque a veces lo sienta. Una parte de tu alma que intuye que hay algo más grande, aunque no siempre lo vea. Y esa confianza, aunque es un regalo, también es un camino. No estás solo. Y no tenés que sentirte conectado todo el tiempo. Lo que sentís no es debilidad. Es amor que sigue fluyendo, incluso cuando ya no encuentra a quién abrazar. Pero ese amor también tiene que fluir hacia vos. Tu alma no necesita que confíes más. Necesita que confíes en que ya está conectada. Tus respuestas dibujan el perfil de alguien que ha aprendido a confiar. Alguien que ha aprendido a ver señales. Alguien que ha aprendido a sentir que hay algo más grande acompañándolo. Y esa confianza, aunque es una forma de sabiduría, también es una forma de entrega. Los ángeles que te acompañan no te piden que confíes más. Te piden que confíes en que ya están con vos. Porque un alma que confía no necesita pruebas. Solo necesita recordar. Hay una imagen que aparece en tus respuestas: la de alguien caminando por un sendero iluminado por la luna. No está solo. La luz lo acompaña. Y aunque no ve el final del camino, sabe que no está perdido. No estás solo. Nunca lo estuviste. Solo necesitás recordar que ya estás conectado. Esta semana, elegí un momento del día para agradecer. No por algo específico, sino por estar acá. Agradecé el aire, el sol, la vida, las personas que amás, los que ya no están. Agradecé en voz alta si podés. Y si al terminarlo sentís paz, aunque sea un poco, es señal de que estás conectado. No estás solo. Los ángeles que te acompañan, los guías que te cuidan, los seres amados que te velan, todos están con vos. No para darte un mapa, sino para recordarte que ya estás en el camino. Confiá. No estás fallando. Estás volviendo a vos.",
  "mensaje_extendido": "El Mundo te invita a celebrar. A reconocer todo lo que has recorrido. A confiar en que estás volviendo a casa. Esta carta te dice que el viaje no termina acá. Que hay un hogar que te espera. Que hay un reencuentro que te aguarda. Permitite celebrar. Permitite agradecer. Permitite confiar. Y recordá que todo lo que viviste, lo bueno y lo malo, te trajo hasta acá. Y acá estás. Entero. Y en camino a casa."
 },
 {
  "id": 13,
  "nombre": "El Que Recuerda",
  "carta": "El Juicio",
  "significado": "Despertar, memoria, llamado interior.",
  "texto": "Hay algo en vos que no quiere olvidar. Una parte de tu alma que se aferra a los recuerdos, a los momentos, a las personas y seres amados que ya no están. Y ese aferrarse, aunque es una forma de amor, también es una forma de dolor. No estás solo. Y no tenés que olvidar para estar bien. Lo que sentís no es debilidad. Es amor que sigue fluyendo, incluso cuando ya no encuentra a quién abrazar. Pero ese amor también tiene que fluir hacia vos. Tu alma no necesita que olvides. Necesita que recuerdes sin dolor. Tus respuestas dibujan el perfil de alguien que ha aprendido a honrar la memoria. Alguien que ha aprendido a mantener vivo el recuerdo. Alguien que ha aprendido a amar lo que ya no está. Y esa memoria, aunque es un regalo, también es una carga. Los ángeles que te acompañan no te piden que olvides. Te piden que recuerdes con paz. Porque un alma que recuerda con dolor no puede avanzar. Y vos necesitás avanzar. Hay una imagen que aparece en tus respuestas: la de alguien guardando una foto en un cajón. No para olvidarla, sino para no verla todo el tiempo. La foto sigue ahí. El amor sigue ahí. Solo cambia la forma de mirarla. No tenés que olvidar para estar bien. A veces, el acto más valiente es recordar con paz. Esta semana, elegí un momento tranquilo y mirá una foto de esa persona o ser amado. No para llorar, sino para recordar. Recordá un momento lindo. Recordá una risa. Recordá algo que te enseñó. Y si al terminarlo sentís paz, aunque sea un poco, es señal de que estás recordando con amor. No estás solo. Los ángeles que te acompañan no te piden que olvides. Te piden que recuerdes con paz. Recordá con paz. No estás fallando. Estás volviendo a vos.",
  "mensaje_extendido": "El Juicio te invita a despertar. A recordar. A escuchar el llamado interior. Esta carta te dice que hay algo en vos que está llamando. Que hay una voz que te invita a volver. A recordar quién sos. A recordar de dónde venís. Permitite escuchar. Permitite recordar. Permitite responder al llamado. Y confiá en que, aunque no sepas a dónde vas, estás siendo guiado."
 },
 {
  "id": 14,
  "nombre": "El Que Se Reencuentra",
  "carta": "La Sacerdotisa",
  "significado": "Intuición, sabiduría interior, conexión profunda.",
  "texto": "Perder a alguien que amás es una de las experiencias más difíciles que existen. No importa si fue una persona o un ser amado: el vínculo era real, y el dolor también. No estás exagerando. No estás \"tardando demasiado\" en superarlo. Estás viviendo uno de los procesos más profundos del alma humana. Lo que sentís no es debilidad. Es amor que sigue fluyendo, incluso cuando ya no encuentra a quién abrazar. Y en ese amor que sigue fluyendo, hay algo sagrado. Porque amar a alguien que ya no está físicamente es una de las formas más puras de amor que existen. No estás solo. Y esa persona o ser amado que amás tampoco desapareció del todo. Está en tu memoria, en tus gestos, en las cosas que te enseñó sin que te dieras cuenta. Tus respuestas dibujan el perfil de alguien que no solo extraña: espera. Hay una parte de vos que intuye que la historia no termina acá. Y esa intuición no es ingenuidad. Es una forma de sabiduría que muchas tradiciones han reconocido a lo largo de la historia. Los pueblos originarios de los Andes hablan de que los que se van no se van del todo: se quedan como guardianes. La tradición cristiana habla de reencuentro. La filosofía oriental habla de ciclos que vuelven a encontrarse. Y muchos médiums que acompañan a personas en duelo sostienen que nuestros seres queridos, incluidos los seres amados, no desaparecen: evolucionan. Pasan a un estado de paz desde donde siguen velando por nosotros. No importa en qué creas. Lo que importa es que tu alma lo sabe. En lo más profundo, intuye que el reencuentro existirá. Hay una imagen que aparece en tus respuestas: la de alguien parado en una orilla, mirando el horizonte. No está desesperado. Está esperando. Y esperar, cuando es con amor, no es una pérdida de tiempo. Es una forma de seguir conectado. El amor es la fuerza más potente del universo. Trasciende estados, planos y seres. Es la fuente de todo, y hacia ahí volveremos. Esta semana, elegí un momento tranquilo y escribí una carta a esa persona o ser amado. No para despedirte, porque no te estás despidiendo. Para contarle cómo estás, qué te pasa, qué te gustaría que supiera. Podés escribirla a mano o en el celular. No hace falta que la envíes a ningún lado. Solo escribila. Y si al terminarla sentís paz, aunque sea un poco, es señal de que el vínculo sigue vivo. No estás solo. Y esa persona o ser amado que amás tampoco desapareció del todo. Está en tu memoria, en tus gestos, en las cosas que te enseñó sin que te dieras cuenta. Muchas personas que han vivido lo mismo que vos describen momentos en los que sintieron su presencia: un aroma, un sueño, una coincidencia. No sabemos exactamente cómo funciona. Pero sabemos que el amor deja huellas que el tiempo no borra.",
  "mensaje_extendido": "La Sacerdotisa te invita a confiar en tu intuición. A escuchar esa voz interior que sabe más de lo que creés. Esta carta te dice que no necesitás respuestas externas. Que la sabiduría está en vos. Permitite escuchar. Permitite confiar. Permitite que tu intuición te guíe. Y recordá que el silencio no es vacío. Es el lugar donde el alma habla."
 },
 {
  "id": 15,
  "nombre": "El Que Aprende a Soltar",
  "carta": "Los Enamorados",
  "significado": "Elección, vínculos, amor que trasciende.",
  "texto": "Hay algo en vos que está aprendiendo a soltar. Una parte de tu alma que ya sabe que tiene que dejar ir, pero que todavía no puede. Y esa tensión entre saber y poder es una de las cosas que más te pesa. No estás solo. Y no tenés que soltar de un día para el otro. Lo que sentís no es debilidad. Es amor que sigue fluyendo, incluso cuando ya no encuentra a quién abrazar. Pero ese amor también tiene que fluir hacia vos. Tu alma no necesita que sueltes ya. Necesita que te permitas soltar de a poco. Tus respuestas dibujan el perfil de alguien que está en proceso. Alguien que ya sabe que tiene que soltar. Alguien que ya intuye que el dolor no es para siempre. Alguien que ya está en paz, pero le cuesta. Y ese proceso, aunque es lento, es sagrado. Los ángeles que te acompañan no te piden que sueltes ya. Te piden que confíes en el proceso. Porque un alma que se apura a soltar termina aferrándose más. Y vos necesitás tiempo. Hay una imagen que aparece en tus respuestas: la de alguien sosteniendo un pájaro en la mano. El pájaro no está herido. Solo está ahí. Y en algún momento, la mano tiene que abrirse. No porque el pájaro se vaya a ir, sino porque la mano necesita descansar. No tenés que soltar de golpe. A veces, el acto más valiente es permitirte soltar de a poco. Esta semana, elegí un momento tranquilo y escribí una carta a esa persona o ser amado. No para despedirte. Para contarle cómo estás, qué te pasa, qué te gustaría que supiera. Y si al terminarla sentís paz, aunque sea un poco, es señal de que estás soltando. No estás solo. Los ángeles que te acompañan no te piden que sueltes ya. Te piden que confíes en el proceso. Soltá de a poco. No estás fallando. Estás volviendo a vos.",
  "mensaje_extendido": "Los Enamorados te invitan a elegir. A confiar en el amor. A reconocer que los vínculos son sagrados. Esta carta te dice que el amor no es debilidad. Es la fuerza más potente del universo. Permitite amar. Permitite elegir. Permitite confiar en los vínculos que te sostienen. Y recordá que el amor verdadero no ata. Libera."
 },
 {
  "id": 16,
  "nombre": "El Que Protege",
  "carta": "El Papa",
  "significado": "Guía, enseñanza, protección espiritual.",
  "texto": "Hay algo en vos que te lleva a proteger. A cuidar. A estar pendiente de los demás. Y eso, aunque es una forma de amor, también es una forma de olvidarte de vos. No estás solo. Y no tenés que proteger a todos. Lo que sentís no es debilidad. Es amor que sigue fluyendo, incluso cuando ya no encuentra a quién abrazar. Pero ese amor también tiene que fluir hacia vos. Tu alma no necesita que cuides más. Necesita que te permitas ser cuidado. Tus respuestas dibujan el perfil de alguien que ha aprendido a proteger. Alguien que ha aprendido a cuidar. Alguien que ha aprendido a estar pendiente de los demás antes que de sí mismo. Y esa protección, aunque es un regalo, también es un agotamiento. Los ángeles que te acompañan no te piden que protejas más. Te piden que te protejas también a vos. Porque un alma que cuida a todos menos a sí misma termina vaciándose. Y vos no viniste a vaciarte. Hay una imagen que aparece en tus respuestas: la de alguien sosteniendo un paraguas para que otros no se mojen. Pero el que sostiene el paraguas también se moja. Y en algún momento, hay que compartir el paraguas. No tenés que proteger a todos para ser digno de amor. A veces, el acto más valiente es permitirte ser protegido. Esta semana, elegí un día para no proteger a nadie. No para ser egoísta, sino para descansar. No te adelantes a lo que puede pasar. No resuelvas problemas. Solo estate. Y si aparece la culpa, respirá y decite: \"Esto también es parte de cuidarme.\" No estás solo. Los ángeles que te acompañan no te piden que protejas más. Te piden que te permitas ser protegido. Permitite ser protegido. No estás fallando. Estás volviendo a vos.",
  "mensaje_extendido": "El Papa te invita a confiar en la guía. A reconocer que hay maestros en tu camino. A permitirte ser enseñado. Esta carta te dice que no estás solo. Que hay guías que te acompañan. Que hay ángeles que te susurran al oído. Permitite escuchar. Permitite aprender. Permitite ser guiado. Y confiá en que, aunque no veas el camino, alguien lo está trazando para vos."
 },
 {
  "id": 17,
  "nombre": "El Que Confía",
  "carta": "La Rueda",
  "significado": "Ciclos, cambio, confianza en el proceso.",
  "texto": "Hay algo en vos que está aprendiendo a confiar. Una parte de tu alma que ya sabe que no puede controlar todo, pero que todavía se resiste. Y esa resistencia, aunque es humana, también es agotadora. No estás solo. Y no tenés que controlar todo. Lo que sentís no es debilidad. Es amor que sigue fluyendo, incluso cuando ya no encuentra a quién abrazar. Pero ese amor también tiene que fluir hacia vos. Tu alma no necesita que controles más. Necesita que confíes. Tus respuestas dibujan el perfil de alguien que ha aprendido a controlar. Alguien que ha aprendido a prever. Alguien que ha aprendido a no soltar. Y ese control, aunque te ha dado seguridad, también te ha quitado paz. Los ángeles que te acompañan no te piden que controles más. Te piden que confíes. Porque un alma que controla todo no puede descansar. Y vos necesitás descansar. Hay una imagen que aparece en tus respuestas: la de alguien sosteniendo un timón con las dos manos. Pero el río ya sabe a dónde va. Y en algún momento, hay que soltar el timón y confiar en la corriente. No tenés que controlar todo. A veces, el acto más valiente es soltar el timón y confiar. Esta semana, elegí una cosa que no podés controlar y soltala. No para ser irresponsable, sino para descansar. No te adelantes a lo que puede pasar. No resuelvas lo que no te toca. Solo estate. Y si aparece la ansiedad, respirá y decite: \"Esto también es parte de confiar.\" No estás solo. Los ángeles que te acompañan no te piden que controles más. Te piden que confíes. Confiá. No estás fallando. Estás volviendo a vos.",
  "mensaje_extendido": "La Rueda te invita a confiar en los ciclos. A reconocer que todo cambia. A soltar el control. Esta carta te dice que no todo depende de vos. Que hay fuerzas más grandes obrando. Que hay un ritmo en el universo. Permitite fluir. Permitite confiar. Permitite que la vida te sorprenda. Y recordá que, aunque hoy estés abajo, mañana vas a estar arriba. Todo pasa. Todo vuelve."
 },
 {
  "id": 18,
  "nombre": "El Que Agradece",
  "carta": "El Mago",
  "significado": "Manifestación, poder personal, gratitud.",
  "texto": "Hay algo en vos que agradece. Una parte de tu alma que encuentra belleza en lo pequeño, en lo cotidiano, en lo que otros no ven. Y esa gratitud, aunque es un regalo, también es un camino. No estás solo. Y no tenés que agradecer todo el tiempo. Lo que sentís no es debilidad. Es amor que sigue fluyendo, incluso cuando ya no encuentra a quién abrazar. Pero ese amor también tiene que fluir hacia vos. Tu alma no necesita que agradezcas más. Necesita que te permitas recibir. Tus respuestas dibujan el perfil de alguien que ha aprendido a agradecer. Alguien que ha aprendido a ver belleza donde otros ven rutina. Alguien que ha aprendido a encontrar paz en lo simple. Y esa gratitud, aunque es una forma de sabiduría, también es una forma de entrega. Los ángeles que te acompañan no te piden que agradezcas más. Te piden que te permitas recibir. Porque un alma que solo da gracias pero no recibe se vacía. Y vos no viniste a vaciarte. Hay una imagen que aparece en tus respuestas: la de alguien mirando el atardecer. No está apurado. No está triste. Solo está mirando. Y en ese mirar, hay una forma de oración. La gratitud no es solo dar gracias. También es permitirte recibir. Esta semana, elegí un momento del día para recibir algo. No para dar, no para agradecer, no para producir. Solo para recibir. Puede ser un halago, un abrazo, un momento de paz. Recibilo sin culpa. Y si aparece la incomodidad, respirá y decite: \"Esto también es parte de agradecer.\" No estás solo. Los ángeles que te acompañan no te piden que agradezcas más. Te piden que te permitas recibir. Recibí. No estás fallando. Estás volviendo a vos.",
  "mensaje_extendido": "El Mago te invita a reconocer tu poder. A confiar en tu capacidad de crear. A manifestar lo que deseás. Esta carta te dice que tenés todo lo que necesitás. Que el poder está en vos. Que podés crear la realidad que soñás. Permitite crear. Permitite manifestar. Permitite confiar en tu magia. Y recordá que el universo conspira a tu favor."
 },
 {
  "id": 19,
  "nombre": "El Que Perdona",
  "carta": "La Justicia",
  "significado": "Equilibrio, perdón, verdad.",
  "texto": "Hay algo en vos que está aprendiendo a perdonar. No a los demás. A vos. Una parte de tu alma que todavía se juzga, que todavía se castiga, que todavía no se permite ser humana. Y ese juicio, aunque parece justo, no te está ayudando. No estás solo. Y no tenés que castigarte. Lo que sentís no es debilidad. Es amor que sigue fluyendo, incluso cuando ya no encuentra a quién abrazar. Pero ese amor también tiene que fluir hacia vos. Tu alma no necesita que te juzgues. Necesita que te perdones. Tus respuestas dibujan el perfil de alguien que ha aprendido a exigirse. Alguien que ha aprendido a juzgarse. Alguien que ha aprendido a no perdonarse. Y esa exigencia, aunque te ha hecho crecer, también te ha hecho daño. Los ángeles que te acompañan no te piden que te juzgues. Te piden que te perdones. Porque un alma que no se perdona no puede avanzar. Y vos necesitás avanzar. Hay una imagen que aparece en tus respuestas: la de alguien mirándose al espejo con dureza. Pero el espejo no muestra la verdad. Muestra lo que creés de vos. Y lo que creés de vos no es lo que sos. No tenés que castigarte para ser mejor. A veces, el acto más valiente es perdonarte. Esta semana, elegí un momento tranquilo y escribí una carta de perdón hacia vos mismo. No para justificar lo que hiciste. Para reconocer que hiciste lo mejor que pudiste con lo que tenías. Escribila a mano. Y si al terminarla sentís paz, aunque sea un poco, es señal de que estás perdonando. No estás solo. Los ángeles que te acompañan no te piden que te juzgues. Te piden que te perdones. Perdonate. No estás fallando. Estás volviendo a vos.",
  "mensaje_extendido": "La Justicia te invita a equilibrar. A perdonar. A soltar la culpa. Esta carta te dice que no estás en deuda. Que no tenés que pagar nada. Que podés soltar el peso. Permitite equilibrar. Permitite perdonar. Permitite perdonarte. Y recordá que el perdón no es para el otro. Es para vos."
 },
 {
  "id": 20,
  "nombre": "El Que Camina",
  "carta": "El Carro",
  "significado": "Avance, dirección, victoria sobre la oscuridad.",
  "texto": "Hay algo en vos que sigue caminando. Aunque no veas el final. Aunque no sepas a dónde vas. Aunque el camino sea largo y a veces solitario. Y esa perseverancia, aunque es una forma de fe, también es una forma de cansancio. No estás solo. Y no tenés que ver el final para seguir. Lo que sentís no es debilidad. Es amor que sigue fluyendo, incluso cuando ya no encuentra a quién abrazar. Pero ese amor también tiene que fluir hacia vos. Tu alma no necesita que veas el final. Necesita que confíes en el camino. Tus respuestas dibujan el perfil de alguien que ha aprendido a caminar sin ver. Alguien que ha aprendido a confiar en el proceso. Alguien que ha aprendido a seguir aunque no sepa a dónde va. Y esa fe, aunque es un regalo, también es un peso. Los ángeles que te acompañan no te piden que veas el final. Te piden que confíes en el camino. Porque un alma que camina sin confiar se agota. Y vos necesitás descansar. Hay una imagen que aparece en tus respuestas: la de alguien caminando por un sendero de noche. No ve el final. No ve el principio. Solo ve el próximo paso. Y el próximo paso es suficiente. No necesitás ver el final del camino. Solo necesitás dar el próximo paso. Esta semana, elegí un momento del día para caminar. No para llegar a algún lado. Solo para caminar. Sentí el aire, el sol, el suelo bajo tus pies. Y si aparece la ansiedad, respirá y decite: \"Esto también es parte del camino.\" No estás solo. Los ángeles que te acompañan no te piden que veas el final. Te piden que confíes en el camino. Caminá. No estás fallando. Estás volviendo a vos.",
  "mensaje_extendido": "El Carro te invita a avanzar. A confiar en tu dirección. A reconocer que ya estás en movimiento. Esta carta te dice que no estás estancado. Que estás avanzando, aunque no lo veas. Que cada paso cuenta. Permitite avanzar. Permitite confiar. Permitite que el carro te lleve. Y recordá que, aunque el camino sea largo, ya estás en movimiento."
 },
 {
  "id": 21,
  "nombre": "El Que Sana",
  "carta": "La Torre",
  "significado": "Transformación abrupta, liberación, reconstrucción.",
  "texto": "Hay algo en vos que está sanando. Una parte de tu alma que está cerrando heridas, que está soltando dolores, que está volviendo a casa. Y esa sanación, aunque es lenta, es sagrada. No estás solo. Y no tenés que sanar de un día para el otro. Lo que sentís no es debilidad. Es amor que sigue fluyendo, incluso cuando ya no encuentra a quién abrazar. Pero ese amor también tiene que fluir hacia vos. Tu alma no necesita que te apures. Necesita que confíes en el proceso. Tus respuestas dibujan el perfil de alguien que está sanando. Alguien que ha atravesado dolores. Alguien que ha aprendido a mirarse. Alguien que ha decidido volver a casa. Y esa sanación, aunque es un proceso, ya empezó. Los ángeles que te acompañan no te piden que sanes más rápido. Te piden que confíes en el proceso. Porque un alma que se apura a sanar termina reabriendo heridas. Y vos necesitás tiempo. Hay una imagen que aparece en tus respuestas: la de una herida que está cerrando. No está cerrada todavía. Pero ya no sangra. Y eso, aunque parece poco, es un milagro. No tenés que sanar de un día para el otro. A veces, el acto más valiente es confiar en el proceso. Esta semana, elegí un momento tranquilo y escribí una carta a tu herida. No para justificarla. Para reconocerla. Decile que la ves. Decile que la estás cuidando. Decile que va a sanar. Y si al terminarla sentís paz, aunque sea un poco, es señal de que estás sanando. No estás solo. Los ángeles que te acompañan no te piden que sanes más rápido. Te piden que confíes en el proceso. Saná. No estás fallando. Estás volviendo a vos.",
  "mensaje_extendido": "La Torre te invita a liberarte. A soltar lo que ya no te sirve. A reconstruir sobre nuevas bases. Esta carta te dice que lo que se cae no es tu vida. Es una versión de tu vida que ya no te representa. Permitite que se caiga. Permitite liberarte. Permitite reconstruir. Y confiá en que lo que viene es más liviano, más verdadero, más tuyo."
 },
 {
  "id": 22,
  "nombre": "El Que Vuelve a Casa",
  "carta": "El Diablo",
  "significado": "Liberación de ataduras, sombra, regreso a la luz.",
  "texto": "Hay algo en vos que sabe que esta vida no es el final. No lo sabés con la mente, lo sabés con el alma. Es una certeza silenciosa que te acompaña incluso en los momentos más difíciles. Y esa certeza, aunque a veces no sepas de dónde viene, es uno de tus mayores tesoros. No estás perdido. No estás lejos. Estás volviendo. Cada paso que das, cada decisión que tomás, cada dolor que atraviesas, te acerca un poco más a casa. A ese lugar del que viniste y al que todos volveremos. Lo que sentís no es debilidad. Es amor que sigue fluyendo, incluso cuando ya no encuentra a quién abrazar. Es tu alma recordándote que todo esto es un viaje, y que el viaje tiene un destino. Tus respuestas dibujan el perfil de alguien que intuye que la vida es más grande de lo que parece. Alguien que, incluso en medio del dolor, siente que hay un orden. No un orden lógico, no un orden que se pueda explicar, sino un orden profundo, un orden del alma. Los ángeles que te acompañan no te hablan de castigos ni de pruebas. Te hablan de regreso. Te hablan de un hogar que nunca dejó de esperarte. Te hablan de un amor que no se rompe, que no se pierde, que no se estanca. Un amor que fluye, que trasciende, que es la fuente de todo. Muchas tradiciones hablan de la vida como un río que vuelve al mar. Y vos, por tus respuestas, parecés intuir eso. Parecés saber que todo lo que vivís, lo bueno y lo malo, es parte de un viaje que tiene un destino. Y el destino es volver a casa. Hay una imagen que aparece en tus respuestas: la de alguien caminando por un sendero largo, con el sol cayendo en el horizonte. No está apurado. No está perdido. Está volviendo. Y sabe que, aunque el camino sea largo, cada paso lo acerca. No estás lejos. Estás volviendo. Y todo lo que vivís es parte del camino a casa. Esta semana, elegí un momento del día para hacer silencio. No para meditar, no para rezar, no para pedir nada. Solo para estar. Sentate en un lugar tranquilo, cerrá los ojos, y respirá. Y mientras respirás, repetí internamente: \"Estoy volviendo a casa.\" No hace falta que sientas nada especial. Solo repetilo. Y dejá que las palabras hagan su trabajo en silencio. No estás solo. Los ángeles que te acompañan, los guías que te cuidan, los seres amados que te velan, todos están con vos. No para darte un mapa, sino para recordarte que el camino ya está trazado. No estás lejos. Estás volviendo. Y el hogar siempre te espera.",
  "mensaje_extendido": "El Diablo te invita a mirar tu sombra. A reconocer tus ataduras. A liberarte. Esta carta te dice que no estás condenado. Que la oscuridad no es tu destino. Que podés salir. Permitite mirar. Permitite reconocer. Permitite liberarte. Y recordá que la luz siempre está más cerca de lo que creés."
 }
];
const MENSAJE_DUELO = "Cada alma tiene su tiempo. El consuelo es difícil, pero la certeza de que nos reencontraremos es la mejor forma de canalizar esa energía. La espera no es eterna. La vida fluye, y es tu momento de vivir y vibrar en armonía. Los que no están, están. Te observan, te apoyan, te acompañan. Y la mejor forma de honrarlos es teniendo una buena vida.";
