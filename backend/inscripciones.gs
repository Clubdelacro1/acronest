/**
 * Acronest 2027 · Receptor de inscripciones
 * Proyecto de Apps Script en la cuenta acronest.fest@gmail.com, publicado como aplicación web.
 * Cada inscripción de la web se guarda como una fila nueva, y llega un mail a la organización.
 * A la persona inscripta le llega un mail con los datos para transferir la seña.
 */
var ORG_EMAIL = 'acronest.fest@gmail.com';
var SHEET_ID = '1ocdqg-a7vCwrWuOomPVuyY1W91WDPQIXpQ4oNOfYOw4'; // planilla "Acronest 2027 · Inscripciones"

var COLUMNS = [
  ['date', 'Fecha'], ['name', 'Nombre'], ['email', 'Email'], ['whatsapp', 'WhatsApp'], ['instagram', 'Instagram'],
  ['from', 'De dónde'], ['ticket', 'Entrada'], ['nest', 'Nido'], ['price', 'Precio (€)'], ['partner', 'Pareja dúo'],
  ['bed', 'Cama'], ['share_with', 'Comparte cama con'], ['disciplines', 'Disciplinas'], ['level', 'Nivel'], ['role', 'Rol'],
  ['experience', 'Experiencia'], ['diet', 'Comida'], ['allergies', 'Alergias'], ['source', 'Cómo se enteró'],
  ['payment', 'Forma de pago'], ['payment_other', 'Pago (otro)'], ['terms', 'Condiciones'], ['language', 'Idioma'],
  ['deposit', 'Seña recibida'], ['notes', 'Notas']
];

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var d = JSON.parse(e.postData.contents);
    d.date = new Date();
    var sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
    if (sheet.getRange(1, 1).getValue() !== 'Fecha') {
      if (sheet.getLastRow() > 0) sheet.insertRowBefore(1);
      sheet.getRange(1, 1, 1, COLUMNS.length).setValues([COLUMNS.map(function (c) { return c[1]; })]);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, COLUMNS.length).setFontWeight('bold');
    }
    sheet.appendRow(COLUMNS.map(function (c) { return d[c[0]] == null ? '' : d[c[0]]; }));

    // Aviso a la organización
    MailApp.sendEmail({
      to: ORG_EMAIL,
      replyTo: d.email,
      subject: 'Nueva inscripción Acronest 2027 · ' + d.name + ' · ' + d.nest,
      body: COLUMNS.filter(function (c) { return d[c[0]]; })
        .map(function (c) { return c[1] + ': ' + d[c[0]]; }).join('\n')
    });

    // Bienvenida a la persona, con los datos para pagar la seña
    var mail = welcomeEmail(d);
    MailApp.sendEmail({
      to: d.email,
      replyTo: ORG_EMAIL,
      name: 'Acronest',
      subject: mail.subject,
      body: mail.text,
      htmlBody: mail.html
    });

    return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

/** Mail de bienvenida (español o inglés) con los datos para pagar la seña. Devuelve {subject, text, html}. */
function welcomeEmail(d) {
  var es = d.language === 'es';
  var BANK = { holder: 'Joaquín Ricardo Castillo Hadler', iban: 'ES45 2100 0365 3202 0022 3160' };
  var ref = 'ACRONEST 2027 · ' + (d.name || '');
  var NESTS_ES = { 'Outdoor nest': 'Nido exterior', '8-person nest': 'Nido de 8', '4-person nest': 'Nido de 4', '2-person nest': 'Nido de 2', 'Love nest': 'Love Nest' };
  var NESTS_EN = { 'Outdoor nest': 'Outdoor Nest', '8-person nest': '8-Person Nest', '4-person nest': '4-Person Nest', '2-person nest': '2-Person Nest', 'Love nest': 'Love Nest' };
  var nest = (es ? NESTS_ES : NESTS_EN)[d.nest] || d.nest || '';
  var price = Number(d.price) || 0;
  var rest = price > 100 ? price - 100 : 0;
  var pay = String(d.payment || '');
  var restText = pay.indexOf('cash') > -1
    ? (es ? 'en efectivo al llegar al evento' : 'in cash when you arrive at the event')
    : pay.indexOf('installments') > -1
      ? (es ? 'en 3 cuotas por transferencia' : 'in 3 instalments by bank transfer')
      : (es ? 'como lo acordemos contigo' : 'as we agree with you');
  var esc = function (s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); };
  var money = function (n) { return es ? n + ' €' : '€' + n; };
  var T = es ? {
    subject: '¡Te damos la bienvenida al Nido! · Acronest 2027',
    title: '¡Bienvenida al Nido!',
    hi: 'Hola ' + esc(d.name) + ',',
    p1: '¡Qué alegría que vengas! Ya eres parte de <b>Acronest 2027</b>, nuestra 4ª edición: cinco días de acro, movimiento, naturaleza y comunidad en el Cortijo Casería del Mercado, en Las Alpujarras (Granada).',
    p2: 'Del <b>7 al 11 de julio de 2027</b> vamos a compartir talleres, jams, masajes, comida rica y noches bajo las estrellas. Lo estamos preparando con muchísimo cariño para que te sientas en casa desde el primer momento.',
    summary: 'Tu inscripción', lNest: 'Nido', lPrice: 'Precio',
    payTitle: 'Reserva tu plaza con la seña de 100 €',
    payText: 'Tu plaza queda confirmada cuando recibimos la seña. Las plazas son limitadas, así que te recomendamos hacer la transferencia en los próximos días.',
    lHolder: 'Titular', lAmount: 'Importe', lRef: 'Concepto',
    restLine: rest ? 'El resto (' + money(rest) + ') lo pagas ' + restText + '.' : '',
    receipt: 'Cuando hayas pagado, envíanos el comprobante por WhatsApp a Rober o a Coni:',
    groupTitle: 'Únete a la Comunidad Acronest', groupText: 'En este grupo de WhatsApp compartimos toda la información del evento: cronograma, profes, cómo llegar y novedades.', groupBtn: 'Unirme al grupo',
    next: 'Y también puedes seguirnos en Instagram <a href="https://instagram.com/acro.nest" style="color:#c9791c">@acro.nest</a>.',
    bye: '¡Nos vemos en el Cortijo!', sign: 'Un abrazo grande,<br>Coni y Rob'
  } : {
    subject: 'Welcome to the Nest! · Acronest 2027',
    title: 'Welcome to the Nest!',
    hi: 'Hi ' + esc(d.name) + ',',
    p1: 'We are so happy you are coming! You are now part of <b>Acronest 2027</b>, our 4th edition: five days of acro, movement, nature and community at Cortijo Casería del Mercado, in Las Alpujarras (Granada, Spain).',
    p2: 'From <b>7 to 11 July 2027</b> we will share workshops, jams, massage, delicious food and nights under the stars. We are preparing it with lots of love so you feel at home from the very first moment.',
    summary: 'Your registration', lNest: 'Nest', lPrice: 'Price',
    payTitle: 'Secure your place with the €100 deposit',
    payText: 'Your place is confirmed once we receive the deposit. Places are limited, so we recommend making the transfer in the next few days.',
    lHolder: 'Account holder', lAmount: 'Amount', lRef: 'Reference',
    restLine: rest ? 'You pay the balance (' + money(rest) + ') ' + restText + '.' : '',
    receipt: 'Once you have paid, send us the receipt via WhatsApp to Rober or Coni:',
    groupTitle: 'Join the Acronest Community', groupText: 'In this WhatsApp group we share all the event information: schedule, teachers, how to get there and news.', groupBtn: 'Join the group',
    next: 'You can also follow us on Instagram <a href="https://instagram.com/acro.nest" style="color:#c9791c">@acro.nest</a>.',
    bye: 'See you at the Cortijo!', sign: 'Big hug,<br>Coni &amp; Rob'
  };
  var strip = function (s) { return String(s).replace(/<br>/g, '\n').replace(/<[^>]+>/g, '').replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"'); };
  var row = function (k, v, strong) {
    return '<tr><td style="padding:7px 0;color:#8a7f73;font-size:14px">' + k + '</td>'
      + '<td style="padding:7px 0;text-align:right;color:#1a1612;font-size:' + (strong ? '17px;font-weight:700' : '15px') + '">' + v + '</td></tr>';
  };
  var GROUP = 'https://chat.whatsapp.com/FnUOCjzmVDw3oRSO2stZeC';
  var wa = '<a href="https://wa.me/5491123357818" style="color:#c9791c;text-decoration:none">Rober · +54 9 11 2335 7818</a><br>'
    + '<a href="https://wa.me/56998419773" style="color:#c9791c;text-decoration:none">Coni · +56 9 9841 9773</a>';

  var html = ''
    + '<div style="background:#efe7dc;padding:24px 12px;font-family:Helvetica,Arial,sans-serif;color:#1a1612">'
    + '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="max-width:600px;margin:0 auto;background:#fffaf3;border-radius:14px;overflow:hidden">'
    + '<tr><td style="background:#030405;padding:30px 28px;text-align:center">'
    + '<img src="https://acronestfest.com/img/logo.png" width="84" alt="Acronest" style="display:inline-block;border:0;height:auto">'
    + '<div style="color:#eaa137;font-size:13px;letter-spacing:4px;text-transform:uppercase;margin-top:14px">Acronest 2027 · ' + (es ? '4ª edición' : '4th edition') + '</div>'
    + '<div style="color:#f6cf7e;font-family:Georgia,serif;font-style:italic;font-size:30px;margin-top:6px">' + T.title + '</div>'
    + '</td></tr>'
    + '<tr><td style="padding:30px 28px 8px;font-size:16px;line-height:1.6">'
    + '<p style="margin:0 0 14px">' + T.hi + '</p><p style="margin:0 0 14px">' + T.p1 + '</p><p style="margin:0">' + T.p2 + '</p>'
    + '</td></tr>'
    + '<tr><td style="padding:18px 28px 0"><div style="border:1px solid #eadcc8;border-radius:10px;padding:14px 18px">'
    + '<div style="font-size:12px;letter-spacing:2px;text-transform:uppercase;color:#c9791c;margin-bottom:4px">' + T.summary + '</div>'
    + '<table role="presentation" width="100%" cellpadding="0" cellspacing="0">' + row(T.lNest, esc(nest)) + (price ? row(T.lPrice, money(price)) : '') + '</table>'
    + '</div></td></tr>'
    + '<tr><td style="padding:18px 28px 0"><div style="background:#1a1612;border-radius:10px;padding:20px 20px 16px;color:#fffaf3">'
    + '<div style="font-family:Georgia,serif;font-size:21px;color:#f6cf7e;margin-bottom:6px">' + T.payTitle + '</div>'
    + '<div style="font-size:14px;line-height:1.5;color:#d9cfc2;margin-bottom:12px">' + T.payText + '</div>'
    + '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#fffaf3;border-radius:8px;padding:6px 14px">'
    + row(T.lHolder, BANK.holder) + row('IBAN', '<span style="font-family:Consolas,Menlo,monospace">' + BANK.iban + '</span>', true)
    + row(T.lAmount, money(100), true) + row(T.lRef, esc(ref)) + '</table>'
    + (T.restLine ? '<div style="font-size:14px;color:#d9cfc2;margin-top:12px">' + T.restLine + '</div>' : '')
    + '</div></td></tr>'
    + '<tr><td style="padding:20px 28px 0;font-size:15px;line-height:1.6">' + T.receipt + '<br>' + wa + '</td></tr>'
    + '<tr><td style="padding:20px 28px 0"><div style="border:1px solid #bfe3c9;background:#eef8f0;border-radius:10px;padding:18px 20px;text-align:center">'
    + '<div style="font-family:Georgia,serif;font-size:20px;color:#1a1612;margin-bottom:6px">' + T.groupTitle + '</div>'
    + '<div style="font-size:14px;line-height:1.5;color:#4a4038;margin-bottom:14px">' + T.groupText + '</div>'
    + '<a href="' + GROUP + '" style="display:inline-block;background:#25d366;color:#ffffff;text-decoration:none;font-weight:700;font-size:15px;padding:12px 26px;border-radius:999px">' + T.groupBtn + ' →</a>'
    + '</div></td></tr>'
    + '<tr><td style="padding:18px 28px 0;font-size:15px;line-height:1.6;color:#4a4038">' + T.next + '</td></tr>'
    + '<tr><td style="padding:22px 28px 30px;font-size:16px;line-height:1.6">' + T.bye + '<br>'
    + '<span style="font-family:Georgia,serif;font-style:italic;font-size:19px;color:#c9791c">' + T.sign + '</span></td></tr>'
    + '<tr><td style="background:#030405;padding:16px;text-align:center;font-size:12px;color:#8a7f73">'
    + '<a href="https://acronestfest.com" style="color:#eaa137;text-decoration:none">acronestfest.com</a> · acronest.fest@gmail.com</td></tr>'
    + '</table></div>';

  var text = strip(T.hi) + '\n\n' + strip(T.p1) + '\n\n' + strip(T.p2) + '\n\n'
    + T.summary + ': ' + nest + (price ? ' · ' + money(price) : '') + '\n\n'
    + T.payTitle + '\n' + T.payText + '\n'
    + T.lHolder + ': ' + BANK.holder + '\nIBAN: ' + BANK.iban + '\n' + T.lAmount + ': ' + money(100) + '\n' + T.lRef + ': ' + ref + '\n'
    + (T.restLine ? T.restLine + '\n' : '') + '\n'
    + T.receipt + '\n• Rober: +54 9 11 2335 7818\n• Coni: +56 9 9841 9773\n\n'
    + T.groupTitle + '\n' + T.groupText + '\n' + GROUP + '\n\n'
    + strip(T.next) + '\n\n' + T.bye + '\n' + strip(T.sign);

  return { subject: T.subject, text: text, html: html };
}
