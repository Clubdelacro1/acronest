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

    // Confirmación a la persona
    var es = d.language === 'es';
    MailApp.sendEmail({
      to: d.email,
      replyTo: ORG_EMAIL,
      name: 'Acronest',
      subject: es ? 'Acronest 2027 · Recibimos tu inscripción' : 'Acronest 2027 · We received your registration',
      body: (es
        ? 'Hola ' + d.name + ',\n\n¡Gracias por inscribirte a Acronest 2027! Tu plaza queda reservada cuando recibamos la seña de 100 €.\n\n'
          + 'Titular: Joaquín Ricardo Castillo Hadler\nIBAN: ES45 2100 0365 3202 0022 3160\nImporte: 100 €\nConcepto: ACRONEST 2027 · ' + d.name + '\n\n'
          + 'Cuando hayas pagado, responde a este correo con el comprobante o envíalo por WhatsApp:\n• Rober: +54 9 11 2335 7818\n• Conitrini: +56 9 9841 9773\n\n'
          + 'Nos vemos en el Nido.\nUn abrazo grande, Coni y Rob.'
        : 'Hi ' + d.name + ',\n\nThank you for registering for Acronest 2027! Your place is reserved once we receive the €100 deposit.\n\n'
          + 'Account holder: Joaquín Ricardo Castillo Hadler\nIBAN: ES45 2100 0365 3202 0022 3160\nAmount: €100\nReference: ACRONEST 2027 · ' + d.name + '\n\n'
          + 'Once you have paid, reply to this email with the receipt or send it via WhatsApp:\n• Rober: +54 9 11 2335 7818\n• Conitrini: +56 9 9841 9773\n\n'
          + 'See you at the Nest.\nBig hug, Coni & Rob.')
    });

    return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}
