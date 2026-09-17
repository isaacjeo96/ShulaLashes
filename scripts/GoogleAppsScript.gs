const CALENDAR_ID = 'primary';
const SHEET_NAME = 'Solicitudes';

function doPost(e) {
  const data = e.parameter;
  if (!data.name || !data.phone || !data.service || !data.date || !data.time) {
    return jsonResponse({ ok: false, error: 'Faltan campos obligatorios.' });
  }
  if (data.website) return jsonResponse({ ok: true });

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error(`No existe la hoja "${SHEET_NAME}".`);
  sheet.appendRow([new Date(), data.name, data.phone, data.service, data.date, data.time, data.notes || '', 'Pendiente']);
  return jsonResponse({ ok: true });
}

function jsonResponse(payload) {
  return ContentService.createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
