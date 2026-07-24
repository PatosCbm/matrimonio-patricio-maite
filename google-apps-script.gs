/**
 * RSVP · Patricio & Maite  →  Google Sheets
 * --------------------------------------------------------------
 * PASOS:
 * 1) Crea un Google Sheet nuevo (ej: "Confirmaciones Boda").
 * 2) En ese Sheet: Extensiones → Apps Script.
 * 3) Borra lo que haya y pega TODO este archivo. Guarda.
 * 4) Implementar → Nueva implementación → tipo "Aplicación web":
 *      - Ejecutar como: Yo (tu cuenta)
 *      - Quién tiene acceso: Cualquier usuario
 *    Implementar → Autorizar permisos.
 * 5) Copia la "URL de la aplicación web" (termina en /exec)
 *    y pásasela a Claude (o pégala en confirmar.html donde dice
 *    REEMPLAZAR_CON_TU_URL_DEL_APPS_SCRIPT).
 *
 * Cada envío agrega/actualiza una fila. Si un invitado confirma
 * dos veces (mismo ID), se ACTUALIZA su fila en vez de duplicar.
 */

var SHEET_NAME = 'Respuestas';
var HEADERS = ['Fecha/hora','ID','Asiste','Nombre invitado','Con acompañante',
               'Nombre acompañante','Preferencia alimentaria','Alergia','Detalle alergia'];

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(15000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) { sheet = ss.insertSheet(SHEET_NAME); }
    if (sheet.getLastRow() === 0) { sheet.appendRow(HEADERS); }

    var d = JSON.parse(e.postData.contents);
    var row = [
      new Date(),
      d.id || '',
      d.asiste || '',
      d.nombre || '',
      d.conAcomp || '',
      d.acompNombre || '',
      d.preferencia || '',
      d.alergia || '',
      d.alergiaDetalle || ''
    ];

    // Upsert por ID (columna B): si ya existe, actualiza esa fila.
    var updated = false;
    if (d.id && sheet.getLastRow() >= 2) {
      var ids = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues();
      for (var i = 0; i < ids.length; i++) {
        if (String(ids[i][0]) === String(d.id)) {
          sheet.getRange(i + 2, 1, 1, row.length).setValues([row]);
          updated = true;
          break;
        }
      }
    }
    if (!updated) { sheet.appendRow(row); }

    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return json({ ok: true, msg: 'RSVP endpoint activo' });
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
