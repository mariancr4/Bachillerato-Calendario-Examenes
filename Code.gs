// =============================================================
// Calendario de Exámenes - Escuelas Salesianas María Auxiliadora
// Backend (Google Apps Script) vinculado a la hoja de cálculo
// =============================================================

// Zona horaria usada para interpretar y formatear fechas.
// Debe coincidir con la zona horaria de la hoja de cálculo
// (Archivo > Configuración de la hoja de cálculo).
var TIMEZONE = "Europe/Madrid";

// Columnas de la hoja, en este orden exacto (fila 1 = cabeceras):
// id | subject | teacher | className | itinerary | date | content | eventId
var COL = { ID: 0, SUBJECT: 1, TEACHER: 2, CLASSNAME: 3, ITINERARY: 4, DATE: 5, CONTENT: 6, EVENT_ID: 7 };

// IDs reales de los 4 calendarios de Google (uno por clase).
var CALENDAR_IDS = {
  "1º Bach A": "c_8c3993fe0d844b4f3fa5d8a18a30ee96242c3576bfc158c11aeb2fa67bd2586a@group.calendar.google.com",
  "1º Bach B": "c_de47733cf4131f0de112e04f086f1a58dcc94387743806a10291ad9c5dae4c1e@group.calendar.google.com",
  "2º Bach A": "c_f156f263656cdc96ad68d5212bbc1a1d4485f52bc70c7a3bca53dd7545f73934@group.calendar.google.com",
  "2º Bach B": "c_44edb9cfb0a13ebac79e26eea7f910afe300725fea3722bc415c10c33e377572@group.calendar.google.com"
};

function getSheet_() {
  // Se usa siempre la primera hoja del libro, independientemente de
  // cuál esté "activa" en el editor en cada momento.
  return SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
}

// Convierte cualquier valor de fecha (string "yyyy-MM-dd" u objeto Date
// que Sheets haya creado automáticamente) a un string "yyyy-MM-dd"
// estable en la zona horaria del centro. Evita el desfase de un día
// que se produce al serializar objetos Date de Sheets como UTC.
function toDateString_(value) {
  if (!value) return "";
  if (Object.prototype.toString.call(value) === "[object Date]") {
    return Utilities.formatDate(value, TIMEZONE, "yyyy-MM-dd");
  }
  var str = String(value).trim();
  if (str.indexOf("T") !== -1) return str.split("T")[0];
  return str.substring(0, 10);
}

// Construye, a partir de un string "yyyy-MM-dd", un Date que representa
// el mediodía de ese día en la zona horaria del centro (evita saltos de
// día al crear eventos de todo el día).
function parseExamDate_(dateStr) {
  var parts = dateStr.split("-").map(Number);
  return new Date(parts[0], parts[1] - 1, parts[2], 12, 0, 0);
}

function doGet() {
  var sheet = getSheet_();
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    return ContentService.createTextOutput(JSON.stringify([]))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var result = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (row[COL.ID] === "" || row[COL.ID] === null) continue;
    result.push({
      id: String(row[COL.ID]),
      subject: row[COL.SUBJECT] || "",
      teacher: row[COL.TEACHER] || "",
      className: row[COL.CLASSNAME] || "",
      itinerary: row[COL.ITINERARY] || "",
      date: toDateString_(row[COL.DATE]),
      content: row[COL.CONTENT] || "",
      eventId: row[COL.EVENT_ID] || ""
    });
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  var locked = lock.tryLock(10000);
  if (!locked) {
    return jsonOutput_({ status: "error", message: "El servidor está ocupado, inténtalo de nuevo en unos segundos." });
  }

  try {
    var contents = JSON.parse(e.postData.contents);
    var action = contents.action;
    var sheet = getSheet_();

    if (action === "create") {
      var exam = normalizeExam_(contents.exam);
      var eventId = syncCalendarEvent_(exam, null);
      sheet.appendRow([
        exam.id, exam.subject, exam.teacher, exam.className,
        exam.itinerary, exam.date, exam.content, eventId || ""
      ]);
      return jsonOutput_({ status: "success", exam: exam, eventId: eventId || "" });

    } else if (action === "update") {
      var exam = normalizeExam_(contents.exam);
      var data = sheet.getDataRange().getValues();
      var found = false;
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][COL.ID]) === String(exam.id)) {
          var oldEventId = data[i][COL.EVENT_ID];
          var newEventId = syncCalendarEvent_(exam, oldEventId);
          sheet.getRange(i + 1, 1, 1, 8).setValues([[
            exam.id, exam.subject, exam.teacher, exam.className,
            exam.itinerary, exam.date, exam.content, newEventId || oldEventId || ""
          ]]);
          found = true;
          break;
        }
      }
      if (!found) {
        return jsonOutput_({ status: "error", message: "No se encontró el examen a actualizar (id " + exam.id + ")." });
      }
      return jsonOutput_({ status: "success", exam: exam });

    } else if (action === "delete") {
      var id = contents.id;
      var data = sheet.getDataRange().getValues();
      var found = false;
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][COL.ID]) === String(id)) {
          deleteCalendarEvent_(data[i][COL.CLASSNAME], data[i][COL.EVENT_ID]);
          sheet.deleteRow(i + 1);
          found = true;
          break;
        }
      }
      if (!found) {
        return jsonOutput_({ status: "error", message: "No se encontró el examen a eliminar (id " + id + ")." });
      }
      return jsonOutput_({ status: "success" });
    }

    return jsonOutput_({ status: "error", message: "Acción desconocida: " + action });

  } catch (err) {
    return jsonOutput_({ status: "error", message: err.toString() });
  } finally {
    lock.releaseLock();
  }
}

function normalizeExam_(exam) {
  return {
    id: String(exam.id),
    subject: exam.subject || "",
    teacher: exam.teacher || "",
    className: exam.className || "",
    itinerary: exam.itinerary || "",
    date: toDateString_(exam.date),
    content: exam.content || ""
  };
}

function jsonOutput_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function syncCalendarEvent_(exam, existingEventId) {
  var calId = CALENDAR_IDS[exam.className];
  if (!calId || !exam.date) return null;

  var calendar = CalendarApp.getCalendarById(calId);
  if (!calendar) return null;

  var examDate = parseExamDate_(exam.date);
  var title = "Examen: " + exam.subject + (exam.itinerary ? " (" + exam.itinerary + ")" : "");
  var description = "Profesor/a: " + exam.teacher + "\nContenido: " + (exam.content || "Sin especificar");

  if (existingEventId) {
    try {
      var event = calendar.getEventById(existingEventId);
      if (event) {
        event.setTitle(title);
        event.setAllDayDate(examDate);
        event.setDescription(description);
        return existingEventId;
      }
    } catch (err) {
      // El evento anterior ya no existe: se crea uno nuevo más abajo.
    }
  }

  var newEvent = calendar.createAllDayEvent(title, examDate, { description: description });
  return newEvent.getId();
}

function deleteCalendarEvent_(className, eventId) {
  if (!eventId) return;
  var calId = CALENDAR_IDS[className];
  if (!calId) return;

  var calendar = CalendarApp.getCalendarById(calId);
  if (!calendar) return;
  try {
    var event = calendar.getEventById(eventId);
    if (event) event.deleteEvent();
  } catch (err) {
    // El evento ya no existía; nada que borrar.
  }
}
