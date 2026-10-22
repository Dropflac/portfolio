/**
 * Google Apps Script для сайта сервиса ремонта.
 * Лист «Заявки»: дата и время, имя, телефон, модель, проблема, статус.
 * Лист «Прайс»: устройство, услуга, срок, цена.
 *
 * Как подключить:
 * 1. В таблице: Расширения, Apps Script, вставить этот код.
 * 2. Заменить NOTIFY_EMAIL на почту владельца.
 * 3. Развернуть как веб-приложение (доступ: всем), адрес вставить на сайт.
 */
var NOTIFY_EMAIL = "owner@example.com";

function doPost(e) {
  var p = e.parameter;
  if (p.site) return json_({ ok: true }); // honeypot: поле заполнил бот
  var sheet = SpreadsheetApp.getActive().getSheetByName("Заявки");
  var now = Utilities.formatDate(new Date(), "Europe/Moscow", "dd.MM.yyyy HH:mm");
  sheet.appendRow([now, p.name, p.phone, p.model, p.issue || "не указано", "Новая"]);
  MailApp.sendEmail(NOTIFY_EMAIL, "Новая заявка на ремонт",
    p.name + "\n" + p.phone + "\n" + p.model + "\n" + (p.issue || ""));
  return json_({ ok: true });
}

function doGet() {
  var rows = SpreadsheetApp.getActive().getSheetByName("Прайс").getDataRange().getValues();
  rows.shift(); // заголовок
  return json_(rows.filter(function (r) { return r[0] !== ""; }));
}

function json_(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
