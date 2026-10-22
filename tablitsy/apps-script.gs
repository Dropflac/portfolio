/**
 * Пример скриптов для учёта заказов в Google Таблице.
 * Лист «Заказы»: A номер, B клиент, C услуга, D менеджер, E срок, F сумма,
 * G оплачено, H остаток, I статус, J почта клиента.
 */

// 1. Остаток считается сам при любой правке суммы или оплаты
function onEdit(e) {
  var sh = e.range.getSheet();
  if (sh.getName() !== "Заказы") return;
  var row = e.range.getRow();
  if (row < 2) return;
  if (e.range.getColumn() === 6 || e.range.getColumn() === 7) {
    var sum = Number(sh.getRange(row, 6).getValue()) || 0;
    var paid = Number(sh.getRange(row, 7).getValue()) || 0;
    sh.getRange(row, 8).setValue(Math.max(0, sum - paid));
  }
  // подсветка просрочки
  var due = sh.getRange(row, 5).getValue();
  var status = sh.getRange(row, 9).getValue();
  var late = due instanceof Date && due < new Date() && status !== "Готов";
  sh.getRange(row, 1, 1, 10).setBackground(late ? "#fff1ee" : null);
}

// 2. Напоминания должникам (запускается по таймеру каждый день в 9:00)
function remindDebtors() {
  var sh = SpreadsheetApp.getActive().getSheetByName("Заказы");
  var data = sh.getDataRange().getValues();
  var sent = 0;
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    var debt = Number(r[7]) || 0;
    if (debt > 0 && r[8] !== "Пауза" && r[9]) {
      MailApp.sendEmail(r[9], "Напоминание об оплате",
        "Здравствуйте! По заказу " + r[0] + " осталось оплатить " + debt + " ₽.");
      sent++;
    }
  }
  Logger.log("Отправлено писем: " + sent);
}

// 3. Создание ежедневного триггера (запустить один раз вручную)
function setupTrigger() {
  ScriptApp.newTrigger("remindDebtors").timeBased().everyDays(1).atHour(9).create();
}
