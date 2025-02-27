function sendReminderEmails() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Tên_Sheet_Của_Bạn"); // Thay tên sheet
  if (!sheet) {
    Logger.log("Lỗi: Không tìm thấy sheet.");
    return;
  }

  var sheetId = sheet.getSheetId();
  var spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  var lastRow = sheet.getLastRow(); // Lấy hàng cuối có dữ liệu
  if (lastRow < 2) return; // Nếu không có dữ liệu thì thoát

  var data = sheet.getRange(2, 1, lastRow - 1, 11).getValues(); // Chỉ lấy dữ liệu cần thiết
  var now = new Date();
  now.setHours(0, 0, 0, 0);

  var subject = "🔔 Nhắc nhở công việc sắp đến hạn";
  var emailQueue = {}; // Gom email theo từng người nhận
  var maxEmailsPerRun = 50;
  var emailCount = 0;
  var cache = CacheService.getScriptCache(); // Lưu trạng thái đã nhắc

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var taskName = row[3];
    var duedate = new Date(row[4]);
    var status = row[5];
    var email = row[6];
    var note = row[9];
    var disableReminder = row[10]; // Cột "Tắt nhắc nhở" (Có/Không)

    if (!email || !duedate || status === "Hoàn thành" || disableReminder === "Có") continue;

    var daysLeft = Math.ceil((duedate - now) / (1000 * 60 * 60 * 24));
    var cellReference = "D" + (i + 2);
    var sheetUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit#gid=${sheetId}&range=${cellReference}`;

    var cacheKey = "reminded_" + taskName;
    var wasReminded = cache.get(cacheKey);

    if (daysLeft <= 7 && daysLeft >= 0 && !wasReminded) {
      var message = `
        Xin chào,<br><br>
        Bạn có nhiệm vụ <b>"${taskName}"</b> cần hoàn thành trước ngày <b>${duedate.toLocaleDateString()}</b>.<br>
        🔗 <a href="${sheetUrl}">Xem chi tiết trong Google Sheet</a> <br>
        🕒 <b>Thời gian còn lại:</b> ${daysLeft} ngày <br>
        📌 <b>Ghi chú:</b> ${note ? note : "Không có"} <br><br>
        Vui lòng hoàn thành sớm để tránh trễ hạn.<br><br>
        Cảm ơn!
      `;

      if (!emailQueue[email]) emailQueue[email] = [];
      emailQueue[email].push(message);

      cache.put(cacheKey, "true", 86400); // Lưu trạng thái trong 24 giờ
    }

    if (daysLeft === 1) {
      var urgentSubject = "⚠️ [Khẩn cấp] Công việc sắp hết hạn!";
      var urgentMessage = `
        🔴 Cảnh báo! <br><br>
        Nhiệm vụ <b>"${taskName}"</b> sẽ hết hạn vào ngày mai <b>${duedate.toLocaleDateString()}</b>.<br>
        🚨 Hãy hoàn thành ngay để tránh bị quá hạn! <br>
        🔗 <a href="${sheetUrl}">Xem chi tiết</a> <br><br>
        📌 <b>Ghi chú:</b> ${note ? note : "Không có"} <br>
        ⏳ <b>Trạng thái:</b> ${status} <br><br>
        Hành động ngay để tránh bị trễ hạn! 🚀
      `;

      MailApp.sendEmail({ to: email, subject: urgentSubject, htmlBody: urgentMessage });
    }
  }

  for (var recipient in emailQueue) {
    if (emailCount >= maxEmailsPerRun) break;
    var combinedMessage = emailQueue[recipient].join("<br><hr><br>");
    MailApp.sendEmail({ to: recipient, subject: subject, htmlBody: combinedMessage });
    emailCount++;
  }
}

function notifyTimeChange() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Tên_Sheet_Của_Bạn");
  if (!sheet) {
    Logger.log("Lỗi: Không tìm thấy sheet.");
    return;
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  var data = sheet.getRange(2, 1, lastRow - 1, 11).getValues();

  var cache = CacheService.getScriptCache();
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var taskName = row[3];
    var duedate = new Date(row[4]);
    var email = row[6];
    var note = row[9];

    if (!email || !duedate) continue;

    var cacheKey = "task_" + i;
    var cachedDate = cache.get(cacheKey);
    var newDueDate = duedate.toISOString().split('T')[0];

    if (cachedDate && cachedDate !== newDueDate) {
      var message = `
        Xin chào,<br><br>
        Nhiệm vụ <b>"${taskName}"</b> của bạn đã có thay đổi thời gian.<br>
        🔄 <b>Thời gian mới:</b> ${duedate.toLocaleDateString()} <br>
        📌 <b>Ghi chú:</b> ${note ? note : "Không có"} <br><br>
        Hãy kiểm tra lại lịch trình của bạn để tránh nhầm lẫn.  
      `;

      MailApp.sendEmail({ to: email, subject: "🔄 Cập nhật thời gian nhiệm vụ", htmlBody: message });
    }

    cache.put(cacheKey, newDueDate, 21600);
  }
}
