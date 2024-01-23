function scrapeNaerAndWriteToSheet() {
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = activeSpreadsheet.getSheetByName("國教院徵才2");

  if (!sheet) {
    console.error("找不到名為 '國教院徵才2' 的工作表");
    return;
  }

  var url = 'https://www.naer.edu.tw/PageDoc?fid=78';
  var response = UrlFetchApp.fetch(url);
  var content = response.getContentText();

  // 更新正則表達式以匹配日期、單位、標題和鏈接
  var dateRegex = /<span class="date">(.*?)<\/span>/g;
  var unitRegex = /<span class="unit">(.*?)<\/span>/g;
  var titleRegex = /<[^>]*class="txt"[^>]*title="([^"]*)"[^>]*>/g;
  var linkRegex = /<a[^>]*href="([^"]*)"[^>]*class="txt"[^>]*title="([^"]*)"[^>]*>/g;

  var dates = extractMatches(content, dateRegex);
  var units = extractMatches(content, unitRegex);
  var titles = extractMatches(content, titleRegex);
  var links = extractMatches(content, linkRegex);

  // Log extracted data for debugging
  console.log("Dates:", dates);
  console.log("Units:", units);
  console.log("Titles:", titles);
  console.log("Links:", links);

  var token = "你的權杖";

  // 遍歷提取的數據
  for (var i = 0; i < dates.length; i++) {
    // 格式化日期，僅顯示日期而不顯示時間
    var formattedDate = Utilities.formatDate(new Date(dates[i]), "GMT+8", "yyyy-MM-dd");

    // 檢查日期是否在過去的15天內
    if (isWithinLastNDays(new Date(dates[i]), 15)) {
      // 處理數值
      var processedDate = "日期：" + formattedDate;
      var processedUnit = "單位：" + units[i];
      var processedInfo = "徵才資訊：" + titles[i];
      var processedLinks = "連結：" + links[i];

      // 發送 LINE 通知
      var message = "\n" + processedDate + "\n" + processedUnit + "\n" + processedInfo + "\n" + processedLinks;
      sendline(message, token);

      // 將數據寫入工作表，僅當標題尚未存在時
      if (!isTitleAlreadyPresent(sheet, titles[i])) {
        sheet.appendRow([formattedDate, units[i], titles[i], links[i]]);
      }
    }
  }
}

function isWithinLastNDays(date, n) {
  var currentDate = new Date();
  var timeDiff = currentDate.getTime() - date.getTime();
  var daysDiff = timeDiff / (1000 * 3600 * 24);
  return daysDiff <= n;
}

function extractMatches(content, regex) {
  var matches = [];
  var match;
  while ((match = regex.exec(content)) !== null) {
    matches.push(match[1].trim());
  }
  return matches;
}

function isTitleAlreadyPresent(sheet, title) {
  var existingTitles = sheet.getRange(1, 3, sheet.getLastRow(), 1).getValues();
  for (var i = 0; i < existingTitles.length; i++) {
    if (existingTitles[i][0] === title) {
      return true; // Title already present
    }
  }
  return false; // Title not present
}

function sendline(message, token) {
  UrlFetchApp.fetch('https://notify-api.line.me/api/notify', {
    'headers': {
      'Authorization': 'Bearer ' + token,
    },
    'method': 'post',
    'payload': {
      'message': message,
    }
  });
}
