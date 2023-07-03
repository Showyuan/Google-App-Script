function doGet(e) {
  if (e.parameter.page === undefined) {
    return loadDashboard(e);
  } else {
    return loadPerformance(e);
  }
}

function loadDashboard(e) {
  // 取得參數 tokenId
  console.log("===開始取得參數===");
  var param = e.parameter;
  var tokenId = param.tokenId;
  var page = param.page;
  console.log("tokenId=" + tokenId);
  console.log("page=" + page);
  console.log("===結束取得參數===");

  console.log("===開始讀取首頁===");
  var indexPage = HtmlService.createTemplateFromFile('Dashboard');
  indexPage.tokenId = tokenId;
  return indexPage.evaluate()
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setTitle('Slowly Richer')
}

function loadPerformance(e) {
  // 取得參數 tokenId
  console.log("===開始取得參數===");
  var param = e.parameter;
  var tokenId = param.tokenId;
  var page = param.page;
  console.log("tokenId=" + tokenId);
  console.log("page=" + page);
  console.log("===結束取得參數===");

  console.log("===開始讀取績效頁===");
  var performancePage = HtmlService.createTemplateFromFile('Performance');
  performancePage.tokenId = tokenId;
  performancePage.name = getCustomer(tokenId);
  return performancePage.evaluate()
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setTitle('Slowly Richer')
}

function getCustomer(tokenId) {
  // 讀取Excel投資人資訊
  console.log("===讀取Excel投資人資訊===");
  var spreadsheet = SpreadsheetApp.openById('xxxxxx');
  var sheet = spreadsheet.getSheetByName("投資人資訊");

  var _tokenId;
  var lastRow = sheet.getLastRow();    // 取LastRow

  for (let i = 2; i <= lastRow; i++) {
    _tokenId = sheet.getRange(i, 3).getValue();  // 取TokenId
    if (_tokenId === tokenId) {
      console.log("===完成讀取Excel投資人資訊===");
      return sheet.getRange(i, 1).getValue();
    }
  }
}

function getCustomerData(tokenId) {
  // 讀取Excel客戶權益
  console.log("===讀取Excel客戶權益===");
  var spreadsheet = SpreadsheetApp.openById('xxxxxx');
  var sheet = spreadsheet.getSheetByName("入金記錄");

  var _tokenId;
  var lastRow = sheet.getLastRow();    // 取LastRow
  var tokenIdColIndex = 5;             // column index of tokenId

  console.log(lastRow);

  // 取的投資記錄 By tokenId
  var records = [];
  for (let i = 2; i <= lastRow; i++) {

    _tokenId = sheet.getRange(i, 6).getValue();  // 取TokenId
    _investMoney = sheet.getRange(i, 8).getValue();  // 取入金單位

    if (_tokenId === tokenId && _investMoney > 0) {
      console.info(_tokenId);

      var a = [];

      var d = new Date(sheet.getRange(i, tokenIdColIndex - 1).getValue());
      var difference = Math.abs(new Date() - d);
      days = parseInt(difference / (1000 * 3600 * 24));
      a.push(d.getFullYear() + "-" + (d.getMonth() + 1) + "-" + d.getDate());
      a.push(sheet.getRange(i, tokenIdColIndex + 2).getValue().toLocaleString('en-US'));
      a.push(sheet.getRange(i, tokenIdColIndex + 5).getValue().toLocaleString('en-US'));
      a.push((sheet.getRange(i, tokenIdColIndex + 6).getValue() * 100).toFixed(2) + '%');
      a.push(days);

      records.push(a);
    }
  }
  console.log("===完成讀取Excel客戶權益===");
  return records;
}

function getScriptUrl() {
  var url = ScriptApp.getService().getUrl();
  console.log(url);
  return url;
}

function getIndexData() {

  var retVal = [];

  // 讀取Excel趨勢資訊
  console.log("===讀取Excel趨勢資訊===");
  var id = 'xxxxxx';
  var spreadsheet = SpreadsheetApp.openById(id);
  var sheet = spreadsheet.getSheetByName("DASHBOARD");

  // 指數趨勢總覽
  console.log("===指數趨勢總覽===");
  var rowsToLoop = 9;  // 指數趨勢總覽資料的LastRow
  var startCell = 1;  // startCell
  var indexOverviewRecords = [];

  for (let i = 3; i <= rowsToLoop; i++) {
    var a = [];
    a.push(sheet.getRange(i, startCell).getValue());
    a.push(sheet.getRange(i, startCell + 1).getValue());
    a.push(sheet.getRange(i, startCell + 5).getValue());
    a.push(sheet.getRange(i, startCell + 6).getValue());
    a.push(sheet.getRange(i, startCell + 7).getValue());
    a.push(sheet.getRange(i, startCell + 8).getValue());
    a.push((sheet.getRange(i, startCell + 9).getValue() * 100).toFixed(2) + '%');
    a.push((sheet.getRange(i, startCell + 10).getValue() * 100).toFixed(2) + '%');
    a.push((sheet.getRange(i, startCell + 11).getValue() * 100).toFixed(2) + '%');
    indexOverviewRecords.push(a);
  }
  retVal.push(indexOverviewRecords);

  // 台股權值
  console.log("===台股權值===");
  rowsToLoop = 22;  // 台股權值總覽資料的LastRow
  var taiWeightedIndexRecords = [];

  for (let i = 13; i <= rowsToLoop; i++) {
    var a = [];
    a.push(sheet.getRange(i, startCell).getValue());
    a.push(sheet.getRange(i, startCell + 1).getValue());
    a.push(sheet.getRange(i, startCell + 5).getValue());
    a.push(sheet.getRange(i, startCell + 6).getValue());
    a.push(sheet.getRange(i, startCell + 7).getValue());
    a.push(sheet.getRange(i, startCell + 8).getValue());
    a.push((sheet.getRange(i, startCell + 9).getValue() * 100).toFixed(2) + '%');
    a.push((sheet.getRange(i, startCell + 10).getValue() * 100).toFixed(2) + '%');
    a.push((sheet.getRange(i, startCell + 11).getValue() * 100).toFixed(2) + '%');

    taiWeightedIndexRecords.push(a);
  }
  retVal.push(taiWeightedIndexRecords);

  // 美股權值
  console.log("===美股權值===");
  rowsToLoop = 37;  // 美股權值總覽資料的LastRow
  var usWeightedIndexRecords = [];

  for (let i = 26; i <= rowsToLoop; i++) {
    var a = [];
    a.push(sheet.getRange(i, startCell).getValue());
    a.push(sheet.getRange(i, startCell + 1).getValue());
    a.push(sheet.getRange(i, startCell + 5).getValue());
    a.push(sheet.getRange(i, startCell + 6).getValue());
    a.push(sheet.getRange(i, startCell + 7).getValue());
    a.push(sheet.getRange(i, startCell + 8).getValue());
    a.push((sheet.getRange(i, startCell + 9).getValue() * 100).toFixed(2) + '%');
    a.push((sheet.getRange(i, startCell + 10).getValue() * 100).toFixed(2) + '%');
    a.push((sheet.getRange(i, startCell + 11).getValue() * 100).toFixed(2) + '%');

    usWeightedIndexRecords.push(a);
  }
  retVal.push(usWeightedIndexRecords);

  console.log("===完成讀取Excel趨勢資訊===");
  return retVal;
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
