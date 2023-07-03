// consts
const SHEET_ID = "";

// doGet
function doGet(e) {
  if (e.parameter.page === undefined) {
    return HtmlService.createTemplateFromFile('index')
      .evaluate();
  } else {
    return HtmlService.createTemplateFromFile('todo')
      .evaluate();
  }
}

// import css/js
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
    .getContent();
}

// get url
function getScriptUrl() {
  var url = ScriptApp.getService().getUrl();
  console.log(url);
  return url;
}

// ===== 下拉選單 start ===================================================

// 公司名稱
function getCompanyName() {
  var sheet = openExcelSheet("公司清單");
  var lastRow = sheet.getLastRow();
  var values = sheet.getRange('A2:A' + lastRow).getValues();
  console.log(values);
  return values;
}

// 工作項目
function getItemsAll() {
  var sheet = openExcelSheet("工作項目");
  var lastRow = sheet.getLastRow();
  var values = sheet.getRange('A2:C' + lastRow).getValues();
  console.log(values);
  return values;
}

// 員工清單
function getEmployees() {
  var sheet = openExcelSheet("員工清單");
  var lastRow = sheet.getLastRow();
  var values = sheet.getRange('A2:A' + lastRow).getValues();
  console.log(values);
  return values;
}

// ===== 下拉選單 end =====================================================
// ===== 新增紀錄 start ===================================================

// 新增一筆紀錄至 Google Sheets 總表
function processForm(form) {
  var sheet = openExcelSheet("總表");
  var id = generateId(sheet);

  sheet.appendRow([id, form.date, form.closeDate, form.deceased, form.item, form.employeeName, form.amount, form.companyName, form.amountCompany
    , -form.adjustment, form.profit, form.funeralDirector, form.note, new Date()]);

  for (var i = 2; i < 8; i++) {
    if (form.hasOwnProperty("employeeName" + i)) {
      id = generateId(sheet);

      var employeeNameKey = "employeeName" + i;
      var amountKey = "amount" + i;
      var adjustmentKey = "adjustment" + i;
      if (form[employeeNameKey]) {
        sheet.appendRow([id, form.date, form.closeDate, form.deceased, form.item, form[employeeNameKey], form[amountKey], form.companyName, form.amountCompany
          , -form[adjustmentKey], "-", form.funeralDirector, form.note, new Date()]);
      }
    }
  }
}

// ===== 新增紀錄 end =====================================================
// ===== 產出報表 start ===================================================

// 產生每月紀錄
function exportMonthTable(start, end) {
  try {
    var startDate = new Date(start);
    startDate.setDate(startDate.getDate() - 1);
    var endDate = new Date(end);

    var year = endDate.getFullYear() - 1911; // 西元年份減去1911即為民國年份
    var month = endDate.getMonth() + 1; // 取得月份，返回值為0到11的數字

    var sheet = openExcelSheet("總表");
    var data = sheet.getDataRange().getValues();
    var filteredData = [];

    filteredData.push(data[0]);

    for (var i = 1; i < data.length; i++) {
      var date = new Date(data[i][1]);
      if (isNaN(date)) {
        return "Error: " + data[i][0] + " 日期為空值！"
      }
      if (date >= startDate && date <= endDate) {
        filteredData.push(data[i]);
      }
    }
    console.log(filteredData);

    if (filteredData.length <= 1) {
      return "Error: 查無資料";
    }

    var fileName = year + '年' + month + '月報表';
    var oriFile = SpreadsheetApp.create(fileName);
    var sheet = oriFile.getActiveSheet();
    sheet.getRange(1, 1, filteredData.length, filteredData[0].length).setValues(filteredData);

    // 取得目標資料夾
    var targetFolder = DriveApp.getFoldersByName("月報表").next();

    // 將檔案移動到目標資料夾
    var oriFileID = oriFile.getId();
    DriveApp.getFileById(oriFileID).moveTo(targetFolder);

    return oriFile.getUrl();

  } catch (error) {

    console.log("產生報表失敗：" + error);
    return "Error: 產生報表失敗：" + error;

  }
}

// 產生每月薪資單
function exportSalaryTable(start, end) {
  try {
    var startDate = new Date(start);
    startDate.setDate(startDate.getDate() - 1);
    var endDate = new Date(end);

    var sheet = openExcelSheet("總表");
    var data = sheet.getDataRange().getValues();
    var filteredData = [];

    for (var i = 1; i < data.length; i++) {
      var date = new Date(data[i][1]);
      if (isNaN(date)) {
        return "Error: " + data[i][0] + " 結案日為空值！"
      }
      if (date >= startDate && date <= endDate) {
        filteredData.push(data[i]);
      }
    }
    console.log(filteredData.length);

    if (filteredData.length == 0) {
      return "Error: 查無資料";
    }

    // Group data by employee name
    var groupedData = filteredData.reduce(function (acc, row) {
      var name = row[5];
      if (!acc[name]) {
        acc[name] = [];
      }
      acc[name].push(row);
      return acc;
    }, {});

    // Create a new spreadsheet for the report
    var reportSpreadsheet = SpreadsheetApp.create(start.slice(0, 7) + ' 薪資');
    var reportSheet = reportSpreadsheet.getActiveSheet();
    var monthTotal = 0;
    // Write headers to report
    reportSheet.appendRow(['員工名', '工作日期', '亡者', '項目', '員工薪資金額', '代收']);

    // Write data to report
    for (var name in groupedData) {
      var employeeData = groupedData[name];
      var positiveTotal = 0;
      var negativeTotal = 0;
      var salaryTotal = 0;

      for (var i = 0; i < employeeData.length; i++) {
        var row = employeeData[i];
        var date = row[1];
        var deceased = row[3];
        var item = row[4];
        var salary = row[6];
        var adjustment = 0;
        if (parseInt(row[9]) < 0) {
          adjustment = row[9];
          negativeTotal += parseInt(adjustment);
        }

        // Update the employee's total salary for this pay period
        positiveTotal += salary;

        // Write the row to the report
        if (adjustment != 0) {
          reportSheet.appendRow([name, date, deceased, item, salary, adjustment]);
        } else {
          reportSheet.appendRow([name, date, deceased, item, salary]);
        }
      }

      // Write the employee's total salary for this pay period to the report
      salaryTotal = positiveTotal + negativeTotal;
      monthTotal += salaryTotal;
      reportSheet.appendRow(['', '', '', '小計', positiveTotal, negativeTotal]);
      reportSheet.appendRow(['', '', '', '總計', salaryTotal]);
      reportSheet.appendRow(['-', '-', '-', '-', '-', '-']);
    }

    reportSheet.appendRow(['', '', '', '月份總計', monthTotal]);
    // 取得目標資料夾
    var targetFolder = DriveApp.getFoldersByName("員工薪資單").next();

    // 將檔案移動到目標資料夾
    var oriFileID = reportSpreadsheet.getId();
    DriveApp.getFileById(oriFileID).moveTo(targetFolder);

    return reportSpreadsheet.getUrl();

  } catch (error) {
    console.log("產生報表失敗：" + error);
    return "Error: 產生報表失敗：" + error;
  }
}

// 產生公司請款單
function exportCompanyTable(start, end) {
  try {
    var startDate = new Date(start);
    startDate.setDate(startDate.getDate() - 1);
    var endDate = new Date(end);

    var year = endDate.getFullYear() - 1911; // 西元年份減去1911即為民國年份
    var month = endDate.getMonth() + 1; // 取得月份，返回值為0到11的數字

    var sheet = openExcelSheet("總表");
    var data = sheet.getDataRange().getValues();
    var filteredData = [];

    for (var i = 1; i < data.length; i++) {
      var date = new Date(data[i][2]);
      if (isNaN(date)) {
        return "Error: " + data[i][0] + " 結案日為空值！"
      }
      if (date >= startDate && date <= endDate) {
        filteredData.push(data[i]);
      }
    }

    if (filteredData.length == 0) {
      console.log("Error: 查無資料");
      return "Error: 查無資料";
    }

    // Group data by company name
    var groupedData = filteredData.reduce(function (acc, row) {
      var companyName = row[7];
      if (!acc[companyName]) {
        acc[companyName] = [];
      }
      acc[companyName].push(row);
      return acc;
    }, {});

    // 當前所在資料夾
    var file = DriveApp.getFileById(SHEET_ID);
    var folder = file.getParents().next();
    var targetfolder = folder.getFoldersByName('公司請款單').next();
    // 進入目標資料夾
    currentFolder = getOrCreateFolder(targetfolder, year + '年' + month + '月請款單');


    for (const key in groupedData) {
      var types = findCompanyTypes(key);
      console.log(key + ":" + types);
      if (types === 1) {
        var spreadsheet = generateCompanyContent1(key, groupedData, year, month);
      } else if (types === 2) {
        var spreadsheet = generateCompanyContent2(key, groupedData, year, month);
      } else if (types === 3) {
        var spreadsheet = generateCompanyContent3(key, groupedData, year, month);
      } else {
        return "Error: 找不到" + key + "的結帳方式";
      }

      // 將檔案移動到目標資料夾
      var oriFileID = spreadsheet.getId();
      DriveApp.getFileById(oriFileID).moveTo(currentFolder);
    }

    // 回傳該資料夾的url
    return currentFolder.getUrl();

  } catch (error) {

    console.log("產生報表失敗：" + error);
    return "Error: 產生報表失敗：" + error;

  }
}

function generateCompanyContent1(key, groupedData, year, month) {
  // 檔案名稱
  var fileName = year + '年-' + key + '-' + month + '月請款單';
  // 公司帳內容
  const companyData = groupedData[key];
  // 新增附件
  const spreadsheet = SpreadsheetApp.create(fileName);
  const sheet = spreadsheet.getActiveSheet();
  // 制式表頭
  sheet.appendRow([year + '年', key, month + '月帳款']);
  sheet.appendRow([' ']);
  sheet.appendRow(['日期', '案名', '項目', '人數', '金額', '代收', '備註', '執案人員']);
  // 將公司帳內容依照案名分類
  var detailData = companyData.reduce(function (acc, row) {
    var deceasedName = row[3];
    if (!acc[deceasedName]) {
      acc[deceasedName] = [];
    }
    acc[deceasedName].push(row);
    return acc;
  }, {});
  // 該案總額
  var _summaryTotal = 0;
  // 該案總額
  var _summaryNegativeTotal = 0;
  // 案名計數器
  var index = 1;
  // 迭代每個案名
  for (const key in detailData) {
    // 公司帳金額
    var positiveTotal = 0;
    // 代收金額
    var negativeTotal = 0;
    // 該案總額
    var salaryTotal = 0;
    // 案名-項目 計數器
    var itemCount = 1;
    // 第一列是否已經出現
    var firstName = false;
    // 暫存總額
    var tempTotal = 0;
    // 暫存代收
    var tempNegativeTotal = 0;

    for (var x = 0; x < detailData[key].length; x++) {
      // 該案名第x項資料
      var deceasedDetail = detailData[key][x];
      // 正規化日期格式
      var dateObj = new Date(deceasedDetail[1]);
      var y = dateObj.getFullYear();
      var m = ('0' + (dateObj.getMonth() + 1)).slice(-2);
      var d = ('0' + dateObj.getDate()).slice(-2);
      var date = y + '-' + m + '-' + d;
      // 先找出下一個項目
      var nextItem;
      if ((x + 1) < detailData[key].length) {
        nextItem = detailData[key][x + 1][4];
      } else {
        nextItem = '';
      }
      // 如果是每個案名的第一筆資料，才要印出案名名稱，反之空白
      if (!firstName) {
        // 如果該筆案名和下筆一樣就++，反之印出後重置1
        if (nextItem == deceasedDetail[4]) {
          tempTotal += deceasedDetail[8];
          if (parseInt(deceasedDetail[9]) < 0) {
            tempNegativeTotal += parseInt(deceasedDetail[9]);
          }
          itemCount++;
          continue;
        }
        tempTotal += deceasedDetail[8];
        if (parseInt(deceasedDetail[9]) < 0) {
          tempNegativeTotal += parseInt(deceasedDetail[9]);
        }
        sheet.appendRow([date, deceasedDetail[3], deceasedDetail[4], itemCount, tempTotal, tempNegativeTotal, deceasedDetail[12], deceasedDetail[11]]);
        itemCount = 1;
        firstName = true;
      } else {
        if (nextItem == deceasedDetail[4]) {
          tempTotal += deceasedDetail[8];
          if (parseInt(deceasedDetail[9]) < 0) {
            tempNegativeTotal += parseInt(deceasedDetail[9]);
          }
          itemCount++;
          continue;
        }
        if (parseInt(deceasedDetail[9]) < 0) {
          tempNegativeTotal += parseInt(deceasedDetail[9]);
        }
        tempTotal += deceasedDetail[8];
        sheet.appendRow([date, '', deceasedDetail[4], itemCount, tempTotal, tempNegativeTotal, deceasedDetail[12], deceasedDetail[11]]);
        itemCount = 1;
      }
      // 看有沒有代收項目，有就+negativeTotal
      if (tempNegativeTotal < 0) {
        negativeTotal += tempNegativeTotal;
      }
      // 將公司帳金額+positiveTotal
      positiveTotal += tempTotal;
      tempTotal = 0;
      tempNegativeTotal = 0;
    }
    // 該案總額
    salaryTotal = positiveTotal + negativeTotal;
    _summaryTotal += positiveTotal;
    // 如果有代收就印小計，反之直接印總計
    if (negativeTotal != 0) {
      sheet.appendRow(['', '', '', '小計', positiveTotal, negativeTotal]);
      _summaryNegativeTotal += negativeTotal;
    }
    sheet.appendRow(['', '', '', '(' + index + ')計', salaryTotal]);
    index++;
  }
  if (_summaryNegativeTotal != 0) {
    sheet.appendRow(['', '', '', '總小計', _summaryTotal, _summaryNegativeTotal]);
  }
  sheet.appendRow(['', '', '', '總計', _summaryTotal + _summaryNegativeTotal]);

  return spreadsheet;
}

function generateCompanyContent2(key, groupedData, year, month) {
  // 檔案名稱
  var fileName = year + '年-' + key + '-' + month + '月請款單';
  // 公司帳內容
  const companyData = groupedData[key];
  // 新增附件
  const spreadsheet = SpreadsheetApp.create(fileName);
  const sheet = spreadsheet.getActiveSheet();
  // 制式表頭
  sheet.appendRow([year + '年', key, month + '月帳款']);
  sheet.appendRow([' ']);
  sheet.appendRow(['日期', '案名', '項目', '人數', '金額', '代收', '備註', '執案人員']);
  // 將公司帳內容依照案名分類
  var detailData = companyData.reduce(function (acc, row) {
    var deceasedName = row[3];
    if (!acc[deceasedName]) {
      acc[deceasedName] = [];
    }
    acc[deceasedName].push(row);
    return acc;
  }, {});
  // 該案總額
  var _summaryTotal = 0;
  // 該案總額
  var _summaryNegativeTotal = 0;
  // 案名計數器
  var index = 1;
  // 迭代每個案名
  for (const key in detailData) {
    // 公司帳金額
    var positiveTotal = 0;
    // 代收金額
    var negativeTotal = 0;
    // 該案總額
    var salaryTotal = 0;
    // 案名-項目 計數器
    var itemCount = 1;
    // 第一列是否已經出現
    var firstName = false;
    // 暫存總額
    var tempTotal = 0;
    // 暫存代收
    var tempNegativeTotal = 0;

    for (var x = 0; x < detailData[key].length; x++) {
      // 該案名第x項資料
      var deceasedDetail = detailData[key][x];
      // 正規化日期格式
      var dateObj = new Date(deceasedDetail[1]);
      var y = dateObj.getFullYear();
      var m = ('0' + (dateObj.getMonth() + 1)).slice(-2);
      var d = ('0' + dateObj.getDate()).slice(-2);
      var date = y + '-' + m + '-' + d;
      // 先找出下一個項目
      var nextItem;
      if ((x + 1) < detailData[key].length) {
        nextItem = detailData[key][x + 1][4];
      } else {
        nextItem = '';
      }
      // 如果是每個案名的第一筆資料，才要印出案名名稱，反之空白
      if (!firstName) {
        // 如果該筆案名和下筆一樣就++，反之印出後重置1
        if (nextItem == deceasedDetail[4]) {
          tempTotal += deceasedDetail[8];
          if (parseInt(deceasedDetail[9]) < 0) {
            tempNegativeTotal += parseInt(deceasedDetail[9]);
          }
          itemCount++;
          continue;
        }
        tempTotal += deceasedDetail[8];
        if (parseInt(deceasedDetail[9]) < 0) {
          tempNegativeTotal += parseInt(deceasedDetail[9]);
        }
        sheet.appendRow([date, deceasedDetail[3], deceasedDetail[4], itemCount, tempTotal, tempNegativeTotal, deceasedDetail[12], deceasedDetail[11]]);
        itemCount = 1;
        firstName = true;
      } else {
        if (nextItem == deceasedDetail[4]) {
          tempTotal += deceasedDetail[8];
          if (parseInt(deceasedDetail[9]) < 0) {
            tempNegativeTotal += parseInt(deceasedDetail[9]);
          }
          itemCount++;
          continue;
        }
        if (parseInt(deceasedDetail[9]) < 0) {
          tempNegativeTotal += parseInt(deceasedDetail[9]);
        }
        tempTotal += deceasedDetail[8];
        sheet.appendRow([date, '', deceasedDetail[4], itemCount, tempTotal, tempNegativeTotal, deceasedDetail[12], deceasedDetail[11]]);
        itemCount = 1;
      }
      // 看有沒有代收項目，有就+negativeTotal
      if (tempNegativeTotal < 0) {
        negativeTotal += tempNegativeTotal;
      }
      // 將公司帳金額+positiveTotal
      positiveTotal += tempTotal;
      tempTotal = 0;
      tempNegativeTotal = 0;
    }
    // 該案總額
    salaryTotal = positiveTotal + negativeTotal;
    _summaryTotal += positiveTotal;
    // 如果有代收就印小計，反之直接印總計
    if (negativeTotal != 0) {
      sheet.appendRow(['', '', '', '小計', positiveTotal, negativeTotal]);
      _summaryNegativeTotal += negativeTotal;
    }
    sheet.appendRow(['', '', '', '(' + index + ')計', salaryTotal]);
    index++;
  }
  if (_summaryNegativeTotal != 0) {
    sheet.appendRow(['', '', '', '總小計', _summaryTotal, _summaryNegativeTotal]);
  }
  sheet.appendRow(['', '', '', '合計', _summaryTotal + _summaryNegativeTotal]);
  var cell2 = Math.ceil((_summaryTotal + _summaryNegativeTotal) * 0.03);
  var cell4 = _summaryTotal + _summaryNegativeTotal - cell2;
  var cell1 = Math.ceil(cell4 / 1.05);
  var cell3 = cell4 - cell1;
  sheet.appendRow([' ']);
  sheet.appendRow(['', '未稅', cell1, '回3%', cell2]);
  sheet.appendRow(['', '稅金', cell3, '總計', cell4]);

  return spreadsheet;
}

function generateCompanyContent3(key, groupedData, year, month) {
  // 檔案名稱
  var fileName = year + '年-' + key + '-' + month + '月請款單';
  // 公司帳內容
  const companyData = groupedData[key];
  // 新增附件
  const spreadsheet = SpreadsheetApp.create(fileName);
  const sheet = spreadsheet.getActiveSheet();
  // 制式表頭
  sheet.appendRow([' ', '廠商名稱：', ' ', key]);
  sheet.appendRow([' ', '請款年月：', ' ', year + '年' + month + '月']);
  sheet.appendRow([' ']);
  sheet.appendRow(['日期', '案名', '項目', '人數', '金額', '代收', '備註', '執案人員', '紅包']);
  // 將公司帳內容依照案名分類
  var detailData = companyData.reduce(function (acc, row) {
    var deceasedName = row[3];
    if (!acc[deceasedName]) {
      acc[deceasedName] = [];
    }
    acc[deceasedName].push(row);
    return acc;
  }, {});
  // 該案總額
  var _summaryTotal = 0;
  // 該案總額
  var _summaryNegativeTotal = 0;
  // 紅包金額
  var _summaryenvelopeTotal = 0;
  // 紅包（未稅）項目
  var dutyFree = findItemsEnvelop();
  // 案名計數器
  // var index = 1;
  // 迭代每個案名
  for (const key in detailData) {
    // 公司帳金額
    var positiveTotal = 0;
    // 代收金額
    var negativeTotal = 0;
    // 案名-項目 計數器
    var itemCount = 1;
    // 第一列是否已經出現
    var firstName = false;
    // 暫存總額
    var tempTotal = 0;
    // 暫存代收
    var tempNegativeTotal = 0;
    // 紅包金額
    var tempEnvelope = 0;

    for (var x = 0; x < detailData[key].length; x++) {

      // 該案名第x項資料
      var deceasedDetail = detailData[key][x];
      // 正規化日期格式
      var dateObj = new Date(deceasedDetail[1]);
      var y = dateObj.getFullYear();
      var m = ('0' + (dateObj.getMonth() + 1)).slice(-2);
      var d = ('0' + dateObj.getDate()).slice(-2);
      var date = y + '-' + m + '-' + d;
      // 先找出下一個項目
      var nextItem;
      if ((x + 1) < detailData[key].length) {
        nextItem = detailData[key][x + 1][4];
      } else {
        nextItem = '';
      }
      // 如果是每個案名的第一筆資料，才要印出案名名稱，反之空白
      if (!firstName) {
        // 如果該筆案名和下筆一樣就++，反之印出後重置1
        if (nextItem == deceasedDetail[4]) {
          tempTotal += deceasedDetail[8];
          if (parseInt(deceasedDetail[9]) < 0) {
            tempNegativeTotal += parseInt(deceasedDetail[9]);
          }
          itemCount++;
          continue;
        }
        tempTotal += deceasedDetail[8];
        if (parseInt(deceasedDetail[9]) < 0) {
          tempNegativeTotal += parseInt(deceasedDetail[9]);
        }
        for (var i = 0; i < dutyFree.length; i++) {
          if (deceasedDetail[4].includes(dutyFree[i][0])) {
            tempEnvelope += dutyFree[i][1] * itemCount;
            tempTotal -= tempEnvelope;
          }
        }
        sheet.appendRow([date, deceasedDetail[3], deceasedDetail[4], itemCount, tempTotal, tempNegativeTotal, deceasedDetail[12], deceasedDetail[11], tempEnvelope]);
        itemCount = 1;
        firstName = true;
      } else {
        if (nextItem == deceasedDetail[4]) {
          tempTotal += deceasedDetail[8];
          if (parseInt(deceasedDetail[9]) < 0) {
            tempNegativeTotal += parseInt(deceasedDetail[9]);
          }
          itemCount++;
          continue;
        }
        tempTotal += deceasedDetail[8];
        if (parseInt(deceasedDetail[9]) < 0) {
          tempNegativeTotal += parseInt(deceasedDetail[9]);
        }
        for (var i = 0; i < dutyFree.length; i++) {
          if (deceasedDetail[4].includes(dutyFree[i][0])) {
            tempEnvelope += dutyFree[i][1] * itemCount;
            tempTotal -= tempEnvelope;
          }
        }
        sheet.appendRow([date, '', deceasedDetail[4], itemCount, tempTotal, tempNegativeTotal, deceasedDetail[12], deceasedDetail[11], tempEnvelope]);
        itemCount = 1;
      }
      // 看有沒有代收項目，有就+negativeTotal
      if (tempNegativeTotal < 0) {
        _summaryNegativeTotal += tempNegativeTotal;
      }
      // 將公司帳金額+positiveTotal
      _summaryTotal += tempTotal;
      _summaryenvelopeTotal += tempEnvelope;
      tempTotal = 0;
      tempNegativeTotal = 0;
      tempEnvelope = 0;
    }
  }
  sheet.appendRow([' ']);
  if (_summaryNegativeTotal != 0) {
    sheet.appendRow(['', '', '', '總小計', _summaryTotal, _summaryNegativeTotal, '', '', _summaryenvelopeTotal]);
  }
  var cell1 = _summaryTotal + _summaryNegativeTotal;
  var cell2 = cell1 * 0.05;
  var cell3 = cell1 + cell2;
  var cell4 = _summaryenvelopeTotal + cell3;
  sheet.appendRow(['', '', '', '未稅', cell1]);
  sheet.appendRow(['', '', '', '稅金', cell2]);
  sheet.appendRow([' ']);
  sheet.appendRow(['', '', '', '合計', cell3]);
  sheet.appendRow(['', '', '', '紅包', _summaryenvelopeTotal]);
  sheet.appendRow(['', '', '', '總計', cell4]);

  return spreadsheet;
}

// 取得計算方式
function findCompanyTypes(companyName) {
  var sheet = openExcelSheet("公司清單");
  var lastRow = sheet.getLastRow();
  var values = sheet.getRange('A2:B' + lastRow).getValues();
  for (var i = 0; i < values.length; i++) {
    if (companyName == values[i][0]) {
      return values[i][1];
    }
  }
  return companyName;
}

// 取得紅包項目
function findItemsEnvelop() {
  var sheet = openExcelSheet("工作項目");
  var lastRow = sheet.getLastRow();
  var values = sheet.getRange('A2:D' + lastRow).getValues();
  var retVal = [];
  for (var i = 0; i < values.length; i++) {
    if (values[i][3] == "Y") {
      retVal.push([values[i][0], values[i][2]]);
    }
  }
  return retVal;
}

// ===== 產出報表 end ===================================================
// ===== 行事曆 start ===================================================

// 找出某筆id的待辦事項
function findDataByID(id) {
  var sheet = openExcelSheet("待辦清單");
  var lastRow = sheet.getLastRow();
  var records = [];
  var dataRange = sheet.getRange('A2:I' + lastRow).getValues();
  for (var i = 0; i < dataRange.length; i++) {
    if (dataRange[i][0] == id) {
      var date = new Date(dataRange[i][1]);
      var year = date.getFullYear();
      var month = ("0" + (date.getMonth() + 1)).slice(-2);
      var day = ("0" + date.getDate()).slice(-2);
      var formattedDate = year + "-" + month + "-" + day;
      records.push(formattedDate);                  // date
      records.push(dataRange[i][2].split("~"));     // time
      records.push(dataRange[i][3]);                // 案名
      records.push(dataRange[i][4].split(", "));    // 項目
      records.push(dataRange[i][9]);                // 備註
      records.push(dataRange[i][6]);                // 公司
      records.push(dataRange[i][7]);                // 禮儀師
      records.push(dataRange[i][8].split(", "));    // 員工
      console.log(records);
      return records;
    }
  }

  return records;
}

// 新增一筆紀錄至 Google Sheets 總表
function newTodoList(form) {
  var sheet = openExcelSheet("待辦清單");
  var id = generateId(sheet);

  sheet.appendRow([id, form.date, form.startTime + '~' + form.endTime, form.deceased, form.item, form.address, form.companyName, form.funeralDirector, form.employeeName
    , form.note, 'Process', new Date()]);

  createCalendarEvent(getRelativeDate(form.date, form.startTime), getRelativeDate(form.date, form.endTime), form.deceased + processString(form.item) + "(" + form.employeeName + ")", form.address, [form.date, form.startTime + '~' + form.endTime, form.deceased, processString(form.item), form.address, form.companyName, form.funeralDirector, form.employeeName, form.note])
}

function createCalendarEvent(start, end, summary, location, details) {
  var calendarId = "primary";
  var description = '';
  description += '日期 : ' + details[0];
  description += '\n時間 : ' + details[1];
  description += '\n案名 : ' + details[2];
  description += '\n項目 : ' + details[3];
  description += '\n地點 : ' + details[4];
  description += '\n公司 : ' + details[5] + details[6];
  description += '\n人員 : ' + details[7];
  description += '\n備註 : ' + details[8];

  let event = {
    summary: summary,
    description: description,
    start: {
      dateTime: start.toISOString()
    },
    end: {
      dateTime: end.toISOString()
    },
    colorId: 11
  };
  try {
    // call method to insert/create new event in provided calandar
    var calendar = CalendarApp.getCalendarById(calendarId);
    calendar.createEvent(event.summary, new Date(event.start.dateTime), new Date(event.end.dateTime), event);
  } catch (err) {
    console.log('Failed with error %s', err.message);
    return err.message;
  }

  Logger.log("Created event: " + event.summary);
}

function getRelativeDate(dateIn, time) {
  var date = new Date(dateIn);
  date.setHours(time.substring(0, 2));
  date.setMinutes(time.substring(time.length - 2));
  date.setSeconds(0);
  date.setMilliseconds(0);
  return date;
}

// ===== 行事曆 end ===================================================
// ===== utils start ===================================================

// util : openExcel
function openExcelSheet(sheetName) {
  var excel = SpreadsheetApp.openById(SHEET_ID);
  var sheet = excel.getSheetByName(sheetName);
  return sheet;
}

// util : generate id
function generateId(sheet) {
  const lastRow = sheet.getLastRow();
  const currentDate = new Date();
  const year = currentDate.getFullYear().toString().substring(2);
  const month = (currentDate.getMonth() + 1).toString().padStart(2, '0');
  const day = currentDate.getDate().toString().padStart(2, '0');
  const serial = (lastRow).toString().padStart(3, '0');
  return `${year}${month}${day}-${serial}`;
}

// util : generate report
function generateReport(folderName, fileName, fileContent, sheetName) {
  var folder = DriveApp.getFoldersByName(folderName).next();// gets first folder with the given foldername
  var newFile = folder.createFile(fileName, fileContent);
  var newSheet = newFile.create(sheetName);
  var sheet = newSheet.getActiveSheet();
  sheet.getRange(1, 1, fileContent.length, fileContent[0].length).setValues(fileContent);
  return newSheet.getUrl();
}

// util : check existing folder
function checkExistingFolder(folder, folderName) {
  var folders = folder.getFoldersByName(folderName);

  if (folders.hasNext()) {
    return folders.next();
  } else {
    return null;
  }
}

// util : check existing file
function checkExistingFile(folder, filename) {
  var files = folder.getFilesByName(filename);

  if (files.hasNext()) {
    return files.next();
  } else {
    return null;
  }
}

// util : get or create folder
function getOrCreateFolder(parentFolder, folderName) {
  var folders = parentFolder.getFoldersByName(folderName);

  if (folders.hasNext()) {
    return folders.next();
  } else {
    return parentFolder.createFolder(folderName);
  }
}

// 處理包含逗號分隔的字串，取代括號的內容後，檢查是否重複後重組字串
function processString(str) {
  // 以逗號分隔字符串
  const strArray = str.split(',');

  // 遍历每个子字符串
  const processedArray = strArray.map((item) => {
    // 取代括号及其内容为空字符串
    const processedItem = item.replace(/\(.*\)/, '');
    // 移除字符串两端的空格
    return processedItem.trim();
  });

  // 去除重复的内容
  const uniqueArray = [...new Set(processedArray)];

  // 将处理后的数组转换为字符串
  const result = uniqueArray.join(', ');
  return result;
}

// ===== utils end =====================================================