# deleteActiveSheet
Google spread sheet delete 删除google表格spread sheet程序

3.给文件命名"DeleteRecentOrders"，将下面一下段代码复制进去，根据实际需要修改代码中的截止日期，并点击保存。
代码中红色日期为表格删除截止的页签名，如07/01，则为从表格第一个页签开始删除，直到指定的日期07/01(包括该页签)为止


function deleteRecentOrders() {

  var date = '07/01';
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  
  var sheet = ss.getSheetByName(date);
  var sheet_id = sheet.getIndex();
  
  for (var i = 0; i < sheet_id; i++) {
    ss.setActiveSheet(sheets[i]);
    ss.deleteActiveSheet();  
  }
}

<br>
将文件名改为DeleteHistoryOrders, 将以下代码复制进去，点击保存。
代码中红色日期为表格删除截止的页签名，如05/09，则为则为从表格最后面一个页签开始删除，直到指定的日期05/09(包括该页签)为止


function deleteHistoryOrders() {

  var date = '05/09';
 
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
 
  var sheet = ss.getSheetByName(date);
  var sheet_id = sheet.getIndex();
 
  Logger.log(sheet_id);
 
  for (var i = sheets.length; i > sheet_id - 1 ; i--) {
    ss.setActiveSheet(sheets[i - 1]);
    ss.deleteActiveSheet(); 
  }
}
