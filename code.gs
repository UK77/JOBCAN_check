var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

function checkJOBCAN() {
  var checkSheet = spreadsheet.insertSheet("JOBCANチェック",0);
  checkSheet.appendRow(["日付","確認事項","名前"]);
  checkSheet.setFrozenRows(1);
  checkSheet.getRange("A1:C1").setFontStyle("bold");
  for (k=0; k<spreadsheet.getSheets().length;k++){
    var sheet = spreadsheet.getSheets()[k];
  
    var lastColumn = sheet.getLastColumn();
    var lastRow = sheet.getLastRow();
    var name = sheet.getRange("A4:B4").getValue();
    
    var ninth = sheet.getRange(9,1).getBackground();
    if (ninth !== "#ffffff"){
      var i = 14;
    }else{
      var i = 11;
    };
  
    for (i; i<= lastRow-1; i++){
      var date = sheet.getRange(i, 1).getValue();
      var columns = {att:2, hol:3, shi:5,com:8 ,lea:9, wor:10, res:14};
      var attSection = sheet.getRange(i, columns.att).getValue();
      var holSection = "休日出勤"; 
      var leaSection = "深夜残業";
      var worSection = "8時間未満労働";
      var resSection = "休憩時間";
      var misSection = "打刻漏れ";
      var attendance = sheet.getRange(i, columns.att).getValue();
      var holiday = sheet.getRange(i, columns.hol).getValue();
      var shift = sheet.getRange(i, columns.shi).getValue();
      var leaving = sheet.getRange(i, columns.lea).getValue();
      var coming = sheet.getRange(i, columns.com).getValue();
      var working = sheet.getRange(i, columns.wor).getValue();
      var rest =  sheet.getRange(i, columns.res).getValue();

      if (attendance){
        checkSheet.appendRow([date, attSection, name]);
      }else if (holiday && working !== "00:00"){
        checkSheet.appendRow([date, holSection, name]);
      }else if (parseInt(leaving,10) >= 22){
        checkSheet.appendRow([date, leaSection, name]);
      }else if (working !== "00:00" && parseInt(working,10)<8){
        checkSheet.appendRow([date, worSection, name]);
      }else if (shift !== "00:00" && coming !== "00:00" && leaving == "00:00"){
        checkSheet.appendRow([date, misSection, name]);
      }else if (rest !== "00:00" && rest !== "00:45" && rest !== "01:00") {
        checkSheet.appendRow([date, resSection, name]);
      }; 
    };
  };
};

