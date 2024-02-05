function onEdit(e) { 
    Logger.log("onedit");
    let range = e.range; 
    // 使用 getSheet 方法，取得了被編輯的單元格所在的工作表（Sheet），並將它存放在變數 sheet 中。
    let sheet = range.getSheet(); 
    let row = range.getRow(); 
    let title = sheet.getRange(row, 4).getValue();
    let lastColumn = 20; 

    if ((e.range.getA1Notation()[0].charCodeAt(0) >= 65 && e.range.getA1Notation()[0].charCodeAt(0) <= 68)|| (e.range.getA1Notation()[0]=='M')) {
      if (e.range.getA1Notation()[0]=='C')
        oldvalue = e.oldValue;

      var date = Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd HH:mm");
      // 從事件物件 e 中取出了被編輯的單元格範圍（Range），並將它存放在變數 range 中
     
      let afterlastColumn = lastColumn+1; 

      // 使用 getRow 方法，取得了被編輯的單元格的「行數」（水平），並將它存放在變數 row 中。
      let postDate = sheet.getRange(row, 3).getValue();
      let person = sheet.getRange(row, 2).getValue();
      let team = sheet.getRange(row, 1).getValue();
      let status = sheet.getRange(row, 13).getValue();
      var month = Utilities.formatDate(postDate, "GMT+8", "M");    
      var datee = Utilities.formatDate(postDate, "GMT+8", "d");          
      var text =  title + "\n@" + team + " " + person;

      console.log(date)

      // 將當前的日期和時間，設置為被編輯的「單元格所在行」的「最後一列」的值。
      sheet.getRange(row, lastColumn).setValue(date);
      sheet.getRange(row, afterlastColumn).setValue(status);
      matchDate(month, datee, text, status)
    }
} 

function matchDate(month, date, value, status)
{
  var sheet = SpreadsheetApp.getActive().getSheetByName('貼文日曆'); 
  var row = 1;
  var real_row, real_col = 0

  for (row = 1; row < 200; row++) {    
    if (+sheet.getRange(row, 2).getValue() == month) 
      break; 
  }     
  row_limit = row + 5
  var cur_date, tmp_date;
  while (row < row_limit)
  {
    cur_date = sheet.getRange(row, 3).getValue()
    console.log("now date :" + cur_date )
    if (+cur_date > date) 
    {
      real_row = row-1;
      real_col = 11 + (date - cur_date)
      break;
    }
    else if (+cur_date < tmp_date)
    {
      real_row = row-1;
      real_col = 4 + (date - tmp_date)
      break;
    }
    row += 1;
    tmp_date = cur_date;
  }

  console.log("row: "+real_row);
  console.log("col: "+real_col);

  var cell = sheet.getRange(real_row, real_col);
  cell.setValue(value);
  var color = matchColor(status)
  if (color!="")
    cell.setBackground(color);

  return (real_row, real_col)
}


function matchColor(status)
{
  if (status == "已審閱")
    color = "#a4c2f4";
  else if (status == "已排程")
    color = "#b7d7a8";
  else if (status == "已發布")
    color = "#cccccc";
  else 
    color = "transparent"
  return color
}
