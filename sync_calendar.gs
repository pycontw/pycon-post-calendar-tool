function onEdit(e) { 
  Logger.log("onedit");

  let date = Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd HH:mm");
  let lastColumn = 21;
  let afterlastColumn = lastColumn + 1;
  
  let range = e.range;          // 從事件物件 e 中取出了被編輯的單元格範圍（Range），並將它存放在變數 range 中
  let sheet = range.getSheet(); // 使用 getSheet 方法，取得了被編輯的單元格所在的工作表（Sheet），並將它存放在變數 sheet 中。
  let row = range.getRow();     // 使用 getRow 方法，取得了被編輯的單元格的「行數」（水平），並將它存放在變數 row 中。

  if ((e.range.getA1Notation()[0].charCodeAt(0) >= 65 && e.range.getA1Notation()[0].charCodeAt(0) <= 68) || (e.range.getA1Notation()[0]=='N')) {
    if (e.range.getA1Notation()[0]=='C')
      oldvalue = e.oldValue;

    let reqId     = sheet.getRange(row, 5).getValue();
    let title     = sheet.getRange(row, 4).getValue();
    let postDate  = sheet.getRange(row, 3).getValue();
    let person    = sheet.getRange(row, 2).getValue();
    let team      = sheet.getRange(row, 1).getValue();
    let status    = sheet.getRange(row, 13).getValue();

    if(
      typeof title === 'undefined' || title.length === 0 ||
      typeof postDate === 'undefined' || postDate.length === 0 ||
      typeof person === 'undefined' || person.length === 0 ||
      typeof team === 'undefined' || team.length === 0
      ) {
        console.log("Element is empty.");
        return;
    }

    if(reqId === 'undefined' || reqId.length === 0) {
      // Add event
      reqId = genReqID();
      addPostEvent(reqId, team, person, postDate, title, status);
      sheet.getRange(row, 5).setValue(reqId);
    } else {
      // Modify event
      editPostEvent(reqId, team, person, postDate, title, status);
    }
  }
}

function addPostEvent(reqId, team, person, postDate, title, status) {
  const logFun = "[addPostEvent]: ";
  var sheet = SpreadsheetApp.getActive().getSheetByName('貼文日曆');
  var text  =  title + "\n@" + team + " " + person + " " + reqId;

  let real_row_col = matchDate(postDate);
  let color = matchColor(status);
  text = addEventContent(text, real_row_col[0], real_row_col[1]);

  console.log(logFun + "row: " + real_row_col[0] + " col: " + real_row_col[1]);
  console.log(logFun + "color: " + color);
  console.log(logFun + "text: " + text);

  var cell = sheet.getRange(real_row_col[0], real_row_col[1]);
  cell.setValue(text);
  
  if (color!="")
    cell.setBackground(color);

  return (real_row_col[0], real_row_col[1]);
}

function editPostEvent(reqId, team, person, postDate, title, status) {
  const logFun = "[editPostEvent]: ";
  var sheet = SpreadsheetApp.getActive().getSheetByName('貼文日曆');
  var text  =  title + "\n@" + team + " " + person + " " + reqId;

  let real_row_col = matchDate(postDate);
  let color = matchColor(status);
  text = modifyEvenContent(reqId, team, person, title, real_row_col[0], real_row_col[1]);

  console.log(logFun + "row: " + real_row_col[0] + " col: " + real_row_col[1]);
  console.log(logFun + "color: " + color);
  console.log(logFun + "text: " + text);

  var cell = sheet.getRange(real_row_col[0], real_row_col[1]);
  cell.setValue(text);
  
  if (color!="")
    cell.setBackground(color);

  return (real_row_col[0], real_row_col[1]);
}

function addEventContent(value, row, col) {
  const logFun = "[addEventContent]: ";
  let sheet = SpreadsheetApp.getActive().getSheetByName('貼文日曆');
  let cell = sheet.getRange(row, col);

  console.log(logFun + " content: " + cell.getValue());
  let cur_data = cell.getValue();

  if (cur_data.length > 0)
    cur_data = cur_data + "\n" + value;
  else 
    cur_data = value;

  return cur_data;
}

function modifyEvenContent(reqId, team, person, title, row, col) {
  const logFun = "[modifyEvenContent]: ";
  let sheet = SpreadsheetApp.getActive().getSheetByName('貼文日曆');
  let cell = sheet.getRange(row, col);

  let cur_data = cell.getValue();
  // 避免既有欄位為空，造成錯誤
  if(cur_data === 'undefine' || cur_data.toString.length === 0) {
    return title + "\n@" + team + " " + person + " " + reqId;
  }

  let arr = cur_data.toString().split("\n");

  var title_arr = [];
  var info_arr = [];
  var team_arr = [];
  var person_arr = [];
  var reqId_arr = [];

  // 分類欄位內各項內容
  for(var i=0; i < arr.length; i++) {
    if(i % 2 == 0) {
      title_arr.push(arr[i]);
    } else {
      info_arr.push(arr[i]);
      let arr2 = arr[i].split(" ");
      team_arr.push(arr2[0]);
      person_arr.push(arr2[1]);
      reqId_arr.push(arr2[2]);
    }
  }

  // 更新正確內容
  for(var i=0; i < info_arr.length; i++) {
    if(reqId_arr[i] === reqId) {
      team_arr[i] = "@" + team;
      person_arr[i] = person;
      title_arr[i] = title;
      break;
    }
  }

  // 重設貼文日曆對應欄位內容
  var result = "";
  for(var i=0; i < title_arr.length; i++) {
    if(result.length === 0) {
      result = title_arr[i] + "\n" + team_arr[i] + " " + person_arr[i] + " " + reqId_arr[i];
    } else {
      result = result + "\n" + title_arr[i] + "\n" + team_arr[i] + " " + person_arr[i] + " " + reqId_arr[i];
    }
  }

  return result;
}

function matchDate(postDate) {
  const logFun = "[matchDate]: ";
  let sheet = SpreadsheetApp.getActive().getSheetByName('貼文日曆');

  let month = Utilities.formatDate(postDate, "GMT+8", "M");
  let date  = Utilities.formatDate(postDate, "GMT+8", "d");

  var row = 1;
  var real_row, real_col = 0;

  for (row = 1; row < 200; row++) {
    if (+sheet.getRange(row, 2).getValue() == month) {
      break; 
    }
  }

  row_limit = row + 5;
  var cur_date, tmp_date;
  while (row < row_limit) {
    cur_date = sheet.getRange(row, 3).getValue();
    console.log(logFun + "now date :" + cur_date + " " + date);
    if (+cur_date > date) {
      real_row = row-1;
      real_col = 11 + (date - cur_date);
      break;
    } else if (+cur_date < tmp_date) {
      real_row = row-1;
      real_col = 4 + (date - tmp_date);
      break;
    }
    row += 1;
    tmp_date = cur_date;
  }

  return [real_row, real_col];
}

function matchColor(status) {
  if (status == "已審閱")
    color = "#a4c2f4";
  else if (status == "已排程")
    color = "#b7d7a8";
  else if (status == "已發布")
    color = "#cccccc";
  else 
    color = "transparent";
  return color;
}

function genReqID() {
  var d = Date.now();
  if (typeof performance !== 'undefined' && typeof performance.now === 'function'){
      d += performance.now(); //use high-precision timer if available
  }
  
  let reqID = 'xxxxxxxx-xxxx-yxxx'.replace(/[xy]/g, function (c) {
      var r = (d + Math.random() * 16) % 16 | 0;
      d = Math.floor(d / 16);
      return (c === 'x' ? r : (r & 0x3 | 0x8)).toString(16);
  });

  return reqID;
}