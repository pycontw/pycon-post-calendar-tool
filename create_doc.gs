/* Create menu */
function onOpen() {
  SpreadsheetApp
    .getUi()
    .createMenu('貼文小工具')
    .addItem('新增貼文 doc (請至貼文表單先選取要新增的文件列)', 'createPostDoc')
    .addItem('test prompt', 'showPrompt')
    .addToUi()
}

/* Create google doc and make sure if user would create a new folder and file */
function createPostDoc(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("貼文表單");
  var row = sheet.getActiveRange().getRow();
  Logger.log('row: %d', row);
  let title = sheet.getRange(row, 4).getValue();

  let postDate = sheet.getRange(row, 3).getValue();
  let team = sheet.getRange(row, 1).getValue();
  var date = Utilities.formatDate(postDate, "GMT+8", "MMdd");    
  // var datee = Utilities.formatDate(postDate, "GMT+8", "d"); 
  docTitle =  date + " - " + team + " - " + title;
  var text = "是否要新增 " + docTitle + " 的 google doc 文件";

  var response = Browser.msgBox('Greetings', text, Browser.Buttons.YES_NO);
  if (response == "yes") {
    Logger.log('The user clicked "Yes."');
    var doc_id = createDoc(docTitle);
    var url = "https://docs.google.com/document/d/"+doc_id
    openUrl(url);
  } else {
    Logger.log('The user clicked "No" or the dialog\'s close button.');
  } 

  const rangeToAddLink = sheet.getRange(row, 11)
  const richText = SpreadsheetApp.newRichTextValue()
      .setText(docTitle)
      .setLinkUrl(url)
      .build();
  rangeToAddLink.setRichTextValue(richText);
}

function createDoc(file_name) {
  var parentFolder=DriveApp.getFolderById('1mI4n-gXhVSmvDqHhzJGfPIT0C7uFEhwA');
  var target_folder_ID = parentFolder.createFolder(file_name).getId();

  let new_doc = DocumentApp.create(file_name);
  let doc_id = new_doc.getId();
  moveFile(doc_id, target_folder_ID);
  var target_folder_link = "https://drive.google.com/drive/folders/" + target_folder_ID;
  template_id = "1JPqaPfXbkzWwXRKbhvPJWwQnfe1RV_s2QpuR0jMfIZc";
  importInDoc(doc_id, template_id, target_folder_link, file_name);

  return doc_id;
}
function openUrl( url ){
  var html = HtmlService.createHtmlOutput('<html><script>'
  +'window.close = function(){window.setTimeout(function(){google.script.host.close()},9)};'
  +'var a = document.createElement("a"); a.href="'+url+'"; a.target="_blank";'
  +'if(document.createEvent){'
  +'  var event=document.createEvent("MouseEvents");'
  +'  if(navigator.userAgent.toLowerCase().indexOf("firefox")>-1){window.document.body.append(a)}'                          
  +'  event.initEvent("click",true,true); a.dispatchEvent(event);'
  +'}else{ a.click() }'
  +'close();'
  +'</script>'
  // Offer URL as clickable link in case above code fails.
  +'<body style="word-break:break-word;font-family:sans-serif;">Failed to open automatically. <a href="'+url+'" target="_blank" onclick="window.close()">Click here to proceed</a>.</body>'
  +'<script>google.script.host.setHeight(40);google.script.host.setWidth(410)</script>'
  +'</html>')
  .setWidth( 90 ).setHeight( 1 );
  SpreadsheetApp.getUi().showModalDialog( html, "Opening ..." );
}

function moveFile(fileId, destinationFolderId) {
  let destinationFolder = DriveApp.getFolderById(destinationFolderId);
  DriveApp.getFileById(fileId).moveTo(destinationFolder);
}

function importInDoc(new_id, template_id,target_folder_link, file_name) {
  var baseDoc = DocumentApp.openById(new_id);
  var body = baseDoc.getBody();
  var paragraph_ele;
  var ele;

  var otherBody = DocumentApp.openById(template_id).getBody();
  var totalElements = otherBody.getNumChildren();
  for( var j = 0; j < totalElements; ++j ) {
    var element = otherBody.getChild(j).copy();
    Logger.log("element.getText(): %s", element.getText());
    var type = element.getType();
    if( type == DocumentApp.ElementType.PARAGRAPH )
    {
      paragraph_ele = body.appendParagraph(element);
      if (paragraph_ele.getText()=="https://drive.google.com/drive/u/1/folders/1i6YmSGFgEpp7vpFGs5IRRxm4cnN03xx-") 
      {
          body.appendParagraph(target_folder_link);
          body.removeChild(paragraph_ele);
      }
      if (j==0) 
      {
          ele = body.appendParagraph(file_name);
          ele.setHeading(DocumentApp.ParagraphHeading.TITLE);
          body.removeChild(paragraph_ele);
      }   
    }
    else if( type == DocumentApp.ElementType.TABLE )
      body.appendTable(element);
    else if( type == DocumentApp.ElementType.LIST_ITEM )
      body.appendListItem(element);
    else if( type == DocumentApp.ElementType.INLINE_IMAGE )
      body.appendImage(element);
    else if( type == DocumentApp.ElementType.TABLE_OF_CONTENTS )
      continue;
    // add other element types as you want

    else
      throw new Error("According to the doc this type couldn't appear in the body: "+type);
  }
}

/* test prompt feature*/
function showPrompt() {
  var ui = SpreadsheetApp.getUi(); // Same variations.

  var result = ui.prompt(
      'Let\'s get to know each other!',
      'Please enter your name:',
      ui.ButtonSet.OK_CANCEL);

  // Process the user's response.
  var button = result.getSelectedButton();
  var text = result.getResponseText();
  if (button == ui.Button.OK) {
    // User clicked "OK".
    ui.alert('Your name is ' + text + '.');
  } else if (button == ui.Button.CANCEL) {
    // User clicked "Cancel".
    ui.alert('I didn\'t get your name.');
  } else if (button == ui.Button.CLOSE) {
    // User clicked X in the title bar.
    ui.alert('You closed the dialog.');
  }
}
