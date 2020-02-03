/*
 * http://www.andrewroberts.net/2016/01/google-apps-script-to-create-and-email-a-pdf/
 */

function getRowsData(sheet, range, columnHeadersRowIndex) {
  columnHeadersRowIndex = columnHeadersRowIndex || range.getRowIndex() - 1;
  var numColumns = range.getEndColumn() - range.getColumn() + 1;
  var headersRange = sheet.getRange(columnHeadersRowIndex, range.getColumn(), 1, numColumns);
  var headers = headersRange.getValues()[0];
  return getObjects(range.getDisplayValues(), normalizeHeaders(headers));
}

function getObjects(data, keys) {
  var objects = [];
  for (var i = 0; i < data.length; ++i) {
    var object = {};
    var hasData = false;
    for (var j = 0; j < data[i].length; ++j) {
      var cellData = data[i][j];
      if (isCellEmpty(cellData)) {
        continue;
      }
      object[keys[j]] = cellData;
      hasData = true;
    }
    if (hasData) {
      objects.push(object);
    }
  }
  return objects;
}

function normalizeHeaders(headers) {
  var keys = [];
  for (var i = 0; i < headers.length; ++i) {
    var key = normalizeHeader(headers[i]);
    if (key.length > 0) {
      keys.push(key);
    }
  }
  return keys;
}

function normalizeHeader(header) {
  var key = '';
  var upperCase = false;
  for (var i = 0; i < header.length; ++i) {
    var letter = header[i];
    if (letter == ' ' && key.length > 0) {
      upperCase = true;
      continue;
    }
    if (!isAlnum(letter)) {
      continue;
    }
    if (key.length == 0 && isDigit(letter)) {
      continue; // first character must be a letter
    }
    if (upperCase) {
      upperCase = false;
      key += letter.toUpperCase();
    } else {
      key += letter.toLowerCase();
    }
  }
  return key;
}

function isCellEmpty(cellData) {
  return typeof(cellData) == 'string' && cellData == '';
}

function isAlnum(char) {
  return char >= 'A' && char <= 'Z' ||
    char >= 'a' && char <= 'z' ||
    isDigit(char);
}

function isDigit(char) {
  return char >= '0' && char <= '9';
}

function fillInTemplateFromObject(template, data) {
  var doc = template;
  // Search for all the variables to be replaced, for instance <<columnName>>
  var templateVars = template.match(/<<[^>]+>>/g);
  // Replace variables from the template with the actual values from the data object.
  for (var i = 0; i < templateVars.length; ++i) {
    var variableData = data[normalizeHeader(templateVars[i])];
    doc = doc.replaceText(templateVars[i], variableData || '');
  }
  return doc;
}

function sendEmails() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dataSheet = ss.getSheets()[0];
  var dataRange = dataSheet.getRange(2, 1, dataSheet.getMaxRows() - 1, 4);
  var templateSheet = ss.getSheets()[1];
  var emailTemplate = templateSheet.getRange('A1').getValue();
  var objects = getRowsData(dataSheet, dataRange);
  // For every row object, create a personalized email from a template and send
  // it to the appropriate person.
  for (var i = 0; i < objects.length; ++i) {
    // Get a row object
    var rowData = objects[i];
    var emailText = fillInTemplateFromObject(emailTemplate, rowData);
    var emailSubject = 'Mail Merge Test';
    var file = DriveApp.getFilesByName('2019_MA_BenefitsGuide.pdf')
    MailApp.sendEmail(rowData.emailAddress, emailSubject, emailText );
    // MailApp.sendEmail({to: emailAddress, subject: subject, body: message, htmlBody: html}).
    // MailApp.sendEmail(emailAddress, "Subject", emailTemplate.getText(), {htmlBody: emailTemplate });
  }
}


// http://www.andrewroberts.net/2014/10/google-apps-script-create-pdf/

// dev: andrewroberts.net

/**  
 * Take the fields from the active row in the active sheet
 * and, using a Google Doc template, create a PDF doc with these
 * fields replacing the keys in the template. The keys are identified
 * by having a % either side, e.g. %Name%.
 *
 * @return {Object} the completed PDF file
 */

function createPdf(config, action) {

  // Set up the docs and the spreadsheet access
  
  var copyFile = DriveApp.getFileById(config.templateID).makeCopy();
  var copyId = copyFile.getId();
  var copyDoc = DocumentApp.openById(copyId);
  var copyBody = copyDoc.getActiveSection();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dataSheet = ss.getSheetByName("Report_Data");
  var dataRange = dataSheet.getRange(3, 1, dataSheet.getLastRow() - 2, dataSheet.getLastColumn());
  var objects = getRowsData(dataSheet, dataRange);
  var season = objects[0]['season'];
  var water_no = objects[0]['wateringNo'];
 
  // For every row object, create a personalized email/document from a template and send
  // it to the appropriate person.
  for (var i = 0; i < objects.length; ++i) {
    // Get a row object
    var rowData = objects[i];
    var copyBody = fillInTemplateFromObject(copyBody, rowData);
  }
    // Put in a page break between each user, but only after the first one
    //if( i > 0) {
    //  var pgBrk = body.appendPageBreak();
    //}
  var adviceNbr = water_no + " " + Utilities.formatDate(new Date(), tz, "yyyy/MM/dd") + " v01"; // get watering number and date
  var doc_name = season + config.DOC_PREFIX + adviceNbr;

  // Create the PDF file, rename it if required and delete the doc copy
    
  copyDoc.saveAndClose()

  var newFile = DriveApp.createFile(copyFile.getAs('application/pdf'))  

  if (doc_name !== '') {
  
    newFile.setName(doc_name)
  } 
  
  copyFile.setTrashed(true)
  
} // createPdf()

// https://stackoverflow.com/questions/16456933/google-apps-script-mailapp-use-document-for-htmlbody

function getDocAsHtml(docId){ 
  var url = 'https://docs.google.com/feeds/download/documents/Export?exportFormat=html&format=html&id=';
  return UrlFetchApp.fetch(url+docId).getContentText();
}

// https://stackoverflow.com/questions/16456933/google-apps-script-mailapp-use-document-for-htmlbody
/**
 * get a String containing the contents of the given document as HTML.
 * Uses DriveApp library, key Mi-n8njGmTTPutNPEIaArGOVJ5jnXUK_T.
 *
 * @param {String} docID ID of a Google Document
 *
 * @returns {String} Content of document, rendered in HTML.
 *
 * @see https://sites.google.com/site/scriptsexamples/new-connectors-to-google-services/driveservice
 */
function getDocAsHTML(docID) {
  var doc = DriveApp.getFileById(docID);
  var html = doc.file.content.src;
  var response = UrlFetchApp.fetch(html);
  var template = response.getContentText();

  return template;
}

function sendEmail(emailAddress, attachment){
    var EMAIL_TEMPLATE_ID = 'SOME_GOOGLE_DOC_ID';
    var emailTemplate = getDocAsHTML(EMAIL_TEMPLATE_ID);
    MailApp.sendEmail(emailAddress, "Subject", emailTemplate.getText(), {  
        htmlBody: emailTemplate });
}


// -- ===================================================================================

// https://www.labnol.org/code/19892-merge-multiple-google-documents

/* -- appendTable() - Creates and appends a new Table - This method will also append an
      empty paragraph after the table, since Google Docs documents cannot end with a table.
*/

function mergeGoogleDocs() {

  var docIDs = ['documentID_1','documentID_2','documentID_3','documentID_4'];
  var baseDoc = DocumentApp.openById(docIDs[0]);

  var body = baseDoc.getActiveSection();

  for (var i = 1; i < docIDs.length; ++i ) {
    var otherBody = DocumentApp.openById(docIDs[i]).getActiveSection();
    var totalElements = otherBody.getNumChildren();
    for( var j = 0; j < totalElements; ++j ) {
      var element = otherBody.getChild(j).copy();
      var type = element.getType();
      if( type == DocumentApp.ElementType.PARAGRAPH )
        body.appendParagraph(element);
      else if( type == DocumentApp.ElementType.TABLE )
        body.appendTable(element);
      else if( type == DocumentApp.ElementType.LIST_ITEM )
        body.appendListItem(element);
      else
        throw new Error("Unknown element type: "+type);
    }
  }
}
// -- ===================================================================================

// https://stackoverflow.com/questions/29032656/google-app-script-merge-multiple-documents-remove-all-line-breaks-and-sent-as

function mergeGoogleDocs2() {
// set folder ID were we should look for files to merge  
  var folder = DriveApp.getFolderById('0BwqMAWnXi8hMmljM3FZpaowb1'); 
  var docIDs = [];
  var files = folder.getFiles();

  while (files.hasNext()){
    file = files.next();
    docIDs.push(file.getId());
  }
// check if we have some ids  
  Logger.log(docIDs); 
// set document id of doc which will contain all merged documents  
  var baseDoc = DocumentApp.openById('0BwqMAWnXi8hMmljM3FZpaowb1');
// clear the whole document and start with empty page
  baseDoc.getBody().clear();
  var body = baseDoc.getActiveSection();

  for (var i = 1; i < docIDs.length; ++i ) {
    var otherBody = DocumentApp.openById(docIDs[i]).getActiveSection();    
    var totalElements = otherBody.getNumChildren();
    for( var j = 0; j < totalElements; ++j ) {
      var element = otherBody.getChild(j).copy();
      var type = element.getType();
      if( type == DocumentApp.ElementType.PARAGRAPH )
        body.appendParagraph(element);
      else if( type == DocumentApp.ElementType.TABLE )
        body.appendTable(element);
      else if( type == DocumentApp.ElementType.LIST_ITEM )
        body.appendListItem(element);
      else
        throw new Error("Unknown element type: "+type);
    }
  }
  // after merging all docs, invoke function to remove all line breaks in the just merged document
  removeMultipleLineBreaks();

  //email document
  emailDocument();
}

function removeMultipleLineBreaks(element) {
  if (!element) {
    // set document id of doc where to remove all line breaks 
    element = DocumentApp.openById('0BwqMAWnXi8hMmljM3FZpaowb1').getBody();
  }
  var parent = element.getParent();
  // Remove empty paragraphs
  if (element.getType() == DocumentApp.ElementType.PARAGRAPH 
      && element.asParagraph().getText().replace(/\s/g, '') == '') {
    if (!(parent.getType() == DocumentApp.ElementType.BODY_SECTION 
         && parent.getChildIndex(element) == parent.getNumChildren() - 1)) {
      element.removeFromParent();
    }
  // Remove duplicate newlines in text
  } else if (element.getType() == DocumentApp.ElementType.TEXT) {
    var text = element.asText();
    var content = text.getText();
    var matches;
    // Remove duplicate carriage returns within text.
    if (matches = content.match(/\r\s*\r/g)) {
      for (var i = matches.length - 1; i >= 0; i--) {
        var match = matches[i];
        var startIndex = content.lastIndexOf(match);
        var endIndexInclusive = startIndex + match.length - 1;
        text.deleteText(startIndex + 1, endIndexInclusive);
      }
    }
    // Grab the text again.
    content = text.getText();
    // Remove carriage returns at the end of the text.
    if (matches = content.match(/\r\s*$/)) {
      var match = matches[0];
      text.deleteText(content.length - match.length, content.length - 1);
    }
    // Remove carriage returns at the start of the text.
    if (matches = content.match(/^\s*\r/)) {
      var match = matches[0];
      text.deleteText(0, match.length - 1);
    }
  // Recursively look in child elements
  } else if (element.getNumChildren) {
    for (var i = element.getNumChildren() - 1; i >= 0; i--) {
      var child = element.getChild(i);
      removeMultipleLineBreaks(child);
    }
  }
}

function emailDocument() {
  //Replace this email address with your own email address
  var email = "sample@email.com"; 

  var fileToAttach = DriveApp.getFileById('Put your file ID here').getAs('application/pdf');

  var message = "This is a test message";
  var subject = "New Merged Document";

 // Send an email with an attachment: a file from Google Drive
 MailApp.sendEmail(email, subject, message, {
     attachments: [fileToAttach]
 });
}

// -- ===================================================================================
// https://gist.github.com/elkusbry/3e16557f1f12a1fa5a32c9e6d69b51ad

function generateShareholderPDF() {
  var sharePrice = 24.85;
  
  var spreadsheetId = '1MdZjiDC44RHh5Z_xWPXKK79B2x9XYDx4Hq3nqgm27ys';
  var rangeName = 'Sheet1!A1:B50';
  var values = Sheets.Spreadsheets.Values.get(spreadsheetId, rangeName).values;
  
  var outputDocId = '1kr5btUNI03l1jEfiRsCBbxzsmTAjwuC7jn5_n87F3hY';
  var shareholderStatementTemplateID = '1smaAifCE_P1K-JJF18mlokzQ9Lq_GIqyvv9GHgrNX7A';
  
  var templateBody = DocumentApp.openById(shareholderStatementTemplateID).getBody();
  var body = DocumentApp.openById(outputDocId).getBody();
  
  body.clear();
  
  if (!values) {
    Logger.log('No data found.');
  } else {
    for (var row = 0; row < values.length; row++) {
      var shareholderName = values[row][0];      
      var shareCount = values[row][1];
      var shareValue = shareCount * sharePrice;
      
      var statementHeader = templateBody.getChild(1).copy();
      var statementInvestor = templateBody.getChild(3).copy();
      var statementContent = templateBody.getChild(5).copy();
      var statementOwnership = templateBody.getChild(7).copy();
      
      body.appendTable(statementHeader);
      body.appendTable(statementInvestor);
      body.appendTable(statementContent);
      body.appendTable(statementOwnership);
      
      body.replaceText('{{shareholderName}}', shareholderName);
      body.replaceText('{{shareCount}}', shareCount);
      body.replaceText('{{shareValue}}', '$' + shareValue.toFixed(2).replace(/\d(?=(\d{3})+\.)/g, '$&,'));
      
      Logger.log(shareCount, shareValue);
      
      body.appendPageBreak()
    }
  }
}



// -- ===================================================================================
function testgetdata() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  //var dataSheet = ss.getSheetByName('Order_Detail');
  //var dataSheet = ss.getSheetByName('Current_Order');
  //var dataSheet = ss.getSheetByName('Water_Charges');
  //var dataSheet = ss.getSheetByName('Water_Users');
  //var dataSheet = ss.getSheetByName('MI_Charges');
  //var dataSheet = ss.getSheetByName('Water_Order');
  //var dataSheet = ss.getSheetByName('Report_Header');
  //var dataRange = dataSheet.getRange(7, 4, dataSheet.getLastRow() - 6, dataSheet.getLastColumn() - 3);
  var dataSheet = ss.getSheetByName("Report_Data");
  var dataRange = dataSheet.getRange(3, 1, dataSheet.getLastRow() - 2, dataSheet.getLastColumn());

  var objects = getRowsData(dataSheet, dataRange);
  for (var i = 0; i < objects.length; ++i) {
    // Get a row object
    var rowData = objects[i];
    Logger.log(rowData);;
  }
}  
  
  

