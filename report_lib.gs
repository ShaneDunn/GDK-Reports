/*
 * http://www.andrewroberts.net/2016/01/google-apps-script-to-create-and-email-a-pdf/
 */

function getRowsData(sheet, range, columnHeadersRowIndex) {
  columnHeadersRowIndex = columnHeadersRowIndex || range.getRowIndex() - 1;
  var numColumns = range.getEndColumn() - range.getColumn() + 1;
  var headersRange = sheet.getRange(columnHeadersRowIndex, range.getColumn(), 1, numColumns);
  var headers = headersRange.getValues()[0];
  return getObjects(range.getValues(), normalizeHeaders(headers));
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

function fillInTemplateFromObject(template, data) {
  var email = template;
  // Search for all the variables to be replaced, for instance ${"Column name"}
  var templateVars = template.match(/\$\{\"[^\"]+\"\}/g);
  // Replace variables from the template with the actual values from the data object.
  for (var i = 0; i < templateVars.length; ++i) {
    var variableData = data[normalizeHeader(templateVars[i])];
    email = email.replace(templateVars[i], variableData || '');
  }
  return email;
}


// http://www.andrewroberts.net/2014/10/google-apps-script-create-pdf/

// dev: andrewroberts.net

// Replace this with ID of your template document.
var TEMPLATE_ID = ''

// var TEMPLATE_ID = '1wtGEp27HNEVwImeh2as7bRNw-tO4HkwPGcAsTrSNTPc' // Demo template
// Demo script - http://bit.ly/createPDF
 
// You can specify a name for the new PDF file here, or leave empty to use the 
// name of the template.
var PDF_FILE_NAME = ''

/**  
 * Take the fields from the active row in the active sheet
 * and, using a Google Doc template, create a PDF doc with these
 * fields replacing the keys in the template. The keys are identified
 * by having a % either side, e.g. %Name%.
 *
 * @return {Object} the completed PDF file
 */

function createPdf() {

  if (TEMPLATE_ID === '') {
    
    SpreadsheetApp.getUi().alert('TEMPLATE_ID needs to be defined in code.gs')
    return
  }

  // Set up the docs and the spreadsheet access
  
  var copyFile = DriveApp.getFileById(TEMPLATE_ID).makeCopy(),
      copyId = copyFile.getId(),
      copyDoc = DocumentApp.openById(copyId),
      copyBody = copyDoc.getActiveSection(),
      activeSheet = SpreadsheetApp.getActiveSheet(),
      numberOfColumns = activeSheet.getLastColumn(),
      activeRowIndex = activeSheet.getActiveRange().getRowIndex(),
      activeRow = activeSheet.getRange(activeRowIndex, 1, 1, numberOfColumns).getValues(),
      headerRow = activeSheet.getRange(1, 1, 1, numberOfColumns).getValues(),
      columnIndex = 0
 
  // Replace the keys with the spreadsheet values
 
  for (;columnIndex < headerRow[0].length; columnIndex++) {
    
    copyBody.replaceText('%' + headerRow[0][columnIndex] + '%', 
                         activeRow[0][columnIndex])                         
  }
  
  // Create the PDF file, rename it if required and delete the doc copy
    
  copyDoc.saveAndClose()

  var newFile = DriveApp.createFile(copyFile.getAs('application/pdf'))  

  if (PDF_FILE_NAME !== '') {
  
    newFile.setName(PDF_FILE_NAME)
  } 
  
  copyFile.setTrashed(true)
  
  SpreadsheetApp.getUi().alert('New PDF file created in the root of your Google Drive')
  
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


