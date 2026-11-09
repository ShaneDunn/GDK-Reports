// https://github.com/gsuitedevs/apps-script-samples/blob/master/gmail/mailmerge/mailmerge.gs 
// http://www.andrewroberts.net/2016/01/google-apps-script-to-create-and-email-a-pdf/

/**
 * Copyright Google LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     https://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

/**
 * Iterates row by row in the input range and returns an array of objects.
 * Each object contains all the data for a given row, indexed by its normalized column name.
 * @param {Sheet} sheet The sheet object that contains the data to be processed
 * @param {Range} range The exact range of cells where the data is stored
 * @param {number} columnHeadersRowIndex Specifies the row number where the column names are stored.
 *   This argument is optional and it defaults to the row immediately above range;
 * @return {object[]} An array of objects.
 */
function getRowsData(sheet, range, columnHeadersRowIndex) {
  columnHeadersRowIndex = columnHeadersRowIndex || range.getRowIndex() - 1;
  var numColumns = range.getEndColumn() - range.getColumn() + 1;
  var headersRange = sheet.getRange(columnHeadersRowIndex, range.getColumn(), 1, numColumns);
  var headers = headersRange.getValues()[0];
  return getObjects(range.getDisplayValues(), normalizeHeaders(headers));
}

/**
 * For every row of data in data, generates an object that contains the data. Names of
 * object fields are defined in keys.
 * @param {object} data JavaScript 2d array
 * @param {object} keys Array of Strings that define the property names for the objects to create
 * @return {object[]} A list of objects.
 */
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

/**
 * Returns an array of normalized Strings.
 * @param {string[]} headers Array of strings to normalize
 * @return {string[]} An array of normalized strings.
 */
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

/**
 * Normalizes a string, by removing all alphanumeric characters and using mixed case
 * to separate words. The output will always start with a lower case letter.
 * This function is designed to produce JavaScript object property names.
 * @param {string} header The header to normalize.
 * @return {string} The normalized header.
 * @example "First Name" -> "firstName"
 * @example "Market Cap (millions) -> "marketCapMillions
 * @example "1 number at the beginning is ignored" -> "numberAtTheBeginningIsIgnored"
 */
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

/**
 * Returns true if the cell where cellData was read from is empty.
 * @param {string} cellData Cell data
 * @return {boolean} True if the cell is empty.
 */
function isCellEmpty(cellData) {
  return typeof(cellData) == 'string' && cellData == '';
}

/**
 * Returns true if the character char is alphabetical, false otherwise.
 * @param {string} char The character.
 * @return {boolean} True if the char is a number.
 */
function isAlnum(char) {
  return char >= 'A' && char <= 'Z' ||
    char >= 'a' && char <= 'z' ||
    isDigit(char);
}

/**
 * Returns true if the character char is a digit, false otherwise.
 * @param {string} char The character.
 * @return {boolean} True if the char is a digit.
 */
function isDigit(char) {
  return char >= '0' && char <= '9';
}

/**
 * Replaces markers in a template string with values define in a JavaScript data object.
 * @param {string} template Contains markers, for instance ${"Column name"}
 * @param {object} data values to that will replace markers.
 *   For instance data.columnName will replace marker ${"Column name"}
 * @return {string} A string without markers. If no data is found to replace a marker,
 *   it is simply removed.
 */
function fillInTemplateFromObject(template, data) {
  var doc = template;
  // Search for all the variables to be replaced, for instance ${"Column name"}
  // var templateVars = template.match(/\$\{\"[^\"]+\"\}/g);
  // Search for all the variables to be replaced, for instance <<columnName>>
  var templateVars = template.getText().match(/<<[^>]+>>/g);
  // Replace variables from the template with the actual values from the data object.
  // If no value is available, replace with the empty string.
  for (var i = 0; templateVars && i < templateVars.length; ++i) {
    // normalizeHeader ignores ${"} so we can call it directly here.
    //var variableData = data[normalizeHeader(templateVars[i])];
    var tidx = templateVars[i].substring(2,templateVars[i].length - 2);
    var variableData = data[tidx];
    doc = doc.replaceText(templateVars[i], variableData || '');
  }
  return doc;
}

// -- ===================================================================================
/**
 * Sends emails from spreadsheet rows.
 */
function sendEmails() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dataSheet = ss.getSheets()[0];
  var dataRange = dataSheet.getRange(2, 1, dataSheet.getMaxRows() - 1, 4);

  var templateSheet = ss.getSheets()[1];
  var emailTemplate = templateSheet.getRange('A1').getValue();

  // Create one JavaScript object per row of data.
  var objects = getRowsData(dataSheet, dataRange);

  // For every row object, create a personalized email from a template and send
  // it to the appropriate person.
  for (var i = 0; i < objects.length; ++i) {
    // Get a row object
    var rowData = objects[i];

    // Generate a personalized email.
    // Given a template string, replace markers (for instance ${"First Name"}) with
    // the corresponding value in a row object (for instance rowData.firstName).
    var emailText = fillInTemplateFromObject(emailTemplate, rowData);
    var emailSubject = 'Tutorial: Simple Mail Merge';

    MailApp.sendEmail(rowData.emailAddress, emailSubject, emailText);
    // var file = DriveApp.getFilesByName('2019_MA_BenefitsGuide.pdf')
    // MailApp.sendEmail({to: emailAddress, subject: subject, body: message, htmlBody: html}).
    // MailApp.sendEmail(emailAddress, "Subject", emailTemplate.getText(), {htmlBody: emailTemplate });

  }
}

// -- ===================================================================================
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

function createPDF(config, action) {

  // Set up the docs and the spreadsheet access
  Logger.log(config);
  // == Setup data for merging
  var gdk_folder = DriveApp.getFolderById(TARGET_FOLDER_ID);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dataSheet = ss.getSheetByName(config.reportTable);
  var dataRange = dataSheet.getRange(config.rT_start_row, config.rT_start_column, dataSheet.getLastRow() - config.rT_start_row + 1, dataSheet.getLastColumn() + config.rT_start_column - 1);
  var objects = getRowsData(dataSheet, dataRange);
  if  ( config.supTable ) {
    var t_dataSheet = ss.getSheetByName(config.supTable);
    var t_dataRange = t_dataSheet.getRange(config.sT_start_row, config.sT_start_column, t_dataSheet.getLastRow() - config.sT_start_row + 1, t_dataSheet.getLastColumn() + config.sT_start_column - 1);
    var t_objects = getRowsData(t_dataSheet, t_dataRange);
  }
  
  // == Initialise variables
  var season = objects[0]['season'];
  var water_no = objects[0]['wateringNo'];
  var adviceNbr = water_no + " " + Utilities.formatDate(new Date(), tz, "yyyy/MM/dd") + " v01"; // get watering number and date
  var doc_name = season + config.DOC_PREFIX + adviceNbr;
  var doc_ids = [];
  
  // For every row object, create a personalized document from a template
  for (var i = 0; i < objects.length; ++i) {
    // Get a row object
    var rowData = objects[i];
    log_('Generating Report for : ' + rowData.name);
    var doc_info = {};
    // == Setup New Documents from Template
    var full_doc_name = doc_name + "-" + rowData.name;
    var tempFile = DriveApp.getFileById(config.templateDocID);
    var copyFile = tempFile.makeCopy(full_doc_name, gdk_folder);
    
    var copyId = copyFile.getId();
    var copyDoc = DocumentApp.openById(copyId);
    var copyBody = copyDoc.getBody(); //getActiveSection();
    
    // if a table exists in the document - fill it first
    if  ( config.supTable ) {
      log_('Generating Report Table for : ' + rowData.name);

      var placeholder = findPlaceholder(copyBody, config.sT_placeholder);
      var table = findContainingTable(placeholder);
      var lines = [];
      
      for (var j = 0; j < t_objects.length; ++j) {
        // Get the row objects for this user
        var t_rowData = t_objects[j];
        if (t_rowData.name == rowData.name) {
          lines.push(t_rowData);
        }
      }
      for (var j = lines.length; j > 0; j-- ) {
        var tableRow = findContainingTableRow(placeholder);
        if (j!=1) {
          tableRow = table.insertTableRow(lines.length - j + 1, tableRow.copy());
        }
        var line = lines[lines.length - j];
        tableRow.replaceText(config.sT_placeholder, config.sT_ph_actual);
        tableRow = fillInTemplateFromObject(tableRow, line);
      }
    }
    copyBody = fillInTemplateFromObject(copyBody, rowData);
    copyDoc.saveAndClose();
    
    doc_info['docID']        = copyId;
    doc_info['emailAddress'] = rowData.emailAddress;
    // doc_info['emailAddress'] = "5han3dunn@gmail.com";
    doc_info['subject']      = full_doc_name;
    doc_ids.push(doc_info); // load array with all the documents created
    // break;
  }
  Logger.log("## Start Email");
  if ( action == "email") {
    emailDocuments(doc_ids);
    // mergeDocuments(doc_ids, doc_name, gdk_folder);
  }
  if (action == "print") {
    mergeDocuments(doc_ids, doc_name, gdk_folder);
  }
  
  // copyFile.setTrashed(true)
  
} // createPdf()

function emailDocuments(docIDs) {
  // go through array of documents to email to individuals
  Logger.log(docIDs);
  for (var i = 0; i < docIDs.length; ++i ) {
    var doc_info = docIDs[i];
    var fileToAttach = DriveApp.getFileById(doc_info.docID).getAs('application/pdf');
    var emailHTML = getDocAsHTML(doc_info.docID);
    var emailText = getDocAsText(doc_info.docID);
    log_('Emailing Report to ' + doc_info.emailAddress);
    Logger.log('## Emailing Report to ' + doc_info.emailAddress);
    MailApp.sendEmail(doc_info.emailAddress, doc_info.subject, "see Attached", {  
                      // htmlBody: emailHTML , 
                      attachments: [fileToAttach]
                    });
  }
}  // emailDocuments(docIDs)

function mergeDocuments(docIDs, doc_name, folder) {
  // create document to merge documents into - get the first one, copy and add to it
  var copyFile = DriveApp.getFileById(docIDs[0].docID).makeCopy(doc_name, folder);
  var copyId = copyFile.getId();
  var baseDoc = DocumentApp.openById(copyId);
  var body = baseDoc.getBody();
  if ( docIDs.length > 0 ) {
    body.appendPageBreak();
  }

  // go through array of documents and add to base document
  for (var i = 1; i < docIDs.length; ++i ) {
    var doc_info = docIDs[i];
    var otherBody = DocumentApp.openById(doc_info.docID).getBody();
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
  // removeMultipleLineBreaks();

}  // mergeDocuments(docIDs, doc_name, folder)

// -- ===================================================================================
/**
 * Create the printWater_Statement for printing
 */
function test1(e) {
  setupLog_();
  var i, config, configName, sheet;
  log_('Running on: ' + now);
  
  var configs = getConfigs_(getOrCreateSheet_(CONFIG_SHEET));
  
  if (!configs.length) {
    log_('No report configurations found');
  } else {
    log_('Found ' + configs.length + ' report configurations.');
    run_report(configs, "Water Statement","email");
  }
  log_('Script done');
    
  // Update the user about the status of the queries.
  if( e === undefined ) {
    displayLog_();
  } 
}

// -- ===================================================================================
// https://stackoverflow.com/questions/16456933/google-apps-script-mailapp-use-document-for-htmlbody

function getDocAsHtml(docId){ 
  var url = 'https://docs.google.com/feeds/download/documents/Export?exportFormat=html&format=html&id=';
  return UrlFetchApp.fetch(url+docId).getContentText();
}

// -- ===================================================================================
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
  var html = doc.getUrl(); // file.content.src;
  var response = UrlFetchApp.fetch(html);
  var template = response.getContentText();
    // nuke the whole head section, including the stylesheet and meta tag
    template = template.replace(/<head>.*<\/head>/, '');
    // remove almost all html attributes
    template = template.replace(/ (id|class|style|start|colspan|rowspan)="[^"]*"/g, '');
    // remove all of the spans, as well as the outer html and body
    template = template.replace(/<(span|\/span|body|\/body|html|\/html)>/g, '');
    // clearly the superior way of denoting line breaks
    template = template.replace(/<br>/g, '<br />');

  return template;
}

function getDocAsText(docID) {
  var doc = DocumentApp.openById(docID);
  var text = doc.getBody().getText();

  return text;
}

// -- ===================================================================================
// https://stackoverflow.com/questions/14663852/get-google-docume-ntas-html
/**
 * Takes in a Google Doc ID, gets that doc in HTML format, cleans up the markup, and returns the resulting HTML string.
 *
 * @param {string} the id of the google doc
 * @param {boolean} [useCaching] enable or disable caching. default true.
 * @return {string} the doc's body in html format
 */
function getContent(id, useCaching) {

  if (!id) {
    throw "Please call this API with a valid Google Doc ID";
  }

  if (useCaching == null) {
    useCaching = true;
  }

  if (typeof useCaching != "boolean") {
    throw "If you're going to specify useCaching, it must be boolean.";
  }

  var cache = CacheService.getScriptCache();
  var cached = cache.get(id); // see if we have a cached version of our parsed html
  if (cached && useCaching) {
    var html = cached;
    Logger.log("Pulling doc html from cache...");
  } else {

    Logger.log("Grabbing and parsing fresh html from the doc...");

    try {
      var doc = DriveApp.getFileById(id);
    } catch (err) {
      throw "Please call this API with a valid Google Doc ID. " + err.message;
    }

    var docName = doc.getName();

    var forDriveScope = DriveApp.getStorageUsed(); // needed to get Drive Scope requested in ScriptApp.getOAuthToken();
    var url = "https://docs.google.com/feeds/download/documents/export/Export?id=" + id + "&exportFormat=html";
    var param = {
      method: "get",
      headers: {"Authorization": "Bearer " + ScriptApp.getOAuthToken()},
      muteHttpExceptions:true,
    };

    var html = UrlFetchApp.fetch(url, param).getContentText();

    // nuke the whole head section, including the stylesheet and meta tag
    html = html.replace(/<head>.*<\/head>/, '');
    // remove almost all html attributes
    html = html.replace(/ (id|class|style|start|colspan|rowspan)="[^"]*"/g, '');
    // remove all of the spans, as well as the outer html and body
    html = html.replace(/<(span|\/span|body|\/body|html|\/html)>/g, '');
    // clearly the superior way of denoting line breaks
    html = html.replace(/<br>/g, '<br />');

    cache.put(id, html, 900) // cache doc contents for 15 minutes, in case we get a lot of requests

  }

  // Logger.log(html);

  return html;

}
// -- ===================================================================================
// https://www.labnol.org/code/19892-merge-multiple-google-documents

/* -- appendTable() - Creates and appends a new Table - This method will also append an
      empty paragraph after the table, since Google Docs documents cannot end with a table.
*/
// -- ===================================================================================
// https://stackoverflow.com/questions/29032656/google-app-script-merge-multiple-documents-remove-all-line-breaks-and-sent-as

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

// -- ===================================================================================
// https://gist.github.com/elkusbry/3e16557f1f12a1fa5a32c9e6d69b51ad

// -- ===================================================================================
function testgetdata() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  //var dataSheet = ss.getSheetByName('Current_Order');
  //var dataSheet = ss.getSheetByName('Water_Charges');
  //var dataSheet = ss.getSheetByName('Water_Users');
  //var dataSheet = ss.getSheetByName('MI_Charges');
  //var dataSheet = ss.getSheetByName('Water_Order');
  //var dataSheet = ss.getSheetByName('Report_Header');
  var dataSheet = ss.getSheetByName('Order_Detail');
  var dataRange = dataSheet.getRange(7, 4, dataSheet.getLastRow() - 6, dataSheet.getLastColumn() - 3);
  var t_objects = getRowsData(dataSheet, dataRange);
  
  dataSheet = ss.getSheetByName("Report_Data");
  dataRange = dataSheet.getRange(3, 1, dataSheet.getLastRow() - 2, dataSheet.getLastColumn());

  var objects = getRowsData(dataSheet, dataRange);
  for (var j = 0; j < objects.length; ++j) {
    // Get a row object
    var rowData = objects[j];
    Logger.log(rowData);
    for (var i = 0; i < t_objects.length; ++i) {
      // Get a row object
      var t_rowData = t_objects[i];
      if (t_rowData.name == rowData.name) {
        Logger.log(t_rowData);
      }
    }
  }
}  
