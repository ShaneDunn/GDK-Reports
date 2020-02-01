/**
 * Logging and configuration functions adapted from the script:
 * 'A Google Apps Script for importing CSV data into a Google Spreadsheet' by Ian Lewis.
 *  https://gist.github.com/IanLewis/8310540
 * @author ianmlewis@gmail.com (Ian Lewis)
 * @author dunn.shane@gmail.com (Shane Dunn)
 * De Bortoli Wines July 2017
*/
/* =========== Globals ======================= */
/**
 * The output text that should be displayed in the log.
 * @private.
 */
var logArray_;

var LOG_SHEET = 'Log';
var ERROR_SHEET = 'Errors';

/* =========== Logging functions ======================= */

/**
 * Clears the in app log.
 * @private.
 */
function setupLog_() {
  logArray_ = [];
}

/**
 * Returns the log as a string.
 * @returns {string} The log.
 */
function getLog_() {
  return logArray_.join('\n'); 
}

/**
 * Appends a string as a new line to the log.
 * @param {String} value The value to add to the log.
 */
function log_(value) {
  logArray_.push(value);
  
  var app = UiApp.getActiveApplication();
  var foo = app.getElementById('log');
  foo.setText(getLog_());
}

/**
 * Displays the log in memory to the user.
 */
function displayLog_() {
  var uiLog = UiApp.createApplication().setTitle('Report Status').setWidth(400).setHeight(500);
  var panel = uiLog.createVerticalPanel();
  uiLog.add(panel);

  var txtOutput = uiLog.createTextArea().setId('log').setWidth('400').setHeight('500').setValue(getLog_());
  panel.add(txtOutput);
  
  SpreadsheetApp.getActiveSpreadsheet().show(uiLog); 
}

function showLogDialog_() {
  var html = HtmlService.createHtmlOutput(getHTMLLog_())
      .setWidth(400)
      .setHeight(500);
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .showModalDialog(html, 'GDK Report Status');
}

function dumpLog_(sheet) {
  var lastRow = sheet.getLastRow() + 1;
  if (logArray_.length != 0) {
    var array = logArray_.map(function (el) {
          return [el];
    });
    sheet.getRange(lastRow,1,array.length,array[0].length).setValues(array);
    lastRow = sheet.getLastRow() + 1;
    sheet.getRange(lastRow,1).setValue(".").setBackground("#ffe2c6");
  }
}

function dumpError_(sheet) {
  var lastRow = sheet.getLastRow() + 1;
  if (errorArray_.length != 0) {
    sheet.getRange(lastRow,1,errorArray_.length,errorArray_[0].length).setValues(errorArray_);
    lastRow = sheet.getLastRow() + 1;
    sheet.getRange(lastRow,1).setValue(".").setBackground("#ffe2c6");
    sheet.getRange(lastRow,2).setValue(" ").setBackground("#ffe2c6");
  }
}

function loadNewLog(sheet) {
  var headerNumber = getLastNumber_(sheet);
  var header = [
    ["Log", "Comment"]
  ];
  //Logger.log(config);
  sheet.getRange(1, sheet.getLastColumn() + 1, header.length, 2).setValues(header);
  sheet.getRange("1:1").setBackground("#efefef")
                       .setFontColor("#000000")
                       .setFontFamily("Verdana")
                       .setFontLine("none")
                       .setFontSize(12.0)
                       .setFontStyle("normal")
                       .setFontWeight("bold")
                       .setNumberFormat("0.###############")
                       .setWrap(true)
                       .setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW)
                       .setHorizontalAlignment("general-left")
                       .setVerticalAlignment("bottom")
                       .setTextDirection(null);
}

function loadNewError(sheet) {
  var header = [
    ["Error", "Comment"]
  ];
  //Logger.log(config);
  sheet.getRange(1, sheet.getLastColumn() + 1, header.length, 2).setValues(header);
  sheet.getRange("1:1").setBackground("#efefef")
                       .setFontColor("#000000")
                       .setFontFamily("Verdana")
                       .setFontLine("none")
                       .setFontSize(12.0)
                       .setFontStyle("normal")
                       .setFontWeight("bold")
                       .setNumberFormat("0.###############")
                       .setWrap(true)
                       .setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW)
                       .setHorizontalAlignment("general-left")
                       .setVerticalAlignment("bottom")
                       .setTextDirection(null);
}
