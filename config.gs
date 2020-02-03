/**
 * Logging and configuration functions adapted from the script:
 * 'A Google Apps Script for importing CSV data into a Google Spreadsheet' by Ian Lewis.
 *  https://gist.github.com/IanLewis/8310540
 * @author ianmlewis@gmail.com (Ian Lewis)
 * @author dunn.shane@gmail.com (Shane Dunn)
 * De Bortoli Wines July 2017
*/
/* =========== Globals ======================= */
var CONFIG_SHEET = 'Configuration';

/* =========== Configuration functions ======================= */
/**
 * Returns the values from 2 columns from the csvconfig sheet starting at
 * colIndex, as key-value pairs. Key-values are only returned if they do
 * not contain the empty string or have a boolean value of false.
 * If the key is start-date or end-date and the value is an instance of
 * the date object, the value will be converted to a string in yyyy-MM-dd.
 * If the key is start-index or max-results and the type of the value is
 * number, the value will be parsed into a string.
 * If value is "ColumnNames", the subsequent key value pairs are treated
 * as an array of values (used for column names)
 * @param {number} colIndex The column index to return values from.
 * @return {object} The values starting in colIndex and the following column
       as key-value pairs.
 */
function getConfigsStartingAtCol_(sheet, colIndex) {
  var config = {}, rowIndex, key, value, dvalue, columnDef, tblName;
  var range = sheet.getRange(1, colIndex, sheet.getLastRow(), 2);

  columnDef = false;

  // The first cell of the first column becomes the name of the query.
  config.report = range.getCell(1,1).getValue();

  for (rowIndex = 2; rowIndex <= range.getLastRow(); ++rowIndex) {
    key = range.getCell(rowIndex, 1).getValue();
    value = range.getCell(rowIndex, 2).getValue();
    dvalue = escapeQuotes(range.getCell(rowIndex, 2).getDisplayValue());
    if (value) {
      if ((key == 'start-date' || key == 'end-date') && value instanceof Date) {
        // Utilities.formatDate is too complicated since it requires a time zone
        // which can be configured by account or per sheet.
        dvalue = formatGaDate_(value);
        
      } else if ((key == 'start-index' || key == 'max-results') && typeof value == 'number') {
        dvalue = value.toString(); 
      }
      var trailNum = getTrailingNumber_(key)
      if ( columnDef && trailNum ) {
        config[tblName][trailNum] = escapeQuotes(value);
      } else {
        columnDef = false;
      }
      if ( columnDef || value == "ColumnNames") {
        if ( value == "ColumnNames") {
          tblName = key;
          columnDef = true;
          config[tblName] = [];
        }
      } else {
        config[key] = dvalue;
      }
    }
  }

  return config;
}

/**
 * Returns an array of config objects. This reads the csvconfig sheet
 * and tries to extract adjacent column names that end with the same
 * number. For example Names1 : Values1. Then both columns are used
 * to define key-value pairs for the coniguration object. The first
 * column defines the keys, and the adjacent column values define
 * each keys values.
 * @param {Sheet} The csvconfig sheet from which to read configurations.
 * @returns {Array} An array of API query configuration object.
 */
function getConfigs_(sheet) {

  var configs = [], colIndex;
  // There must be at least 2 columns.
  if (sheet.getLastColumn() < 2) {
    return configs;
  }

  var headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  var firstColValue, firstColNum, secondColValue, secondColNum;

  // Test the name of each column to see if it has an adjacent column that ends
  // in the same number. ie xxxx555 : yyyy555.
  // Since we check 2 columns at a time, we don't need to check the last column,
  // as there is no second column to also check. 
  for (colIndex = 1; colIndex <= headerRange.getNumColumns() - 1; ++colIndex) {
    firstColValue = headerRange.getCell(1, colIndex).getValue();
    firstColNum = getTrailingNumber_(firstColValue);
    
    secondColValue = headerRange.getCell(1, colIndex + 1).getValue();
    secondColNum = getTrailingNumber_(secondColValue);
  
    if (firstColNum && secondColNum && firstColNum === secondColNum) {
      configs.push(getConfigsStartingAtCol_(sheet, colIndex)); 
    }
  }
  return configs;  
}

/**
 * Returns 1 greater than the largest trailing number in the header row.
 * @param {Object} sheet The sheet in which to find the last number.
 * @returns {Number} The next largest trailing number.
 */
function getLastNumber_(sheet) {
  var maxNumber = 0;
  
  var lastColIndex = sheet.getLastColumn();

  if (lastColIndex > 0) {
    var range = sheet.getRange(1, 1, 1, lastColIndex);

    for (var colIndex = 1; colIndex < sheet.getLastColumn(); ++colIndex) {
      var value = range.getCell(1, colIndex).getValue();
      var headerNumber = getTrailingNumber_(value);
      if (headerNumber) {
        var number = parseInt(headerNumber, 10);
        maxNumber = number > maxNumber ? number : maxNumber;                                  
      }
    }
  }
  return maxNumber + 1;
}

/**
 * Returns the trailing number on a string. For example the
 * input: xxxx555 will return 555. Inputs with no trailing numbers
 * return undefined. Trailing whitespace is not ignored.
 * @param {string} input The input to parse.
 * @resturns {number} The trailing number on the input as a string.
 *     undefined if no number was found.
 */
function getTrailingNumber_(input) {
  // Match at one or more digits at the end of the string.
  var pattern = /(\d+)$/;
  var result = pattern.exec(input);
  if (result) {
    // Return the matched number.
    return result[0];
  }
  
  return undefined;
}

/**
 * Returns the dateInput object in yyyy-MM-dd.
 * @param {Date} dateInput The object to convert.
 * @returns {string} The date object as yyyy-MM-dd.
 */
function formatGaDate_(inputDate) {
  var output = [];
  var year = inputDate.getFullYear();

  var month = inputDate.getMonth() + 1;
  if (month < 10) {
    month = '0' + month; 
  }

  var day = inputDate.getDate();
  if (day < 10) {
    day = '0' + day; 
  }
  return [year, month, day].join('-');
}

function escapeQuotes(value) {
  if (!value) {
    return "";
  }
  if (typeof value != 'string') {
    value = value.toString();
  }
  return value.replace(/\\/g, '\\\\').replace(/'/g, "\\\'");
}
    

function getOrCreateSheet_(sheet_name) {
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = activeSpreadsheet.getSheetByName(sheet_name);
  if (!sheet) {
    sheet = activeSpreadsheet.insertSheet(sheet_name, 0);
  }
  return sheet;
}

/**
 * Adds a configuration to the spreadsheet.
 */
function createConfig() {
  loadNewConfig(getOrCreateSheet_(CONFIG_SHEET));
}

function loadNewConfig(sheet) {
  var headerNumber = getLastNumber_(sheet);
  var config = [
    ["Report_name_" + headerNumber, "value_" + headerNumber],
    ['Report_Name', 'Water Statement'],
    ['templateDocID', '10OlSE9c8__vLaqPydUnZQbSqNvQmxvogttx-wsWiVsE'],
    ['DOC_PREFIX', 'Water Statement - '],
    ['DUMMY_PARA', 'Remove'],
    ['WS_TABLE', 'Watering Record']
  ];
  //Logger.log(config);
  sheet.getRange(1, sheet.getLastColumn() + 1, config.length, 2).setValues(config);
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

/*
  var config = [
    ['ConfigKey', 'Value'],
    ['data-table-name', 'Flocom Data - Data'],
    ['event-table-name', 'Flocom Data - Events'],
    ['load-log-sheet-name', 'Log'],
    ['error-sheet-name', 'errors'],
    ['data-table-name-columns', 'ColumnNames'],
    ['data-table-column1', 'Reading Date/Time'],
    ['data-table-column2', 'Status'],
    ['data-table-column3', 'Flow Rate (Ml/day)'],
    ['data-table-column4', 'Actual Date/Time'],
    ['data-table-column5', 'Last Total'],
    ['data-table-column6', 'Ml/min'],
    ['data-table-column7', 'Interval (min)'],
    ['data-table-column8', 'Flow (Ml)'],
    ['data-table-column9', 'Adjusted Total'],
    ['data-table-column10', 'Cumulative Total'],
    ['event-table-name-columns', 'ColumnNames'],
    ['event-table-column1', 'Event Date/Time'],
    ['event-table-column2', 'Event'],
    ['event-table-column3', 'Result'],
  ];
*/

function testConfig() {
  var vConfig = getConfigs_(getOrCreateSheet_(CONFIG_SHEET));
  var vC1 = vConfig[0];
  Logger.log(vC1);
}

/*

[
{Report_Name=Water Statement,
  DUMMY_PARA=Remove,
  WS_TABLE=Watering Record,
  templateDocID=10OlSE9c8__vLaqPydUnZQbSqNvQmxvogttx-wsWiVsE,
  DOC_PREFIX= Water Statement - },

{Report_Name=Water Delivery Advice,
  DUMMY_PARA=Remove,
  WS_TABLE=Watering Schedule,
  templateDocID=13Wsgk7jmjXtjfyNorlX_OlBJLZoy3u8PuRarozETwOo,
  DOC_PREFIX= Water Delivery Advice - },

{Report_Name=Water Delivery Request,
  DUMMY_PARA=Remove,
  templateDocID=1qdC6s7Jgt1F8xauXyEu_khlTBYh8IExpNQaLMu_NPig,
  DOC_PREFIX=Water Allocation Request -}]
*/
function testGetFormating() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var thiscell = sheet.getRange("A1");
  Logger.log(thiscell.getA1Notation());          // String        Returns a string description of the range, in A1 notation.
  Logger.log(thiscell.getColumn());              // Integer       Returns the starting column position for this range.
  Logger.log(thiscell.getLastColumn());          // Integer       Returns the end column position.
  Logger.log(thiscell.getLastRow());             // Integer       Returns the end row position.
  Logger.log(thiscell.getNumColumns());          // Integer       Returns the number of columns in this range.
  Logger.log(thiscell.getNumRows());             // Integer       Returns the number of rows in this range.
  Logger.log(thiscell.getRow());                 // Integer       Returns the row position for this range.
  Logger.log(thiscell.getRowIndex());            // Integer       Returns the row position for this range.
  Logger.log(thiscell.getHeight());              // Integer       Returns the height of the range.
  Logger.log(thiscell.getWidth());               // Integer       Returns the width of the range in columns.
  Logger.log(thiscell.getFormula());             // String        Returns the formula (A1 notation) for the top-left cell of the range, or an empty string if the cell is empty or doesn't contain a formula.
  Logger.log(thiscell.getFormulaR1C1());         // String        Returns the formula (R1C1 notation) for a given cell, or null if none.
  Logger.log(thiscell.getGridId());              // Integer       Returns the grid ID of the range's parent sheet.
  Logger.log(thiscell.getDataSourceUrl());       // String        Returns a URL for the data in this range, which can be used to create charts and queries.

  Logger.log(thiscell.getDisplayValue());        // String        Returns the displayed value of the top-left cell in the range.
  Logger.log(thiscell.getValue());               // Object        Returns the value of the top-left cell in the range.
  Logger.log(thiscell.getNote());                // String        Returns the note associated with the given range.

  Logger.log(thiscell.getBackground());          // String        Returns the background color of the top-left cell in the range (for example, '#ffffff').
  Logger.log(thiscell.getFontColor());           // String        Returns the font color of the cell in the top-left corner of the range, in CSS notation (such as '#ffffff' or 'white').
  Logger.log(thiscell.getFontFamily());          // String        Returns the font family of the cell in the top-left corner of the range.
  Logger.log(thiscell.getFontLine());            // String        Gets the line style of the cell in the top-left corner of the range ('underline', 'line-through', or 'none').
  Logger.log(thiscell.getFontSize());            // Integer       Returns the font size in point size of the cell in the top-left corner of the range.
  Logger.log(thiscell.getFontStyle());           // String        Returns the font style ('italic' or 'normal') of the cell in the top-left corner of the range.
  Logger.log(thiscell.getFontWeight());          // String        Returns the font weight (normal/bold) of the cell in the top-left corner of the range.
  Logger.log(thiscell.getNumberFormat());        // String        Get the number or date formatting of the top-left cell of the given range.
  Logger.log(thiscell.getWrap());                // Boolean       Returns the wrapping policy of the cell in the top-left corner of the range.
  Logger.log(thiscell.getWrapStrategy());        // WrapStrategy  Returns the text wrapping strategy for the top left cell of the range.
  Logger.log(thiscell.getHorizontalAlignment()); // String        Returns the horizontal alignment of the text (left/center/right) of the cell in the top-left corner of the range.
  Logger.log(thiscell.getVerticalAlignment());   // String        Returns the vertical alignment (top/middle/bottom) of the cell in the top-left corner of the range.
  Logger.log(thiscell.getTextDirection());       // TextDirection Returns the text direction for the top left cell of the range.
}

/*
getA1Notation()           = "A1"
getColumn()               = 1.0
getLastColumn()           = 1.0
getLastRow()              = 1.0
getNumColumns()           = 1.0
getNumRows()              = 1.0
getRow()                  = 1.0
getRowIndex()             = 1.0
getHeight()               = 1.0
getWidth()                = 1.0
getFormula()              = ""
getFormulaR1C1()          = ""
getGridId()               = 8.70277097E8
getDataSourceUrl()        = "https://docs.google.com/spreadsheets/d/tIwWhRF6PsuT3kaz8WTMDXQ/gviz/tq?headers=-1&transpose=0&merge=rows&gid=870277097&range=A1"

getDisplayValue()         = "Error"
getValue()                = Error
getNote()                 = ""

setBackground("#efefef")
setFontColor("#000000")
setFontFamily("Verdana")
setFontLine("none")
setFontSize(12.0)
setFontStyle("normal")
setFontWeight("bold")
setNumberFormat("0.###############")
setWrap(true)
setWrapStrategy(OVERFLOW)
setHorizontalAlignment("general-left")
setVerticalAlignment("bottom")
setTextDirection(null)

setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

 sheet.getActiveRangeList().setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
        .setVerticalAlignment('middle')
        .setHorizontalAlignment('left')
        .setFontFamily('Roboto')
        .setFontSize(10)
        .setBorder(true, true, true, true, true, true, '#b7b7b7', SpreadsheetApp.BorderStyle.SOLID);
    }
*/
