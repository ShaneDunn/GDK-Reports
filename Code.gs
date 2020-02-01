/**
 * == A Set of scripts to create reports and email them when required. ==
 *
 * Logging and configuration functions adapted from the script:
 * 'A Google Apps Script for importing CSV data into a Google Spreadsheet' by Ian Lewis.
 *  https://gist.github.com/IanLewis/8310540
 * @author ianmlewis@gmail.com (Ian Lewis)
 * @author dunn.shane@gmail.com (Shane Dunn)
 * July 2017
 *
 * Document merge and pdf creation functions adapted from 'Google Apps Script to create PDF' and 
 * 'Google Apps Script to create and email a PDF' by Andrew Roberts
 * @author ianmlewis@gmail.com (Ian Lewis)
 * @author dunn.shane@gmail.com (Shane Dunn)
 * January 2019
*/

/* =========== Globals ======================= */
var SPREADSHEET_ID = "1mHqk3VWYTkG7AqvqQ_UErGTky3QfYwVAUWwRtdPRHNA"; // "GDK Reports" spreadsheet
var FLOWCOM_SS_ID  = "1mlpCfUHtG1HY4fNDCiwKI92FWqMiQHKRcrJ2WlywT64"; // "Flocom Data" spreadsheet
var WATERSYN_SS_ID = "1ou_tJmY4t-e0umPLEMbGHaGXEFi8wcee1dAAUwiqmNQ"; // "Water Syndicate" spreadsheet
// var CONFIG_SHEET = 'Configuration';
// var LOG_SHEET = 'Log';
// var ERROR_SHEET = 'Errors';
var tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
var now = new Date();

/* =========== Setup Menu ======================= */
/**
/**
 * Eventhandler for spreadsheet opening
 * Create a Menu when the script loads. Adds a new configuration sheet if
 * one doesn't exist.
 */
function onOpen(e) {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('GDK Reports Menu')
    .addSubMenu(ui.createMenu('Print')
      .addItem('Water Statement', 'printWater_Statement')
      .addItem('Water Advice',    'printWater_Advice')
      .addItem('Water Request',   'printWater_Request')
    )
    .addSeparator()
    .addSubMenu(ui.createMenu('E-mail')
      .addItem('Water Statement', 'emailWater_Statement')
      .addItem('Water Advice',    'emailWater_Advice')
      .addItem('Water Request',   'emailWater_Request')
    )
    .addToUi();

  var sheet = getOrCreateSheet_(CONFIG_SHEET);
  if (sheet.getRange(1, 1).getValue() == ""){
    loadNewConfiguration(sheet);
  }
  sheet = getOrCreateSheet_(LOG_SHEET);
  if (sheet.getRange(1, 1).getValue() == ""){
    loadNewLog(sheet);
  }
  sheet = getOrCreateSheet_(ERROR_SHEET);
  if (sheet.getRange(1, 1).getValue() == ""){
    loadNewError(sheet);
  }
}


/**
 * Do-nothing method to trigger the authorization dialog if not already done.
 */
function checkAuthorization() {
}

/**
 * Create the printWater_Statement for printing
 */
function printWater_Statement(e) {
  setupLog_();
  var i, config, configName, sheet;
  log_('Running on: ' + now);
  
  var configs = getConfigs_(getOrCreateSheet_(CONFIG_SHEET));
  
  if (!configs.length) {
    log_('No report configurations found');
  } else {
    log_('Found ' + configs.length + ' report configurations.');
    run_report("Water Statement","Print");
  }
  log_('Script done');
    
  // Update the user about the status of the queries.
  if( e === undefined ) {
    displayLog_();
  } 
}

/**
 * Create the printWater_Advice for printing
 */
function printWater_Advice() {
}

/**
 * Create the printWater_Request for printing
 */
function printWater_Request() {
}

/**
 * Create the emailWater_Statement for email
 */
function emailWater_Statement() {
}

/**
 * Create the emailWater_Advicexx for email
 */
function emailWater_Advice() {
}

/**
 * Create the emailWater_Request for email
 */
function emailWater_Request() {
}

function get_report_config(report, configs) {
  var i, config, configName, sheet;
  for (i = 0; config = configs[i]; ++i) {
    configName = config.report;
    if (config['Report_Name'] === report) {
      log_('Using configuration from: ' + configName);
      return config;
    }
  }
  log_('No Report_Name found: ' + report);
}

function run_report(report, action) {
  var configs = getConfigs_(getOrCreateSheet_(CONFIG_SHEET));
  var config = get_report_config(report, configs);
  if (config['Report_Name']) {
    if (config.templateDocID) {
      try {
        log_('Creating Report: ' + config['Report_Name']);
        // sheet = getOrCreateSheet_(config['sheet-name']);
        // populateSheetWithCSV_(sheet, config.url, config['http-username'], config['http-password']);
      } catch (error) {
        log_('Error executing ' + config['Report_Name'] + ': ' + error.message);
      }
    } else {
      log_('No Template found: ' + config.templateDocID);
    }
  } else {
    log_('No Report_Name found: ' + report);
  }
}
