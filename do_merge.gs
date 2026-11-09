
var selectedTemplateId = null;
var selectedSpreadsheetId = null;
var spreadsheetDocPicker = null;
var templateDocPicker = null;

function mailMerge(app) {
  var app = UiApp.createApplication().setTitle("Mail Merge");
  templateDocPicker = createFilePicker(app, "Choose template", 
         UiApp.FileType.DOCUMENTS, "templateSelectionHandler"); 
  templateDocPicker.showDocsPicker();
  return app;
};

function createFilePicker(app, title, fileType, selectionHandlerName) {
  Logger.log("Creating file picker for " + fileType);
  var docPicker = app.createDocsListDialog();
  docPicker.setDialogTitle(title);
  docPicker.setInitialView(fileType);
  var selectionHandler = app.createServerHandler(selectionHandlerName);
  docPicker.addSelectionHandler(selectionHandler);
  return docPicker;
}

function templateSelectionHandler(e) {
  var app = UiApp.getActiveApplication();
  selectedTemplateId = e.parameter.items[0].id;
  UserProperties.setProperty("templateId", e.parameter.items[0].id);
  Logger.log("Selected template: " + selectedTemplateId);
  var spreadsheetDocPicker = createFilePicker(app, "Choose spreadsheet", 
        UiApp.FileType.SPREADSHEETS, "spreadsheetSelectionHandler");
  spreadsheetDocPicker.showDocsPicker();
  return app;
}

function spreadsheetSelectionHandler(e) {
  var app = UiApp.getActiveApplication();
  UserProperties.setProperty("spreadsheetId", e.parameter.items[0].id);
  selectedSpreadsheetId = e.parameter.items[0].id;
  Logger.log("Selected spreadsheet: " + selectedSpreadsheetId);
  doMerge();
  return app;
}

function doMerge() {
  var selectedSpreadsheetId = UserProperties.getProperty("spreadsheetId");
  var selectedTemplateId = UserProperties.getProperty("templateId");
  Logger.log("Selected spreadsheet: " + selectedSpreadsheetId);
  var sheet = SpreadsheetApp.openById(selectedSpreadsheetId);
  Logger.log("Spreadsheet opened");
  Logger.log("Opening template: " + selectedTemplateId);
  var template = DocumentApp.openById(selectedTemplateId);
  Logger.log("Template opened");
  var templateFile = DocsList.getFileById(selectedTemplateId);
  var templateDoc = DocumentApp.openById(templateFile.getId());
  //var mergedFile = templateFile.makeCopy();
  var mergedDoc = DocumentApp.create("Result of mail merge");
  var bodyCopy = templateDoc.getActiveSection().copy();
  Logger.log("Copy made");
  var rows = sheet.getDataRange();
  var numRows = rows.getNumRows();
  var values = rows.getValues();
  var fieldNames = values[0];

  for (var i = 1; i < numRows; i++) {
    var row = values[i];
    Logger.log("Processing row " + i + " " + row);
    var body = bodyCopy.copy();
    for (var f = 0; f < fieldNames.length; f++) {
      Logger.log("Processing field " + f + " " + fieldNames[f]);
      Logger.log("Replacing [" + fieldNames[f] + "] with " + row[f]);
      body.replaceText("\\[" + fieldNames[f] + "\\]", row[f]);
    }
    var numChildren = body.getNumChildren();
    for (var c = 0; c < numChildren; c++) {
      var child = body.getChild(c);
      child = child.copy();
      if (child.getType() == DocumentApp.ElementType.HORIZONTALRULE) {
        mergedDoc.appendHorizontalRule(child);
      } else if (child.getType() == DocumentApp.ElementType.INLINEIMAGE) {
        mergedDoc.appendImage(child);
      } else if (child.getType() == DocumentApp.ElementType.PARAGRAPH) {
        mergedDoc.appendParagraph(child);
      } else if (child.getType() == DocumentApp.ElementType.LISTITEM) {
        mergedDoc.appendListItem(child);
      } else if (child.getType() == DocumentApp.ElementType.TABLE) {
        mergedDoc.appendTable(child);
      } else {
        Logger.log("Unknown element type: " + child);
      }
   }
   Logger.log("Appending page break");
   mergedDoc.appendPageBreak();
   Logger.log("Result is now " + mergedDoc.getActiveSection().getText());
  }
}

function testMerge() {
  UserProperties.setProperty("templateId", 
    "1pAXWE0uklZ8z-O_Tejuv3pWSTiSv583ptUTGPt2Knm8");
  UserProperties.setProperty("spreadsheetId", 
    "0Avea1NXBTibYdFo5QkZzWWlMYUhkclNSaFpRWUZOTUE");
  doMerge();
}


function doGet() {
  return mailMerge();
}
