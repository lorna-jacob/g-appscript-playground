//@OnlyCurrentDoc
function onOpen(e) {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Analyze CC Transactions 👉️")
    .addItem("Import from Drive", "importCSVFromDrive")
    .addItem("Process Data", "processData")
    .addToUi();
}

function processData() {
  var ss = SpreadsheetApp.getActive();
  
  // make sure sheet is not empty
  if(ss.getDataRange().isBlank()) {
    displayToastAlert("Sheet is empty. Please import data from csv first.")
  } else {
    displayToastAlert("Processing data...")

    // delete unnecessary columns 6-8
    ss.deleteColumns(6, 3);
  
    // if debit, change amount to minus
    changeAmountBasedOnType();

    // add category
    categorizeTransactions();

    displayToastAlert("Finished processing data.");
  }
}

//Displays an alert as a Toast message
function displayToastAlert(message) {
  SpreadsheetApp.getActive().toast(message, "⚠️ Alert"); 
}

function promptUserForInput(promptText) {
  var ui = SpreadsheetApp.getUi();
  var prompt = ui.prompt(promptText);
  var response = prompt.getResponseText();
  return response;
}

function importCSVFromUrl() {
  var url = promptUserForInput("Please enter the URL of the CSV file:");
  var contents = Utilities.parseCsv(UrlFetchApp.fetch(url));
  var sheetName = writeDataToSheet(contents);
  displayToastAlert("The CSV file was successfully imported into " + sheetName + ".");
}

//Returns files in Google Drive that have a certain name.
function findFilesInDrive(filename) {
  var files = DriveApp.getFilesByName(filename);
  var result = [];
  while(files.hasNext())
    result.push(files.next());
  return result;
}

function importCSVFromDrive() {
  var fileName = promptUserForInput("Please enter the name of the CSV file to import from Google Drive:");
  var files = findFilesInDrive(fileName);
  if(files.length === 0) {
    displayToastAlert("No files with name \"" + fileName + "\" were found in Google Drive.");
    return;
  } else if(files.length > 1) {
    displayToastAlert("Multiple files with name " + fileName +" were found. This program does not support picking the right file yet.");
    return;
  }
  var file = files[0];
  var contents = Utilities.parseCsv(file.getBlob().getDataAsString());
  var sheetName = writeDataToSheet(contents);
  displayToastAlert("The CSV file was successfully imported into " + sheetName + ".");
}

function writeDataToSheet(data) {
  var ss = SpreadsheetApp.getActive();
  sheet = ss.insertSheet();
  sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
  return sheet.getName();
}

function changeAmountBasedOnType() {
  var ss = SpreadsheetApp.getActive();
  var numRows = ss.getLastRow();
  
  ss.setActiveRange(ss.getRange('B:B'))

  for (var i = 1; i <= numRows; i++) {
    range = ss.getRange('B' + i );
    if (range.getValue() == "C") {
      var nextCell = range.offset(0, 1);
      var origValue = nextCell.getValue();
      nextCell.setValue("-" + origValue);
    }
  }
}

function categorizeTransactions() {
  var ss = SpreadsheetApp.getActive();
  var numRows = ss.getLastRow();
  
  ss.insertColumnAfter(5);
  ss.setCurrentCell(ss.getRange('F1'));
  ss.getCurrentCell().setValue("Category");
  ss.setActiveRange(ss.getRange('F:F'));

  for (var i = 1; i <= numRows; i++) {
    var range = ss.getRange('D' + i );

    var merchantName = range.getValue().toString();
    
    if (merchantName.indexOf('Apple.Com') > -1 || 
    merchantName.indexOf('Microsoft') > -1 ||
    merchantName.indexOf('Amazon Web Services') > -1 ||
    merchantName.indexOf('Trainerroad') > -1)
    {

      // Online Subscriptions
      range.offset(0, 2).setValue("Online Subscription");

    } else if (merchantName.indexOf('Bp Connect') > -1 || 
    merchantName.indexOf('Bp Connect') > -1 ||
    merchantName.indexOf('Rolleston Fuelstop') > -1 ||
    merchantName.indexOf('Z Rolleston') > -1) {

      // Petrol
      range.offset(0, 2).setValue("Petrol");

    } else if (merchantName.indexOf('New World') > -1 || 
    merchantName.indexOf('Countdown') > -1 ||
    merchantName.indexOf('Kosco') > -1 ||
    merchantName.indexOf('Japan Mart') > -1 ||
    merchantName.indexOf('Raeward Fresh') > -1 ||
    merchantName.indexOf('Sunson') > -1) {

      // Grocery
      range.offset(0, 2).setValue("Grocery");

    } else if (merchantName.indexOf('Mcdo') > -1 || 
    merchantName.indexOf('Kfc') > -1 ||
    merchantName.indexOf('Cj') > -1 ||
    merchantName.indexOf('Gongcha') > -1 ||
    merchantName.indexOf('Panadero') > -1 ||
    merchantName.indexOf('Ramen') > -1 ||
    merchantName.indexOf('Chatime') > -1 ||
    merchantName.indexOf('Ben Gong') > -1 ||
    merchantName.indexOf('Columbus') > -1 ||
    merchantName.indexOf('Coffee') > -1 ||
    merchantName.indexOf('Sushi') > -1){

      // Takeaways
      range.offset(0, 2).setValue("Takeaways");

    } else if (merchantName.indexOf('Aia') > -1 || 
    merchantName.indexOf('Southern Cross') > -1 ||
    merchantName.indexOf('Aa Insurance') > -1) {

      // Insurance
      range.offset(0, 2).setValue("Insurance");

    } else if (merchantName.indexOf('Bunnings') > -1 || 
    merchantName.indexOf('Mitre') > -1 ||
    merchantName.indexOf('Placemakers') > -1 ||
    merchantName.indexOf('Machineryhouse') > -1 ||
    merchantName.indexOf('Hub') > -1 ||
    merchantName.indexOf('Eb Games') > -1 ||
    merchantName.indexOf('JB Hi-Fi') > -1 ||
    merchantName.indexOf('The Warehouse') > -1 ||
    merchantName.indexOf('Kmart') > -1) {

      // Shopping
      range.offset(0, 2).setValue("Shopping");

    } else if (merchantName.indexOf('Metro Card') > -1  || 
    merchantName.indexOf('Parking') > -1 || 
    merchantName.indexOf('Zilch') > -1 || 
    merchantName.indexOf('Uber') > -1) {

      // Transportation
      range.offset(0, 2).setValue("Transportation");

    } else if (merchantName.indexOf('Bigpipe') > -1) {

      // Internet
      range.offset(0, 2).setValue("Internet");

    } else if (merchantName.indexOf('Skinny') > -1) {

      // Mobile Data
      range.offset(0, 2).setValue("Mobile Data");

    }
  }
}