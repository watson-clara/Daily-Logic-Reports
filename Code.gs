function main() {
  Logger.log("STARTING loadXLSX()");
  var date, id = loadXLSX();
  Logger.log("STARTING wrap(id)");
  wrap(id);
  Logger.log("STARTING copyDataToTemplate(date)");
  copyDataToTemplate(date, id);
}

function loadXLSX() {
  // calls specific folder in google drive
  const fldr = DriveApp.getFolderById("1-0WibmmZpk7xqOHvdw_S5UHtL784GgMX");
  Logger.log("got xlsx folder");
  // gets the xlsx files in the specific folder
  const xlsx_files = fldr.getFiles();
  Logger.log("got xlsx files");
  // creates bool to know if it is the first loop and to create file 
  var first = true;
  var count = 0;
  // loop throough every xlsx file in folder
  while (xlsx_files.hasNext()) {
    count = count + 1;
    Logger.log("NEW FILE #" + count);
    // sleep
    Utilities.sleep(1000);
    // find the next xlsx file 
    var xlsx = xlsx_files.next();
    // get the xlsx file name and id
    var xlsx_name = xlsx.getName();
    var id = xlsx.getId();
    Logger.log("xlsx file name:  " + xlsx_name);
    // splits the file name to get the name and date
    var name = xlsx_name.split("_")[1];
    var date = xlsx_name.split("_")[2];
    date = date.split(".")[0];
    Logger.log("new SS name:  " + name);
    Logger.log("date:  " +date);
    // copies data in xlsx file to copy to new google sheet file 
    var blob = xlsx.getBlob();
    // sleep
    Utilities.sleep(1000);
    // if statement to check if folder needs to be made 
    if (first == true){
      // finds the "daily reports file in my drive"
      var parentFolder = DriveApp.getFolderById("1DqUi0qe65wNACRP1BZF1AhlEDrW1Zbex");
      var newFolderID = makeFolder(date, parentFolder);
      first = false; 
    }
    // creates new google sheet file to convert xlsx to 
    var newFile = {
      title: name,
      parents: [{ id: newFolderID }],
      mimeType: MimeType.GOOGLE_SHEETS
    };
    Logger.log("new SS file created");
   
    // copies the xlsx content to the new file 
    Drive.Files.insert(newFile, blob);
    // deletes the xlsx file 
    Drive.Files.remove(id);
    Logger.log("old files deleted")
    Logger.log("done")
    // sleep 
    Utilities.sleep(1000);
  }
  var combinedSSID = combineToSingleSS( date, newFolderID);
  return date, combinedSSID;
}

function combineToSingleSS(date, folderID) {
  // calls specific folder in google drive
  const fldr = DriveApp.getFolderById(folderID);
  Logger.log("got folder");
  // gets the converted google sheet files in the specific folder
  const files = fldr.getFiles();
  Logger.log("got files");
  Logger.log(files);
  // creates new spreadsheet for data in individual xlsx sheets to be uploaded to
  var combinedSS = SpreadsheetApp.create("AS DIALER " + date);
  // gets new spreadsheet id 
  var id = combinedSS.getId();
  Logger.log(combinedSS.getName());
  // opens new spreadsheet so it can be edited
  var newSS = SpreadsheetApp.openById(id);
  // loops through GS files in folder to add to new spreadsheet
  while (files.hasNext()) {
    // sleep
    Utilities.sleep(1000);
    // gets next file
    var file = files.next();
    Logger.log(file);
    // gets next file name 
    var file_name = file.getName();
    Logger.log(file_name)
    // opens source spreadsheet
    var source = SpreadsheetApp.openById(file.getId());
    // gets sheet in that spreadsheets
    var target_sh = source.getSheets()[0];
    // copies the sheet from the source spreadsheet to the combined spreadsheeyt
    target_sh.copyTo(newSS).setName(file_name);
  }
  return id;
}

function makeFolder(name, parentFolder) {
  // creates a new file by the date to hold raw data
  var newFolder = parentFolder.createFolder(name);
  // saves created folders id 
  var newFolderID = newFolder.getId();
  // adds permissons
  newFolder.addEditor('cwatson@taxsupportteam.com');
  Logger.log("new folder for new SS's created");
  return newFolderID;
}

function copyDataToTemplate(date, id) {
  // raw data spreadsheet akak data scrapped from logics
  var ssRD = SpreadsheetApp.openById(id);
  // gets individual sheets within raw data spreadsheet
  var sheetsRD = ssRD.getSheets();
  // sorts sheets so that they are in the same order as the template
  var source = sortArray(sheetsRD);
  // report template spreadsheet
  var template = SpreadsheetApp.openById("1pzouN-Yr9lkKUtk8l8nbEteplY2XiUZeKRTrxgA3VIo");
  // makes copy of report template spreadsheet and renames it 
  var ssT = template.copy("TRAN REP VIA ACTIVITY " + date);
  Logger.log(ssT.getName());
  // open new final spreadsheet 
  var ssF = SpreadsheetApp.openByUrl(ssT.getUrl());
  //gets individual sheets within template spreadsheet
  var sheetsF = ssF.getSheets();
  //sorts sheets to be same order as raw data spreadsheet
  var destination = sortArray(sheetsF);
  // loop to copy data from 
  for (var s = 0 in source) {
    // gets individual sheet in raw data spreadsheet by index in array
    var sheetR = ssRD.getSheetByName(source[s]);
    Logger.log(sheetR.getName());
    // gets range of data to transfer
    var last_col = sheetR.getLastColumn();
    var last_row = sheetR.getLastRow();
    var range_input = sheetR.getRange(1, 1, last_row, last_col);
    // copies the values in the the raw data spreadsheet in the range 
    var values = range_input.getValues();
    // gets the sheets from the copied template 
    var sheetF = ssF.getSheetByName(destination[s]);
    Logger.log(sheetF.getName());
    // populates the raw data into the copied template spreadsheet
    sheetF.getRange(1, 1, last_row, last_col).setValues(values);
    Logger.log("");
  }
}

function removeEmptyColumns(sheet) {
  // gets how many columns are on the sheet
  var maxColumns = sheet.getMaxColumns();
  // gets how many used columbs on the sheet
  var lastColumn = sheet.getLastColumn();
  // if there are extra columns delete them
  if (maxColumns - lastColumn != 0) {
    sheet.deleteColumns(lastColumn + 1, maxColumns - lastColumn);
  }
}

function deleteBlankRows(sh) {
  // gets last used row 
  var lastRow = sh.getLastRow();
  // for each row check if the row is empty
  for (var raw = 1; raw < lastRow; raw++) {
    // if row has no data
    if (sh.getRange('A' + raw).getValue() == '') {
      // then delete it 
      sh.deleteRow(raw);
    }
  }
}

function sortArray(allsheets) {
  // creates empty array to store sheet names
  var sheetNameArray = [];
  // loop through sheets and add each name to array
  for (var i = 0; i < allsheets.length; i++) {
    sheetNameArray.push(allsheets[i].getName());
  }
  // sort the array 
  sheetNameArray.sort(function (a, b) {
    return a.localeCompare(b);
  });
  Logger.log(sheetNameArray);
  // return sorted array
  return sheetNameArray;
}

function wrap(id) {
  // opens raw data spreadsheet by id
  var sheet = SpreadsheetApp.openById(id);
  // gets sheet1 and deletes it 
  var s = sheet.getSheetByName("Sheet1");
  s.activate;
  sheet.deleteActiveSheet();
  // gets all sheets in spreadsheet
  var sheets = sheet.getSheets();
  // loops through each sheet in the spreasdsheet
  for (var i = 0; i < sheets.length; i++) {
    // gets first sheet
    var sheet = sheets[i];
    // gets data range of sheet
    var range = sheet.getDataRange();
    // sets the wrap strategy 
    range.setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW);
    Logger.log(i);
    Logger.log(sheet.getName());
    removeEmptyColumns(sheet);
    deleteBlankRows(sheet);
  }
}
