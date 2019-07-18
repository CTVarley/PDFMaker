function onOpen(){
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('PDF Publisher')
 .addItem('Publish Week','publishWeek')
 .addSeparator()
 .addSubMenu(SpreadsheetApp.getUi().createMenu('Publish Day')
                  .addItem('Monday', 'Monday')
                  .addItem('Tuesday', 'Tuesday')
                  .addItem('Wednesday', 'Wednesday')
                  .addItem('Thursday', 'Thursday')
                  .addItem('Friday', 'Friday'))
 .addToUi();
}

function sleepLength() {
  // If you get an Error Number 429, increase this number (in milliseconds):
  return 38000;
}

function Monday(){
  publishDay("Monday");
}
function Tuesday(){
  publishDay("Tuesday");
}
function Wednesday(){
  publishDay("Wednesday");
}
function Thursday(){
  publishDay("Thursday");
}
function Friday(){
  publishDay("Friday");
}

function publishDay(day) {  
  // TODO: function out redundancies between publishDay and publishWeek  
  
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  SpreadsheetApp.setActiveSheet(spreadsheet.getSheetByName("School_Delivery"));
  var sheetWithDate = SpreadsheetApp.getActiveSheet();
  sheetWithDate.getRange('E2').setValue(day);
  var targetDate = sheetWithDate.getRange("E3").getValue();
  targetDate = Utilities.formatDate(targetDate, "GMT", "yyyy-MM-dd-EEEE");  
  sheetWithDate.getRange('E2').setValue('Monday');
  var mondayDate = sheetWithDate.getRange('E3').getValue();
  mondayDate = Utilities.formatDate(mondayDate, 'GMT', 'yyyy-MM-dd');
  var foldersIds = findTargetFolders(day, mondayDate, sheetWithDate);
  generatePdf(targetDate, foldersIds);
  saveSnapshot(spreadsheet, foldersIds[4]);
}

// finding target folder IDs for exporting one day at a time
function findTargetFolders(weekday, mondayDate, sheetWithDate) {
  var cellWithCycleWeek = 'E4';
  var strCycleWeek = String(sheetWithDate.getRange(cellWithCycleWeek).getDisplayValue());
  return createFolderStructure(mondayDate, strCycleWeek);
}

function publishWeek() {  
  options = {muteHttpExceptions: false};  
  days = ["Monday","Tuesday","Wednesday","Thursday","Friday"]
  for (var i = 0; i < days.length; i++) {    
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    SpreadsheetApp.setActiveSheet(spreadsheet.getSheetByName("School_Delivery"));    
    var sheetWithDate = SpreadsheetApp.getActiveSheet();
    sheetWithDate.getRange('E2').setValue(days[i]);
    var targetDate = sheetWithDate.getRange("E3").getValue();
    var mondayDate = Utilities.formatDate(targetDate, 'GMT', 'yyyy-MM-dd');
    targetDate = Utilities.formatDate(targetDate, "GMT", "yyyy-MM-dd-EEEE");
    if (days[i] == "Monday") {
      var cellWithCycleWeek = 'E4';
      var strCycleWeek = String(sheetWithDate.getRange(cellWithCycleWeek).getDisplayValue());
      // There might be a more elegant way to pass this ID back so that it's clear this function is being called for its side effects:
      var foldersIds = createFolderStructure(mondayDate, strCycleWeek); 
    }
    
    generatePdf(targetDate, foldersIds);
    
    Logger.log("Pausing for 20 seconds, to avoid 429 error while processing the whole week.");
    Utilities.sleep(sleepLength());
  }
  saveSnapshot(spreadsheet, foldersIds[4])
  // TODO: function that saves a version of the whole file still in sheet form
  Browser.msgBox("All done! You can find your files in the target folders.");    
}



function generatePdf(targetDate, foldersIds) {
  
  var sheetsToConvert = ["School_Delivery", "COLD_Delivery", "Prep"];
  var sheetNames = ['UAECS_Delivery', 'Propel_Delivery', 'Prep'];

  
  for (var i = 0; i < sheetsToConvert.length; i++) {
    
    SpreadsheetApp.flush();    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    SpreadsheetApp.setActiveSheet(ss.getSheetByName(sheetsToConvert[i]));
    var sheet = ss.getActiveSheet();
    var sheetName = sheetNames[i];
    
    //Date set with format in EST (NYC) used in subject and PDF name 
    var currentDate = "generated on " + Utilities.formatDate(new Date(), "GMT+5", "yyyy.MM.dd");
    Logger.log(currentDate);
    var url = ss.getUrl();
    
    var weekFolderID = foldersIds[4]; // The source of this is the "createFolderStructure" function.
    var folder = DriveApp.getFolderById(foldersIds[i]);
    
    
    
    //remove the trailing 'edit' from the url
    url = url.replace(/edit$/, '');
    
    // An if statement: if this is the Prep sheet
    // a loop: save a file for each of three ranges / names
    
    if (sheetsToConvert[i] == 'Prep') {
      var rangesToPrint = ['&r1=0&c1=0&r2=19&c2=14', '&r1=0&c1=14&r2=19&c2=24', '&r1=0&c1=24&r2=19&c2=39'];
      var prepTableNames = ['Lunch', 'Breakfast', 'Cold'];

      for (var i = 0; i < rangesToPrint.length; i++) {
        var fileName = nameFile(targetDate, (sheetName + '_' + (i + 1) + '_' + prepTableNames[i]), folder);
        //additional parameters for exporting the sheet as a pdf (if this block has an outside source, can we credit it?)
        var url_ext = 'export?exportFormat=pdf&format=pdf' + //export as pdf
          //below parameters are optional...
          '&size=letter' + //paper size
            '&portrait=false' + //orientation, false for landscape
              '&fitw=true' + //fit to width, false for actual size
                '&fith=true' + 
                  rangesToPrint[i] + 
                    '&sheetnames=true&printtitle=true&pagenumbers=true' + //hide optional headers and footers
                      '&printnotes=false' +
                        '&gridlines=false' + //hide gridlines
                          '&fzr=false' + //do not repeat row headers (frozen rows) on each page
                            '&gid=' + sheet.getSheetId(); //the sheet's Id
        
        var token = ScriptApp.getOAuthToken();  
        var response = UrlFetchApp.fetch(url + url_ext, {headers: {'Authorization': 'Bearer ' + token}});
        //newFolder = newFolder.hasNext() ? newFolder.next() : parentFolder.createFolder(childFolderName);    
        var blob = response.getBlob().setName(fileName);
        var newFile = folder.createFile(blob);
      }
        
      // EXPORT RANGE OPTIONS FOR PDF
      //need all the below to export a range
      //gid=sheetId                must be included. The first sheet will be 0. others will have a uniqe ID
      //ir=false                   seems to be always false
      //ic=false                   same as ir
      //r1=Start Row number - 1        row 1 would be 0 , row 15 wold be 14
      //c1=Start Column number - 1     column 1 would be 0, column 8 would be 7   
      //r2=End Row number
      //c2=End Column number
      
      
    } else {
      var fileName = nameFile(targetDate, sheetName, folder);
      var rangeToPrint = '';
      //additional parameters for exporting the sheet as a pdf (if this block has an outside source, can we credit it?)
      var url_ext = 'export?exportFormat=pdf&format=pdf' + //export as pdf
        //below parameters are optional...
        '&size=letter' + //paper size
          '&portrait=false' + //orientation, false for landscape
            '&fitw=true' + //fit to width, false for actual size
              rangeToPrint + 
                '&sheetnames=true&printtitle=true&pagenumbers=true' + //hide optional headers and footers
                  '&printnotes=false' +
                    '&gridlines=false' + //hide gridlines
                      '&fzr=false' + //do not repeat row headers (frozen rows) on each page
                        '&gid=' + sheet.getSheetId(); //the sheet's Id
      
      var token = ScriptApp.getOAuthToken(); 
      var response = UrlFetchApp.fetch(url + url_ext, {headers: {'Authorization': 'Bearer ' + token}}); 
      //newFolder = newFolder.hasNext() ? newFolder.next() : parentFolder.createFolder(childFolderName);
      var blob = response.getBlob().setName(fileName);
      var newFile = folder.createFile(blob);
    }
  }
}

// Deprecated
function findBreakRows() {
  sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("School_Delivery");
  value = "Signature:";
  
  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();
  
  var rowsBeforeBreaks = [];
  
  for (var i = 0; i < values.length; i++) {
    for (var j = 0; j < values[i].length; j++) {     
      if (values[i][j] == value) {
        rowsBeforeBreaks.push((i+3)); // The +3 gets us the row two after the search term.  This works for the pages with signature, may need to be adjusted for others.
      }
    }    
  }
  Logger.log("Rows that could be expanded to account for page breaks:  " + rowsBeforeBreaks)
  return rowsBeforeBreaks;
}

function nameFile(targetDate, sheetName, folder) {
  var fileName = targetDate + "_" + sheetName;  
  var oldFile = folder.getFilesByName(fileName);
  fileName = oldFile.hasNext() ? fileName + '_UPDATED_' + Utilities.formatDate(new Date(), localTimeZone(), 'MM-dd-hh:mm') : fileName;
  Logger.log('Exporting ' + fileName + '...');
  return fileName;  
}

function getCycleWeekNum(strCycleWeek) {
  strCycleWeek = String(strCycleWeek);
  for (var i = 1; i < 5; i++) {
    // if strCycleWeek includes i, then return i
    if (strCycleWeek.indexOf(i) > -1) {
      return i;
    }
  }
}

// Could be reused in loop for subfolders?
function createSubFolder(parentFolderID, childFolderName) {
  var parentFolder = DriveApp.getFolderById(parentFolderID);
  var newFolder = parentFolder.getFoldersByName(childFolderName);
  newFolder = newFolder.hasNext() ? newFolder.next() : parentFolder.createFolder(childFolderName);
  return newFolder;
}

function saveSnapshot(spreadsheet, folderID){ 
  var destFolder = DriveApp.getFolderById(folderID); 
  var fileName = "Snapshot generated on " + Utilities.formatDate(new Date(), localTimeZone(), "yyyy.MM.dd, h:mm a")
  DriveApp.getFileById(spreadsheet.getId()).makeCopy(fileName, destFolder); 
}


function createFolderStructure(mondayDate, cycleWeek){    
  // var PARENT_FOLDER_ID = '1vnSexVhPkCqJWhi3worsjHF2o28qBWkA'; // TODO: move to config block
  var rootFolder = DriveApp.getRootFolder();
  var parentFolder = createSubFolder(rootFolder.getId(), "Weekly_Sheets");
  
  
  cycleWeek = getCycleWeekNum(cycleWeek);
  var weekFolderName = mondayDate + "_w" + cycleWeek; // Call function to name Folder for whole week
  var weekFolder = createSubFolder(parentFolder.getId(), weekFolderName);  
  // Would like to change this array to something more coherent but for now this is a "get it to green" type implementation until we determine what constants can be global configuration.
  var foldersIds = [];
  foldersIds[4] = weekFolder.getId();
  
  var CHILD_FOLDER_NAMES = ['UAECS_Delivery', 'Propel_Delivery', 'Prep_Sheets', 'Orders'];
  for (var i = 0; i < CHILD_FOLDER_NAMES.length; i++) {
    var newFolderName = CHILD_FOLDER_NAMES[i];
    var newFolder = weekFolder.getFoldersByName(newFolderName);
    newFolder = newFolder.hasNext() ? newFolder.next() : weekFolder.createFolder(newFolderName);
    foldersIds[i] = newFolder.getId();
    Logger.log(CHILD_FOLDER_NAMES[i] + foldersIds[i]);
  }
  return foldersIds;
}

function localTimeZone() { 
  var localTimeOffset = (new Date()).getTimezoneOffset() / 60;
  return "GMT-" + localTimeOffset;
}