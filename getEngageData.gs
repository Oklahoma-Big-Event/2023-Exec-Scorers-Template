function pullEngageData(){ 
  // IF YOU NEED HELP CHANGING THIS, CALL KELSEY DEWBRE AT (405)693-9924 OR HOPE ANDERSON AT (469)233-1891
  
  // this list contains all headers of the Engage sheet matched with corresponding headers of the application scoring sheet (engage sheet:application scoring sheet)
  // If headers change, this is where you will change the code!!!
  var HEADER_PAIRS = [["First Name", "First Name"],
                      ["Last Name", "Last Name"], 
                      ["Email", "Email"], 
                      ["Phone Number", "Phone Number"], 
                      ["Classification", "Classification"],
                      ["First Position Choice", "First Position"],
                      ["Second Position Choice", "Second Position"],
                      ["Third Position Choice", "Third Position"], 
                      ["Hours enrolled for Fall 2022", "Fall Hours"],
                      ["Anticipated hours enrolled for Spring 2023", "Spring Hours"]];
  
  // If the index of the first row with data in the 'Application' sheet for scoring changes, change below!
  var FIRST_DATA_ROW_INDEX = 4;
  // If the index of the first column with data in the 'Application' sheet for scoring changes, change below!
  var FIRST_DATA_COL_INDEX = 3;
  // If the index of the row containing the headers in the 'Application' scoring sheet changes, change below!
  var SCORE_HEADER_ROW_INDEX = 2;
  // If the column name in the orgsync export containing Submission Date changes, change below!
  var SUBMISSION_DATE_COL_NAME = "DateSubmitted";
  // If the index of the row containing headers in the Engage document changes, change below! 
  var ENGAGE_SHEET_HEADERS_INDEX = 3 

  
  var file_name = Browser.inputBox("What is the name of the document containing the application data? (This document must be in the Google drive and saved as a Google Sheet document.)");
  var file_iterator = DriveApp.getFilesByName(file_name);
  
  var count = 0;
  while (file_iterator.hasNext()) {
    var file = file_iterator.next();
    ++count;
  }
  
  if (count == 0) {
    Browser.msgBox("There were no files with that name. Please check file name and try again.");
    return;
  }
  if (count > 1) {
    Browser.msgBox("There are too many files with that name. Please rename the file and try again.");
    return;
  }
  
  // gets the vice chair position
 // var vice_chair = Browser.inputBox("What is your committee name? ");
  // these committee names and associated positions may need to be changed depending on admistrative structure
  //var committees = {"database": ["database coordinator"], "sponsorship": ["sponsorship"], "public relations": ["graphic designer", "social media specialist", "photographer/videographer", "publicity", "special events"], "jobsite recruitment": ["service project coordinator", "jobsite coordinator"], "operations":["operations staff coordinator", "supplies"], "outreach":["high school expansion coordinator", "international expansion coordinator", "college expansion coordinator", "alumni coordinator"]};  
  
  // var database = ["database coordinator"];
  // var outreach = ["high school expansion coordinator", "international expansion coordinator", "college expansion coordinator", "alumni coordinator"];
  // var campus_engagement = ["campus engagement"]; 
  // var sponsorship = ["sponsorship"]; 
  // var public_relations = ["graphic designer", "social media specialist", "photographer/videographer", "publicity", "special events", ]; 
  // var jobsite_recruitment = ["jobsite coordinator", "service project coordinator"];
  // var operations = ["operations staff coordinator", "supplies"]
  
  // open template and orgsync sheets, get data, and get headers
  var score_sheet = SpreadsheetApp.getActiveSheet();
  var score_headers = score_sheet.getDataRange().getValues()[SCORE_HEADER_ROW_INDEX - 1];
  var orgsync_sheet = SpreadsheetApp.open(file).getSheets()[0];
  var orgsync_headers = orgsync_sheet.getDataRange().getValues()[ENGAGE_SHEET_HEADERS_INDEX - 1];
  console.log(orgsync_headers);
  
  // Find how many orgsync apps
  var num_apps = orgsync_sheet.getRange("A4:A").getValues().filter(String).length;
  // Find how many entries are already inserted
  var num_entries = score_sheet.getRange(FIRST_DATA_ROW_INDEX,
                                         FIRST_DATA_COL_INDEX, 
                                         score_sheet.getLastRow(), 
                                         1)
                               .getValues().filter(String).length;
  
  var num_test = num_apps - num_entries; 
  
  // return if no new entries
  if (num_apps == num_entries) {
    Browser.msgBox("There are no new entries to import.")
    return; 
  }
  
  // Sort orgysnc sheet on submission date
  var submission_datetime_col = orgsync_headers.indexOf(SUBMISSION_DATE_COL_NAME) + 1;
  console.log(submission_datetime_col);
  orgsync_sheet.getRange(4, submission_datetime_col, orgsync_sheet.getLastRow() - 1, 1).setNumberFormat('yyyy-MM-dd hh:mm');
  var orgsync_data = orgsync_sheet.getRange(4, 1, orgsync_sheet.getLastRow(), orgsync_sheet.getLastColumn());
  orgsync_data.sort(submission_datetime_col);
             
  // Go through each header, pasting data into the sheet
  for(var i = 0; i < HEADER_PAIRS.length; ++i) {
    var header_test = orgsync_headers.indexOf(HEADER_PAIRS[i][0]); 
    var orgsync_col = (orgsync_headers.indexOf(HEADER_PAIRS[i][0])) + 1;
    var source_row_test = ENGAGE_SHEET_HEADERS_INDEX + 1 + num_entries; 
    var source = orgsync_sheet.getRange(source_row_test, orgsync_col, num_test, 1).getValues(); 
    var target_col = score_headers.indexOf(HEADER_PAIRS[i][1]) + 1;
    var target = score_sheet.getRange(4 + num_entries, target_col, num_test, 1); 
    target.setValues(source);
   }
} 
