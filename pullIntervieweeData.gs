function pull_interviewee_data(){ 
  // IF YOU NEED HELP CHANGING THIS, CALL HOPE ANDERSON AT (469)233-1891
 
  // Ensure "Application" sheet headers from "Name" to "Previous Involvement" match.
  // If the name of the application scoring sheet changes, edit the statement below!
  var APP_SCORING_SHEET_NAME = "Application";
  // If the index of the first row containing data on the "Application" sheet changes, edit statement below!
  var APP_SHEET_FIRST_DATA_ROW_INDEX = 4;
  // If the index of the row containing headers on the "Application" sheet changes, edit statement below!
  var APP_SHEET_HEADER_ROW_INDEX = 2;
  // If the index of the row containing headers on the "Interview" sheet changes, edit statement below!
  var INTERVIEW_SHEET_HEADER_ROW_INDEX = 2;
  // If the name of the column containing emails on the "Application" and "Interview" sheet changes, edit the statement below! 
  var SHEET_EMAIL_COL = "Email"
  // If the name of the column containing names on the "Application" and "Interview" sheet changes, edit statement below!
  // var SHEET_NAME_COL = "Name"; 
  // If the name of the column containing date and times of interview slots on the "Interview" sheet changes, edit statement below!
  var INTERVIEW_SHEET_DATETIME_COL_NAME = "Date and Time";
  
  
  
  
  
  
  // get the file name
  var file_name = Browser.inputBox("What is the name of the document containing the Sign Up Genius data? (This document must be in the Google drive.)");
  var file_iterator = DriveApp.getFilesByName(file_name);
  
  // count how many files have that name
  var count = 0;
  while (file_iterator.hasNext()) {
    var file = file_iterator.next();
    ++count;
  }
  
  // if there are no files with that name
  if (count == 0) {
    Browser.msgBox("There are no files with that name. Please check file name and try again.");
    return;
  }
  // if there is more than one file with that name
  if (count > 1) {
    Browser.msgBox("There are too many files with that name. Please rename the file and try again.");
    return;
  }
  
  // open sheets, get data, and get headers
  var signup_sheet = SpreadsheetApp.open(file).getSheets()[0]
  var signup_headers = signup_sheet.getDataRange().getValues()[0];
  var signup_data = signup_sheet.getRange(2, 1, signup_sheet.getLastRow() - 3, signup_sheet.getLastColumn()).getValues();
  
  var app_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(APP_SCORING_SHEET_NAME);
  var app_data = app_sheet.getRange(APP_SHEET_FIRST_DATA_ROW_INDEX,
                                    1,
                                    app_sheet.getLastRow() - APP_SHEET_FIRST_DATA_ROW_INDEX + 1, 
                                    app_sheet.getLastColumn())
                          .getValues();
  var app_headers = app_sheet.getDataRange().getValues()[APP_SHEET_HEADER_ROW_INDEX - 1];
  
  var interview_sheet = SpreadsheetApp.getActiveSheet();
  var interview_headers = interview_sheet.getDataRange().getValues()[INTERVIEW_SHEET_HEADER_ROW_INDEX - 1];
  
  // get emails of signups
  var email_col = signup_headers.indexOf("Email"); 
  var interviewee_emails = []; 
  for (i = 0; i < signup_data.length; ++i) {
    interviewee_emails[i] = []; 
    interviewee_emails[i][0] = signup_data[i][email_col]
  }
  Logger.log("interviewee_emails"); 
  Logger.log(interviewee_emails); 
  
  // var interviewee_emails = signup_data.getRange(2, email_col + 1, signup_sheet.getLastRow() - 3, 1).getValues(); 
  
  // get names of signups
//  var first_name_col = signup_headers.indexOf("First Name");
//  var last_name_col = signup_headers.indexOf("Last Name");
//  var source_names = [];
//  for(var i = 0; i < signup_data.length; ++i){
//    source_names[i] = [];
//    source_names[i][0] = signup_data[i][first_name_col] +" "+ signup_data[i][last_name_col];
//  }
  
  // input the emails of signups
  var target_email_col = interview_headers.indexOf(SHEET_EMAIL_COL) + 1; 
  var target = interview_sheet.getRange(4, target_email_col, interviewee_emails.length, 1); 
  target.setValues(interviewee_emails); 
  
  // input the names of signups
//  var target_name_col = interview_headers.indexOf(SHEET_NAME_COL) + 1;
//  var target = interview_sheet.getRange(4, target_name_col, source_names.length, 1); 
//  target.setValues(source_names);

  // get datetime and change to datetime format
  var signup_datetimes_col = signup_headers.indexOf("Start Date/Time (mm/dd/yyyy)") + 1;
  var datetimes = signup_sheet.getRange(2, signup_datetimes_col, signup_sheet.getLastRow() - 3, 1).getValues();
  for (datetime in datetimes) {
    datetimes[datetime][0] = Utilities.formatDate(datetimes[datetime][0], 'GMT-7', 'MM/dd/yyyy hh:mm a')
  }
  // input the datetimes of signups
  var target_col = interview_headers.indexOf(INTERVIEW_SHEET_DATETIME_COL_NAME) + 1;
  var target = interview_sheet.getRange(4, target_col, signup_sheet.getLastRow() - 3, 1); 
  target.setValues(datetimes);

  // copy other data from app sheet
  var app_email_col = app_headers.indexOf(SHEET_EMAIL_COL); 
  for (var i = 0; i < interviewee_emails.length; ++i) {
    for (var j = 0; j < app_data.length; ++j) {
      if (interviewee_emails[i][0].toLowerCase() == app_data[j][app_email_col].toLowerCase()) {
        interview_sheet.getRange(4 + i, 4, 1, 10).setValues([app_data[j].slice(app_email_col - 2, app_email_col + 8)]); 
  
  // copy other data from app sheet
//  var app_name_col = app_headers.indexOf(SHEET_NAME_COL);
//  for (var i = 0; i < source_names.length; ++i) {
//    for (var j = 0; j < app_data.length; ++j) {
//      if (source_names[i][0].toLowerCase() == app_data[j][app_name_col].toLowerCase()) {
//        interview_sheet.getRange(4 + i, 4, 1, 8).setValues([app_data[j].slice(app_name_col, app_name_col + 8)]);
      }
    }
  }
}
