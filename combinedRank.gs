function combined_rank() {
  // gets the rank sheet and headers
  var rank_sheet = SpreadsheetApp.getActiveSheet();
  var rank_headers = rank_sheet.getRange(3, 1, 1, 7).getValues()[0]; 
  // var rank_headers = rank_sheet.getDataRange().getValues()[1];
  // opens scored application sheet and gets data and headers
  var app_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Application");
  var app_data = app_sheet.getRange(4, 1, app_sheet.getRange("C4:C").getValues().filter(String).length, app_sheet.getLastColumn()).getValues();
  var app_headers = app_sheet.getDataRange().getValues()[1];
  // opens scored interview sheet and gets data and headers
  var interview_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Interview");
  var interview_headers = interview_sheet.getDataRange().getValues()[1];
  var interview_first_name_col = interview_headers.indexOf("First Name");
  var interview_last_name_col = interview_headers.indexOf("Last Name"); 
  
  var first_position_index = 7; 
  var second_position_index = 8; 
  var third_position_index = 9; 

  function not_empty(row) {
    return /\S/.test(row[interview_first_name_col]);
  }
  var interview_data = interview_sheet.getRange(4, 1, interview_sheet.getLastRow() - 3, interview_sheet.getLastColumn()).getValues().filter(not_empty);
    
  // get relevant headers
  var rank_first_name_col = rank_headers.indexOf("First Name");
  var rank_last_name_col = rank_headers.indexOf("Last Name"); 
  var app_first_name_col = app_headers.indexOf("First Name");
  var app_last_name_col = app_headers.indexOf("Last Name"); 
  
  var rank_appaverage_col = rank_headers.indexOf("Average App Score");
  var app_appaverage_col = app_headers.indexOf("Average");
  
  var rank_interviewaverage_col = rank_headers.indexOf("Average Interview Score");
  
  var interview_interviewaverage_col = interview_headers.indexOf("Average");
  
  var interview_email_col = interview_headers.indexOf("Email");
  var app_email_col = app_headers.indexOf("Email"); 
  // var rank_email_col = rank_headers.indexOf("Email");
  
  var rank_average_col = rank_headers.indexOf("Average of App and Interview");
  
  var rank_col = rank_headers.indexOf("Rank");
  
  // combine interview and app score data
  var combined_data = [];
  for (var i = 0; i < interview_data.length; ++i) {
    combined_data[i] = "person's name does not match";
    for (var j = 0; j < app_data.length; ++j) {
      if (interview_data[i][interview_email_col].toLowerCase() == app_data[j][app_email_col].toLowerCase()) {
        var average = (app_data[j][app_appaverage_col] + interview_data[i][interview_interviewaverage_col])/2;
        combined_data[i] = [interview_data[i][interview_first_name_col], 
                            interview_data[i][interview_last_name_col], 
                            app_data[j][app_appaverage_col], 
                            interview_data[i][interview_interviewaverage_col], 
                            average, 
                            interview_data[i][first_position_index], 
                            interview_data[i][second_position_index],
                            interview_data[i][third_position_index]];
      }
    }
    if (combined_data[i] == "person's name does not match") {
      combined_data[i] = [interview_data[i][interview_first_name_col], 
                          interview_data[i][interview_last_name_col], 
                          "", 
                          interview_data[i][interview_interviewaverage_col], 
                          interview_data[i][interview_interviewaverage_col], 
                          interview_data[i][first_position_index], 
                          interview_data[i][second_position_index],
                          interview_data[i][third_position_index]
                          ]
    }
  }
  
  combined_data.sort(function(a, b){return b[4] - a[4]});
  Logger.log("combined data"); 
  Logger.log(combined_data); 
  
  Logger.log("test indexing of combined data"); 
  Logger.log(combined_data[0][5]); 
  
  for (var i = 0; i < combined_data.length; ++i) {
    rank_sheet.getRange(i + 4, 2, 1, 1).setValue(i + 1);
    // rank_sheet.getRange(i + 4, rank_col + 1, 1, 1).setValue(i + 1);
    rank_sheet.getRange(i + 4, rank_first_name_col + 1, 1, 1).setValue(combined_data[i][0]);
    rank_sheet.getRange(i + 4, rank_last_name_col + 1, 1, 1).setValue(combined_data[i][1]); 
    rank_sheet.getRange(i + 4, rank_appaverage_col + 1, 1, 1).setValue(combined_data[i][2]);
    rank_sheet.getRange(i + 4, rank_interviewaverage_col + 1, 1, 1).setValue(combined_data[i][3]);
    rank_sheet.getRange(i + 4, rank_average_col + 1, 1, 1).setValue(combined_data[i][4]);
  }
  
  var position_name_indeces = [["Overall", 2, 7 ], 
                               ["Service Project Coordinator", 9, 14], 
                               ["Jobsite Coordinator", 16, 21], 
                               ["Operations Staff Coordinator", 23, 28], 
                               ["Supplies", 30, 35], 
                               ["High School Expansion Coordinator", 37, 42], 
                               ["International Expansion Coordinator", 44, 49], 
                               ["College Expansion Coordinator", 51, 56], 
                               ["Alumni Coordinator", 58, 63], 
                               ["Campus Engagement", 65, 70], 
                               ["Sponsorship", 72, 77], 
                               ["Social Media Specialist", 79,84], 
                               ["Graphic Designer", 86, 91], 
                               ["Publicity", 93, 98], 
                               ["Special Events", 100, 105], 
                               ["Photographer/Videographer", 107, 112], 
                               ["Database Coordinator", 114, 119]]; 
    
    for (var i = 0; i < position_name_indeces.length; ++i) {
      
      // get position name
    var position_name = position_name_indeces[i][0]; 
    
    // get headers for particular position
    var rank_sheet = SpreadsheetApp.getActiveSheet();
    var rank_headers = rank_sheet.getRange(3, 
                                         position_name_indeces[i][1], 
                                         1, 
                                         position_name_indeces[i][2] - position_name_indeces[i][1] + 1).getValues()[0];
    var p = 0
      
      for (var j = 0; j < combined_data.length; ++j) {
        if (combined_data[j][5] == position_name || combined_data[j][6] == position_name || combined_data[j][7] == position_name) {
          rank_sheet.getRange(p + 4, position_name_indeces[i][1], 1, 1).setValue(p + 1);
          rank_sheet.getRange(p + 4, position_name_indeces[i][1] + 1, 1, 1).setValue(combined_data[j][0]);
          rank_sheet.getRange(p + 4, position_name_indeces[i][1] + 2, 1, 1).setValue(combined_data[j][1]); 
          rank_sheet.getRange(p + 4, position_name_indeces[i][1] + 3, 1, 1).setValue(combined_data[j][2]);
          rank_sheet.getRange(p + 4, position_name_indeces[i][1] + 4, 1, 1).setValue(combined_data[j][3]);
          rank_sheet.getRange(p + 4, position_name_indeces[i][1] + 5, 1, 1).setValue(combined_data[j][4]);
          
          ++p
  
        }
      }
    }

}
