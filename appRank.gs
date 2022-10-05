function app_rank(){ 
  // IF YOU NEED HELP CHANGING THIS, CALL KELSEY DEWBRE AT (405)693-9924 OR HOPE ANDERSON AT (469)233-1891
 
  // this list contains all headers of the application scoring sheet matched with corresponding headers of the rank sheet (scoring sheet:rank sheet).
  // If headers change, this is where you will change the code!!!
  var header_pairs = [["First Name", "First Name"],
              ["Last Name", "Last Name"],         
              ["Average", "Average App Score"]];
  
  // this list contains the names of each position, the beginning column index (should be rank number), and the ending column index (should be average score)
  // if the positions names or placements change, this is where you change the code
  var position_name_indeces = [["Overall", 2, 5 ], 
                               ["Service Project Coordinator", 7, 10], 
                               ["Jobsite Coordinator", 12, 15], 
                               ["Operations Staff Coordinator", 17, 20], 
                               ["Supplies", 22, 25], 
                               ["High School Expansion Coordinator", 27, 30], 
                               ["International Expansion Coordinator", 32, 35], 
                               ["College Expansion Coordinator", 37, 40], 
                               ["Alumni Coordinator", 42, 45], 
                               ["Campus Engagement", 47, 50], 
                               ["Sponsorship", 52, 55], 
                               ["Social Media Specialist", 57,60], 
                               ["Graphic Designer", 62, 65], 
                               ["Publicity", 67, 70], 
                               ["Special Events", 72, 75], 
                               ["Photographer/Videographer", 77, 80], 
                               ["Database Coordinator", 82, 85]]; 
  
  // If the name of the sheet containing the scores changes, edit the following statement!
  var SCORE_SHEET_NAME = "Application";
  // If the index of the first row containing data on the score sheet changes, edit the following statement!
  var FIRST_DATA_ROW_INDEX = 4;
  // If the index of the first column containing data on the score sheet changes, edit the following statement!
  var FIRST_DATA_COL_INDEX = 3;
  // If the index of the row containing headers on the score sheet changes, edit the following statement!
  var SCORE_SHEET_HEADER_ROW_INDEX = 2;
  // If the index of the row containing headers on the rank sheet changes, edit the following statement!
  var RANK_SHEET_HEADER_ROW_INDEX = 3;
  // If the name of the column containing the rank number info on the rank sheet changes, edit the following statement!
  var RANK_NO_COL_NAME = "Rank";
  // if the name of the score sheet column containing the FIRST POSITION CHOICE changes, edit the following statement! 
  var FIRST_CHOICE_COL_NAME = "First Position"; 
  // if the name of the score sheet column containing the SECOND POSITION CHOICE changes, edit the following statement! 
  var SECOND_CHOICE_COL_NAME = "Second Position"; 
  // if the name of the score sheet column containing the THIRD POSITION CHOICE changes, edit the following statement! 
  var THIRD_CHOICE_COL_NAME = "Third Position"; 
  
  var final_row_index = parseInt(Browser.inputBox( "What is the number of the row containing the last applicant entry on the Application Scoring Sheet? ")); 
  
  // open scored application sheet and get its data and headers
  var score_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SCORE_SHEET_NAME);
  var score_data = score_sheet.getRange(FIRST_DATA_ROW_INDEX, 
                                        FIRST_DATA_COL_INDEX, 
                                        final_row_index - FIRST_DATA_ROW_INDEX + 1, 
                                        score_sheet.getLastColumn()).getValues(); 
  
  var score_headers = score_sheet.getRange(SCORE_SHEET_HEADER_ROW_INDEX, 
                                           FIRST_DATA_COL_INDEX, 
                                           1, 
                                           score_sheet.getLastColumn()).getValues()[0]; 
  
  // sort the data on "Average" from the scoring sheet
  // if the index of the "average" list in header_pairs changes, the indexOf argument will need to be changed
  var scored_col_to_rank_on = score_headers.indexOf(header_pairs[2][0])
  score_data.sort(function(a, b){return b[scored_col_to_rank_on] - a[scored_col_to_rank_on]});
  
  // POPULATE OVERALL RANK
  
  // get the rank sheet and its headers
  var rank_sheet = SpreadsheetApp.getActiveSheet();
  var rank_headers = rank_sheet.getRange(RANK_SHEET_HEADER_ROW_INDEX, 
                                         position_name_indeces[0][1], 
                                         1, 
                                         position_name_indeces[0][2] - position_name_indeces[0][1] + 1).getValues()[0];
  
  // column with rank number
  var rank_no_col = position_name_indeces[0][1]; 
  
  // go through each row and each cell, inserting data
  for (var i = 0; i < score_data.length; ++i) {
    // populate rank numbers
    rank_sheet.getRange(i + RANK_SHEET_HEADER_ROW_INDEX + 1, rank_no_col, 1, 1).setValue(i + 1); 
    
    // cycle through the header pairs and insert data
    for (var j = 0; j < header_pairs.length; ++j) {
      // get columns numbers
      var score_col = score_headers.indexOf(header_pairs[j][0]);  
      var rank_col = rank_headers.indexOf(header_pairs[j][1]) + position_name_indeces[0][1];
      // set the value of the cell
      rank_sheet.getRange(i + FIRST_DATA_ROW_INDEX, rank_col, 1, 1).setValue(score_data[i][score_col]);
      
    }
  }
      
 // POPULATE POSITIONS RANKS
  for (var f = 1; f < position_name_indeces.length; ++f) {
    
    // get position name
    var position_name = position_name_indeces[f][0]; 
    
    // get headers for particular position
    var rank_sheet = SpreadsheetApp.getActiveSheet();
    var rank_headers = rank_sheet.getRange(RANK_SHEET_HEADER_ROW_INDEX, 
                                         position_name_indeces[f][1], 
                                         1, 
                                         position_name_indeces[f][2] - position_name_indeces[f][1] + 1).getValues()[0];
    
    var p = 0
    
    // cycle through score data and record position choices for each applicant
    for (var k = 0; k < score_data.length; ++k) {
      var app_first_choice_index = score_headers.indexOf(FIRST_CHOICE_COL_NAME); 
      var app_first_choice_name = score_data[k][app_first_choice_index] ; 
      var app_second_choice_index = score_headers.indexOf(SECOND_CHOICE_COL_NAME); 
      var app_second_choice_name = score_data[k][app_second_choice_index]; 
      var app_third_choice_index = score_headers.indexOf(THIRD_CHOICE_COL_NAME);
      var app_third_choice_name = score_data[k][app_third_choice_index]; 
      
      // organize rankings by position
      if (app_first_choice_name == position_name || app_second_choice_name == position_name || app_third_choice_name == position_name) {
  
        // populate ranking
        rank_sheet.getRange(p + 1 + RANK_SHEET_HEADER_ROW_INDEX, position_name_indeces[f][1], 1, 1).setValue(p + 1); 
        
        // cycle through rank headers and each remaining columns
        for (var a = 0; a < header_pairs.length; ++a) {
          var score_col = score_headers.indexOf(header_pairs[a][0]); 
          rank_sheet.getRange(p + 1 + RANK_SHEET_HEADER_ROW_INDEX, position_name_indeces[f][1] + 1 + a, 1, 1).setValue(score_data[k][score_col]); 

        }
        
        ++p
        
      }
        
    }  
    }
      
      
      
 

}
