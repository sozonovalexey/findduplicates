/**
 * Copyright 2014 Google Inc. All Rights Reserved.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

/**
 * Adds a custom menu with items to show the sidebar and dialog.
 *
 * @param {Object} e The event parameter for a simple onOpen trigger.
 */
function onOpen(e) {
  SpreadsheetApp.getUi()
      .createAddonMenu()
      .addItem('Find', 'findDuplicates')
      .addToUi();
}

/**
 * Runs when the add-on is installed; calls onOpen() to ensure menu creation and
 * any other initializion work is done immediately.
 *
 * @param {Object} e The event parameter for a simple onInstall trigger.
 */
function onInstall(e) {
  onOpen(e);
}

function getSheets() {
  var list = [];
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var spreadSheetsInA = ss.getSheets();

  list = spreadSheetsInA.map(function(sheet) {
    return sheet.getName();
  });
  
  return list;
}


// This function gets the full column Range like doing 'A1:A9999' in excel
// @param {String} column The column name to get ('A', 'G', etc)
// @param {Number} startIndex The row number to start from (1, 5, 15)
// @return {Range} The "Range" object containing the full column: https://developers.google.com/apps-script/class_range
function getFullColumn(spreadSheetName){
  var column = 'A';
  var startIndex = 1;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(spreadSheetName);
  var lastRow = sheet.getLastRow();
  return sheet.getRange(column+startIndex+':'+column+lastRow).getValues();
}


function findDuplicates() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var currentSheet = ss.getActiveSheet();
  var currentSheetName = currentSheet.getName();
  var currentColumn = getFullColumn(currentSheetName);
  
  var sheets = getSheets();
  for (i in sheets) {
    
    var sheetName = sheets[i];
    if (sheetName === currentSheetName) {
      break;
    }
    
    var column = getFullColumn(sheetName);
    for (k in currentColumn) {
      var currentRow = currentColumn[k];
      var searchPhrase = currentRow[0];

      for (j in column) {
        var row = column[j];
        var value = row[0];
        
        if (value === searchPhrase) {   
          var cellIndex = parseInt(k)+1;
          var a1Notation = 'A' + cellIndex;
          currentSheet.getRange(a1Notation).setBackground("yellow");
        }
        
      }
      
    }
  }
}

