// Custom function to create a custom menu in the Google Sheets UI.
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Menu')
    .addItem('Highlight Duplicates', 'colorDuplicates')
    .addItem('Delete Duplicates', 'deleteDuplicates')
    .addToUi();
}

// Function to highlight duplicate rows or delete them based on user choice.
function processDuplicates(isHighlight, columnName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var allSheets = ss.getSheets();

  var columnIndex = getColumnIndexByName(ss, columnName);

  if (columnIndex < 0) {
    throw new Error("Invalid column name or column not found in the header row of any subsheet.");
  }

  var duplicateValues = new Set();

  for (var i = 0; i < allSheets.length; i++) {
    var sheet = allSheets[i];
    var headerRow = getHeaderRow(sheet);

    if (headerRow.includes(columnName)) {
      var data = sheet.getDataRange().getValues();

      for (var j = 1; j < data.length; j++) {
        var value = data[j][columnIndex];

        if (value === "") {
          continue; // Skip empty cells.
        }

        if (duplicateValues.has(value)) {
          if (isHighlight) {
            sheet.getRange(j + 1, 1, 1, data[0].length).setBackground("yellow");
          } else {
            sheet.deleteRow(j + 1);
            // Since we delete the current row, we need to adjust the index and data length.
            j--;
            data = sheet.getDataRange().getValues();
          }
        } else {
          duplicateValues.add(value);
        }
      }
    }
  }
}

// Function to highlight duplicate rows.
function colorDuplicates() {
  var columnName = getColumnNameFromInput();
  processDuplicates(true, columnName);
}

// Function to delete duplicate rows.
function deleteDuplicates() {
  var columnName = getColumnNameFromInput();
  processDuplicates(false, columnName);
}

// Function to get the user-selected column name for comparison.
function getColumnNameFromInput() {
  var ui = SpreadsheetApp.getUi();
  var result = ui.prompt('Enter the column name for comparison:', ui.ButtonSet.OK_CANCEL);
  var button = result.getSelectedButton();
  var userInput = result.getResponseText().trim();

  if (button !== ui.Button.OK || userInput === "") {
    throw new Error("Invalid input or column name not provided.");
  }

  return userInput;
}

// Function to get the index of the column using the exact column name in the header row of any subsheet.
function getColumnIndexByName(ss, columnName) {
  var allSheets = ss.getSheets();

  for (var i = 0; i < allSheets.length; i++) {
    var sheet = allSheets[i];
    var headerRow = getHeaderRow(sheet);
    var columnIndex = headerRow.indexOf(columnName);

    if (columnIndex >= 0) {
      return columnIndex;
    }
  }

  return -1; // Column not found in the header row of any subsheet.
}

// Function to get the header row for a sheet.
function getHeaderRow(sheet) {
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
}
