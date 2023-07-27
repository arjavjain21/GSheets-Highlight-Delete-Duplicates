# Google Apps Script - Highlight/Delete Duplicates
This repository contains a Google Apps Script (GAS) that enhances the functionality of Google Sheets by providing an easy-to-use custom menu for highlighting or deleting duplicate rows based on a user-selected column name. Additionally, it includes an HTML file that serves as an informative user interface (UI) for the add-on.

**Features:**
1. **Custom Menu:** The script creates a custom menu named "Custom Menu" in the Google Sheets UI. This menu includes two options: "Highlight Duplicates" and "Delete Duplicates."
2. **Highlight Duplicates:** When the user selects this option, a prompt appears, asking the user to enter the column name for comparison. The script then identifies duplicate rows based on the provided column name and highlights them with a yellow background.
3. **Delete Duplicates:** This option also prompts the user to enter the column name for comparison. The script identifies duplicate rows based on the given column name and deletes them from the sheet. When deleting duplicates, the script adjusts the row index accordingly to avoid skipping rows during iteration.

**Functioning:**
- The `onOpen()` function creates a custom menu and adds the "Highlight Duplicates" and "Delete Duplicates" options to it.
- The `colorDuplicates()` and `deleteDuplicates()` functions are called when the user selects the respective options from the custom menu.
- The `getColumnNameFromInput()` function prompts the user to input a column name for comparison.
- The `processDuplicates()` function is the core function responsible for handling both highlighting and deleting duplicates. It iterates through all sheets in the active spreadsheet, finds the user-selected column using the `getColumnIndexByName()` function, and then identifies and processes the duplicates based on the provided choice.
- The `getColumnIndexByName()` function gets the index of the column using the exact column name from the header row of any subsheet.
- The `getHeaderRow()` function retrieves the header row of a sheet.

**HTML File:**
The `index.html` file is an HTML user interface that provides information about the Google Apps Script add-on. It serves as a landing page when the add-on is accessed from the Google Workspace Marketplace or Google Sheets add-ons store. It contains a title and a short description of the add-on's functionality.

**Usage:**
1. Install the Google Apps Script in your Google Sheets by copying and pasting the provided script into the Google Apps Script editor.
2. Save the script and reload your Google Sheets to see the "Custom Menu" added to the UI.
3. Click on "Custom Menu" and select either "Highlight Duplicates" or "Delete Duplicates."
4. When prompted, enter the column name for comparison, and the script will perform the selected action accordingly.

**Note:**
- Ensure that the script has access to the necessary Google Sheets API permissions.
- Carefully review the duplicates before selecting the "Delete Duplicates" option, as deleted data cannot be easily restored.

**License:**
This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

**GitHub Repository:**
For more information and access to the latest version of the Google Apps Script, you can visit the original GitHub repository at [https://github.com/YourUsername/YourRepository](https://github.com/YourUsername/YourRepository).
