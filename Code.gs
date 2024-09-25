// Function to create the custom menu
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  
  // Create a new menu called "Range Namer" with two sub-options
  ui.createMenu('Range Namer')
    .addItem('Auto Name Ranges', 'autoNameRanges')
    .addItem('Update Named Ranges', 'updateNamedRanges')
    .addToUi();
}

// Function to sanitize the header name to create a valid named range
function sanitizeName(name) {
  // Replace dashes with underscores
  name = name.replace(/-/g, '_');
  
  // Remove any text inside brackets
  name = name.replace(/\[.*?\]/g, '');
  
  // Handle text inside quotation marks and replace with underscores
  name = name.replace(/"([^"]*)"/g, function(match, p1) {
    return p1.replace(/\s+/g, '_') + '_';
  });

  // Replace any remaining special characters (non-alphanumeric and non-underscore) with nothing
  name = name.replace(/[^A-Za-z0-9_]/g, '');
  
  // Remove any leading or trailing underscores
  name = name.replace(/^_+|_+$/g, '');
  
  return name;
}

// Function to automatically name the selected ranges based on the first row
function autoNameRanges() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();
  var selection = sheet.getActiveRange();  // Get the currently selected range
  var firstRowValues = selection.getValues()[0];  // Get the first row from the selected range
  
  var numColumns = selection.getNumColumns();
  var numRows = selection.getNumRows();
  
  if (numRows <= 1) {
    SpreadsheetApp.getUi().alert('Please select a range with at least two rows (header and data).');
    return;
  }

  for (let i = 0; i < numColumns; i++) {
    let rawHeaderName = firstRowValues[i] ? firstRowValues[i].toString() : '';
    let headerName = sanitizeName(rawHeaderName);  // Sanitize the header name
    
    if (headerName) {  // Only proceed if there's a valid name left after sanitization
      // Define the range for the entire column excluding the header
      let columnRange = sheet.getRange(selection.getRow() + 1, selection.getColumn() + i, numRows - 1, 1);
      
      // Create the named range using the sanitized header name
      spreadsheet.setNamedRange(headerName, columnRange);
    }
  }
  
  SpreadsheetApp.getUi().alert('Named ranges created successfully.');
}

// Function to update the existing named ranges based on the selected range
function updateNamedRanges() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();
  var selection = sheet.getActiveRange();  // Get the currently selected range
  var firstRowValues = selection.getValues()[0];  // Get the first row from the selected range
  
  var numColumns = selection.getNumColumns();
  var numRows = selection.getNumRows();
  
  if (numRows <= 1) {
    SpreadsheetApp.getUi().alert('Please select a range with at least two rows (header and data).');
    return;
  }

  // Remove existing named ranges first
  var namedRanges = spreadsheet.getNamedRanges();
  for (let i = 0; i < namedRanges.length; i++) {
    namedRanges[i].remove();
  }
  
  // Create new named ranges for the updated selection
  for (let i = 0; i < numColumns; i++) {
    let rawHeaderName = firstRowValues[i] ? firstRowValues[i].toString() : '';
    let headerName = sanitizeName(rawHeaderName);  // Sanitize the header name
    
    if (headerName) {  // Only proceed if there's a valid name left after sanitization
      // Define the range for the entire column excluding the header
      let columnRange = sheet.getRange(selection.getRow() + 1, selection.getColumn() + i, numRows - 1, 1);
      
      // Create the named range using the sanitized header name
      spreadsheet.setNamedRange(headerName, columnRange);
    }
  }
  
  SpreadsheetApp.getUi().alert('Named ranges updated successfully.');
}
