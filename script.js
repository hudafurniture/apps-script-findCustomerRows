function findCustomerRows(customerNumber) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  let results = [];
  
  sheets.forEach(sheet => {
    const data = sheet.getDataRange().getValues();
    
    const customerNumberIndex = 2;
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][customerNumberIndex] == customerNumber) {
        results.push(data[i])
        //results.push([sheet.getName(), i + 1].concat(data[i]));
      }
    }
  });
  
  return results;
}

function exportResultsToNewSheet(results, customerNumber) {
  const newSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet();
  newSheet.setName('Search Results: ' + customerNumber );
  
  const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  const firstSheet = sheets[0]
  const numColumns = firstSheet.getLastColumn();
  const headerRow = numColumns > 0 ? firstSheet.getRange(1, 1, 1, numColumns).getValues()[0] : [];
  
  if (headerRow.length > 0) {
    newSheet.appendRow(headerRow);
  }
  
  results.forEach(result => {
    newSheet.appendRow(result);
  });

  newSheet.getRange(1, 1, newSheet.getLastRow(), newSheet.getLastColumn()).copyFormatToRange(firstSheet, 1, newSheet.getLastColumn(), 1, newSheet.getLastRow());
  
  newSheet.setRightToLeft(true);
  
  newSheet.getRange(1, 1, 1, newSheet.getLastColumn()).setBackground('orange');
  newSheet.setTabColor('orange'); 

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setActiveSheet(newSheet);
  ss.moveActiveSheet(ss.getNumSheets());

  SpreadsheetApp.getUi().alert('Search completed. Results exported to a new sheet.');
}

function showCustomerRows(customerNumber) {
  const results = findCustomerRows(customerNumber);
  
  if (results.length === 0) {
    Logger.log('No rows found for customer number: ' + customerNumber);
    SpreadsheetApp.getUi().alert('No rows found for customer number: ' + customerNumber);
  } else {
    exportResultsToNewSheet(results, customerNumber );
  }
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Customer Search')
    .addItem('Search Customer', 'showCustomerSearchPrompt')
    .addToUi();
}

function showCustomerSearchPrompt() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Enter Customer Number', ui.ButtonSet.OK_CANCEL);
  
  if (response.getSelectedButton() == ui.Button.OK) {
    const customerNumber = response.getResponseText();
    showCustomerRows(customerNumber);
  }
}
