function findCustomerRows(customerNumber) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  let results = [];

  const filteredSheets = sheets.filter(sheet => !sheet.getName().includes("Search"));

  filteredSheets.forEach(sheet => {
    const data = sheet.getDataRange().getValues();

    const customerNumberIndex = 2;

    for (let i = 1; i < data.length; i++) {
      if (data[i][customerNumberIndex] == customerNumber) {
        //results.push(data[i].concat(sheet.getName(), i+1))
        
        const startIndex = 18; 

        const paddedRow = data[i].length < startIndex ?
          data[i].concat(Array(startIndex - data[i].length).fill('')) :
          data[i];

        results.push([
          ...paddedRow.slice(0, startIndex),
          sheet.getName(),
          i + 1, 
          ...paddedRow.slice(startIndex) 
        ]);
        

      }
    }
  });

  return results;
}

function exportResultsToNewSheet(results, customerNumber) {
  //const newSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let newSheet = ss.getSheetByName('Search Results: ' + customerNumber);
  if (newSheet) {
    newSheet.clear();
  } else {
    newSheet = ss.insertSheet();
    newSheet.setName('Search Results: ' + customerNumber);
  }

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


  // design
  newSheet.setRightToLeft(true);
  newSheet.getRange(1, 1, 1, newSheet.getLastColumn()).setBackground('#aeeaf5');
  newSheet.setTabColor('#1687cb');
  newSheet.setFrozenRows(1);

  const cellSourceSheet = newSheet.getRange(1, 19);
  cellSourceSheet.setValue("גליון");
  cellSourceSheet.setBackground("#58d4ea")

  const cellSourceRow = newSheet.getRange(1, 20);
  cellSourceRow.setValue("שורה");
  cellSourceRow.setBackground("#58d4ea")
  
  //----------------------------------------

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
    exportResultsToNewSheet(results, customerNumber);
  }
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Functions')
    .addItem('Search Customer', 'showCustomerSearchPrompt')
    .addToUi();
}

function showCustomerSearchPrompt() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.prompt('הכנס קוד לקוח', ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() == ui.Button.OK) {
    const customerNumber = response.getResponseText();
    showCustomerRows(customerNumber);
  }
}
