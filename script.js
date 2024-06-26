function findCustomerRows(customerNumber) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  let results = [];

  const filteredSheets = sheets.filter(sheet => !sheet.getName().includes("Search")  && !sheet.getName().includes("מקורי"));

  filteredSheets.forEach(sheet => {
    const data = sheet.getDataRange().getValues();

    const customerNumberIndex = 1;

    for (let i = 1; i < data.length; i++) {
      if (data[i][customerNumberIndex] == customerNumber) {
        
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


function convertDateFormat(dateString) {
  var parts = dateString.split(/[\/-]/);
  var year = parts[2].length === 2 ? '20' + parts[2] : parts[2];
  var date = new Date(year, parts[1] - 1, parts[0]);
  var formattedDate = date.getFullYear() + '-' +
                      ('0' + (date.getMonth() + 1)).slice(-2) + '-' +
                      ('0' + date.getDate()).slice(-2);
  return formattedDate;
}



function findByProductionDate(date) {
  Logger.log("date");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  let results = [];

  const filteredSheets = sheets.filter(sheet => !sheet.getName().includes("Search")  && !sheet.getName().includes("מקורי"));
  
    filteredSheets.forEach(sheet => {
    const data = sheet.getDataRange().getValues();

    const prodDateIndex = 14;

    for (let i = 1; i < data.length; i++) {
      rowDate= new Date(data[i][prodDateIndex]);
      const rowDateAdjusted = new Date(rowDate.getTime() + (180 * 60000));
      targetDate = new Date(convertDateFormat(date));

      if (rowDateAdjusted.getTime() === targetDate.getTime()) {
        
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

function findBySupplyDate(date) {
  Logger.log("date");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  let results = [];

  const filteredSheets = sheets.filter(sheet => !sheet.getName().includes("Search")  && !sheet.getName().includes("מקורי"));
  
    filteredSheets.forEach(sheet => {
    const data = sheet.getDataRange().getValues();

    const supplyDateIndex = 15;

    for (let i = 1; i < data.length; i++) {
      rowDate= new Date(data[i][supplyDateIndex]);
      const rowDateAdjusted = new Date(rowDate.getTime() + (180 * 60000));
      targetDate = new Date(convertDateFormat(date));

      if (rowDateAdjusted.getTime() === targetDate.getTime()) {
        
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


function findByCoordinationDate(date) {
  Logger.log("date");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  let results = [];

  const filteredSheets = sheets.filter(sheet => !sheet.getName().includes("Search")  && !sheet.getName().includes("מקורי"));
  
    filteredSheets.forEach(sheet => {
    const data = sheet.getDataRange().getValues();

    const coordinationDateIndex = 16;

    for (let i = 1; i < data.length; i++) {
      rowDate= new Date(data[i][coordinationDateIndex]);
      const rowDateAdjusted = new Date(rowDate.getTime() + (180 * 60000));
      targetDate = new Date(convertDateFormat(date));

      if (rowDateAdjusted.getTime() === targetDate.getTime()) {
        
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


function findByUrgency() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  let results = [];

  const filteredSheets = sheets.filter(sheet => !sheet.getName().includes("Search") && !sheet.getName().includes("מקורי") );
  
    filteredSheets.forEach(sheet => {
    const data = sheet.getDataRange().getValues();

    const urgencyIndex = 11;

    for (let i = 1; i < data.length; i++) {
      //Logger.log(data[i][urgencyIndex])
      if (data[i][urgencyIndex] !== "") {
        
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

function findAddressRows(address) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  let results = [];

  const filteredSheets = sheets.filter(sheet => !sheet.getName().includes("Search")  && !sheet.getName().includes("מקורי"));

  filteredSheets.forEach(sheet => {
    const data = sheet.getDataRange().getValues();

    const addressIndex = 9;

    for (let i = 1; i < data.length; i++) {
      if (data[i][addressIndex] == address) {
        
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



function findDriverRows(driver) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  let results = [];

  const filteredSheets = sheets.filter(sheet => !sheet.getName().includes("Search")  && !sheet.getName().includes("מקורי"));

  filteredSheets.forEach(sheet => {
    const data = sheet.getDataRange().getValues();

    const driverIndex = 10;

    for (let i = 1; i < data.length; i++) {
      if (data[i][driverIndex] == driver) {
        
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


function findByOrderDate(date) {
  Logger.log("date");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  let results = [];

  const filteredSheets = sheets.filter(sheet => !sheet.getName().includes("Search")  && !sheet.getName().includes("מקורי"));
  
    filteredSheets.forEach(sheet => {
    const data = sheet.getDataRange().getValues();

    const orderDateIndex = 2;

    for (let i = 1; i < data.length; i++) {
      rowDate= new Date(data[i][orderDateIndex]);
      const rowDateAdjusted = new Date(rowDate.getTime() + (180 * 60000));
      targetDate = new Date(convertDateFormat(date));

      if (rowDateAdjusted.getTime() === targetDate.getTime()) {
        
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



//-------------------------------

function exportResultsToNewSheet(results, input) {
  //const newSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let newSheet = ss.getSheetByName('Search Results: ' + input);
  if (newSheet) {
    newSheet.clear();
  } else {
    newSheet = ss.insertSheet();
    newSheet.setName('Search Results: ' + input);
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

function showProdDateRows(date) {
  const results = findByProductionDate(date);

  if (!results || results.length === 0) {
    Logger.log('No rows found for date: ' + date);
    SpreadsheetApp.getUi().alert('No rows found for date: ' + date);
  } else {
    exportResultsToNewSheet(results, date);
  }
}

function showSupplyDateRows(date) {
  const results = findBySupplyDate(date);

  if (!results || results.length === 0) {
    Logger.log('No rows found for date: ' + date);
    SpreadsheetApp.getUi().alert('No rows found for date: ' + date);
  } else {
    exportResultsToNewSheet(results, date);
  }
}

function showCoordinationDateRows(date) {
  const results = findByCoordinationDate(date);

  if (!results || results.length === 0) {
    Logger.log('No rows found for date: ' + date);
    SpreadsheetApp.getUi().alert('No rows found for date: ' + date);
  } else {
    exportResultsToNewSheet(results, date);
  }
}

function showUrgencyRows() {
  const results = findByUrgency();

  if (!results || results.length === 0) {
    Logger.log('No rows found');
    SpreadsheetApp.getUi().alert('No rows found');
  } else {
    exportResultsToNewSheet(results, "Urgent");
  }
}

function showAdressRows(address) {
  const results = findAddressRows(address);

  if (results.length === 0) {
    Logger.log('No rows found for address: ' + address);
    SpreadsheetApp.getUi().alert('No rows found for address: ' + address);
  } else {
    exportResultsToNewSheet(results, address);
  }
}

function showDriverRows(driver) {
  const results = findDriverRows(driver);

  if (results.length === 0) {
    Logger.log('No rows found for driver: ' + driver);
    SpreadsheetApp.getUi().alert('No rows found for driver: ' + driver);
  } else {
    exportResultsToNewSheet(results, driver);
  }
}

function showOrderDateRows(date) {
  const results = findByOrderDate(date);

  if (!results || results.length === 0) {
    Logger.log('No rows found for date: ' + date);
    SpreadsheetApp.getUi().alert('No rows found for date: ' + date);
  } else {
    exportResultsToNewSheet(results, date);
  }
}



function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Functions')
    .addItem('Search Customer', 'showCustomerSearchPrompt')
    .addItem('Search By Prodcution Date', 'showProductionDatePrompt')
    .addItem('Search By Supply Date', 'showSupplyDatePrompt')
    .addItem('Search By Coordination Date', 'showCoordinationDatePrompt')
    .addItem('Search By Urgency', 'showUrgencyRows')
    .addItem('Search By Address', 'showAddressSearchPrompt')
    .addItem('Search By Driver', 'showDriverSearchPrompt')
    .addItem('Search By Order Date', 'showOrderDatePrompt')
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

function showProductionDatePrompt(){
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.prompt('הכנס תאריך כניסה ליצור', ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() == ui.Button.OK) {
    const searchDate = response.getResponseText().trim();
    showProdDateRows(searchDate);
  }

}

function showSupplyDatePrompt(){
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.prompt('הכנס תאריך משלוח למחסן', ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() == ui.Button.OK) {
    const searchDate = response.getResponseText().trim();
    showSupplyDateRows(searchDate);
  }

}

function showCoordinationDatePrompt(){
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.prompt('הכנס תאריך תיאום', ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() == ui.Button.OK) {
    const searchDate = response.getResponseText().trim();
    showCoordinationDateRows(searchDate);
  }

}


function showAddressSearchPrompt() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.prompt('הכנס עיר', ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() == ui.Button.OK) {
    const address = response.getResponseText();
    showAdressRows(address);
  }
}

function showDriverSearchPrompt() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.prompt('הכנס מרכיב', ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() == ui.Button.OK) {
    const driver = response.getResponseText();
    showDriverRows(driver);
  }
}


function showOrderDatePrompt(){
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.prompt('הכנס תאריך הפקה', ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() == ui.Button.OK) {
    const searchDate = response.getResponseText().trim();
    showOrderDateRows(searchDate);
  }

}





