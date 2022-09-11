// load custom menu
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Spending Categories')
      .addItem('Categorize All', 'categorizeAll')
      .addToUi();
}

// private: loads vendors from the specified sheet
// returns [vendors, error], where vendors is an array of {name, category, pattern} objects, and error is a string.
function loadVendors(vendorSheetName = "Vendors") {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const vendorSheet = ss.getSheets().find(s => s.getName() == vendorSheetName)
  if (!vendorSheet) {
    error = "Cannot find vendors sheet. Create a sheet named " + vendorSheetName + ' to use this function.'
    Logger.log(error)
    return [null, error];
  }
 
  const vendorData = vendorSheet.getDataRange().getValues()
  const header = vendorData[0]
  if(header[0] != 'Vendor' || header[1] != 'Category' || header[2] != 'Pattern') {
    error = "Vendors sheet does not have the expected format. Ensure the first row has 'Vendor', 'Category', and 'Pattern' in the first 3 columns."
    Logger.log(error)
    return [null, error];
  }

  const vendors = vendorData.slice(1)
    .filter(row => row[0] && row[1] && row[2])
    .map(row => {
      return {
        name: row[0],
        category: row[1],
        pattern: new RegExp(row[2], "i")
      }
    })

  return [vendors, null];
}

// private: categorize a transactions array given vendors and a config.
function categorizeTransactions(transactions, transactionConfig, vendors) {
  const header = transactions[0]
  const descIndex = header.indexOf(transactionConfig.descColName)
  const vendorIndex = header.indexOf(transactionConfig.vendorColName)
  const catIndex = header.indexOf(transactionConfig.catColName)

  if(descIndex < 0) {
    return 'Missing column titled: "' + transactionConfig.descColName + '" on transaction sheet.';
  }
  if(vendorIndex < 0) {
    return 'Missing column titled: "' + transactionConfig.vendorColName + '" on transaction sheet.';
  }
  if(catIndex < 0) {
    return 'Missing column titled: "' + transactionConfig.catColName + '" on transaction sheet.';
  }

  for(let r = 1; r < transactions.length; r++) {
    row = transactions[r];
    currVendorCell = row[vendorIndex]
    currCatCell = row[catIndex]
    desc = row[descIndex]
    if(!desc || currVendorCell || currCatCell) {
      continue;
    }
    
    vendor = vendors.find(v => v.pattern.test(desc))
    if(vendor) {
      row[vendorIndex] = vendor.name
      row[catIndex] = vendor.category
    }
  }
}

// Public: Categorize all values, unless category or vendor is already set.
function categorizeAll(descColName = 'Description', vendorColName = 'Vendor', catColName = 'Category', vendorSheetName = 'Vendors') {  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet =  ss.getActiveSheet();
  const dataRange = sheet.getDataRange();
  const transactions = dataRange.getValues();
  
  let [vendors, error] = loadVendors(vendorSheetName)
  if(error) {
    SpreadsheetApp.getUi().alert("Error loading vendors: " + error);
    return;
  }
  
  let transactionConfig = {
    descColName: descColName,
    vendorColName: vendorColName,
    catColName: catColName
  }
  error = categorizeTransactions(transactions, transactionConfig, vendors)
  if(error) {
    SpreadsheetApp.getUi().alert('Error categorizing transactions: ' + error);
  }
  
  dataRange.setValues(transactions)
}