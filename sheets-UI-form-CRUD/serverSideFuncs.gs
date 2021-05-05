function getDataForSearch() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName("Customers");
  return ws.getRange(2, 1, ws.getLastRow()-1, 3).getValues();
}

function deleteById(id){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName("Customers");
  const custIds = ws.getRange(2, 1, ws.getLastRow()-1, 1).getValues().map(r => r[0].toString().toLowerCase());
  const posIndex = custIds.indexOf(id.toString().toLowerCase());
  const rowNumber = posIndex === -1 ? 0 : posIndex + 2;
  ws.deleteRow(rowNumber);
}

function getCustomerById(id){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName("Customers");
  const custIds = ws.getRange(2, 1, ws.getLastRow()-1, 1).getValues().map(r => r[0].toString().toLowerCase());
  const posIndex = custIds.indexOf(id.toString().toLowerCase());
  const rowNumber = posIndex === -1 ? 0 : posIndex + 2;  
  const customerInfo = ws.getRange(rowNumber, 1, 1, 4).getValues()[0];
  //[[45,"Joe"]]
  return { custID: customerInfo[0], firstName: customerInfo[1], lastName: customerInfo[2], phone: customerInfo[3]  }
}


function editCustomerById(id,customerInfo){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName("Customers");
  const custIds = ws.getRange(2, 1, ws.getLastRow()-1, 1).getValues().map(r => r[0].toString().toLowerCase());
  const posIndex = custIds.indexOf(id.toString().toLowerCase());
  const rowNumber = posIndex === -1 ? 0 : posIndex + 2;
  ws.getRange(rowNumber, 2,1,3).setValues([[
                                            customerInfo.firstName,
                                            customerInfo.lastName,
                                            customerInfo.phone
                                          ]]);
  return true;
}

function addCustomer(customerInfo){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName("Customers");
  const uniqueIDs = ws.getRange(2, 1, ws.getLastRow()-1, 1).getValues();
  var maxNum = 0;
  uniqueIDs.forEach(r => {
     maxNum = r[0] > maxNum ? r[0] : maxNum
  });
  var newID = maxNum + 1;
  
  ws.appendRow([
                 newID,
                 customerInfo.firstName,
                 customerInfo.lastName,
                 customerInfo.phone
               ]);
  
}

