
// Activate sheet by Name
function activateSheetByName(sheetName) {
  var sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
  sheet.activate();
  return sheet;
}


// Test Firestore Connection & Authentication

function testFirestore() {
 var email = "** Paste email here **";
  var key = "** Paste private key here **";
  var projectId = "** Paste project id here **";
var firestore = FirestoreApp.getFirestore(email, key, projectId);
const data = {
"name": "test!"
}
firestore.createDocument("TestCollection/FirstDocument", data)
Logger.log(data);
}



// --------------- Write Employee Data to Firestore ---------------------------------

function writeEmployeeDataToFirebase() {
  var email = "** Paste email here **";
  var key = "** Paste private key here **";
  var projectId = "** Paste project id here **";
  var firestore = FirestoreApp.getFirestore(email, key, projectId);
  // var url = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  
  // Employees
  // EmployeeID, LastName, FirstName, Title, TitleOfCourtesy, BirthDate, HireDate, Address, City, Region, PostalCode, Country, HomePhone, Extension, Photo, Notes, ReportsTo

  var sheetName = 'employees';
  var url = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  var ss = SpreadsheetApp.openByUrl(url);
  var sheet = activateSheetByName(sheetName);
Logger.log(sheetName);
  
  var data = sheet.getDataRange().getValues();
  var dataToImport = {};
  
// Logger.log(data.length); // the number of rows in the sheet
  
  // Use this loop code if you want all rows in the sheet: 
  //   for(var i = 1; i < data.length; i++) {
  
for(var i = 1; i < 10; i++) {
 var EmployeeID = data[i][0];
 var LastName = data[i][1];
Logger.log(EmployeeID + '-' + LastName);
 dataToImport[EmployeeID + '-' + LastName] = {

   EmployeeID: data[i][0],
   LastName:data[i][1],
   FirstName:data[i][2],
   Title:data[i][3],
   TitleOfCourtesy:data[i][4],
   BirthDate:data[i][5],
   HireDate:data[i][6],
   Address:data[i][7],
   City:data[i][8],
   Region:data[i][9],
   PostalCode:data[i][10],
   Country:data[i][11],
   HomePhone:data[i][12],
   Extension:data[i][13],
   Photo:data[i][14],
   Notes:data[i][15],
   ReportsTo:data[i][6]
};

firestore.createDocument("Employees/", dataToImport[EmployeeID + '-' + LastName]);
// Logger.log(dataToImport);
}
}
