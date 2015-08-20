# excel-api-example
This repo demonstrates the usage of JavaScript Excel API provided by Openfin.

Note: This demo is intentionally coded in plain JavaScript so that its easy to follow,
without any need to understand other technologies/ frameworks.

# How to run the demo

1) Download and run the installer.
[openfin installer download](https://dl.openfin.co/services/download?fileName=excel-api-example-installer&config=http://openfin.github.io/excel-api-example/app.json)

2) Download the [add-in.zip](http://openfin.github.io/excel-api-example/add-in.zip)
extract the zip and load the FinDesktopAddin.xll (or FinDesktopAddin64.xll for 64bit Excel)
by double clicking it.
Once its loaded correctly you should see a message in status bar saying "Connected to Openfin", which means
the add-in is loaded and working correctly.

3) At this point you should be able to operate the Excel from either side (Demo App or Excel), and you should
see it mirrored on the other side

4) If you initially don't see workbooks on Openfin side, refresh the HTML page.

# API Documentation

##fin.desktop.Excel:
**methods:**
```javascript
init(); // this function is required to be executed before using the rest of the API
getWorkbooks(callback); // an array of workbooks will be passed as an argument to the callback
addWorkbook(); // creates a new workbook in Excel
getWorkbookByName(name); // to use this function, you need to call getWorkbooks at least once.
getConnectionStatus(callback); //  passes true to the callback if its connected to Excel
addEventListener(type, listener);
removeEventListener(type, listener);
```
**events:**
```javascript
"connected",  "workbookAdded", "workbookClosed" 
```

##fin.desktop.ExcelWorkbook:
**properties:**
```javascript
name: String // name of the workbook
```

**methods:**
```javascript
getWorksheets(callback); // passes an array of worksheet objects to the callback.
getWorksheetByName(name); //returns the worksheet object with the specified name.
addWorksheet(callback); // creates a new worksheet and passes the worksheet object to the callback
activate(); // activates or brings focus to the workbook
```
**events:**
```javascript
"sheetAdded", "sheetRemoved", "workbookActivated", "workbookDeactivated"
```

##fin.desktop.ExcelWorksheet:

**properties:**
```javascript
name: String // name of the worksheet
workbook: fin.desktop.ExcelWorkbook // workbook object that worksheet belongs to.
```
**methods:**
```javascript
setCells(values, offset);// populates the cells with the values that is two dimensional array starting from the provided offset.
getCells(start, offsetWidth, offsetHeight, callback); // passes a two dimensional array of cell values to the callback
activate(); // activates or brings focus to the worksheet.
activateCell(cellAddress); // selects the given cell. cellAddress: (A1, A2 etc)
```
**example:**
```javascript
sheet.getCells("A5", 5, 10, function(values){....}); the values are objects of following form {value: --, formula: --}
```
**events:** 
```javascript
"sheetChanged", "selectionChanged", "sheetActivated", "sheetDeactivated"
```
