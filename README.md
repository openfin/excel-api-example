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

3) At this point you should be able to interact with Excel(create workbooks, worksheets, update cells etc) from either side and you should
see it mirrored on the other side

4) If you initially don't see workbooks on Openfin side, refresh the HTML page.

## Custom Functions:
The excel plugin also allows you to call custom JS functions from excel.
You can call a custom function by entering following in the formula bar:
=CustomFunction("nameOfTheJSFunction", "comma,separated,arguments")
the above function call will call a function in JavaScript app as following:  nameOfTheJSFunction("comma", "separated", "arguments");
There are two sample functions included in the demo app for demonstration.
To try the custom functions:
1) Enter some numbers in a column formation.

2) In any empty cell make the following call =CustomFunction("averageColumn", "startingCellAddress,columnHeight,resultDestination")
e.g  =CustomFunction("averageColumn", "A1,7,A8")
the above example will call a function defined in JavaScript app (averageColumn("A1", 7, "A8"))
and the defined JavaScript function will take the average of column values from A1 to A7 and update the result in cell A8.

3) The second function is =CustomFunction("averageRow", "A1,7,H1")
the above example will take the average of row values from A1 to G1 and update the cell at H1 with the result.

4) You can also define your own custom functions on the fly and call them from excel.
 To try your own custom function you can open the JS console from the demo app and define a test function by typing in the console.
e.g var test = function(){console.log("excel just called me with these arguments", arguments)}

5) Now go ahead and call your defined function from excel.
e.g =CustomFunction("test","a,b,c")
you will see following string printed out in the console. "excel just called me with these arguments ['a', 'b', 'c']"


# API Documentation

The Excel API is composition based object model. Where Excel is the top most level which has workbooks which have worksheets and worksheets have cells.
To use the API you will need to include ExcelAPI.js in your project and it will extend Openfin API with Excel API included.
Once included you will be able to use following API calls.


##fin.desktop.Excel:
**methods:**

``` javascript
/*
init();
this function is required to be executed before using the rest of the API.
*/
fin.desktop.init();

/*
getWorkbooks(callback);
Retrieves currently opened workbooks from excel and passes an array of workbook objects as an argument to the callback.
*/
fin.desktop.Excel.getWorkbooks(function(workbooks){...});

/*
addWorkbook();
creates a new workbook in Excel
*/
fin.desktop.Excel.addWorkbook();

/*
getWorkbookByName(name);
returns workbook object that represents the workbook with supplied name.
Note: to use this function, you need to call getWorkbooks at least once.
*/
var workbook = fin.desktop.Excel.getWorkbookByName("workbook1");

/*
getConnectionStatus(callback);
Passes true to the callback if its connected to Excel
*/
fin.desktop.Excel.getConnectionStatus(function(isConnected){...});

/*
addEventListener(type, listener);
Adds event handler to handle events from Excel
*/
fin.desktop.Excel.addEventListener("workbookAdded", function(event){...})

/*
addEventListener(type, listener);
removes an attached event handler from Excel
*/
removeEventListener("workbookAdded", handler);
```
**events:**
```javascript
"connected": is fired when excel connects to Openfin.
"workbookAdded": is fired when a new workbook is added in excel (this includes adding workbooks using API).
"workbookClosed": is fired when a workbook is closed.
```

##fin.desktop.ExcelWorkbook:
Represents an Excel workbook.
Note: New workbooks are not supposed to be created using new or Object.create().
Workbook objects can only be retrieved using API calls like fin.desktop.Excel.getWorkbooks() fin.desktop.Excel.getWorkbookByName() and fin.desktop.Excel.addWorkbook() etc.

**properties:**
```javascript
name: String // name of the workbook that the object represents.
```

**methods:**
```javascript
/*
getWorksheets(callback);
Passes an array of worksheet objects to the callback.
*/
var workbook = fin.desktop.Excel.getWorkbookByName("workbook1");
workbook.getWorksheets(function(worksheets){...});

/*
getWorksheetByName(name);
returns the worksheet object with the specified name.
Note: you have to at least use getWorksheets() once before using this function.
*/

var workbook = fin.desktop.Excel.getWorkbookByName("workbook1");
var sheet = workbook.getWorksheetByName("sheet1");


/*
addWorksheet(callback);
creates a new worksheet and passes the worksheet object to the callback
*/
var workbook = fin.desktop.Excel.getWorkbookByName("workbook1");
workbook.addWorksheet(function(sheet){...});

/*
activate();
activates or brings focus to the workbook
*/
var workbook = fin.desktop.Excel.getWorkbookByName("workbook1");
workbook.activate();


```
**events:**
```javascript
"sheetAdded": fired when a new sheet is added to the workbook
"sheetRemoved": fired when a sheet is closed/removed
"workbookActivated": fired when a workbook is activated/focused
"workbookDeactivated": fired when a workbook si deactivated/blurred
```

##fin.desktop.ExcelWorksheet:
Represents a worksheet in excel.
Note: New sheets are not supposed to be created using new or Object.create().
new sheets can be created only using workbook.addWorksheet() or existing sheet objects can be retrieved using workbook.getWorksheets() and workbook.getWorksheetByName();

**properties:**
```javascript
name: String // name of the worksheet
workbook: fin.desktop.ExcelWorkbook // workbook object that worksheet belongs to.
```
**methods:**
```javascript
/*
setCells(values, offset);
Populates the cells with the values that is a two dimensional array(array of rows) starting from the provided offset.
*/
workbook.addWorksheet(function(sheet){

   sheet.setCells([["a", "b", "c"], [1, 2, 3]], "A1");
});

/*
getCells(start, offsetWidth, offsetHeight, callback);
Passes a two dimensional array of objects that have following format {value: --, formula: --}
*/
var sheet = workbook.getSheetByName("sheet1");
sheet.getCells("A1", 3, 2, function(cells){ // cell: {value: --, formula: --}});

/*
activate();
activates or brings focus to the worksheet.
*/

var sheet = workbook.getSheetByName("sheet1");
sheet.activate();

/*
activateCell(cellAddress);
selects the given cell. cellAddress: (A1, A2 etc)
*/

var sheet = workbook.getSheetByName("sheet1");
sheet.activateCell("A1");

```
**events:**
```javascript
"sheetChanged": fired when any cell value in the sheet has changed.
"selectionChanged": fired when a selection on the sheet has changed.
"sheetActivated": fired when the sheet gets into focus.
"sheetDeactivated": fired when the sheet gets out of focus due to a different sheet getting in focus.
```

##Custom Functions:
Custom function allows you to call functions defined in your JavaScript app from Excel just like calling an Excel formula.

e.g =CustomFunction("nameOfTheJSFunction", "comma,separated,arguments")

the above will call a function in JavaScript app as following:  nameOfTheJSFunction("comma", "separated", "arguments");

**example:**
```javascript
\\in JavaScript
function averageColumn(start, height, resultDestination){
    ....
}
```
You could call the above function defined in your JavaScript app by entering following formula in Excel
=CustomFunction("averageColumn", "A1,7,A8")

