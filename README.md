# excel-api-example
This repo demonstrates the usage of JavaScript Excel API provided by Openfin.

Note: This demo is intentionally coded in plain JavaScript so that its easy to follow,
without any need to understand other technologies/ frameworks.

# How to run the demo

1) Download and run the installer.
[openfin installer download](https://dl.openfin.co/services/download?fileName=excel-api-example-installer&config=http://openfin.github.io/excel-api-example/app.json)

2) Download the [add-in.zip](http://openfin.github.io/excel-api-example/add-in.zip)
, extract the zip to disk (do not simply open it in place) and load the FinDesktopAddin.xll 
(or FinDesktopAddin64.xll for 64bit Excel) by double clicking it.
Once its loaded correctly you should see a message in status bar saying "Connected to Openfin", which means
the add-in is loaded and working correctly.

3) At this point you should be able to interact with Excel(create workbooks, worksheets, update cells etc) from either side and you should
see it mirrored on the other side

4) If you initially don't see workbooks on Openfin side, refresh the HTML page.

##How to run locally, for development
For development purposes you may wish to clone this repository and run on a local computer. OpenFin is currently only Windows compatible, but may be run on a Virtual Machine on Mac. 

Node must be installed ( [https://nodejs.org/en/](https://nodejs.org/en/) ).

Clone the repositry and, in the Command Prompt, navigate into the 'excel-api-example' directory created.

In the Command Prompt run:

```
$ npm install
```
Once the Node packages have installed, in the same 'excel-api-example' directory, run:

```
$ node server
```
This will start a simple Express server at http://localhost:5040.

Install the local app via this url:

[openfin excel-localhost installer download](https://dl.openfin.co/services/download?fileName=excel-api-example-installer&config=http://localhost:5040/app.json) 

The Node server must be running at http://localhost:5040 for you to be able to run the app locally.
   
## Custom Functions:
The Excel plugin also allows you to call custom JS functions from excel.
You can use them to do some complicated calculations using your connected app and feed the
result back to Excel, or you can simply use them to register cells that you would like to be
updated by your connected app.

You can call a custom function by entering following in the formula bar:

```
=CustomFunction("nameOfTheJSFunction", "comma,separated,arguments")
```
the above function call will call a function in JavaScript app as following:  

```
nameOfTheJSFunction("comma", "separated", "arguments");
```
There are two sample functions included in the demo app for demonstration.

To try the custom functions:

1) Enter some numbers in a column formation.

2) In any empty cell make the following call 
```
=CustomFunction("averageColumn", "startingCellAddress,columnHeight,resultDestination")
e.g  =CustomFunction("averageColumn", "A1,7,A8")
```

the above example will call a function defined in JavaScript app (averageColumn("A1", 7, "A8"))
and the defined JavaScript function will take the average of column values from A1 to A7 and update the result in cell A8.

3) The second function is =CustomFunction("averageRow", "A1,7,H1").

The above example will take the average of row values from A1 to G1 and update the cell at H1 with the result.

4) You can also define your own custom functions on the fly and call them from excel.
To try your own custom function you can open the JS console from the demo app and define a test function by typing in the console.
e.g 

```var test = function(){console.log("excel just called me with these arguments", arguments)}```

5) Now go ahead and call your defined function from excel.

```
e.g =CustomFunction("test","a,b,c")
```
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
addWorkbook(callback);
creates a new workbook in Excel
*/
fin.desktop.Excel.addWorkbook(function(workbook){...});

/*
openWorkbook(path, callback);
opens workbook from the specified path and passes the workbook object to the callback
*/
fin.desktop.Excel.openWorkbook(function(workbook){...});

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
getCalculationMode(callback);
Passes calculation mode information to the callback if its connected to Excel
*/
fin.desktop.Excel.getCalculationMode(function(info){...});

/*
calculateAll();
forces calculation on all sheets
*/
fin.desktop.Excel.calculateAll();

/*
addEventListener(type, listener);
Adds event handler to handle events from Excel
*/
fin.desktop.Excel.addEventListener("workbookAdded", function(event){...})

/*
removeEventListener(type, listener);
removes an attached event handler from Excel
*/
removeEventListener("workbookAdded", handler);
```
**events:**
```javascript

{type: "connected"};
// is fired when excel connects to Openfin.
//Example:
fin.desktop.Excel.addEventListener("connected", function(){ console.log("Connected to Excel"); })

{type: "workbookAdded", workbook: ExcelWorkbook};
//is fired when a new workbook is added in excel (this includes adding workbooks using API).
//Example:
fin.desktop.Excel.addEventListener("workbookAdded",
function(event){
    console.log("New workbook added; Name:", event.workbook.name);
});

{type: "workbookClosed", workbook: ExcelWorkbook};
//is fired when a workbook is closed.
//Example:
fin.desktop.Excel.addEventListener("workbookClosed",
function(event){
    console.log("Workbook closed; Name:", event.workbook.name);
});

{type: "afterCalculation"};
//is fired when calculation is complete on any sheet.
//Example:
fin.desktop.Excel.addEventListener("afterCalculation",
function(event){
    console.log("calculation is complete.";
});

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

/*
save();
Saves the changes to the workbook.
*/
var workbook = fin.desktop.Excel.getWorkbookByName("workbook1");
workbook.save();

/*
close();
Closes the workbook.
*/
var workbook = fin.desktop.Excel.getWorkbookByName("workbook1");
workbook.close();



```
**events:**
```javascript
{type: "sheetAdded", target: ExcelWorkbook, worksheet: ExcelWorksheet};
//fired when a new sheet is added to the workbook
//Example:
var workbook = fin.desktop.Excel.getWorkbookByName("workbook1");
workbook.addEventListener("sheetAdded",
function(event){
    console.log("New sheet", event.worksheet.name, "was added to the workbook", event.worksheet.workbook.name)
});

{type: "sheetRemoved", target: ExcelWorkbook, worksheet: ExcelWorksheet};
//fired when a sheet is closed/removed
//Example:
var workbook = fin.desktop.Excel.getWorkbookByName("workbook1");
workbook.addEventListener("sheetRemoved",
function(event){
    console.log("Sheet", event.worksheet.name, "was removed from workbook", event.worksheet.workbook.name)
});

{type: "workbookActivated", target: ExcelWorkbook};
//fired when a workbook is activated/focused
//Example:
var workbook = fin.desktop.Excel.getWorkbookByName("workbook1");
workbook.addEventListener("workbookActivated",
function(event){
    console.log("Workbook", event.target.name, "was activated");
});

{type: "workbookDeactivated", target: ExcelWorkbook};
//fired when a workbook is deactivated/blurred
//Example:
var workbook = fin.desktop.Excel.getWorkbookByName("workbook1");
workbook.addEventListener("workbookDeactivated",
function(event){
    console.log("Workbook", event.target.name, "was deactivated");
});

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
setFilter(start, offsetWidth, offsetHeight, field, criteria1[, operator, criteria2, visibleDropDown]);
sets filter on selected range in a worksheet.

arguments:
start: starting address of the range. e.g "A1"
offsetWidth: width of the range.
offsetHeight: height of the range.
field: integer representing the field number. starts with 1.
criteria1: The criteria (a string; for example, "101"). Use "=" to find blank fields, or use "<>" to find nonblank fields. If this argument is omitted, the criteria is All. If Operator is xlTop10Items, Criteria1 specifies the number of items (for example, "10").
operator: Optional. Can be one of the following:
          and
          bottom10items
          bottom10percent
          or
          top10items
          top10percent
criteria2: Optional. The second criteria (a string). Used with Criteria1 and Operator to construct compound criteria.
visibleDropDown: Optional. true to display the AutoFilter drop-down arrow for the filtered field; false to hide the AutoFilter drop-down arrow for the filtered field. true by default.

*/
var sheet = workbook.getSheetByName("sheet1");
sheet.setCells([["Column1", "Column2"], ["TRUE", "1"], ["TRUE", "2"],["FALSE", ""]], "A1");
sheet.setFilter("A1", 2, 4, 1, "TRUE");

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

/*
clearAllCells();
clears all the cell values and formatting in the worksheet.
*/

var sheet = workbook.getSheetByName("sheet1");
sheet.clearAllCells();

/*
clearAllCellContents();
clears all the cell values in the worksheet.
*/

var sheet = workbook.getSheetByName("sheet1");
sheet.clearAllCellContents();

/*
clearAllCellFormats();
clears all the cell formatting in the worksheet.
*/

var sheet = workbook.getSheetByName("sheet1");
sheet.clearAllCellFormats();

/*
clearRange();
clears all the cell values and formatting in the specified range.
*/

var sheet = workbook.getSheetByName("sheet1");
sheet.clearRange();

/*
clearRangeContents();
clears all the cell values in the specified range.
*/

var sheet = workbook.getSheetByName("sheet1");
sheet.clearRangeContents();

/*
clearRangeFormats();
clears all the cell formatting in the specified range.
*/

var sheet = workbook.getSheetByName("sheet1");
sheet.clearRangeFormats();

/*
setCellName(cellAddress, cellName);
sets a name for the cell which can be referenced to get values or in formulas
*/

var sheet = workbook.getSheetByName("sheet1");
sheet.setCellName("A1", "TheCellName");

/*
calculate();
forces calculation on the sheet.
*/

var sheet = workbook.getSheetByName("sheet1");
sheet.calculate();


/*
getCellByName(name, callback);
returns cell info of the cell with the name provided.
*/

var sheet = workbook.getSheetByName("sheet1");
sheet.getCellByName("TheCellName", function(cellInfo){...});

/*
protect();
makes all cells in the sheet read only, except the ones marked as locked:false
*/

var sheet = workbook.getSheetByName("sheet1");
sheet.protect();

/*

formatRange(rangeCode, format);
formats the specified range.
var sheet = workbook.getSheetByName("sheet1");
sheet.formatRange("A1:E:10", {
                                    border: {color:"0,0,0,1", style: "continuous"}, //dash, dashDot, dashDotDot, dot, double, none, slantDashDot
                                    border-right: {color:"0,0,0,1", style: "continuous"},
                                    border-left: {color:"0,0,0,1", style: "continuous"},
                                    border-top: {color:"0,0,0,1", style: "continuous"},
                                    border-bottom: {color:"0,0,0,1", style: "continuous"},
                                    horizontalLines: {color:"255,255,255,1", style: "none"}, // horizontal lines between cells
                                    verticalLines: {color:"255,255,255,1", style: "none"}, // vertical lines between cell rows
                                    font: {color: "100,100,100,1", size: 12, bold: true, italic: true, name: "Verdana"},
                                    mergeCells: true, // merges the given range into one big cell
                                    shrinkToFit: true, // the text will shrink to fit the cell
                                    locked: false // specifies if the cell is readonly or not in protect mode, default is true
                                });


*/


```
**events:**
```javascript
{type: "sheetChanged", target: ExcelWorksheet,  data: {column: int, row: int, formula: String, sheetName: String, value:String}};
//fired when any cell value in the sheet has changed.
//Example:
var sheet = workbook.getSheetByName("sheet1");
sheet.addEventListener("sheetChanged",
function(event){
    console.log("sheet values were changed. column:", event.data.column, "row:", event.data.row, "value:", event.data.value, "formula", event.data.formula);
});

{type: "selectionChanged", target: ExcelWorksheet, data: {column: int, row: int, value: String}};
//fired when a selection on the sheet has changed.
//Example:
var sheet = workbook.getSheetByName("sheet1");
sheet.addEventListener("selectionChanged",
function(event){
    console.log("sheet selection was changed. column:", event.data.column, "row:", event.data.row, "value:", event.data.value);
});

{type: "sheetActivated", target: ExcelWorksheet};
//fired when the sheet gets into focus.
//Example:
var sheet = workbook.getSheetByName("sheet1");
sheet.addEventListener("sheetActivated",
function(event){
    console.log("sheet activated. Sheet", event.target.name, "Workbook:", event.target.workbook.name);
});

{type: "sheetDeactivated", target: ExcelWorksheet};
//fired when the sheet gets out of focus due to a different sheet getting in focus.
//Example:
var sheet = workbook.getSheetByName("sheet1");
sheet.addEventListener("sheetDeactivated",
function(event){
    console.log("sheet deactivated. Sheet", event.target.name, "Workbook:", event.target.workbook.name);
});

{type: "sheetRenamed", target: ExcelWorksheet};
//fired when the sheet is renamed.
//Example:
var sheet = workbook.getSheetByName("sheet1");
sheet.addEventListener("sheetRenamed",
function(event){
    console.log("sheet", event.data.sheetName, "was renamed to: ", event.data.newName
    );
});

```

