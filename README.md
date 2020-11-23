# Excel Service API Demo

This repo provides a demonstration of the OpenFin Excel Service and its JavaScript API.

Note: This main source code for this demo is intentionally coded in plain JavaScript so that its easy to follow, without any need to understand other technologies/ frameworks. The API libraries are generated from TypeScript, and the end-product utilizes webpack to achieve a single-file web application.

This demo uses [ExcelDna](https://github.com/Excel-DNA/ExcelDna) to create the Excel addin.

# Running the Demo

## Quick Start

1) Download and run the installer:

[openfin installer download](https://install.openfin.co/download/?config=http%3A%2F%2Fopenfin.github.io%2Fexcel-api-example%2Fdemo%2Fapp.json&fileName=excel-api-example-installer)

2) After the installer runs, the OpenFin application should launch and either connect to Excel if it already running or present the option to launch Excel.

3) Once Excel is running you can create workbooks or open existing workbooks and observe two-way data synchronization between Excel and the demo app.

## Modifying and Building Locally

For development purposes you may wish to clone this repository and run on a local computer. The Excel Add-In is only compatible with Excel for Windows.

Pre-requisite: Node and NPM must be installed ( [https://nodejs.org/en/](https://nodejs.org/en/) ).

Clone the repository and, in the Command Prompt, navigate into the _excel-api-example_ directory created.

In the Command Prompt run:


```
> npm install
```

Once the Node packages have installed, it is now possible to make modifications to files in the _excel-api-example\src_ folder and rebuild the project by running:


```
> npm run webpack
```

After rebuilding, start the application by running:

```
> npm start
```

This will start a simple HTTP server on port 8080 and launch the OpenFin App automatically.


# Including the API in Your Own Project

## Manifest Declaration

Declare the Excel Service by including the following declaration in your application manifest:

```
"services":
[
   { "name": "excel" }
]
```

## Including the Client

Unlike other services, currently the Excel API client is only provided as a script tag. Include the following script tag on each page that requires API access:

```
<script src="https://openfin.github.io/excel-api-example/client/fin.desktop.Excel.js"></script>
```

## Waiting for the Excel Service to be Running

During startup, an application which wishes to utilize the Excel Service should ensure the service is running and ready to receive commands by invoking:

```
await fin.desktop.ExcelService.init();
```

It is advisable to place this call before any calls on the `fin.desktop.Excel` namespace.

# Getting Started with the API

## Writing to and Reading from a Spreadsheet:

After a connection has been established between Excel and the OpenFin application, pushing data to a spreadsheet and reading back the calculated values can be performed as follows:

```
var sheet1 = fin.desktop.Excel.getWorkbookByName('Book1').getWorksheetByName('Sheet1');

// A little fun with Pythagorean triples:
sheet1.setCells([
  ["A", "B", "C"],
  [  3,   4, "=SQRT(A2^2+B2^2)"],
  [  5,  12, "=SQRT(A3^2+B3^2)"],
  [  8,  15, "=SQRT(A4^2+B4^2)"],
], "A1");

// Write the computed values to console:
sheet1.getCells("C2", 0, 2, cells => {
  console.log(cells[0][0].value);
  console.log(cells[1][0].value);
  console.log(cells[2][0].value);
});

```

## Subscribing to Events:

Monitoring various application, workbook, and sheet events are done via the `addEventListener` functions on their respective objects. For example:

```
sheet1.getCells("C2", 0, 2, cells => {
  var lastValue = cells[0][0].value;
  
  fin.desktop.Excel.addEventListener('afterCalculation', () => {
    sheet1.getCells("C2", 0, 2, cells => {
      if(cells[0][0].value !== lastValue) {
        console.log('Value Changed!');
      }

      lastValue = cells[0][0].value;
    });
  });
})
```

## Legacy Callback-Based API

Version 1 of the Demo API utilized callbacks to handle asynchronous actions and return values. Starting in Version 2, all asynchronous functions are instead handled via Promises, however, for reverse compatibility the legacy callback-style calls are still supported.

All functions which return promises can also take a callback as the final argument. The following three calls are identical:

```
// Version 1 - Callback style [deprecated]
fin.desktop.Excel.getWorkbooks(workbooks => {
  console.log('Number of open workbooks: ', workbooks.length);
});

// Version 2 - Promise then callback
fin.desktop.Excel.getWorkbooks().then(workbooks => {
  console.log('Number of open workbooks: ', workbooks.length);
});

// Version 2 - Promise await
var workbooks = await fin.desktop.Excel.getWorkbooks();
console.log('Number of open workbooks: ', workbooks.length);

```

## Full API Documentation

The complete Excel Service API Documentation is available [here](DOCS.md).

In the future, type definition files will be available for public consumption via an NPM types package.

## License
MIT

The code in this repository is covered by the included license.

However, if you run this code, it may call on the OpenFin RVM or OpenFin Runtime, which are covered by OpenFin’s Developer, Community, and Enterprise licenses. You can learn more about OpenFin licensing at the links listed below or just email us at support@openfin.co with questions.

https://openfin.co/developer-agreement/ <br/>
https://openfin.co/licensing/

## Support
Please enter an issue in the repo for any questions or problems. Alternatively, please contact us at support@openfin.co 
