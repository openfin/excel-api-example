# OpenFin Excel API Logs

## From Version 4

### Client Logs

From version 4.0 onwards you can pass true to the init function and it will enable logging (you will see the excel api console.log entries). 

```javascript
await fin.desktop.ExcelService.init(true);
```

If you wish to add these logs to your own logging setup you can pass an object with the following implementation (the following is the Typescript interface to give you an idea):

```javascript
export interface ILog {
    name?: string;
    trace?: (message, ...args) => void;
    debug?: (message, ...args) => void;
    info?: (message, ...args) => void;
    warn?: (message, ...args) => void;
    error?: (message, error, ...args) => void;
    fatal?: (message, error, ...args) => void;
}
```
If you pass an object but do not provide a full implementation then we will use our default implementation that logs to the console for the functions that are not provided.

This is to help with debugging if you wish to verify something.

### Excel Provider Logs

The excel provider will log information to a log file. If you go to your OpenFin/apps directory you should find the Excel-Service-Manager app and this should have an app.log file within it's folder.

### Excel Helper/Plugin Logs

These produce logs to help with debugging and can be found by going to your OpenFin\shared\assets\excel-api-addin\logging folder.