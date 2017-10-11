"use strict";
const RpcDispatcher_1 = require('./RpcDispatcher');
const ExcelWorkbook_1 = require('./ExcelWorkbook');
const ExcelWorksheet_1 = require('./ExcelWorksheet');
class ExcelApplication extends RpcDispatcher_1.RpcDispatcher {
    constructor(connectionUuid) {
        super();
        this.workbooks = {};
        this.processExcelEvent = (data, uuid) => {
            var eventType = data.event;
            var workbook = this.workbooks[data.workbookName];
            var worksheets = workbook && workbook.worksheets;
            var worksheet = worksheets && worksheets[data.sheetName];
            switch (eventType) {
                case "connected":
                    this.connected = true;
                    this.dispatchEvent(eventType);
                    break;
                case "sheetChanged":
                    if (worksheet) {
                        worksheet.dispatchEvent(eventType, { data: data });
                    }
                    break;
                case "sheetRenamed":
                    var oldWorksheetName = data.oldSheetName;
                    var newWorksheetName = data.sheetName;
                    worksheet = worksheets && worksheets[oldWorksheetName];
                    if (worksheet) {
                        delete worksheets[oldWorksheetName];
                        worksheet.worksheetName = newWorksheetName;
                        worksheets[worksheet.worksheetName] = worksheet;
                        workbook.dispatchEvent(eventType, { worksheet: worksheet.toObject(), oldWorksheetName: oldWorksheetName });
                    }
                    break;
                case "selectionChanged":
                    if (worksheet) {
                        worksheet.dispatchEvent(eventType, { data: data });
                    }
                    break;
                case "sheetActivated":
                case "sheetDeactivated":
                    if (worksheet) {
                        worksheet.dispatchEvent(eventType);
                    }
                    break;
                case "workbookDeactivated":
                case "workbookActivated":
                    if (workbook) {
                        workbook.dispatchEvent(eventType);
                    }
                    break;
                case "sheetAdded":
                    var newWorksheet = worksheet || new ExcelWorksheet_1.ExcelWorksheet(data.sheetName, workbook);
                    worksheets[newWorksheet.worksheetName] = newWorksheet;
                    workbook.dispatchEvent(eventType, { worksheet: newWorksheet.toObject() });
                    break;
                case "sheetRemoved":
                    delete workbook.worksheets[worksheet.worksheetName];
                    worksheet.dispatchEvent(eventType);
                    workbook.dispatchEvent(eventType, { worksheet: worksheet.toObject() });
                    break;
                case "workbookAdded":
                case "workbookOpened":
                    var newWorkbook = workbook || new ExcelWorkbook_1.ExcelWorkbook(this, data.workbookName);
                    this.workbooks[newWorkbook.workbookName] = newWorkbook;
                    this.dispatchEvent(eventType, { workbook: newWorkbook.toObject() });
                    break;
                case "workbookClosed":
                    delete this.workbooks[workbook.workbookName];
                    workbook.dispatchEvent(eventType);
                    this.dispatchEvent(eventType, { workbook: workbook.toObject() });
                    break;
                case "workbookSaved":
                    var oldWorkbookName = data.oldWorkbookName;
                    var newWorkbookName = data.workbookName;
                    workbook = this.workbooks[oldWorkbookName];
                    if (workbook) {
                        delete this.workbooks[oldWorkbookName];
                        workbook.workbookName = newWorkbookName;
                        this.workbooks[workbook.workbookName] = workbook;
                        this.dispatchEvent(eventType, { workbook: workbook.toObject(), oldWorkbookName: oldWorkbookName });
                    }
                    break;
                case "afterCalculation":
                default:
                    this.dispatchEvent(eventType);
                    break;
            }
        };
        this.processExcelResult = (result) => {
            var callbackData = {};
            var workbook = this.workbooks[result.target.workbookName];
            var worksheets = workbook && workbook.worksheets;
            var worksheet = worksheets && worksheets[result.target.sheetName];
            var resultData = result.data;
            switch (result.action) {
                case "getWorkbooks":
                    var workbookNames = resultData;
                    var oldworkbooks = this.workbooks;
                    this.workbooks = {};
                    workbookNames.forEach(workbookName => {
                        this.workbooks[workbookName] = oldworkbooks[workbookName] || new ExcelWorkbook_1.ExcelWorkbook(this, workbookName);
                    });
                    callbackData = workbookNames.map(workbookName => this.workbooks[workbookName].toObject());
                    break;
                case "getWorksheets":
                    var worksheetNames = resultData;
                    var oldworksheets = worksheets;
                    workbook.worksheets = {};
                    worksheetNames.forEach(worksheetName => {
                        workbook.worksheets[worksheetName] = oldworksheets[worksheetName] || new ExcelWorksheet_1.ExcelWorksheet(worksheetName, workbook);
                    });
                    callbackData = worksheetNames.map(worksheetName => workbook.worksheets[worksheetName].toObject());
                    break;
                case "addWorkbook":
                case "openWorkbook":
                    var newWorkbookName = resultData;
                    var newWorkbook = this.workbooks[newWorkbookName] || new ExcelWorkbook_1.ExcelWorkbook(this, newWorkbookName);
                    this.workbooks[newWorkbook.workbookName] = newWorkbook;
                    callbackData = newWorkbook.toObject();
                    break;
                case "addSheet":
                    var newWorksheetName = resultData;
                    var newWorksheet = workbook[newWorkbookName] || new ExcelWorksheet_1.ExcelWorksheet(newWorksheetName, workbook);
                    worksheets[newWorksheet.worksheetName] = newWorksheet;
                    callbackData = newWorksheet.toObject();
                    break;
                default:
                    callbackData = resultData;
                    break;
            }
            if (RpcDispatcher_1.RpcDispatcher.callbacks[result.messageId]) {
                RpcDispatcher_1.RpcDispatcher.callbacks[result.messageId](callbackData);
                delete RpcDispatcher_1.RpcDispatcher.callbacks[result.messageId];
            }
        };
        this.connectionUuid = connectionUuid;
    }
    init() {
        if (!this.initialized) {
            fin.desktop.InterApplicationBus.subscribe("*", "excelEvent", this.processExcelEvent);
            fin.desktop.InterApplicationBus.subscribe("*", "excelResult", this.processExcelResult);
            this.monitorDisconnect();
            this.initialized = true;
        }
    }
    monitorDisconnect() {
        fin.desktop.ExternalApplication.wrap(this.connectionUuid).addEventListener('disconnected', () => {
            this.connected = false;
            this.dispatchEvent('disconnected');
        });
    }
    run(callback) {
        if (this.connected) {
            callback();
        }
        else {
            var connectedCallback = () => {
                this.removeEventListener('connected', connectedCallback);
                callback();
            };
            this.addEventListener('connected', connectedCallback);
            fin.desktop.System.launchExternalProcess({
                target: 'excel',
                uuid: this.connectionUuid
            });
        }
    }
    getWorkbooks(callback) {
        this.invokeExcelCall("getWorkbooks", null, callback);
    }
    getWorkbookByName(name) {
        return this.workbooks[name];
    }
    addWorkbook(callback) {
        this.invokeExcelCall("addWorkbook", null, callback);
    }
    openWorkbook(path, callback) {
        this.invokeExcelCall("openWorkbook", { path: path }, callback);
    }
    getConnectionStatus(callback) {
        callback(this.connected);
    }
    getCalculationMode(callback) {
        this.invokeExcelCall("getCalculationMode", null, callback);
    }
    calculateAll(callback) {
        this.invokeExcelCall("calculateFull", null, callback);
    }
    toObject() {
        return {
            connectionUuid: this.connectionUuid,
            addEventListener: this.addEventListener.bind(this),
            dispatchEvent: this.dispatchEvent.bind(this),
            removeEventListener: this.removeEventListener.bind(this),
            addWorkbook: this.addWorkbook.bind(this),
            calculateAll: this.calculateAll.bind(this),
            getCalculationMode: this.getCalculationMode.bind(this),
            getConnectionStatus: this.getConnectionStatus.bind(this),
            getWorkbookByName: name => this.getWorkbookByName(name).toObject(),
            getWorkbooks: this.getWorkbooks.bind(this),
            init: this.init.bind(this),
            openWorkbook: this.openWorkbook.bind(this)
        };
    }
}
exports.ExcelApplication = ExcelApplication;
//# sourceMappingURL=ExcelApplication.js.map