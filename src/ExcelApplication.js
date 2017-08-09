"use strict";
const RpcDispatcher_1 = require('./RpcDispatcher');
const ExcelWorkbook_1 = require('./ExcelWorkbook');
const ExcelWorksheet_1 = require('./ExcelWorksheet');
class Excel extends RpcDispatcher_1.RpcDispatcher {
    constructor() {
        super();
        this.workbooks = {};
        this.worksheets = {};
        this.processExcelEvent = (data, uuid) => {
            switch (data.event) {
                case "connected":
                    this.connected = true;
                    this.monitorDisconnect(uuid);
                    this.dispatchEvent({ type: data.event });
                    break;
                case "sheetChanged":
                    var sheets = this.worksheets[data.workbookName];
                    if (sheets && sheets[data.sheetName]) {
                        sheets[data.sheetName].dispatchEvent({ type: data.event, data: data });
                    }
                    break;
                case "sheetRenamed":
                    var sheets = this.worksheets[data.workbookName];
                    if (sheets && sheets[data.sheetName]) {
                        var sheet = sheets[data.sheetName];
                        sheets[data.sheetName] = null;
                        sheet.name = data.newName;
                        sheets[data.newName] = sheet;
                        sheet.dispatchEvent({ type: data.event, data: data });
                    }
                    break;
                case "selectionChanged":
                    var sheets = this.worksheets[data.workbookName];
                    if (sheets && sheets[data.sheetName]) {
                        sheets[data.sheetName].dispatchEvent({ type: data.event, data: data });
                    }
                    break;
                case "sheetActivated":
                    var sheets = this.worksheets[data.workbookName];
                    if (sheets && sheets[data.sheetName]) {
                        sheets[data.sheetName].dispatchEvent({ type: data.event });
                    }
                    break;
                case "sheetDeactivated":
                    var sheets = this.worksheets[data.workbookName];
                    if (sheets && sheets[data.sheetName]) {
                        sheets[data.sheetName].dispatchEvent({ type: data.event });
                    }
                    break;
                case "sheetAdded":
                    var workbook = this.getWorkbookByName(data.workbookName);
                    if (!this.worksheets[data.workbookName])
                        this.worksheets[data.workbookName] = {};
                    var sheets = this.worksheets[data.workbookName];
                    var sheet = sheets[data.sheetName] ? sheets[data.sheetName] : sheets[data.sheetName] = new ExcelWorksheet_1.ExcelWorksheet(data.sheetName, workbook);
                    workbook.dispatchEvent({ type: data.event, worksheet: sheet.toObject() });
                    break;
                case "sheetRemoved":
                    var workbook = this.getWorkbookByName(data.workbookName);
                    var sheet = this.worksheets[data.workbookName][data.sheetName];
                    delete this.worksheets[data.workbookName][data.sheetName];
                    workbook.dispatchEvent({ type: data.event, worksheet: sheet.toObject() });
                    break;
                case "workbookAdded":
                case "workbookOpened":
                    var workbook = new ExcelWorkbook_1.ExcelWorkbook(this, data.workbookName);
                    this.workbooks[data.workbookName] = workbook;
                    this.dispatchEvent({ type: data.event, workbook: workbook.toObject() });
                    break;
                case "afterCalculation":
                    this.dispatchEvent({ type: data.event });
                    break;
                case "workbookDeactivated":
                case "workbookActivated":
                    var workbook = this.getWorkbookByName(data.workbookName);
                    if (workbook)
                        workbook.dispatchEvent({ type: data.event });
                    break;
                case "workbookClosed":
                    var workbook = this.getWorkbookByName(data.workbookName);
                    delete this.workbooks[data.workbookName];
                    delete this.worksheets[data.workbookName];
                    workbook.dispatchEvent({ type: data.event });
                    this.dispatchEvent({ type: data.event, workbook: workbook.toObject() });
                    break;
                default:
                    this.dispatchEvent({ type: data.event });
                    break;
            }
        };
        this.processExcelResult = (data) => {
            var callbackData = {};
            switch (data.action) {
                case "getWorkbooks":
                    var workbookNames = data.data;
                    var _workbooks = [];
                    for (var i = 0; i < workbookNames.length; i++) {
                        var name = workbookNames[i];
                        if (!this.workbooks[name]) {
                            this.workbooks[name] = new ExcelWorkbook_1.ExcelWorkbook(this, name);
                        }
                        _workbooks.push(this.workbooks[name]);
                    }
                    callbackData = _workbooks.map(wb => wb.toObject());
                    break;
                case "getWorksheets":
                    var worksheetNames = data.data;
                    var _worksheets = [];
                    var worksheet = null;
                    for (var i = 0; i < worksheetNames.length; i++) {
                        if (!this.worksheets[data.workbook]) {
                            this.worksheets[data.workbook] = {};
                        }
                        worksheet = this.worksheets[data.workbook][worksheetNames[i]] ? this.worksheets[data.workbook][worksheetNames[i]] : this.worksheets[data.workbook][worksheetNames[i]] = new ExcelWorksheet_1.ExcelWorksheet(worksheetNames[i], this.workbooks[data.workbook]);
                        _worksheets.push(worksheet);
                    }
                    callbackData = _worksheets.map(ws => ws.toObject());
                    break;
                case "getCells":
                case "getCellsColumn":
                case "getCellsRow":
                    callbackData = data.data;
                    break;
                case "addWorkbook":
                case "openWorkbook":
                    if (!this.workbooks[data.workbookName]) {
                        var workbook = new ExcelWorkbook_1.ExcelWorkbook(this, data.workbook);
                        this.workbooks[data.workbook] = workbook;
                    }
                    else {
                        var workbook = this.workbooks[data.workbookName];
                    }
                    callbackData = workbook.toObject();
                case "addSheet":
                    if (!this.worksheets[data.workbookName])
                        this.worksheets[data.workbookName] = {};
                    var sheets = this.worksheets[data.workbookName];
                    var worksheet = sheets[data.sheetName] ? sheets[data.sheetName] : sheets[data.sheetName] = new ExcelWorksheet_1.ExcelWorksheet(data.sheetName, this.workbooks[data.workbookName]);
                    callbackData = worksheet.toObject();
                    break;
                case "getStatus":
                    callbackData = data.status;
                    break;
                case "getCalculationMode":
                case "getCellByName":
                    callbackData = data;
                    break;
            }
            if (RpcDispatcher_1.RpcDispatcher.callbacks[data.messageId]) {
                RpcDispatcher_1.RpcDispatcher.callbacks[data.messageId](callbackData);
                delete RpcDispatcher_1.RpcDispatcher.callbacks[data.messageId];
            }
        };
        this.processExcelServiceResult = (data) => {
            if (RpcDispatcher_1.RpcDispatcher.callbacks[data.messageId]) {
                RpcDispatcher_1.RpcDispatcher.callbacks[data.messageId](data.result);
                delete RpcDispatcher_1.RpcDispatcher.callbacks[data.messageId];
            }
        };
    }
    init() {
        fin.desktop.InterApplicationBus.subscribe("*", "excelEvent", this.processExcelEvent);
        fin.desktop.InterApplicationBus.subscribe("*", "excelResult", this.processExcelResult);
        fin.desktop.InterApplicationBus.subscribe("*", "excelServiceCallResult", this.processExcelServiceResult);
    }
    monitorDisconnect(uuid) {
        fin.desktop.ExternalApplication.wrap(uuid).addEventListener('disconnected', () => {
            this.connected = false;
            this.dispatchEvent({ type: 'disconnected' });
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
            fin.desktop.System.launchExternalProcess({ target: 'excel' });
        }
    }
    install(callback) {
        this.invokeServiceCall("install", null, callback);
    }
    getInstallationStatus(callback) {
        this.invokeServiceCall("getInstallationStatus", null, callback);
    }
    getWorkbooks(callback) {
        this.invokeExcelCall("getWorkbooks", null, callback);
    }
    getWorkbookByName(name) {
        return this.workbooks[name];
    }
    getWorksheetByName(workbookName, worksheetName) {
        if (this.worksheets[workbookName])
            return this.worksheets[workbookName][worksheetName] ? this.worksheets[workbookName][worksheetName] : null;
        return null;
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
exports.Excel = Excel;
Object.defineProperty(exports, "__esModule", { value: true });
exports.default = new Excel();
//# sourceMappingURL=ExcelApplication.js.map