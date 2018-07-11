"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
Object.defineProperty(exports, "__esModule", { value: true });
const RpcDispatcher_1 = require("./RpcDispatcher");
const ExcelWorkbook_1 = require("./ExcelWorkbook");
const ExcelWorksheet_1 = require("./ExcelWorksheet");
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
            var executor = RpcDispatcher_1.RpcDispatcher.promiseExecutors[result.messageId];
            delete RpcDispatcher_1.RpcDispatcher.promiseExecutors[result.messageId];
            if (result.error) {
                executor.reject(result.error);
                return;
            }
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
            executor.resolve(callbackData);
        };
        this.connectionUuid = connectionUuid;
        this.connected = false;
    }
    init() {
        return __awaiter(this, void 0, void 0, function* () {
            if (!this.initialized) {
                yield this.subscribeToExcelMessages();
                yield this.monitorDisconnect();
                this.initialized = true;
            }
            return;
        });
    }
    release() {
        return __awaiter(this, void 0, void 0, function* () {
            if (this.initialized) {
                yield this.unsubscribeToExcelMessages();
                //TODO: Provide external means to stop monitoring disconnect
            }
            return;
        });
    }
    subscribeToExcelMessages() {
        return Promise.all([
            new Promise(resolve => fin.desktop.InterApplicationBus.subscribe(this.connectionUuid, "excelEvent", this.processExcelEvent, resolve)),
            new Promise(resolve => fin.desktop.InterApplicationBus.subscribe(this.connectionUuid, "excelResult", this.processExcelResult, resolve))
        ]);
    }
    unsubscribeToExcelMessages() {
        return Promise.all([
            new Promise(resolve => fin.desktop.InterApplicationBus.unsubscribe(this.connectionUuid, "excelEvent", this.processExcelEvent, resolve)),
            new Promise(resolve => fin.desktop.InterApplicationBus.unsubscribe(this.connectionUuid, "excelResult", this.processExcelResult, resolve))
        ]);
    }
    monitorDisconnect() {
        return new Promise((resolve, reject) => {
            var excelApplicationConnection = fin.desktop.ExternalApplication.wrap(this.connectionUuid);
            var onDisconnect;
            excelApplicationConnection.addEventListener('disconnected', onDisconnect = () => {
                excelApplicationConnection.removeEventListener('disconnected', onDisconnect);
                this.connected = false;
                this.dispatchEvent('disconnected');
            }, resolve, reject);
        });
    }
    run(callback) {
        var runPromise = this.connected ? Promise.resolve() : new Promise(resolve => {
            var connectedCallback = () => {
                this.removeEventListener('connected', connectedCallback);
                resolve();
            };
            if (this.connectionUuid !== undefined) {
                this.addEventListener('connected', connectedCallback);
            }
            fin.desktop.System.launchExternalProcess({
                target: 'excel',
                uuid: this.connectionUuid
            });
        });
        return this.applyCallbackToPromise(runPromise, callback);
    }
    getWorkbooks(callback) {
        return this.invokeExcelCall("getWorkbooks", null, callback);
    }
    getWorkbookByName(name) {
        return this.workbooks[name];
    }
    addWorkbook(callback) {
        return this.invokeExcelCall("addWorkbook", null, callback);
    }
    openWorkbook(path, callback) {
        return this.invokeExcelCall("openWorkbook", { path: path }, callback);
    }
    getConnectionStatus(callback) {
        return this.applyCallbackToPromise(Promise.resolve(this.connected), callback);
    }
    getCalculationMode(callback) {
        return this.invokeExcelCall("getCalculationMode", null, callback);
    }
    calculateAll(callback) {
        return this.invokeExcelCall("calculateFull", null, callback);
    }
    toObject() {
        return this.objectInstance || (this.objectInstance = {
            connectionUuid: this.connectionUuid,
            addEventListener: this.addEventListener.bind(this),
            removeEventListener: this.removeEventListener.bind(this),
            addWorkbook: this.addWorkbook.bind(this),
            calculateAll: this.calculateAll.bind(this),
            getCalculationMode: this.getCalculationMode.bind(this),
            getConnectionStatus: this.getConnectionStatus.bind(this),
            getWorkbookByName: name => this.getWorkbookByName(name).toObject(),
            getWorkbooks: this.getWorkbooks.bind(this),
            openWorkbook: this.openWorkbook.bind(this),
            run: this.run.bind(this)
        });
    }
}
ExcelApplication.defaultInstance = undefined;
exports.ExcelApplication = ExcelApplication;
//# sourceMappingURL=ExcelApplication.js.map