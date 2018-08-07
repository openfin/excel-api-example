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
/**
 * @class Represents the Excel application itself
 */
class ExcelApplication extends RpcDispatcher_1.RpcDispatcher {
    /**
     * @constructor Constructor for the class
     * @param connectionUuid The connection uuid of the openfin application
     */
    constructor(connectionUuid) {
        super();
        /**
         * @private
         * @description A key value pair container that holds name of the workbook as key
         * and the workbook object itself as the value
         */
        this.workbooks = {};
        this.connectionUuid = connectionUuid;
        this.mConnected = false;
        this.initialized = false;
    }
    /**
     * @public
     * @property Flag to indicate whether excel is connected to openfin
     * @returns {boolean} Connected or not
     */
    get connected() {
        return this.mConnected;
    }
    /**
     * @public
     * @function init Initialises the application
     * @returns {Promise<void>} A promise
     */
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
    /**
     * @public
     * @function release Release all connection from the excel application to the openfin app
     * @returns {Promise<void>} A promise
     */
    release() {
        return __awaiter(this, void 0, void 0, function* () {
            if (this.initialized) {
                yield this.unsubscribeToExcelMessages();
                //TODO: Provide external means to stop monitoring disconnect
            }
            return;
        });
    }
    /**
     * @private
     * @function processExcelEvent Process events coming from excel to be handled by the openfin app
     * @param data The data being sent over from the excel app
     * @param uuid The uuid of the sender
     */
    processExcelEvent(data, uuid) {
        var eventType = data.event;
        var workbook = this.workbooks[data.workbookName];
        var worksheets = workbook && workbook.worksheets;
        var worksheet = worksheets && worksheets[data.sheetName];
        switch (eventType) {
            case "connected":
                this.mConnected = true;
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
            case "rowDeleted":
                console.log(data);
                if (!worksheet) {
                    console.error('No worksheet could be found');
                    return;
                }
                worksheet.dispatchEvent(eventType, { data: data });
                break;
            case "rowInserted":
                console.log(data);
                if (!worksheet) {
                    console.error('No worksheet could be found');
                    return;
                }
                worksheet.dispatchEvent(eventType, { data: data });
            case "afterCalculation":
            default:
                this.dispatchEvent(eventType);
                break;
        }
    }
    /**
     * @private
     * @function processExcelResult Process results coming from excel application
     * @param result The result of the call being made in the excel application
     */
    processExcelResult(result) {
        var callbackData = {};
        var executor = RpcDispatcher_1.RpcDispatcher.promiseExecutors[result.messageId];
        delete RpcDispatcher_1.RpcDispatcher.promiseExecutors[result.messageId];
        if (result.error) {
            executor.reject(result.error);
            return;
        }
        const workbook = this.workbooks[result.target.workbookName];
        const worksheets = workbook && workbook.worksheets;
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
    }
    /**
     * @private
     * @function subscribeToExelMessages Subscribes to messages from Excel application
     * @returns {Promise<[void, void]>} A promise
     */
    subscribeToExcelMessages() {
        return Promise.all([
            new Promise(resolve => fin.desktop.InterApplicationBus.subscribe(this.connectionUuid, "excelEvent", this.processExcelEvent.bind(this), resolve)),
            new Promise(resolve => fin.desktop.InterApplicationBus.subscribe(this.connectionUuid, "excelResult", this.processExcelResult.bind(this), resolve))
        ]);
    }
    /**
     * @private
     * @function unsubscribeToExcelMessages Unsubscribes from Excel application
     * @returns {Promise<[void, void]>} A promise
     */
    unsubscribeToExcelMessages() {
        return Promise.all([
            new Promise(resolve => fin.desktop.InterApplicationBus.unsubscribe(this.connectionUuid, "excelEvent", this.processExcelEvent.bind(this), resolve)),
            new Promise(resolve => fin.desktop.InterApplicationBus.unsubscribe(this.connectionUuid, "excelResult", this.processExcelResult.bind(this), resolve))
        ]);
    }
    /**
     * @private
     * @function monitorDisconnect Monitors disconnection event when openfin disconnects from excel
     * @returns {Promise<void>} A promise
     */
    monitorDisconnect() {
        return new Promise((resolve, reject) => {
            var excelApplicationConnection = fin.desktop.ExternalApplication.wrap(this.connectionUuid);
            var onDisconnect;
            excelApplicationConnection.addEventListener('disconnected', onDisconnect = () => {
                excelApplicationConnection.removeEventListener('disconnected', onDisconnect);
                this.mConnected = false;
                this.dispatchEvent('disconnected');
            }, resolve, reject);
        });
    }
    /**
     * @public
     * @function run Runs Excel application
     * @param callback The callback to be applied
     */
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
    /**
     * @public
     * @function getWorkbooks Gets the workbooks within the excel application
     * @returns {Promise<any>} A promise
     */
    getWorkbooks() {
        return this.invokeExcelCall("getWorkbooks", null);
    }
    /**
     * @public
     * @function getWorkbookByName Gets the registered workbook with the specified name
     * @param name The name of the workbook
     */
    getWorkbookByName(name) {
        return this.workbooks[name];
    }
    /**
     * @function addWorkbook adds a workbook to the Excel application
     * @returns {Promise<any>} A promise with a result
     */
    addWorkbook() {
        return this.invokeExcelCall("addWorkbook", null);
    }
    /**
     * @public
     * @function openWorkbook Opens the workbook specified at the path
     * @param path The path of the workbook
     * @returns {Promise<any>} Returns a promise with a result
     */
    openWorkbook(path) {
        return this.invokeExcelCall("openWorkbook", { path: path });
    }
    /**
     * @public
     * @function getConnectionStatus Gets the connection status of of the Excel application
     * @returns {Promise<any>} A promise with a result
     */
    getConnectionStatus(callback) {
        return this.applyCallbackToPromise(Promise.resolve(this.connected), callback);
    }
    /**
     * @public
     * @function getCalculationMode Gets the calculation mode from Excel application
     * @returns {Promise<any>} A promise with a result
     */
    getCalculationMode() {
        return this.invokeExcelCall("getCalculationMode", null);
    }
    /**
     * @public
     * @function calculateAll Calculates all formulas on the workbook
     * @returns {Promise<any>} A promise with a result
     */
    calculateAll() {
        return this.invokeExcelCall("calculateFull", null);
    }
    /**
     * @private
     * @function getInitialised Gets whether or not the ExcelApplicationServer has been initialised
     * @returns {Promise<any>} A promise with a result
     */
    getInitialised() {
        return this.invokeExcelCall("getInitialised", null);
    }
    /**
     * @public
     * @function toObject Returns an object with only the methods and properties to be exposed
     * @returns {any} An object with only the methods and properties to be exposed
     */
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
/**
 * @public
 * @static
 * @description The default excel application instance
 */
ExcelApplication.defaultInstance = undefined;
exports.ExcelApplication = ExcelApplication;
//# sourceMappingURL=ExcelApplication.js.map