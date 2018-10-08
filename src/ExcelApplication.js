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
const ExcelWorkbook_1 = require("./ExcelWorkbook");
const ExcelWorksheet_1 = require("./ExcelWorksheet");
const RpcDispatcher_1 = require("./RpcDispatcher");
/**
 * @description Wraps an external application
 */
const externalApplicationWrap = fin.desktop.ExternalApplication.wrap;
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
         * @description A key value pair container that holds name of the workbook as
         * key and the workbook object itself as the value
         */
        this.workbooks = {};
        this.connectionUuid = connectionUuid;
        this.mConnected = false;
        this.initialized = false;
        this.objectInstance = undefined;
    }
    /**
     * @public
     * @property
     * @description Flag to indicate whether excel is connected to openfin
     * @returns {boolean} Connected or not
     */
    get connected() {
        return this.mConnected;
    }
    /**
     * @public
     * @function init
     * @description Initialises the application
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
     * @function release
     * @description Release all connection from the excel application to the
     * openfin app
     * @returns {Promise<void>} A promise
     */
    release() {
        return __awaiter(this, void 0, void 0, function* () {
            if (this.initialized) {
                yield this.unsubscribeToExcelMessages();
                // TODO: Provide external means to stop monitoring disconnect
            }
            return;
        });
    }
    /**
     * @private
     * @function processExcelEvent
     * @description Process events coming from excel to be handled
     * by the openfin app
     * @param {Readonly<Partial<ExcelEventData>>} data The data being sent over from the excel app
     */
    processExcelEvent(data) {
        const eventType = data.event;
        let workbook = this.workbooks[data.workbookName];
        const worksheets = workbook && workbook.worksheets;
        let worksheet = worksheets && worksheets[data.sheetName];
        switch (eventType) {
            case 'connected':
                this.mConnected = true;
                this.dispatchEvent(eventType);
                break;
            case 'sheetChanged':
                if (worksheet) {
                    worksheet.dispatchEvent(eventType, { data });
                }
                break;
            case 'sheetRenamed':
                const oldWorksheetName = data.oldSheetName;
                const newWorksheetName = data.sheetName;
                worksheet =
                    (worksheets && worksheets[oldWorksheetName]);
                if (worksheet) {
                    delete worksheets[oldWorksheetName];
                    worksheet.name = newWorksheetName;
                    worksheets[worksheet.name] = worksheet;
                    workbook.dispatchEvent(eventType, { worksheet: worksheet.toObject(), oldWorksheetName });
                }
                break;
            case 'selectionChanged':
                if (worksheet) {
                    worksheet.dispatchEvent(eventType, { data });
                }
                break;
            case 'sheetActivated':
            case 'sheetDeactivated':
                if (worksheet) {
                    worksheet.dispatchEvent(eventType);
                }
                break;
            case 'workbookDeactivated':
            case 'workbookActivated':
                if (workbook) {
                    workbook.dispatchEvent(eventType);
                }
                break;
            case 'sheetAdded':
                const newWorksheet = worksheet || new ExcelWorksheet_1.ExcelWorksheet(data.sheetName, workbook);
                worksheets[newWorksheet.name] = newWorksheet;
                workbook.dispatchEvent(eventType, { worksheet: newWorksheet.toObject() });
                break;
            case 'sheetRemoved':
                delete workbook.worksheets[worksheet.name];
                worksheet.dispatchEvent(eventType);
                workbook.dispatchEvent(eventType, { worksheet: worksheet.toObject() });
                break;
            case 'workbookAdded':
            case 'workbookOpened':
                const newWorkbook = workbook || new ExcelWorkbook_1.ExcelWorkbook(this, data.workbookName);
                this.workbooks[newWorkbook.name] = newWorkbook;
                this.dispatchEvent(eventType, { workbook: newWorkbook.toObject() });
                break;
            case 'workbookClosed':
                delete this.workbooks[workbook.name];
                workbook.dispatchEvent(eventType);
                this.dispatchEvent(eventType, { workbook: workbook.toObject() });
                break;
            case 'workbookSaved':
                const oldWorkbookName = data.oldWorkbookName;
                const newWorkbookName = data.workbookName;
                workbook = this.workbooks[oldWorkbookName];
                if (workbook) {
                    delete this.workbooks[oldWorkbookName];
                    workbook.name = newWorkbookName;
                    this.workbooks[workbook.name] = workbook;
                    this.dispatchEvent(eventType, { workbook: workbook.toObject(), oldWorkbookName });
                }
                break;
            case 'rowDeleted':
                console.log(data);
                if (!worksheet) {
                    console.error('No worksheet could be found');
                    return;
                }
                worksheet.dispatchEvent(eventType, { data });
                break;
            case 'rowInserted':
                console.log(data);
                if (!worksheet) {
                    console.error('No worksheet could be found');
                    return;
                }
                worksheet.dispatchEvent(eventType, { data });
                break;
            case 'afterCalculation':
                break;
            default:
                this.dispatchEvent(eventType);
                break;
        }
    }
    /**
     * @private
     * @function processExcelResult
     * @description Process results coming from excel application
     * @param {Readonly<ExcelResultData>} result The result of the call being made in the excel application
     */
    processExcelResult(result) {
        let callbackData;
        const executor = RpcDispatcher_1.RpcDispatcher.promiseExecutors[result.messageId];
        delete RpcDispatcher_1.RpcDispatcher.promiseExecutors[result.messageId];
        if (result.error) {
            executor.reject(result.error);
            return;
        }
        const workbook = this.workbooks[result.target.workbookName];
        const worksheets = workbook && workbook.worksheets;
        const resultData = result.data;
        switch (result.action) {
            case 'getWorkbooks':
                const workbookNames = resultData;
                const oldworkbooks = this.workbooks;
                this.workbooks = {};
                workbookNames.forEach(workbookName => {
                    this.workbooks[workbookName] = oldworkbooks[workbookName] ||
                        new ExcelWorkbook_1.ExcelWorkbook(this, workbookName);
                });
                callbackData = workbookNames.map((workbookName) => this.workbooks[workbookName].toObject());
                break;
            case 'getWorksheets':
                const worksheetNames = resultData;
                const oldworksheets = worksheets;
                workbook.worksheets = {};
                worksheetNames.forEach((worksheetName) => {
                    worksheets[worksheetName] = oldworksheets[worksheetName] ||
                        new ExcelWorksheet_1.ExcelWorksheet(worksheetName, workbook);
                });
                callbackData = worksheetNames.map((worksheetName) => worksheets[worksheetName].toObject());
                break;
            case 'addWorkbook':
            case 'openWorkbook':
                const newWorkbookName = resultData;
                const newWorkbook = this.workbooks[newWorkbookName] ||
                    new ExcelWorkbook_1.ExcelWorkbook(this, newWorkbookName);
                this.workbooks[newWorkbook.name] = newWorkbook;
                callbackData = newWorkbook.toObject();
                break;
            case 'addSheet':
                const newWorksheetName = resultData;
                const newWorksheet = worksheets[newWorksheetName] ||
                    new ExcelWorksheet_1.ExcelWorksheet(newWorksheetName, workbook);
                worksheets[newWorksheet.name] = newWorksheet;
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
     * @function subscribeToExelMessages
     * @description Subscribes to messages from Excel
     * application
     * @returns {Promise<[void, void]>} A promise
     */
    subscribeToExcelMessages() {
        return Promise.all([
            new Promise(resolve => fin.desktop.InterApplicationBus.subscribe(this.connectionUuid, 'excelEvent', this.processExcelEvent.bind(this), resolve)),
            new Promise(resolve => fin.desktop.InterApplicationBus.subscribe(this.connectionUuid, 'excelResult', this.processExcelResult.bind(this), resolve))
        ]);
    }
    /**
     * @private
     * @function unsubscribeToExcelMessages
     * @description Unsubscribes from Excel application
     * @returns {Promise<[void, void]>} A promise
     */
    unsubscribeToExcelMessages() {
        return Promise.all([
            new Promise(resolve => fin.desktop.InterApplicationBus.unsubscribe(this.connectionUuid, 'excelEvent', this.processExcelEvent.bind(this), resolve)),
            new Promise(resolve => fin.desktop.InterApplicationBus.unsubscribe(this.connectionUuid, 'excelResult', this.processExcelResult.bind(this), resolve))
        ]);
    }
    /**
     * @private
     * @function monitorDisconnect
     * @description Monitors disconnection event when openfin
     * disconnects from excel
     * @returns {Promise<void>} A promise
     */
    monitorDisconnect() {
        return new Promise((resolve, reject) => {
            const excelApplicationConnection = externalApplicationWrap(this.connectionUuid);
            let onDisconnect;
            excelApplicationConnection.addEventListener('disconnected', onDisconnect = () => {
                excelApplicationConnection.removeEventListener('disconnected', onDisconnect);
                this.mConnected = false;
                this.dispatchEvent('disconnected');
            }, resolve, reject);
        });
    }
    /**
     * @public
     * @function run
     * @description Runs Excel application
     * @returns {Promise<void>} A promise
     */
    run() {
        return this.connected ? Promise.resolve() : new Promise(resolve => {
            const connectedCallback = () => {
                this.removeEventListener('connected', connectedCallback);
                resolve();
            };
            if (this.connectionUuid !== undefined) {
                this.addEventListener('connected', connectedCallback);
            }
            const launchOptions = {
                target: 'excel',
                uuid: this.connectionUuid
            };
            fin.desktop.System.launchExternalProcess(launchOptions);
        });
    }
    /**
     * @public
     * @function getWorkbooks
     * @description Gets the workbooks within the excel application
     * @returns {Promise<Workbooks>} A promise
     */
    getWorkbooks() {
        return this.invokeExcelCall('getWorkbooks', null);
    }
    /**
     * @public
     * @function getWorkbookByName
     * @description Gets the registered workbook with the specified
     * name
     * @param {string} name The name of the workbook
     */
    getWorkbookByName(name) {
        if (!this.workbooks[name]) {
            console.error(`No workbooks with the name ${name}`);
            return;
        }
        return this.workbooks[name].toObject();
    }
    /**
     * @public
     * @function addWorkbook
     * @description adds a workbook to the Excel application
     * @returns {Promise<Workbook>} A promise with a result
     */
    addWorkbook() {
        return this.invokeExcelCall('addWorkbook', null);
    }
    /**
     * @public
     * @function openWorkbook
     * @description Opens the workbook specified at the path
     * @param {string} path The path of the workbook
     * @returns {Promise<void>} Returns a promise with a result
     */
    openWorkbook(path) {
        return this.invokeExcelCall('openWorkbook', { path });
    }
    /**
     * @public
     * @function getConnectionStatus
     * @description Gets the connection status of of the Excel
     * application
     * @returns {Promise<boolean>} A promise with a result
     */
    getConnectionStatus() {
        return Promise.resolve(this.connected);
    }
    /**
     * @public
     * @function getCalculationMode
     * @description Gets the calculation mode from Excel
     * application
     * @returns {Promise<CalculationMode>} A promise with a result
     */
    getCalculationMode() {
        return this.invokeExcelCall('getCalculationMode', null);
    }
    /**
     * @public
     * @function calculateAll
     * @description Calculates all formulas on the workbook
     * @returns {Promise<void>} A promise with a result
     */
    calculateAll() {
        return this.invokeExcelCall('calculateFull', null);
    }
    /**
     * @public
     * @function toObject
     * @description Returns an object with only the methods and properties
     * to be exposed
     * @returns {Application} An object with only the methods and properties to be exposed
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
            getWorkbookByName: (name) => {
                return this.getWorkbookByName(name);
            },
            getWorkbooks: this.getWorkbooks.bind(this),
            openWorkbook: this.openWorkbook.bind(this),
            run: this.run.bind(this),
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