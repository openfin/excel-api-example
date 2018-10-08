/******/ (function(modules) { // webpackBootstrap
/******/ 	// The module cache
/******/ 	var installedModules = {};
/******/
/******/ 	// The require function
/******/ 	function __webpack_require__(moduleId) {
/******/
/******/ 		// Check if module is in cache
/******/ 		if(installedModules[moduleId]) {
/******/ 			return installedModules[moduleId].exports;
/******/ 		}
/******/ 		// Create a new module (and put it into the cache)
/******/ 		var module = installedModules[moduleId] = {
/******/ 			i: moduleId,
/******/ 			l: false,
/******/ 			exports: {}
/******/ 		};
/******/
/******/ 		// Execute the module function
/******/ 		modules[moduleId].call(module.exports, module, module.exports, __webpack_require__);
/******/
/******/ 		// Flag the module as loaded
/******/ 		module.l = true;
/******/
/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}
/******/
/******/
/******/ 	// expose the modules object (__webpack_modules__)
/******/ 	__webpack_require__.m = modules;
/******/
/******/ 	// expose the module cache
/******/ 	__webpack_require__.c = installedModules;
/******/
/******/ 	// identity function for calling harmony imports with the correct context
/******/ 	__webpack_require__.i = function(value) { return value; };
/******/
/******/ 	// define getter function for harmony exports
/******/ 	__webpack_require__.d = function(exports, name, getter) {
/******/ 		if(!__webpack_require__.o(exports, name)) {
/******/ 			Object.defineProperty(exports, name, {
/******/ 				configurable: false,
/******/ 				enumerable: true,
/******/ 				get: getter
/******/ 			});
/******/ 		}
/******/ 	};
/******/
/******/ 	// getDefaultExport function for compatibility with non-harmony modules
/******/ 	__webpack_require__.n = function(module) {
/******/ 		var getter = module && module.__esModule ?
/******/ 			function getDefault() { return module['default']; } :
/******/ 			function getModuleExports() { return module; };
/******/ 		__webpack_require__.d(getter, 'a', getter);
/******/ 		return getter;
/******/ 	};
/******/
/******/ 	// Object.prototype.hasOwnProperty.call
/******/ 	__webpack_require__.o = function(object, property) { return Object.prototype.hasOwnProperty.call(object, property); };
/******/
/******/ 	// __webpack_public_path__
/******/ 	__webpack_require__.p = "";
/******/
/******/ 	// Load entry module and return exports
/******/ 	return __webpack_require__(__webpack_require__.s = 5);
/******/ })
/************************************************************************/
/******/ ([
/* 0 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

Object.defineProperty(exports, "__esModule", { value: true });
/**
 * @abstract
 * @class
 * @description Top level class that communicates with the Excel application
 */
class RpcDispatcher {
    constructor() {
        /**
         * @public
         * @description The connectionUuid of the excel application
         */
        this.connectionUuid = '';
        /**
         * @private
         * @description Holds event listeners
         */
        this.listeners = {};
    }
    /**
     * @public
     * @function addEventListener Adds event listener to listen to events coming
     * from Excel application
     * @param {string} type The type of the event to listen to
     * @param {EventListenerOrEventListenerObject} listener The method to execute when the event has been fired
     */
    addEventListener(type, listener) {
        if (this.hasEventListener(type, listener)) {
            return;
        }
        if (!this.listeners[type]) {
            this.listeners[type] = [];
        }
        this.listeners[type].push(listener);
    }
    /**
     * @public
     * @function removeEventListener
     * @description Removes the event from the local store
     * @param {string} type The type of the event to listen to
     * @param {EventListenerOrEventListenerObject} listener The method to execute when the event has been fired
     */
    removeEventListener(type, listener) {
        if (!this.hasEventListener(type, listener)) {
            return;
        }
        const callbacksOfType = this.listeners[type];
        callbacksOfType.splice(callbacksOfType.indexOf(listener), 1);
    }
    /**
     * @private
     * @function hasEventListener
     * @description Check whether an event listener has been
     * registered
     * @param {string} type The type of the event
     * @param {EventListenerOrEventListenerObject} listener The method to execute when the event has been fired
     * @returns {boolean} True or false depending on if the listener exists
     */
    hasEventListener(type, listener) {
        if (!this.listeners[type]) {
            return false;
        }
        if (!listener) {
            return true;
        }
        return (this.listeners[type].indexOf(listener) >= 0);
    }
    /**
     * @public
     * @function dispatchEvent
     * @description Sends event over to the correct entity e.g.
     * Workbook, worksheet
     * @param {string|Event} evtOrTypeArg Pass either an event or event type as a string
     * @param {T} data The data to be passed to the receiving entity
     * @returns {boolean} Whether or not the events default behaviour has been prevented
     */
    dispatchEvent(evtOrTypeArg, data) {
        let event;
        if (typeof evtOrTypeArg === 'string') {
            event =
                Object.assign({
                    target: this.toObject(),
                    type: evtOrTypeArg,
                    defaultPrevented: false,
                }, data);
        }
        else {
            event = evtOrTypeArg;
        }
        const callbacks = this.listeners[event.type] || [];
        callbacks.forEach((callback) => {
            callback(event);
        });
        return event.defaultPrevented;
    }
    /**
     * @private
     * @function getDefaultMessage
     * @description Get the default message when invoking a remote
     * call
     * @returns {object} Returns an empty object to be populated
     */
    getDefaultMessage() {
        return {};
    }
    /**
     * @protected
     * @function invokeExcelCall
     * @description Invokes a call in excel application via RPC
     * @param {string} functionName The name of the function to invoke
     * @param {RemoteData?} data Any data to be sent over as part of the invocation
     * @returns {Promise<T>} A Promise with generic data depending on which function calls it
     */
    invokeExcelCall(functionName, data) {
        return this.invokeRemoteCall('excelCall', functionName, data);
    }
    /**
     * @protected
     * @function invokeServiceCall Invokes a call in the excel service process via
     * RPC
     * @param {string} functionName The name of the function to invoke
     * @param {ExcelData|null} data Any data to be sent over as part of the invocation
     * @returns {Promise<T>} A Promise with generic data depending on which function calls it
     */
    invokeServiceCall(functionName, data) {
        return this.invokeRemoteCall('excelServiceCall', functionName, data);
    }
    /**
     * @private
     * @function invokeRemoteCall
     * @description Invokes a remote procedure call
     * @param {string} topic Topic to send on
     * @param {string} functionName The name of the function to invoke
     * @param {RemoteData?} data The data to be sent over as part of the invocation
     * @returns {Promise<T>} A Promise with generic data depending on which function calls it
     */
    invokeRemoteCall(topic, functionName, data) {
        const message = this.getDefaultMessage();
        const args = data || {};
        const invoker = this;
        const workbook = (invoker.workbook) ||
            (invoker) || null;
        const worksheet = (invoker) || null;
        Object.assign(message, {
            messageId: RpcDispatcher.messageId,
            target: {
                connectionUuid: this.connectionUuid,
                workbookName: workbook.name ? workbook.name : null,
                worksheetName: worksheet.name ? worksheet.name : null,
                rangeCode: args.rangeCode
            },
            action: functionName,
            data
        });
        let executor;
        const promise = new Promise((resolve, reject) => {
            executor = { resolve, reject };
        });
        const currentMessageId = RpcDispatcher.messageId;
        RpcDispatcher.messageId++;
        if (this.connectionUuid !== undefined) {
            fin.desktop.InterApplicationBus.send(this.connectionUuid, topic, message, () => {
                RpcDispatcher.promiseExecutors[currentMessageId] = executor;
            }, (nak) => {
                executor.reject(new Error(nak));
            });
        }
        else {
            // @ts-ignore
            executor.reject(new Error('The target UUID of the remote call is undefined.'));
        }
        return promise;
    }
}
/**
 * @protected
 * @static
 * @description The message id of the action being sent
 */
RpcDispatcher.messageId = 1;
/**
 * @protected
 * @static
 * @description Promises to be executed
 */
RpcDispatcher.promiseExecutors = {};
exports.RpcDispatcher = RpcDispatcher;
//# sourceMappingURL=RpcDispatcher.js.map

/***/ }),
/* 1 */
/***/ (function(module, exports, __webpack_require__) {

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
const ExcelApplication_1 = __webpack_require__(2);
const RpcDispatcher_1 = __webpack_require__(0);
/**
 * @description Gets the uuid of the current app
 */
const getUuid = fin.desktop.getUuid;
/**
 * @description Wraps an external application
 */
const externalApplicationWrap = fin.desktop.ExternalApplication.wrap;
/**
 * @constant {string} excelServiceUuid Uuid for the excel service
 */
const excelServiceUuid = '886834D1-4651-4872-996C-7B2578E953B9';
/**
 * @class
 * @description Class for interacting with the .NET ExcelService process
 */
class ExcelService extends RpcDispatcher_1.RpcDispatcher {
    /**
     * @constructor
     * @description Constructor for ExcelService
     */
    constructor() {
        super();
        this.connectionUuid = excelServiceUuid;
        this.mInitialized = false;
        this.mApplications = {};
        this.mDefaultApplicationUuid = '';
        this.defaultApplicationObj = null;
    }
    /**
     * @public
     * @async
     * @function init
     * @description Initialises the ExcelService
     * @returns {Promise<void>} A promise
     */
    init() {
        return __awaiter(this, void 0, void 0, function* () {
            if (!this.mInitialized) {
                yield this.subscribeToServiceMessages();
                yield this.monitorDisconnect();
                // await fin.desktop.Service.connect({ uuid: excelServiceUuid })
                yield this.registerWindowInstance();
                yield this.getExcelInstances();
                this.mInitialized = true;
            }
            return;
        });
    }
    /**
     * @private
     * @async
     * @function processExcelServiceEvent
     * @description Processes events coming from the Excel
     * application
     * @param {ExcelServiceEventData} data Payload passed from the Excel Service
     * @returns {Promise<void>} A promise
     */
    processExcelServiceEvent(data) {
        return __awaiter(this, void 0, void 0, function* () {
            console.log(data);
            const eventType = data.event;
            let eventData = null;
            switch (eventType) {
                case 'started':
                    break;
                case 'registrationRollCall':
                    if (this.mInitialized) {
                        this.registerWindowInstance();
                    }
                    break;
                case 'excelConnected':
                    yield this.processExcelConnectedEvent(data);
                    eventData = { connectionUuid: data.uuid };
                    break;
                case 'excelDisconnected':
                    yield this.processExcelDisconnectedEvent(data).catch(console.error);
                    eventData = { connectionUuid: data.uuid };
                    break;
                default:
                    console.error('Event type not supported: ' + eventType);
                    break;
            }
            this.dispatchEvent(eventType, eventData);
        });
    }
    /**
     * @private
     * @async
     * @function processExcelServiceResult
     * @description Processes results from excel service
     * @param {ExcelResultData} result The result from the service
     * @returns {Promise<void>} A promise
     */
    processExcelServiceResult(result) {
        return __awaiter(this, void 0, void 0, function* () {
            const executor = RpcDispatcher_1.RpcDispatcher.promiseExecutors[result.messageId];
            if (!executor) {
                console.warn('No executors matching the messageId: ' + result.messageId, result);
                return;
            }
            delete RpcDispatcher_1.RpcDispatcher.promiseExecutors[result.messageId];
            if (result.error) {
                executor.reject(result.error);
                return;
            }
            // Internal processing
            switch (result.action) {
                case 'getExcelInstances':
                    yield this.processGetExcelInstancesResult(result.data);
                    break;
                default:
                    break;
            }
            executor.resolve(result.data);
        });
    }
    /**
     * @private
     * @function subscribeToServiceMessages
     * @description function to subscribe to topics
     * @returns {Promise<[void, void]>} A list of promises
     */
    subscribeToServiceMessages() {
        return Promise.all([
            new Promise(resolve => fin.desktop.InterApplicationBus.subscribe(excelServiceUuid, 'excelServiceEvent', this.processExcelServiceEvent.bind(this), resolve)),
            new Promise(resolve => fin.desktop.InterApplicationBus.subscribe(excelServiceUuid, 'excelServiceCallResult', this.processExcelServiceResult.bind(this), resolve))
        ]);
    }
    /**
     * @private
     * @function monitorDisconnect
     * @description Subscribes to the disconnected event and
     * dispatches to the excel application
     * @returns {Promnise<void>} A promise
     */
    monitorDisconnect() {
        return new Promise((resolve, reject) => {
            const excelServiceConnection = externalApplicationWrap(excelServiceUuid);
            let onDisconnect;
            excelServiceConnection.addEventListener('disconnected', onDisconnect = () => {
                excelServiceConnection.removeEventListener('disconnected', onDisconnect);
                this.dispatchEvent('stopped');
            }, resolve, reject);
        });
    }
    /**
     * @private
     * @async
     * @function registerWindowInstance
     * @description This registers a new Excel instance to a
     * new workbook domain
     * @returns {Promise<void>} A promise
     */
    registerWindowInstance() {
        return __awaiter(this, void 0, void 0, function* () {
            return this.invokeServiceCall('registerOpenfinWindow', { domain: document.domain });
        });
    }
    /**
     * @private
     * @async
     * @function configureDefaultApplication
     * @description Configures the default application
     * when the application first starts
     * @returns {Promise<void>} A promise
     */
    configureDefaultApplication() {
        return __awaiter(this, void 0, void 0, function* () {
            const defaultAppObjUuid = this.defaultApplicationObj && this.defaultApplicationObj.connectionUuid;
            const defaultAppEntry = this.mApplications[defaultAppObjUuid];
            const defaultAppObjConnected = defaultAppEntry ? defaultAppEntry.connected : false;
            if (defaultAppObjConnected) {
                return;
            }
            const connectedAppUuid = Object.keys(this.mApplications)
                .find((appUuid) => this.mApplications[appUuid]
                .connected);
            if (connectedAppUuid) {
                delete this.mApplications[defaultAppObjUuid];
                this.defaultApplicationObj =
                    this.mApplications[connectedAppUuid].toObject();
                return;
            }
            if (defaultAppEntry === undefined) {
                const disconnectedAppUuid = getUuid();
                const disconnectedApp = new ExcelApplication_1.ExcelApplication(disconnectedAppUuid);
                yield disconnectedApp.init();
                this.mApplications[disconnectedAppUuid] = disconnectedApp;
                this.defaultApplicationObj = disconnectedApp.toObject();
            }
            return;
        });
    }
    // Internal Event Handlers
    /**
     * @private
     * @async
     * @function processExcelConnectedEvent
     * @description Process the connected event
     * @param {ExcelConnectionEventData} data payload that holds uuid of the connected application
     * @returns {Promise<void>} A promise
     */
    processExcelConnectedEvent(data) {
        return __awaiter(this, void 0, void 0, function* () {
            const applicationInstance = this.mApplications[data.uuid] ||
                new ExcelApplication_1.ExcelApplication(data.uuid);
            yield applicationInstance.init();
            this.mApplications[data.uuid] = applicationInstance;
            // Synthetically raise connected event
            applicationInstance.processExcelEvent({ event: 'connected' });
            yield this.configureDefaultApplication();
            return;
        });
    }
    /**
     * @public
     * @async
     * @function processExcelDisconnectedEvent
     * @description Processes event when excel is
     * disconnected
     * @param {ExcelConnectionEventData} data The data from excel
     * @returns {Promise<void>} A promise
     */
    processExcelDisconnectedEvent(data) {
        return __awaiter(this, void 0, void 0, function* () {
            const applicationInstance = this.mApplications[data.uuid];
            if (applicationInstance === undefined) {
                return;
            }
            delete this.mApplications[data.uuid];
            console.log('configuring default application in disconnect event');
            this.configureDefaultApplication()
                .then(applicationInstance.release)
                .catch(console.error);
            return;
        });
    }
    // Internal API Handlers
    /**
     * @private
     * @async
     * @function processGetExcelInstancesResult
     * @description Gets Excel instance
     * @param {string[]} connectionUuids THe connection Uuids the Excel service is holding
     * @returns {Promise<void>} A promise
     */
    processGetExcelInstancesResult(connectionUuids) {
        return __awaiter(this, void 0, void 0, function* () {
            const existingInstances = this.mApplications;
            this.mApplications = {};
            yield Promise.all(connectionUuids.map((connectionUuid) => __awaiter(this, void 0, void 0, function* () {
                const existingInstance = existingInstances[connectionUuid];
                if (existingInstance === undefined) {
                    // Assume that since the ExcelService is aware of the instance it,
                    // is connected and to simulate the the connection event
                    yield this.processExcelServiceEvent({ event: 'excelConnected', uuid: connectionUuid });
                }
                else {
                    this.mApplications[connectionUuid] = existingInstance;
                }
                return;
            })));
            yield this.configureDefaultApplication();
        });
    }
    // API Calls
    /**
     * @public
     * @function install
     * @description Get Excel instance
     * @returns {Promise<void>} A promise
     */
    install() {
        return this.invokeServiceCall('install', null);
    }
    /**
     * @public
     * @function getInstallationStatus
     * @description Checks the installation status
     * @returns {Promise<void>} A promise
     */
    getInstallationStatus() {
        return this.invokeServiceCall('getInstallationStatus', null);
    }
    /**
     * @public
     * @function getExcelInstances
     * @description Returns all the excel instances that are open
     * @returns {Promise<void>} A promsie
     */
    getExcelInstances() {
        return this.invokeServiceCall('getExcelInstances', null);
    }
    /**
     * @public
     * @function toObject
     * @description Creates an empty object
     * @returns {object} An empty object
     */
    toObject() {
        return {};
    }
}
exports.ExcelService = ExcelService;
//# sourceMappingURL=ExcelApi.js.map

/***/ }),
/* 2 */
/***/ (function(module, exports, __webpack_require__) {

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
const ExcelWorkbook_1 = __webpack_require__(3);
const ExcelWorksheet_1 = __webpack_require__(4);
const RpcDispatcher_1 = __webpack_require__(0);
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

/***/ }),
/* 3 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

Object.defineProperty(exports, "__esModule", { value: true });
const RpcDispatcher_1 = __webpack_require__(0);
/**
 * @class
 * @description Class that represents a workbook
 */
class ExcelWorkbook extends RpcDispatcher_1.RpcDispatcher {
    /**
     * @constructor
     * @description Constructor for the ExcelWorkbook class
     * @param {Application} application The Application this workbook belongs to
     * @param {string}name The name of the workbook
     */
    constructor(application, name) {
        super();
        this.connectionUuid = application.connectionUuid;
        this.application = application;
        this.mWorksheets = {};
        this.mWorkbookName = name;
        this.objectInstance = null;
    }
    /**
     * @private
     * @function getDefaultMessage
     * @description Gets the default message to be sent over the
     * wire
     * @returns {object} An object with the workbook name in as default
     */
    getDefaultMessage() {
        return { workbook: this.mWorkbookName };
    }
    /**
     * @public
     * @property
     * @description Worksheets tied to this workbook
     * @returns {Worksheets} The worksheets tied to this workbook
     */
    get worksheets() {
        return this.mWorksheets;
    }
    /**
     * @public
     * @property
     * @description Set the worksheets that are tied to this workbook
     */
    set worksheets(worksheets) {
        this.mWorksheets = worksheets;
    }
    /**
     * @public
     * @property
     * @description workbookName property
     * @returns {string} The name of the workbook
     */
    get name() {
        return this.mWorkbookName;
    }
    /**
     * @public
     * @property
     * @description Sets the workbook name
     * @param {string} name Set the name of the workbook
     */
    set name(name) {
        this.mWorkbookName = name;
    }
    /**
     * @public
     * @function getWorksheets
     * @description Gets the worksheets tied to this workbook
     * @returns {Promise<Worksheets>} A promise with worksheets as the result
     */
    getWorksheets() {
        return this.invokeExcelCall('getWorksheets', null);
    }
    /**
     * @public
     * @function getWorksheetByName
     * @description Gets the worksheet by name
     * @param {string} name The name of the worksheet
     * @returns {ExcelWorksheet} The excel worksheet with the specified name
     */
    getWorksheetByName(name) {
        const worksheet = this.worksheets[name];
        if (!worksheet) {
            console.error(`No worksheet found with the name: ${name}`);
            return;
        }
        return this.worksheets[name].toObject();
    }
    /**
     * @public
     * @function addWorksheet Adds a new worksheet to the workbook
     * @description Adds a new worksheet to the workbook
     * @returns {Promise<Worksheet>} A promise
     */
    addWorksheet() {
        return this.invokeExcelCall('addSheet', null);
    }
    /**
     * @public
     * @function activate
     * @description Activates the workbook
     * @returns {Promise<void>} A promise
     */
    activate() {
        return this.invokeExcelCall('activateWorkbook');
    }
    /**
     * @public
     * @function save
     * @description Save the current workbook
     * @returns {Promise<void>} A promise
     */
    save() {
        return this.invokeExcelCall('saveWorkbook');
    }
    /**
     * @public
     * @function close
     * @description Closes the workbook
     * @returns {Promise<void>} A promise
     */
    close() {
        return this.invokeExcelCall('closeWorkbook');
    }
    /**
     * @public
     * @function toObject
     * @description Returns only the methods exposed
     * @returns {Workbook} Returns only the methods exposed
     */
    toObject() {
        return this.objectInstance || (this.objectInstance = {
            addEventListener: this.addEventListener.bind(this),
            removeEventListener: this.removeEventListener.bind(this),
            name: this.name,
            activate: this.activate.bind(this),
            addWorksheet: this.addWorksheet.bind(this),
            close: this.close.bind(this),
            getWorksheetByName: (name) => {
                return this.getWorksheetByName(name);
            },
            getWorksheets: this.getWorksheets.bind(this),
            save: this.save.bind(this),
            toObject: this.toObject.bind(this)
        });
    }
}
exports.ExcelWorkbook = ExcelWorkbook;
//# sourceMappingURL=ExcelWorkbook.js.map

/***/ }),
/* 4 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

Object.defineProperty(exports, "__esModule", { value: true });
const RpcDispatcher_1 = __webpack_require__(0);
/**
 * @class
 * @description Class that represents a worksheet
 */
class ExcelWorksheet extends RpcDispatcher_1.RpcDispatcher {
    /**
     * @constructor
     * @description Constructor for the ExcelWorksheet class
     * @param {string} name The name of the worksheet
     * @param {Workbook} workbook The ExcelWorkbook this worksheet is tied to
     */
    constructor(name, workbook) {
        super();
        this.connectionUuid = workbook.connectionUuid;
        this.mWorkbook = workbook;
        this.mWorksheetName = name;
        this.objectInstance = null;
    }
    /**
     * @protected
     * @function getDefaultMessage
     * @description Returns the default message
     * @returns {object} Returns the default message
     */
    getDefaultMessage() {
        return { workbook: this.workbook.name, worksheet: this.mWorksheetName };
    }
    /**
     * @public
     * @property
     * @description Returns worksheet name
     * @returns {string} The name of the worksheet
     */
    get name() {
        return this.mWorksheetName;
    }
    /**
     * @public
     * @property
     * @description Returns worksheet name
     * @param {string} name The name of the worksheet
     */
    set name(name) {
        this.mWorksheetName = name;
    }
    /**
     * @public
     * @property
     * @description Returns the workbook that this worksheet is tied to
     * @returns {Workbook} Returns the workbook that this worksheet is tied to
     */
    get workbook() {
        return this.mWorkbook;
    }
    /**
     * @public
     * @function setCells
     * @description Sets the content for the cells
     * @param {(string|number)[][]} values values for the cell
     * @param {string} offset The cell address
     * @returns {Promise<void>} A promise
     */
    setCells(values, offset) {
        if (!offset) {
            offset = 'A1';
        }
        const payload = { offset, values };
        return this.invokeExcelCall('setCells', payload);
    }
    /**
     * @public
     * @function getCells
     * @description Gets cell values from the range specified
     * @param {string} start The start cell address
     * @param {number} offsetWidth The number of columns in the openfin app
     * @param {number} offsetHeight The number of rows in the openfin app
     * @returns {Promise<(string | number)[][]>} A promise containing the cells
     */
    getCells(start, offsetWidth, offsetHeight) {
        const payload = { start, offsetHeight, offsetWidth };
        return this.invokeExcelCall('getCells', payload);
    }
    /**
     * @public
     * @function activateRow
     * @description This mirrors the row selected in the openfin
     * application to Excel
     * @param {string} cellAddress THe address of the first cell of the row
     * @returns {Promise<void>} A promise
     */
    activateRow(cellAddress) {
        const payload = { address: cellAddress };
        return this.invokeExcelCall('activateRow', payload);
    }
    /**
     * @public
     * @function insertRow
     * @description This inserts a row just before the selected row
     * @param {number} rowNumber The address of the first cell in the row
     * @returns {Promise<void>} A promise
     */
    insertRow(rowNumber) {
        return this.invokeExcelCall('insertRow', { rowNumber });
    }
    /**
     * @public
     * @function deleteRow
     * @description This deletes the selected row
     * @param {number} rowNumber The address of the first cell in the row
     * @returns {Promise<any>} A promise
     */
    deleteRow(rowNumber) {
        return this.invokeExcelCall('deleteRow', { rowNumber });
    }
    /**
     * @public
     * @function
     * @description activate Activates the current worksheet
     * @returns {Promise<any>} A promise
     */
    activate() {
        return this.invokeExcelCall('activateSheet');
    }
    /**
     * @public
     * @function
     * @description activateCell Activates the selected cell
     * @param {string} cellAddress The address of the cell
     * @returns {Promise<void>} A promise
     */
    activateCell(cellAddress) {
        return this.invokeExcelCall('activateCell', { address: cellAddress });
    }
    // public addButton(name: string, caption: string, cellAddress: string):
    // Promise<void> {
    //    return this.invokeExcelCall("addButton", { address: cellAddress,
    //    buttonName: name, buttonCaption: caption });
    //}
    // public setFilter(start: string, offsetWidth: number, offsetHeight: number,
    // field: number, criteria1: string, op: string, criteria2: string,
    // visibleDropDown: string): Promise<void> {
    //    return this.invokeExcelCall("setFilter", {
    //        start,
    //        offsetWidth,
    //        offsetHeight,
    //        field,
    //        criteria1,
    //        op,
    //        criteria2,
    //        visibleDropDown
    //    });
    //}
    ///**
    // * @public
    // * @function formatRange Formats the range selected
    // * @param {string} rangeCode The selected range
    // * @param {any} format The formatting to be applied to the range
    // * @returns {Promise<void>} A promise
    // */
    // public formatRange(rangeCode: string, format: any): Promise<void> {
    //    return this.invokeExcelCall("formatRange", { rangeCode, format });
    //}
    /**
     * @public
     * @function clearRange
     * @description Clear the range of formatting and content
     * @param {string} rangeCode The range selected
     * @returns {Promise<void>} A promise
     */
    clearRange(rangeCode) {
        return this.invokeExcelCall('clearRange', { rangeCode });
    }
    /**
     * @public
     * @function clearRangeContents
     * @description Clears the contents in the specified range
     * @param {string} rangeCode The selected range
     * @returns {Promise<void>} A promise
     */
    clearRangeContents(rangeCode) {
        return this.invokeExcelCall('clearRangeContents', { rangeCode });
    }
    ///**
    // * @public
    // * @function clearRangeFormats Clears the formatting in the range specified
    // * @param rangeCode The selected range
    // */
    // public clearRangeFormats(rangeCode: string): Promise<void> {
    //    return this.invokeExcelCall("clearRangeFormats", { rangeCode });
    //}
    /**
     * @public
     * @function clearAllCells
     * @description Clear all cells and their formatting
     * @returns {Promise<any>} A promise
     */
    clearAllCells() {
        return this.invokeExcelCall('clearAllCells', null);
    }
    /**
     * @public
     * @function clearAllCellContents
     * @description Clears all the cells content
     * @returns {Promise<void>} A promise
     */
    clearAllCellContents() {
        return this.invokeExcelCall('clearAllCellContents', null);
    }
    ///**
    // * @public
    // * @function clearAllCellFormats Clear all formatting in every cell
    // * @returns {Promise<any>} A promise
    // */
    // public clearAllCellFormats(): Promise<void> {
    //    return this.invokeExcelCall("clearAllCellFormats", null);
    //}
    /**
     * @public
     * @function setCellName
     * @description Sets a name for the cell address
     * @param {string} cellAddress The address of the cell e.g. A1
     * @param {string} cellName The name of the cell
     * @returns {Promise<any>} A promise
     */
    setCellName(cellAddress, cellName) {
        return this.invokeExcelCall('setCellName', { address: cellAddress, cellName });
    }
    /**
     * @public
     * @function calculate
     * @description Calculates all formula on teh sheet
     * @returns {Promise<void>} A promise
     */
    calculate() {
        return this.invokeExcelCall('calculateSheet');
    }
    /**
     * @public
     * @function getCellByName
     * @description Gets a cell by its name
     * @param {string} cellName The name of the cell
     * @returns {Promise<any>} A promise
     */
    getCellByName(cellName) {
        return this.invokeExcelCall('getCellByName', { cellName });
    }
    /**
     * @public
     * @function protect
     * @description Password protects the sheet
     * @param {string} password Password used to protect the sheet
     * @returns {Promise<any>} A promise
     */
    protect(password) {
        return this.invokeExcelCall('protectSheet', { password });
    }
    /**
     * @public
     * @function
     * @description toObject Returns only the functions that should be exposed by
     * this class
     * @returns {Worksheet} Public methods in ExcelWorksheet
     */
    toObject() {
        return this.objectInstance || (this.objectInstance = {
            addEventListener: this.addEventListener.bind(this),
            removeEventListener: this.removeEventListener.bind(this),
            name: this.name,
            workbook: this.workbook,
            activate: this.activate.bind(this),
            activateCell: this.activateCell.bind(this),
            activateRow: this.activateRow.bind(this),
            calculate: this.calculate.bind(this),
            clearAllCellContents: this.clearAllCellContents.bind(this),
            clearAllCells: this.clearAllCells.bind(this),
            clearRange: this.clearRange.bind(this),
            clearRangeContents: this.clearRangeContents.bind(this),
            getCellByName: this.getCellByName.bind(this),
            getCells: this.getCells.bind(this),
            protect: this.protect.bind(this),
            setCellName: this.setCellName.bind(this),
            setCells: this.setCells.bind(this),
            insertRow: this.insertRow.bind(this),
            deleteRow: this.deleteRow.bind(this),
            toObject: this.toObject.bind(this)
        });
    }
}
exports.ExcelWorksheet = ExcelWorksheet;
//# sourceMappingURL=ExcelWorksheet.js.map

/***/ }),
/* 5 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

Object.defineProperty(exports, "__esModule", { value: true });
// This is the entry point of the Plugin script
const ExcelApi_1 = __webpack_require__(1);
const excelService = new ExcelApi_1.ExcelService();
// Attach ExcelService to the window
window.fin.desktop.ExcelService = excelService;
// Attach the Excel api to the window
Object.defineProperty(window.fin.desktop, 'Excel', {
    get() {
        return excelService.defaultApplicationObj;
    }
});
fin.desktop.main(() => {
    // For dev purposes
    fin.desktop.System.deleteCacheOnExit();
    function init(message) {
        console.log(message);
        excelService.init()
            .then(() => {
            fin.desktop.InterApplicationBus.unsubscribe('886834D1-4651-4872-996C-7B2578E953B9', 'init', init, () => {
                console.log('Successfully unsubscribed from initialisation');
            }, (reason) => {
                console.error(reason);
            });
        })
            .catch((err) => {
            console.log('This error might be ok', err);
        });
    }
    fin.desktop.InterApplicationBus.subscribe('886834D1-4651-4872-996C-7B2578E953B9', 'init', init);
    fin.desktop.InterApplicationBus.send('886834D1-4651-4872-996C-7B2578E953B9', 'init-multi-window', 'initial fire');
});
//# sourceMappingURL=plugin.js.map

/***/ })
/******/ ]);