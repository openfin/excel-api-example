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
const ExcelApplication_1 = require("./ExcelApplication");
/**
 * @constant {string} excelServiceUuid Uuid for the excel service
 */
const excelServiceUuid = "886834D1-4651-4872-996C-7B2578E953B9";
/**
 * @class Class for interacting with the .NET ExcelService process
 */
class ExcelService extends RpcDispatcher_1.RpcDispatcher {
    /**
     * @constructor Constructor for ExcelService
     */
    constructor() {
        super();
        this.connectionUuid = excelServiceUuid;
        this.mInitialized = false;
        this.mApplications = {};
        this.mDefaultApplicationUuid = undefined;
        this.defaultApplicationObj = undefined;
        this.getInitialized();
    }
    /**
     * @public
     * @function init Initialises the ExcelService
     * @returns {Promise<void>} A promise
     */
    init() {
        return __awaiter(this, void 0, void 0, function* () {
            if (!this.mInitialized) {
                yield this.subscribeToServiceMessages();
                yield this.monitorDisconnect();
                //await fin.desktop.Service.connect({ uuid: excelServiceUuid })
                yield this.registerWindowInstance();
                yield this.getExcelInstances();
                this.mInitialized = true;
            }
            return;
        });
    }
    /**
     * @private
     * @function processExcelServiceEvent Processes events coming from the Excel application
     * @param {any} data Payload passed from the Excel Service
     * @returns {Promise<void>} A promise
     */
    processExcelServiceEvent(data) {
        return __awaiter(this, void 0, void 0, function* () {
            let eventType = data.event;
            let eventData;
            switch (eventType) {
                case "started":
                    break;
                case "registrationRollCall":
                    if (this.mInitialized) {
                        this.registerWindowInstance();
                    }
                    break;
                case "excelConnected":
                    yield this.processExcelConnectedEvent(data);
                    eventData = { connectionUuid: data.uuid };
                    break;
                case "excelDisconnected":
                    yield this.processExcelDisconnectedEvent(data).catch(console.error);
                    eventData = { connectionUuid: data.uuid };
                    break;
            }
            this.dispatchEvent(eventType, eventData);
        });
    }
    /**
     * @private
     * @function processExcelServiceResult Processes results from excel service
     * @param {any} result The result from the service
     * @returns {Promise<void>} A promise
     */
    processExcelServiceResult(result) {
        return __awaiter(this, void 0, void 0, function* () {
            var executor = RpcDispatcher_1.RpcDispatcher.promiseExecutors[result.messageId];
            if (!executor) {
                console.warn("No executors matching the messageId: " + result.messageId, result);
                return;
            }
            delete RpcDispatcher_1.RpcDispatcher.promiseExecutors[result.messageId];
            if (result.error) {
                executor.reject(result.error);
                return;
            }
            // Internal processing
            switch (result.action) {
                case "getExcelInstances":
                    yield this.processGetExcelInstancesResult(result.data);
                    break;
                case "getInitialized":
                    this.mInitialized = result.data;
                    break;
                default:
                    break;
            }
            executor.resolve(result.data);
        });
    }
    /**
     * @private
     * @function subscribeToServiceMessages function to subscribe to topics ExcelService will send to
     * @returns {Promise<[void, void]>} A list of promises
     */
    subscribeToServiceMessages() {
        return Promise.all([
            new Promise(resolve => fin.desktop.InterApplicationBus.subscribe(excelServiceUuid, "excelServiceEvent", this.processExcelServiceEvent.bind(this), resolve)),
            new Promise(resolve => fin.desktop.InterApplicationBus.subscribe(excelServiceUuid, "excelServiceCallResult", this.processExcelServiceResult.bind(this), resolve))
        ]);
    }
    /**
     * @private
     * @function monitorDisconnect Subscribes to the disconnected event and dispatches to the excel application
     * @returns {Promnise<void>} A promise
     */
    monitorDisconnect() {
        return new Promise((resolve, reject) => {
            var excelServiceConnection = fin.desktop.ExternalApplication.wrap(excelServiceUuid);
            var onDisconnect;
            excelServiceConnection.addEventListener("disconnected", onDisconnect = () => {
                excelServiceConnection.removeEventListener("disconnected", onDisconnect);
                this.dispatchEvent("stopped");
            }, resolve, reject);
        });
    }
    /**
     * @private
     * @function registerWindowInstance This registers a new Excel instance to a new workbook domain
     * @returns {Promise<void>} A promise
     */
    registerWindowInstance() {
        return __awaiter(this, void 0, void 0, function* () {
            return this.invokeServiceCall("registerOpenfinWindow", { domain: document.domain });
        });
    }
    /**
     * @private
     * @function configureDefaultApplication Configures the default application when the application first starts
     * @returns {Promise<void>} A promise
     */
    configureDefaultApplication() {
        return __awaiter(this, void 0, void 0, function* () {
            var defaultAppObjUuid = this.defaultApplicationObj && this.defaultApplicationObj.connectionUuid;
            var defaultAppEntry = this.mApplications[defaultAppObjUuid];
            var defaultAppObjConnected = defaultAppEntry ? defaultAppEntry.connected : false;
            if (defaultAppObjConnected) {
                return;
            }
            var connectedAppUuid = Object.keys(this.mApplications).find(appUuid => this.mApplications[appUuid].connected);
            if (connectedAppUuid) {
                delete this.mApplications[defaultAppObjUuid];
                this.defaultApplicationObj = this.mApplications[connectedAppUuid].toObject();
                return;
            }
            if (defaultAppEntry === undefined) {
                var disconnectedAppUuid = fin.desktop.getUuid();
                var disconnectedApp = new ExcelApplication_1.ExcelApplication(disconnectedAppUuid);
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
     * @function processExcelConnectedEvent Process the connected event
     * @param {any} data payload that holds uuid of the connected application
     * @returns {Promise<void>} A promise
     */
    processExcelConnectedEvent(data) {
        return __awaiter(this, void 0, void 0, function* () {
            var applicationInstance = this.mApplications[data.uuid] || new ExcelApplication_1.ExcelApplication(data.uuid);
            yield applicationInstance.init();
            this.mApplications[data.uuid] = applicationInstance;
            // Synthetically raise connected event
            applicationInstance.processExcelEvent({ event: "connected" }, data.uuid);
            yield this.configureDefaultApplication();
            return;
        });
    }
    /**
     * @public
     * @function processExcelDisconnectedEvent Processes event when excel is disconnected
     * @param data The data from excel
     * @returns {Promise<void>} A promise
     */
    processExcelDisconnectedEvent(data) {
        return __awaiter(this, void 0, void 0, function* () {
            var applicationInstance = this.mApplications[data.uuid];
            if (applicationInstance === undefined) {
                return;
            }
            delete this.mApplications[data.uuid];
            console.log('configuring default application in disconnect event');
            this.configureDefaultApplication().then(applicationInstance.release).catch(console.error);
            return;
        });
    }
    // Internal API Handlers
    /**
     * @private
     * @function processGetExcelInstancesResult Get Excel instance
     * @param {string[]} connectionUuids THe connection Uuids the Excel service is holding
     * @returns {Promise<void>} A promise
     */
    processGetExcelInstancesResult(connectionUuids) {
        return __awaiter(this, void 0, void 0, function* () {
            var existingInstances = this.mApplications;
            this.mApplications = {};
            yield Promise.all(connectionUuids.map((connectionUuid) => __awaiter(this, void 0, void 0, function* () {
                var existingInstance = existingInstances[connectionUuid];
                if (existingInstance === undefined) {
                    // Assume that since the ExcelService is aware of the instance it,
                    // is connected and to simulate the the connection event
                    yield this.processExcelServiceEvent({ event: "excelConnected", uuid: connectionUuid });
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
     * @function install Installs the addin
     * @returns {Promise<any>} A promise
     */
    install() {
        return this.invokeServiceCall("install", null);
    }
    /**
     * @public
     * @function getInstallationStatus Checks the installation status
     * @returns {Promise<any>} A promise
     */
    getInstallationStatus() {
        return this.invokeServiceCall("getInstallationStatus", null);
    }
    /**
     * @public
     * @function getExcelInstances Returns all the excel instances that are open
     * @returns {Promise<any>} A promsie
     */
    getExcelInstances() {
        return this.invokeServiceCall("getExcelInstances", null);
    }
    /**
     * @public
     * @function getInitialized Returns whether or not the service has been initialised or not
     * @returns {Promise<any>} A promise
     */
    getInitialized() {
        return this.invokeServiceCall("getInitialized", null);
    }
    /**
     * @public
     * @function toObject Creates an empty object
     * @returns {object} An empty object
     */
    toObject() {
        return {};
    }
}
exports.ExcelService = ExcelService;
//# sourceMappingURL=ExcelApi.js.map