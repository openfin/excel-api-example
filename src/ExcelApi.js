"use strict";
const RpcDispatcher_1 = require('./RpcDispatcher');
const ExcelApplication_1 = require('./ExcelApplication');
const excelServiceUuid = "886834D1-4651-4872-996C-7B2578E953B9";
class ExcelApi extends RpcDispatcher_1.RpcDispatcher {
    constructor() {
        super();
        this.applications = {};
        this.processExcelServiceEvent = (data) => {
            var preventDefault = false;
            var eventPayload = { type: data.event };
            switch (data.event) {
                case "started":
                    break;
                case "registrationRollCall":
                    this.registerAppInstance();
                    break;
                case "excelConnected":
                    this.processExcelConnectedEvent(data);
                    break;
                case "excelDisconnected":
                    this.processExcelDisconnectedEvent(data);
                    break;
            }
            if (!preventDefault) {
                this.dispatchEvent(eventPayload);
            }
        };
        this.processExcelServiceResult = (data) => {
            // Internal processing
            switch (data.action) {
                case "getExcelInstances":
                    this.processGetExcelInstancesResult(data.result);
                    break;
            }
            // Dispatch result to callbacks
            if (RpcDispatcher_1.RpcDispatcher.callbacks[data.messageId]) {
                RpcDispatcher_1.RpcDispatcher.callbacks[data.messageId](data.result);
                delete RpcDispatcher_1.RpcDispatcher.callbacks[data.messageId];
            }
        };
        this.registerAppInstance = () => {
            this.invokeServiceCall("registerAppInstance");
        };
        this.connectionUuid = excelServiceUuid;
    }
    init() {
        if (!this.initialized) {
            fin.desktop.InterApplicationBus.subscribe("*", "excelServiceEvent", this.processExcelServiceEvent);
            fin.desktop.InterApplicationBus.subscribe("*", "excelServiceCallResult", this.processExcelServiceResult);
            this.registerAppInstance();
            this.getExcelInstances();
            this.monitorDisconnect();
            this.initialized = true;
        }
    }
    monitorDisconnect() {
        fin.desktop.ExternalApplication.wrap(excelServiceUuid).addEventListener("disconnected", () => {
            this.dispatchEvent({ type: "stopped" });
        });
    }
    connectLegacyApi(connectedUuid) {
        if (!ExcelApi.legacyApi) {
            ExcelApi.legacyApi = ExcelApi.instance.applications[connectedUuid];
        }
    }
    disconnectLegacyApi(disconnectedUuid) {
        if (ExcelApi.legacyApi.connectionUuid === disconnectedUuid) {
            ExcelApi.legacyApi = undefined;
            for (var connectionUuid in ExcelApi.instance.applications) {
                ExcelApi.legacyApi = ExcelApi.instance.applications[connectionUuid];
                break;
            }
        }
    }
    // Internal Event Handlers
    processExcelConnectedEvent(data) {
        var applicationInstance = this.applications[data.uuid] || new ExcelApplication_1.ExcelApplication(data.uuid);
        this.applications[data.uuid] = applicationInstance;
        applicationInstance.init();
        // Synthetically raise connected event
        applicationInstance.processExcelEvent({ event: "connected" }, data.uuid);
        this.connectLegacyApi(data.uuid);
    }
    processExcelDisconnectedEvent(data) {
        delete this.applications[data.uuid];
        this.disconnectLegacyApi(data.uuid);
    }
    // Internal API Handlers
    processGetExcelInstancesResult(connectionUuids) {
        var oldInstances = this.applications;
        this.applications = {};
        connectionUuids.forEach(connectionUuid => {
            var applicationInstance = oldInstances[connectionUuid] || new ExcelApplication_1.ExcelApplication(connectionUuid);
            this.applications[connectionUuid] = applicationInstance;
            applicationInstance.init();
            this.connectLegacyApi(connectionUuid);
        });
    }
    // API Calls
    install(callback) {
        this.invokeServiceCall("install", null, callback);
    }
    getInstallationStatus(callback) {
        this.invokeServiceCall("getInstallationStatus", null, callback);
    }
    getExcelInstances(callback) {
        this.invokeServiceCall("getExcelInstances", null, callback);
    }
    // Legacy API / Single-Application Functions
    static init() {
        ExcelApi.instance.init();
    }
    static addEventListener(type, listener) {
        ExcelApi.legacyApi.addEventListener(type, listener);
    }
    static removeEventListener(type, listener) {
        ExcelApi.legacyApi.removeEventListener(type, listener);
    }
    static run(callback) {
        if (ExcelApi.legacyApi && callback) {
            callback();
        }
        else {
            var connectedCallback = () => {
                ExcelApi.instance.removeEventListener("excelConnected", connectedCallback);
                callback && callback();
            };
            ExcelApi.instance.addEventListener("excelConnected", connectedCallback);
            fin.desktop.System.launchExternalProcess({
                target: "excel"
            });
        }
    }
    static install(callback) {
        ExcelApi.instance.install(callback);
    }
    static getInstallationStatus(callback) {
        ExcelApi.instance.getInstallationStatus(callback);
    }
    static getWorkbooks(callback) {
        ExcelApi.legacyApi.getWorkbooks(callback);
    }
    static getWorkbookByName(name) {
        return ExcelApi.legacyApi.getWorkbookByName(name);
    }
    static getWorksheetByName(workbookName, worksheetName) {
        return ExcelApi.legacyApi.getWorksheetByName(workbookName, worksheetName);
    }
    static addWorkbook(callback) {
        ExcelApi.legacyApi.addWorkbook(callback);
    }
    static openWorkbook(path, callback) {
        ExcelApi.legacyApi.openWorkbook(path, callback);
    }
    static getConnectionStatus(callback) {
        ExcelApi.legacyApi.getConnectionStatus(callback);
    }
    static getCalculationMode(callback) {
        ExcelApi.legacyApi.getCalculationMode(callback);
    }
    static calculateAll(callback) {
        ExcelApi.legacyApi.calculateAll(callback);
    }
    static toObject() {
        return ExcelApi.legacyApi.toObject();
    }
}
ExcelApi.instance = new ExcelApi();
ExcelApi.legacyApi = undefined;
exports.ExcelApi = ExcelApi;
exports.LegacyApi = ExcelApi;
Object.defineProperty(exports, "__esModule", { value: true });
exports.default = ExcelApi.instance;
//# sourceMappingURL=ExcelApi.js.map