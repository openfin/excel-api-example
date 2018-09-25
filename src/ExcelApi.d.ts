import { Application } from './ExcelApplication';
import { RpcDispatcher } from './RpcDispatcher';
/**
 * @class Class for interacting with the .NET ExcelService process
 */
export declare class ExcelService extends RpcDispatcher {
    /**
     * @private
     * @description Handle to the default application uuid
     */
    private mDefaultApplicationUuid;
    /**
     * @public
     * @description Handle to the default application
     */
    defaultApplicationObj: Application;
    /**
     * @private
     * @description Checks whether the ExcelService is initialised
     */
    private mInitialized;
    /**
     * @private
     * @description Keeps track of the excel instances running
     */
    private mApplications;
    /**
     * @constructor Constructor for ExcelService
     */
    constructor();
    /**
     * @public
     * @function init Initialises the ExcelService
     * @returns {Promise<void>} A promise
     */
    init(): Promise<void>;
    /**
     * @private
     * @function processExcelServiceEvent Processes events coming from the Excel
     * application
     * @param {any} data Payload passed from the Excel Service
     * @returns {Promise<void>} A promise
     */
    private processExcelServiceEvent;
    /**
     * @private
     * @function processExcelServiceResult Processes results from excel service
     * @param {any} result The result from the service
     * @returns {Promise<void>} A promise
     */
    private processExcelServiceResult;
    /**
     * @private
     * @function subscribeToServiceMessages function to subscribe to topics
     * ExcelService will send to
     * @returns {Promise<[void, void]>} A list of promises
     */
    private subscribeToServiceMessages;
    /**
     * @private
     * @function monitorDisconnect Subscribes to the disconnected event and
     * dispatches to the excel application
     * @returns {Promnise<void>} A promise
     */
    private monitorDisconnect;
    /**
     * @private
     * @function registerWindowInstance This registers a new Excel instance to a
     * new workbook domain
     * @returns {Promise<void>} A promise
     */
    private registerWindowInstance;
    /**
     * @private
     * @function configureDefaultApplication Configures the default application
     * when the application first starts
     * @returns {Promise<void>} A promise
     */
    private configureDefaultApplication;
    /**
     * @private
     * @function processExcelConnectedEvent Process the connected event
     * @param {ExcelConnectionEventData} data payload that holds uuid of the connected application
     * @returns {Promise<void>} A promise
     */
    private processExcelConnectedEvent;
    /**
     * @public
     * @function processExcelDisconnectedEvent Processes event when excel is
     * disconnected
     * @param data The data from excel
     * @returns {Promise<void>} A promise
     */
    private processExcelDisconnectedEvent;
    /**
     * @private
     * @function processGetExcelInstancesResult Get Excel instance
     * @param {string[]} connectionUuids THe connection Uuids the Excel service is holding
     * @returns {Promise<void>} A promise
     */
    private processGetExcelInstancesResult;
    /**
     * @public
     * @function install Installs the addin
     * @returns {Promise<void>} A promise
     */
    install(): Promise<void>;
    /**
     * @public
     * @function getInstallationStatus Checks the installation status
     * @returns {Promise<void>} A promise
     */
    getInstallationStatus(): Promise<void>;
    /**
     * @public
     * @function getExcelInstances Returns all the excel instances that are open
     * @returns {Promise<void>} A promsie
     */
    getExcelInstances(): Promise<void>;
    /**
     * @public
     * @function toObject Creates an empty object
     * @returns {object} An empty object
     */
    toObject(): {};
}
