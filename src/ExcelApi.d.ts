import { Application } from './ExcelApplication';
import { RpcDispatcher } from './RpcDispatcher';
/**
 * @class
 * @description Class for interacting with the .NET ExcelService process
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
    defaultApplicationObj: Application | null;
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
     * @constructor
     * @description Constructor for ExcelService
     */
    constructor();
    /**
     * @public
     * @async
     * @function init
     * @description Initialises the ExcelService
     * @returns {Promise<void>} A promise
     */
    init(): Promise<void>;
    /**
     * @private
     * @async
     * @function processExcelServiceEvent
     * @description Processes events coming from the Excel
     * application
     * @param {ExcelServiceEventData} data Payload passed from the Excel Service
     * @returns {Promise<void>} A promise
     */
    private processExcelServiceEvent;
    /**
     * @private
     * @async
     * @function processExcelServiceResult
     * @description Processes results from excel service
     * @param {ExcelResultData} result The result from the service
     * @returns {Promise<void>} A promise
     */
    private processExcelServiceResult;
    /**
     * @private
     * @function subscribeToServiceMessages
     * @description function to subscribe to topics
     * @returns {Promise<[void, void]>} A list of promises
     */
    private subscribeToServiceMessages;
    /**
     * @private
     * @function monitorDisconnect
     * @description Subscribes to the disconnected event and
     * dispatches to the excel application
     * @returns {Promnise<void>} A promise
     */
    private monitorDisconnect;
    /**
     * @private
     * @async
     * @function registerWindowInstance
     * @description This registers a new Excel instance to a
     * new workbook domain
     * @returns {Promise<void>} A promise
     */
    private registerWindowInstance;
    /**
     * @private
     * @async
     * @function configureDefaultApplication
     * @description Configures the default application
     * when the application first starts
     * @returns {Promise<void>} A promise
     */
    private configureDefaultApplication;
    /**
     * @private
     * @async
     * @function processExcelConnectedEvent
     * @description Process the connected event
     * @param {ExcelConnectionEventData} data payload that holds uuid of the connected application
     * @returns {Promise<void>} A promise
     */
    private processExcelConnectedEvent;
    /**
     * @public
     * @async
     * @function processExcelDisconnectedEvent
     * @description Processes event when excel is
     * disconnected
     * @param {ExcelConnectionEventData} data The data from excel
     * @returns {Promise<void>} A promise
     */
    private processExcelDisconnectedEvent;
    /**
     * @private
     * @async
     * @function processGetExcelInstancesResult
     * @description Gets Excel instance
     * @param {string[]} connectionUuids THe connection Uuids the Excel service is holding
     * @returns {Promise<void>} A promise
     */
    private processGetExcelInstancesResult;
    /**
     * @public
     * @function install
     * @description Get Excel instance
     * @returns {Promise<void>} A promise
     */
    install(): Promise<void>;
    /**
     * @public
     * @function getInstallationStatus
     * @description Checks the installation status
     * @returns {Promise<void>} A promise
     */
    getInstallationStatus(): Promise<void>;
    /**
     * @public
     * @function getExcelInstances
     * @description Returns all the excel instances that are open
     * @returns {Promise<void>} A promsie
     */
    getExcelInstances(): Promise<void>;
    /**
     * @public
     * @function toObject
     * @description Creates an empty object
     * @returns {object} An empty object
     */
    toObject(): {};
}
