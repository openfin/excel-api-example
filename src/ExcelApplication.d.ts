import { RpcDispatcher } from './RpcDispatcher';
import { ExcelWorkbook } from './ExcelWorkbook';
/**
 * @class Represents the Excel application itself
 */
export declare class ExcelApplication extends RpcDispatcher {
    /**
     * @public
     * @static
     * @description The default excel application instance
     */
    static defaultInstance: ExcelApplication;
    /**
     * @private
     * @description A key value pair container that holds name of the workbook as key
     * and the workbook object itself as the value
     */
    private workbooks;
    /**
     * @private
     * @description Holds the state of the connection between the Excel instance and the openfin app
     */
    private mConnected;
    /**
     * @private
     * @description Flag to check whether or not the Application has been initialised
     */
    private initialized;
    /**
     * @private
     * @description Instance of the ExcelApplication object itself
     */
    private objectInstance;
    /**
     * @constructor Constructor for the class
     * @param connectionUuid The connection uuid of the openfin application
     */
    constructor(connectionUuid: string);
    /**
     * @public
     * @property Flag to indicate whether excel is connected to openfin
     * @returns {boolean} Connected or not
     */
    readonly connected: boolean;
    /**
     * @public
     * @function init Initialises the application
     * @returns {Promise<void>} A promise
     */
    init(): Promise<void>;
    /**
     * @public
     * @function release Release all connection from the excel application to the openfin app
     * @returns {Promise<void>} A promise
     */
    release(): Promise<void>;
    /**
     * @private
     * @function processExcelEvent Process events coming from excel to be handled by the openfin app
     * @param data The data being sent over from the excel app
     * @param uuid The uuid of the sender
     */
    processExcelEvent(data: any, uuid: string): void;
    /**
     * @private
     * @function processExcelResult Process results coming from excel application
     * @param result The result of the call being made in the excel application
     */
    private processExcelResult(result);
    /**
     * @private
     * @function subscribeToExelMessages Subscribes to messages from Excel application
     * @returns {Promise<[void, void]>} A promise
     */
    private subscribeToExcelMessages();
    /**
     * @private
     * @function unsubscribeToExcelMessages Unsubscribes from Excel application
     * @returns {Promise<[void, void]>} A promise
     */
    private unsubscribeToExcelMessages();
    /**
     * @private
     * @function monitorDisconnect Monitors disconnection event when openfin disconnects from excel
     * @returns {Promise<void>} A promise
     */
    private monitorDisconnect();
    /**
     * @public
     * @function run Runs Excel application
     * @param callback The callback to be applied
     */
    run(callback?: Function): Promise<void>;
    /**
     * @public
     * @function getWorkbooks Gets the workbooks within the excel application
     * @returns {Promise<any>} A promise
     */
    getWorkbooks(): Promise<any>;
    /**
     * @public
     * @function getWorkbookByName Gets the registered workbook with the specified name
     * @param name The name of the workbook
     */
    getWorkbookByName(name: string): ExcelWorkbook;
    /**
     * @function addWorkbook adds a workbook to the Excel application
     * @returns {Promise<any>} A promise with a result
     */
    addWorkbook(): Promise<any>;
    /**
     * @public
     * @function openWorkbook Opens the workbook specified at the path
     * @param path The path of the workbook
     * @returns {Promise<any>} Returns a promise with a result
     */
    openWorkbook(path: string): Promise<any>;
    /**
     * @public
     * @function getConnectionStatus Gets the connection status of of the Excel application
     * @returns {Promise<any>} A promise with a result
     */
    getConnectionStatus(callback?: Function): Promise<any>;
    /**
     * @public
     * @function getCalculationMode Gets the calculation mode from Excel application
     * @returns {Promise<any>} A promise with a result
     */
    getCalculationMode(): Promise<any>;
    /**
     * @public
     * @function calculateAll Calculates all formulas on the workbook
     * @returns {Promise<any>} A promise with a result
     */
    calculateAll(): Promise<any>;
    /**
     * @public
     * @function toObject Returns an object with only the methods and properties to be exposed
     * @returns {any} An object with only the methods and properties to be exposed
     */
    toObject(): any;
}
