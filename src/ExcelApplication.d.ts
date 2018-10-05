/// <reference types="openfin" />
import { Workbook } from './ExcelWorkbook';
import { RpcDispatcher } from './RpcDispatcher';
export interface LaunchExternalProcessMeta extends fin.ExternalProcessLaunchInfo {
    target: string;
    uuid: string;
}
export interface CalculationMode {
    calculationMode: string;
    calculationState: string;
}
export interface ExcelEventData {
    event: string;
    workbookName: string;
    oldWorkbookName: string;
    sheetName: string;
    oldSheetName: string;
    range: string;
    column: number;
    height: number;
    width: number;
    row: number;
    value: string;
}
export interface Application {
    connectionUuid: string;
    addEventListener(type: string, listener: EventListenerOrEventListenerObject): void;
    removeEventListener(type: string, listener: EventListenerOrEventListenerObject): void;
    addWorkbook(): Promise<Workbook>;
    calculateAll(): Promise<void>;
    getCalculationMode(): Promise<CalculationMode>;
    getConnectionStatus(): Promise<boolean>;
    getWorkbookByName(name: string): Workbook | undefined;
    getWorkbooks(): Promise<Workbooks>;
    openWorkbook(path: string): Promise<Workbook>;
    run(): Promise<void>;
}
export interface Workbooks {
    [workbookName: string]: Workbook;
}
/**
 * @class Represents the Excel application itself
 */
export declare class ExcelApplication extends RpcDispatcher implements Application {
    /**
     * @public
     * @static
     * @description The default excel application instance
     */
    static defaultInstance: ExcelApplication | undefined;
    /**
     * @private
     * @description A key value pair container that holds name of the workbook as
     * key and the workbook object itself as the value
     */
    private workbooks;
    /**
     * @private
     * @description Holds the state of the connection between the Excel instance
     * and the openfin app
     */
    private mConnected;
    /**
     * @private
     * @description Flag to check whether or not the Application has been
     * initialised
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
     * @property
     * @description Flag to indicate whether excel is connected to openfin
     * @returns {boolean} Connected or not
     */
    readonly connected: boolean;
    /**
     * @public
     * @function init
     * @description Initialises the application
     * @returns {Promise<void>} A promise
     */
    init(): Promise<void>;
    /**
     * @public
     * @function release
     * @description Release all connection from the excel application to the
     * openfin app
     * @returns {Promise<void>} A promise
     */
    release(): Promise<void>;
    /**
     * @private
     * @function processExcelEvent
     * @description Process events coming from excel to be handled
     * by the openfin app
     * @param {Readonly<Partial<ExcelEventData>>} data The data being sent over from the excel app
     */
    processExcelEvent(data: Readonly<Partial<ExcelEventData>>): void;
    /**
     * @private
     * @function processExcelResult
     * @description Process results coming from excel application
     * @param {Readonly<ExcelResultData>} result The result of the call being made in the excel application
     */
    private processExcelResult;
    /**
     * @private
     * @function subscribeToExelMessages
     * @description Subscribes to messages from Excel
     * application
     * @returns {Promise<[void, void]>} A promise
     */
    private subscribeToExcelMessages;
    /**
     * @private
     * @function unsubscribeToExcelMessages
     * @description Unsubscribes from Excel application
     * @returns {Promise<[void, void]>} A promise
     */
    private unsubscribeToExcelMessages;
    /**
     * @private
     * @function monitorDisconnect
     * @description Monitors disconnection event when openfin
     * disconnects from excel
     * @returns {Promise<void>} A promise
     */
    private monitorDisconnect;
    /**
     * @public
     * @function run
     * @description Runs Excel application
     * @returns {Promise<void>} A promise
     */
    run(): Promise<void>;
    /**
     * @public
     * @function getWorkbooks
     * @description Gets the workbooks within the excel application
     * @returns {Promise<Workbooks>} A promise
     */
    getWorkbooks(): Promise<Workbooks>;
    /**
     * @public
     * @function getWorkbookByName
     * @description Gets the registered workbook with the specified
     * name
     * @param {string} name The name of the workbook
     */
    getWorkbookByName(name: string): Workbook | undefined;
    /**
     * @public
     * @function addWorkbook
     * @description adds a workbook to the Excel application
     * @returns {Promise<Workbook>} A promise with a result
     */
    addWorkbook(): Promise<Workbook>;
    /**
     * @public
     * @function openWorkbook
     * @description Opens the workbook specified at the path
     * @param {string} path The path of the workbook
     * @returns {Promise<void>} Returns a promise with a result
     */
    openWorkbook(path: string): Promise<Workbook>;
    /**
     * @public
     * @function getConnectionStatus
     * @description Gets the connection status of of the Excel
     * application
     * @returns {Promise<boolean>} A promise with a result
     */
    getConnectionStatus(): Promise<boolean>;
    /**
     * @public
     * @function getCalculationMode
     * @description Gets the calculation mode from Excel
     * application
     * @returns {Promise<CalculationMode>} A promise with a result
     */
    getCalculationMode(): Promise<CalculationMode>;
    /**
     * @public
     * @function calculateAll
     * @description Calculates all formulas on the workbook
     * @returns {Promise<void>} A promise with a result
     */
    calculateAll(): Promise<void>;
    /**
     * @public
     * @function toObject
     * @description Returns an object with only the methods and properties
     * to be exposed
     * @returns {Application} An object with only the methods and properties to be exposed
     */
    toObject(): Application;
}
