import { Application } from './ExcelApplication';
import { Worksheet } from './ExcelWorksheet';
import { RpcDispatcher } from './RpcDispatcher';
/**
 * @description Interface for the workbook
 */
export interface Workbook {
    addEventListener(type: string, listener: EventListenerOrEventListenerObject): void;
    removeEventListener(type: string, listener: EventListenerOrEventListenerObject): void;
    name: string;
    worksheets?: Worksheets;
    activate: () => Promise<void>;
    addWorksheet(): Promise<Worksheet>;
    close(): Promise<void>;
    getWorksheetByName(name: string): Worksheet | undefined;
    getWorksheets(): Promise<Worksheets>;
    save(): Promise<void>;
    toObject(): Workbook | undefined;
}
/**
 * @description Worksheets object
 */
export interface Worksheets {
    [worksheetName: string]: Worksheet;
}
/**
 * @class
 * @description Class that represents a workbook
 */
export declare class ExcelWorkbook extends RpcDispatcher implements Workbook, EventTarget {
    /**
     * @private
     * @description The application instance itself
     */
    private application;
    /**
     * @private
     * @description The name of the workbook
     */
    private mWorkbookName;
    /**
     * @private
     * @description The worksheets tied to this workbook
     */
    private mWorksheets;
    /**
     * @private
     * @description An instance of this object
     */
    private objectInstance;
    /**
     * @constructor
     * @description Constructor for the ExcelWorkbook class
     * @param {Application} application The Application this workbook belongs to
     * @param {string}name The name of the workbook
     */
    constructor(application: Application, name: string);
    /**
     * @private
     * @function getDefaultMessage
     * @description Gets the default message to be sent over the
     * wire
     * @returns {object} An object with the workbook name in as default
     */
    protected getDefaultMessage(): object;
    /**
     * @public
     * @property
     * @description Worksheets tied to this workbook
     * @returns {Worksheets} The worksheets tied to this workbook
     */
    /**
    * @public
    * @property
    * @description Set the worksheets that are tied to this workbook
    */
    worksheets: Worksheets;
    /**
     * @public
     * @property
     * @description workbookName property
     * @returns {string} The name of the workbook
     */
    /**
    * @public
    * @property
    * @description Sets the workbook name
    * @param {string} name Set the name of the workbook
    */
    name: string;
    /**
     * @public
     * @function getWorksheets
     * @description Gets the worksheets tied to this workbook
     * @returns {Promise<Worksheets>} A promise with worksheets as the result
     */
    getWorksheets(): Promise<Worksheets>;
    /**
     * @public
     * @function getWorksheetByName
     * @description Gets the worksheet by name
     * @param {string} name The name of the worksheet
     * @returns {ExcelWorksheet} The excel worksheet with the specified name
     */
    getWorksheetByName(name: string): Worksheet | undefined;
    /**
     * @public
     * @function addWorksheet Adds a new worksheet to the workbook
     * @description Adds a new worksheet to the workbook
     * @returns {Promise<Worksheet>} A promise
     */
    addWorksheet(): Promise<Worksheet>;
    /**
     * @public
     * @function activate
     * @description Activates the workbook
     * @returns {Promise<void>} A promise
     */
    activate(): Promise<void>;
    /**
     * @public
     * @function save
     * @description Save the current workbook
     * @returns {Promise<void>} A promise
     */
    save(): Promise<void>;
    /**
     * @public
     * @function close
     * @description Closes the workbook
     * @returns {Promise<void>} A promise
     */
    close(): Promise<void>;
    /**
     * @public
     * @function toObject
     * @description Returns only the methods exposed
     * @returns {Workbook} Returns only the methods exposed
     */
    toObject(): Workbook | undefined;
}
