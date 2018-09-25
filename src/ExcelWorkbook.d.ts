import { ExcelApplication } from './ExcelApplication';
import { Worksheet } from './ExcelWorksheet';
import { RpcDispatcher } from './RpcDispatcher';
export interface Workbook {
    addEventListener(type: string, listener: EventListenerOrEventListenerObject): void;
    removeEventListener(type: string, listener: EventListenerOrEventListenerObject): void;
    name: string;
    worksheets?: Worksheets;
    activate: () => Promise<void>;
    addWorksheet(): Promise<Worksheet>;
    close(): Promise<void>;
    getWorksheetByName(name: string): Worksheet;
    getWorksheets(): Promise<Worksheets>;
    save(): Promise<void>;
    toObject(): Workbook;
}
/**
 * Worksheets object
 */
export interface Worksheets {
    [worksheetName: string]: Worksheet;
}
/**
 * @class Class that represents a workbook
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
     * @constructor Constructor for the ExcelWorkbook class
     * @param application The Application this workbook belongs to
     * @param name The name of the workbook
     */
    constructor(application: ExcelApplication, name: string);
    /**
     * @private
     * @function getDefaultMessage Gets the default message to be sent over the
     * wire
     * @returns {any} An object with the workbook name in as default
     */
    protected getDefaultMessage(): object;
    /**
     * @public
     * @property Worksheets tied to this workbook
     * @returns {Worksheets}
     */
    worksheets: Worksheets;
    /**
     * @public
     * @property workbookName property
     * @returns {string} The name of the workbook
     */
    /**
    * @public
    * @property Sets the workbook name
    */
    name: string;
    /**
     * @public
     * @function getWorksheets Gets the worksheets tied to this workbook
     * @returns A promise with worksheets as the result
     */
    getWorksheets(): Promise<Worksheets>;
    /**
     * @public
     * @function getWorksheetByName Gets the worksheet by name
     * @param name The name of the worksheet
     * @returns {ExcelWorksheet} The excel worksheet with the specified name
     */
    getWorksheetByName(name: string): Worksheet;
    /**
     * @public
     * @function addWorksheet Adds a new worksheet to the workbook
     * @returns {Promise<any>} A promise
     */
    addWorksheet(): Promise<Worksheet>;
    /**
     * @public
     * @function activate Activates the workbook
     * @returns {Promise<any>} A promise
     */
    activate(): Promise<void>;
    /**
     * @public
     * @function save Save the workbook
     * @returns {Promise<void>} A promise
     */
    save(): Promise<void>;
    /**
     * @public
     * @function close Closes the workbook
     * @returns {Promise<void>} A promise
     */
    close(): Promise<void>;
    /**
     * @public
     * @function toObject Returns only the methods exposed
     * @returns {Workbook} Returns only the methods exposed
     */
    toObject(): Workbook;
}
