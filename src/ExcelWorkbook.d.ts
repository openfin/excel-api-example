import { RpcDispatcher } from './RpcDispatcher';
import { ExcelApplication } from './ExcelApplication';
import { ExcelWorksheet } from './ExcelWorksheet';
/**
 * @class Class that represents a workbook
 */
export declare class ExcelWorkbook extends RpcDispatcher {
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
     * @function getDefaultMessage Gets the default message to be sent over the wire
     * @returns {any} An object with the workbook name in as default
     */
    protected getDefaultMessage(): any;
    /**
     * @public
     * @property Worksheets tied to this workbook
     * @returns {{ [worksheetName: string]: ExcelWorksheet }}
     */
    worksheets: {
        [worksheetName: string]: ExcelWorksheet;
    };
    /**
     * @public
     * @property workbookName property
     * @returns {string} The name of the workbook
     */
    /**
     * @public
     * @property Sets the workbook name
     */
    workbookName: string;
    /**
     * @public
     * @function getWorksheets Gets the worksheets tied to this workbook
     * @returns A promise with worksheets as the result
     */
    getWorksheets(): Promise<any>;
    /**
     * @public
     * @function getWorksheetByName Gets the worksheet by name
     * @param name The name of the worksheet
     * @returns {ExcelWorksheet} The excel worksheet with the specified name
     */
    getWorksheetByName(name: string): ExcelWorksheet;
    /**
     * @public
     * @function addWorksheet Adds a new worksheet to the workbook
     * @returns {Promise<any>} A promise
     */
    addWorksheet(): Promise<any>;
    /**
     * @public
     * @function activate Activates the workbook
     * @returns {Promise<any>} A promise
     */
    activate(): Promise<any>;
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
     * @returns {any} Returns only the methods exposed
     */
    toObject(): any;
}
