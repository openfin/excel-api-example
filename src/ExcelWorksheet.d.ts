import { RpcDispatcher } from './RpcDispatcher';
import { ExcelWorkbook } from './ExcelWorkbook';
/**
 * @class Class that represents a worksheet
 */
export declare class ExcelWorksheet extends RpcDispatcher {
    /**
     * @private
     * @description Handle to the ExcelWorkbook
     */
    private workbook;
    /**
     * @private
     * @description Name of the worksheet
     */
    private mWorksheetName;
    /**
     * @private
     * @description An instance to the ExcelWorksheet itself
     */
    private objectInstance;
    /**
     * @constructor Constructor for the ExcelWorksheet class
     * @param name The name of the worksheet
     * @param workbook The ExcelWorkbook this worksheet is tied to
     */
    constructor(name: string, workbook: ExcelWorkbook);
    /**
     * @protected
     * @function getDefaultMessage Returns the default message
     * @returns {any} Returns the default message
     */
    protected getDefaultMessage(): any;
    /**
     * @public
     * @property Returns worksheet name
     * @returns {string} The name of the worksheet
     */
    /**
     * @public
     * @property Returns worksheet name
     * @returns {string} The name of the worksheet
     */
    worksheetName: string;
    /**
     * @public
     * @function setCells Sets the content for the cells
     * @param values values for the cell
     * @param offset The cell address
     */
    setCells(values: any[][], offset: string): Promise<any>;
    /**
     * @public
     * @function getCells Gets cell values from the range specified
     * @param start The start cell address
     * @param offsetWidth The number of columns in the openfin app
     * @param offsetHeight The number of rows in the openfin app
     */
    getCells(start: string, offsetWidth: number, offsetHeight: number): Promise<any>;
    /**
     * @function activateRow This mirrors the row selected in the openfin application to Excel
     * @param {string} cellAddress THe address of the first cell of the row
     */
    activateRow(cellAddress: string): Promise<void>;
    /**
     * @function insertRow This inserts a row just before the selected row
     * @param {number} rowNumber The address of the first cell in the row
     * @returns {Promise<any>} A promise
     */
    insertRow(rowNumber: number): Promise<any>;
    /**
     * @function deleteRow This deletes the selected row
     * @param {number} rowNumber The address of the first cell in the row
     * @returns {Promise<any>} A promise
     */
    deleteRow(rowNumber: string): Promise<any>;
    /**
     * @public
     * @function activate Activates the current worksheet
     * @returns {Promise<any>} A promise
     */
    activate(): Promise<any>;
    /**
     * @public
     * @function activateCell Activates the selected cell
     * @param cellAddress The address of the cell
     * @returns {Promise<any>} A promise
     */
    activateCell(cellAddress: string): Promise<any>;
    addButton(name: string, caption: string, cellAddress: string): Promise<any>;
    setFilter(start: string, offsetWidth: number, offsetHeight: number, field: number, criteria1: string, op: string, criteria2: string, visibleDropDown: string): Promise<any>;
    /**
     * @public
     * @function formatRange Formats the range selected
     * @param rangeCode The selected range
     * @param format The formatting to be applied to the range
     */
    formatRange(rangeCode: string, format: any): Promise<any>;
    /**
     * @public
     * @function clearRange Clear the range of formatting and content
     * @param rangeCode The range selected
     */
    clearRange(rangeCode: string): Promise<any>;
    /**
     * @public
     * @function clearRangeContents Clears the contents in the specified range
     * @param rangeCode The selected range
     */
    clearRangeContents(rangeCode: string): Promise<any>;
    /**
     * @public
     * @function clearRangeFormats Clears the formatting in the range specified
     * @param rangeCode The selected range
     */
    clearRangeFormats(rangeCode: string): Promise<any>;
    /**
     * @public
     * @function clearAllCells Clear all cells and their formatting
     * @returns {Promise<any>} A promise
     */
    clearAllCells(): Promise<any>;
    /**
     * @public
     * @function clearAllCellContents Clears all the cells content
     * @returns {Promise<any>} A promise
     */
    clearAllCellContents(): Promise<any>;
    /**
     * @public
     * @function clearAllCellFormats Clear all formatting in every cell
     * @returns {Promise<any>} A promise
     */
    clearAllCellFormats(): Promise<any>;
    /**
     * @public
     * @function setCellName Sets a name for the cell address
     * @param cellAddress The address of the cell e.g. A1
     * @param cellName The name of the cell
     * @returns {Promise<any>} A promise
     */
    setCellName(cellAddress: string, cellName: string): Promise<any>;
    /**
     * @public
     * @function calculate Calculates all formula on teh sheet
     * @returns {Promise<any>} A promise
     */
    calculate(): Promise<any>;
    /**
     * @public
     * @function getCellByName Gets a cell by its name
     * @param cellName The name of the cell
     * @returns {Promise<any>} A promise
     */
    getCellByName(cellName: string): Promise<any>;
    /**
     * @public
     * @function protect Password protects the sheet
     * @param password Password used to protect the sheet
     * @returns {Promise<any>} A promise
     */
    protect(password: string): Promise<any>;
    /**
     * @public
     * @function toObject Returns only the functions that should be exposed by this class
     * @returns {object} Public methods in ExcelWorksheet
     */
    toObject(): object;
}
