import { Workbook } from './ExcelWorkbook';
import { RpcDispatcher } from './RpcDispatcher';
/**
 * Worksheet public functions
 */
export interface Worksheet {
    addEventListener(type: string, listener: EventListenerOrEventListenerObject): void;
    removeEventListener(type: string, listener: EventListenerOrEventListenerObject): void;
    name: string;
    workbook: Workbook;
    activate(): Promise<void>;
    activateCell(cellAddress: string): Promise<void>;
    activateRow(cellAddress: string): Promise<void>;
    calculate(): Promise<void>;
    clearAllCellContents(): Promise<void>;
    clearAllCells(): Promise<void>;
    clearRange(rangeCode: string): Promise<void>;
    clearRangeContents(rangeCode: string): Promise<void>;
    getCellByName(cellName: string): Promise<Cell>;
    getCells(start: string, offsetWidth: number, offsetHeight: number): Promise<(string | number)[][]>;
    protect(password: string): Promise<void>;
    setCellName(cellAddress: string, cellName: string): Promise<void>;
    setCells(values: (string | number)[][], offset: string): Promise<void>;
    insertRow(rowNumber: number): Promise<void>;
    deleteRow(rowNumber: number): Promise<void>;
    toObject(): Worksheet;
}
/**
 * Information about the cell
 */
export interface Cell extends CellValue {
    address: string;
    column: number;
    row: number;
}
/**
 * Value of the cell
 */
export interface CellValue {
    formula: string;
    value: string | number;
}
export interface CellArrayRange {
    arrayOfCells: CellValue[];
}
/**
 * Payload for set cells payload
 */
export interface SetCellsPayload {
    offset: string;
    values: (string | number)[][];
}
export interface GetCellsPayload {
    start: string;
    offsetWidth: number;
    offsetHeight: number;
}
export interface CellAddress {
    address: string;
}
/**
 * @class Class that represents a worksheet
 */
export declare class ExcelWorksheet extends RpcDispatcher implements Worksheet, EventTarget {
    /**
     * @private
     * @description Handle to the ExcelWorkbook
     */
    private mWorkbook;
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
    constructor(name: string, workbook: Workbook);
    /**
     * @protected
     * @function getDefaultMessage Returns the default message
     * @returns {object} Returns the default message
     */
    protected getDefaultMessage(): object;
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
    name: string;
    readonly workbook: Workbook;
    /**
     * @public
     * @function setCells Sets the content for the cells
     * @param values values for the cell
     * @param offset The cell address
     */
    setCells(values: (string | number)[][], offset: string): Promise<void>;
    /**
     * @public
     * @function getCells Gets cell values from the range specified
     * @param start The start cell address
     * @param offsetWidth The number of columns in the openfin app
     * @param offsetHeight The number of rows in the openfin app
     */
    getCells(start: string, offsetWidth: number, offsetHeight: number): Promise<(string | number)[][]>;
    /**
     * @function activateRow This mirrors the row selected in the openfin
     * application to Excel
     * @param {string} cellAddress THe address of the first cell of the row
     */
    activateRow(cellAddress: string): Promise<void>;
    /**
     * @function insertRow This inserts a row just before the selected row
     * @param {number} rowNumber The address of the first cell in the row
     * @returns {Promise<void>} A promise
     */
    insertRow(rowNumber: number): Promise<void>;
    /**
     * @function deleteRow This deletes the selected row
     * @param {number} rowNumber The address of the first cell in the row
     * @returns {Promise<any>} A promise
     */
    deleteRow(rowNumber: number): Promise<void>;
    /**
     * @public
     * @function activate Activates the current worksheet
     * @returns {Promise<any>} A promise
     */
    activate(): Promise<void>;
    /**
     * @public
     * @function activateCell Activates the selected cell
     * @param cellAddress The address of the cell
     * @returns {Promise<void>} A promise
     */
    activateCell(cellAddress: string): Promise<void>;
    /**
     * @public
     * @function clearRange Clear the range of formatting and content
     * @param rangeCode The range selected
     */
    clearRange(rangeCode: string): Promise<void>;
    /**
     * @public
     * @function clearRangeContents Clears the contents in the specified range
     * @param rangeCode The selected range
     */
    clearRangeContents(rangeCode: string): Promise<void>;
    /**
     * @public
     * @function clearAllCells Clear all cells and their formatting
     * @returns {Promise<any>} A promise
     */
    clearAllCells(): Promise<void>;
    /**
     * @public
     * @function clearAllCellContents Clears all the cells content
     * @returns {Promise<void>} A promise
     */
    clearAllCellContents(): Promise<void>;
    /**
     * @public
     * @function setCellName Sets a name for the cell address
     * @param cellAddress The address of the cell e.g. A1
     * @param cellName The name of the cell
     * @returns {Promise<any>} A promise
     */
    setCellName(cellAddress: string, cellName: string): Promise<void>;
    /**
     * @public
     * @function calculate Calculates all formula on teh sheet
     * @returns {Promise<void>} A promise
     */
    calculate(): Promise<void>;
    /**
     * @public
     * @function getCellByName Gets a cell by its name
     * @param cellName The name of the cell
     * @returns {Promise<any>} A promise
     */
    getCellByName(cellName: string): Promise<Cell>;
    /**
     * @public
     * @function protect Password protects the sheet
     * @param password Password used to protect the sheet
     * @returns {Promise<any>} A promise
     */
    protect(password: string): Promise<void>;
    /**
     * @public
     * @function toObject Returns only the functions that should be exposed by
     * this class
     * @returns {object} Public methods in ExcelWorksheet
     */
    toObject(): Worksheet;
}
