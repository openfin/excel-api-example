import { Workbook } from './ExcelWorkbook';
import { RpcDispatcher } from './RpcDispatcher';
/**
 * @description Worksheet public functions
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
 * @description Information about the cell
 */
export interface Cell extends CellValue {
    address: string;
    column: number;
    row: number;
}
/**
 * @description Value of the cell
 */
export interface CellValue {
    formula: string;
    value: string | number;
}
/**
 * @description Range of cells
 */
export interface CellArrayRange {
    arrayOfCells: CellValue[];
}
/**
 * @description Payload for set cells function
 */
export interface SetCellsPayload {
    offset: string;
    values: (string | number)[][];
}
/**
 * @description Payload for get cells function
 */
export interface GetCellsPayload {
    start: string;
    offsetWidth: number;
    offsetHeight: number;
}
/**
 * @description The address of the cell
 */
export interface CellAddress {
    address: string;
}
/**
 * @class
 * @description Class that represents a worksheet
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
     * @constructor
     * @description Constructor for the ExcelWorksheet class
     * @param {string} name The name of the worksheet
     * @param {Workbook} workbook The ExcelWorkbook this worksheet is tied to
     */
    constructor(name: string, workbook: Workbook);
    /**
     * @protected
     * @function getDefaultMessage
     * @description Returns the default message
     * @returns {object} Returns the default message
     */
    protected getDefaultMessage(): object;
    /**
     * @public
     * @property
     * @description Returns worksheet name
     * @returns {string} The name of the worksheet
     */
    /**
    * @public
    * @property
    * @description Returns worksheet name
    * @param {string} name The name of the worksheet
    */
    name: string;
    /**
     * @public
     * @property
     * @description Returns the workbook that this worksheet is tied to
     * @returns {Workbook} Returns the workbook that this worksheet is tied to
     */
    readonly workbook: Workbook;
    /**
     * @public
     * @function setCells
     * @description Sets the content for the cells
     * @param {(string|number)[][]} values values for the cell
     * @param {string} offset The cell address
     * @returns {Promise<void>} A promise
     */
    setCells(values: (string | number)[][], offset: string): Promise<void>;
    /**
     * @public
     * @function getCells
     * @description Gets cell values from the range specified
     * @param {string} start The start cell address
     * @param {number} offsetWidth The number of columns in the openfin app
     * @param {number} offsetHeight The number of rows in the openfin app
     * @returns {Promise<(string | number)[][]>} A promise containing the cells
     */
    getCells(start: string, offsetWidth: number, offsetHeight: number): Promise<(string | number)[][]>;
    /**
     * @public
     * @function activateRow
     * @description This mirrors the row selected in the openfin
     * application to Excel
     * @param {string} cellAddress THe address of the first cell of the row
     * @returns {Promise<void>} A promise
     */
    activateRow(cellAddress: string): Promise<void>;
    /**
     * @public
     * @function insertRow
     * @description This inserts a row just before the selected row
     * @param {number} rowNumber The address of the first cell in the row
     * @returns {Promise<void>} A promise
     */
    insertRow(rowNumber: number): Promise<void>;
    /**
     * @public
     * @function deleteRow
     * @description This deletes the selected row
     * @param {number} rowNumber The address of the first cell in the row
     * @returns {Promise<any>} A promise
     */
    deleteRow(rowNumber: number): Promise<void>;
    /**
     * @public
     * @function
     * @description activate Activates the current worksheet
     * @returns {Promise<any>} A promise
     */
    activate(): Promise<void>;
    /**
     * @public
     * @function
     * @description activateCell Activates the selected cell
     * @param {string} cellAddress The address of the cell
     * @returns {Promise<void>} A promise
     */
    activateCell(cellAddress: string): Promise<void>;
    /**
     * @public
     * @function clearRange
     * @description Clear the range of formatting and content
     * @param {string} rangeCode The range selected
     * @returns {Promise<void>} A promise
     */
    clearRange(rangeCode: string): Promise<void>;
    /**
     * @public
     * @function clearRangeContents
     * @description Clears the contents in the specified range
     * @param {string} rangeCode The selected range
     * @returns {Promise<void>} A promise
     */
    clearRangeContents(rangeCode: string): Promise<void>;
    /**
     * @public
     * @function clearAllCells
     * @description Clear all cells and their formatting
     * @returns {Promise<any>} A promise
     */
    clearAllCells(): Promise<void>;
    /**
     * @public
     * @function clearAllCellContents
     * @description Clears all the cells content
     * @returns {Promise<void>} A promise
     */
    clearAllCellContents(): Promise<void>;
    /**
     * @public
     * @function setCellName
     * @description Sets a name for the cell address
     * @param {string} cellAddress The address of the cell e.g. A1
     * @param {string} cellName The name of the cell
     * @returns {Promise<any>} A promise
     */
    setCellName(cellAddress: string, cellName: string): Promise<void>;
    /**
     * @public
     * @function calculate
     * @description Calculates all formula on teh sheet
     * @returns {Promise<void>} A promise
     */
    calculate(): Promise<void>;
    /**
     * @public
     * @function getCellByName
     * @description Gets a cell by its name
     * @param {string} cellName The name of the cell
     * @returns {Promise<any>} A promise
     */
    getCellByName(cellName: string): Promise<Cell>;
    /**
     * @public
     * @function protect
     * @description Password protects the sheet
     * @param {string} password Password used to protect the sheet
     * @returns {Promise<any>} A promise
     */
    protect(password: string): Promise<void>;
    /**
     * @public
     * @function
     * @description toObject Returns only the functions that should be exposed by
     * this class
     * @returns {Worksheet} Public methods in ExcelWorksheet
     */
    toObject(): Worksheet;
}
