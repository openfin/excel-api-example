import { RpcDispatcher } from './RpcDispatcher';
import { ExcelWorkbook } from './ExcelWorkbook';
export declare class ExcelWorksheet extends RpcDispatcher {
    workbook: ExcelWorkbook;
    worksheetName: string;
    private objectInstance;
    constructor(name: string, workbook: ExcelWorkbook);
    getDefaultMessage(): any;
    setCells(values: any[][], offset: string): Promise<any>;
    getCells(start: string, offsetWidth: number, offsetHeight: number, callback: Function): Promise<any>;
    getRow(start: string, width: number, callback: Function): Promise<any>;
    /**
     * @function activateRow This mirrors the row selected in the openfin application to Excel
     * @param {string} cellAddress THe address of the first cell of the row
     */
    activateRow(cellAddress: string): Promise<void>;
    getColumn(start: string, offsetHeight: number, callback: Function): Promise<any>;
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
    deleteRow(rowNumber: string): Promise<void>;
    activate(): Promise<any>;
    activateCell(cellAddress: string): Promise<any>;
    addButton(name: string, caption: string, cellAddress: string): Promise<any>;
    setFilter(start: string, offsetWidth: number, offsetHeight: number, field: number, criteria1: string, op: string, criteria2: string, visibleDropDown: string): Promise<any>;
    formatRange(rangeCode: string, format: any, callback: Function): Promise<any>;
    clearRange(rangeCode: string, callback: Function): Promise<any>;
    clearRangeContents(rangeCode: string, callback: Function): Promise<any>;
    clearRangeFormats(rangeCode: string, callback: Function): Promise<any>;
    clearAllCells(callback: Function): Promise<any>;
    clearAllCellContents(callback: Function): Promise<any>;
    clearAllCellFormats(callback: Function): Promise<any>;
    setCellName(cellAddress: string, cellName: string): Promise<any>;
    calculate(): Promise<any>;
    getCellByName(cellName: string, callback: Function): Promise<any>;
    protect(password: string): Promise<any>;
    toObject(): any;
}
