"use strict";
const RpcDispatcher_1 = require("./RpcDispatcher");
class ExcelWorksheet extends RpcDispatcher_1.RpcDispatcher {
    constructor(name, workbook) {
        super();
        this.name = name;
        this.workbook = workbook;
    }
    getDefaultMessage() {
        return {
            workbook: this.workbook.name,
            worksheet: this.name
        };
    }
    setCells(values, offset) {
        if (!offset)
            offset = "A1";
        this.invokeExcelCall("setCells", { offset: offset, values: values });
    }
    getCells(start, offsetWidth, offsetHeight, callback) {
        this.invokeExcelCall("getCells", { start: start, offsetWidth: offsetWidth, offsetHeight: offsetHeight }, callback);
    }
    getRow(start, width, callback) {
        this.invokeExcelCall("getCellsRow", { start: start, offsetWidth: width }, callback);
    }
    getColumn(start, offsetHeight, callback) {
        this.invokeExcelCall("getCellsColumn", { start: start, offsetHeight: offsetHeight }, callback);
    }
    activate() {
        this.invokeExcelCall("activateSheet");
    }
    activateCell(cellAddress) {
        this.invokeExcelCall("activateCell", { address: cellAddress });
    }
    addButton(name, caption, cellAddress) {
        this.invokeExcelCall("addButton", { address: cellAddress, buttonName: name, buttonCaption: caption });
    }
    setFilter(start, offsetWidth, offsetHeight, field, criteria1, op, criteria2, visibleDropDown) {
        this.invokeExcelCall("setFilter", {
            start: start,
            offsetWidth: offsetWidth,
            offsetHeight: offsetHeight,
            field: field,
            criteria1: criteria1,
            op: op,
            criteria2: criteria2,
            visibleDropDown: visibleDropDown
        });
    }
    formatRange(rangeCode, format, callback) {
        this.invokeExcelCall("formatRange", { rangeCode: rangeCode, format: format }, callback);
    }
    clearRange(rangeCode, callback) {
        this.invokeExcelCall("clearRange", { rangeCode: rangeCode }, callback);
    }
    clearRangeContents(rangeCode, callback) {
        this.invokeExcelCall("clearRangeContents", { rangeCode: rangeCode }, callback);
    }
    clearRangeFormats(rangeCode, callback) {
        this.invokeExcelCall("clearRangeFormats", { rangeCode: rangeCode }, callback);
    }
    clearAllCells(callback) {
        this.invokeExcelCall("clearAllCells", null, callback);
    }
    clearAllCellContents(callback) {
        this.invokeExcelCall("clearAllCellContents", null, callback);
    }
    clearAllCellFormats(callback) {
        this.invokeExcelCall("clearAllCellFormats", null, callback);
    }
    setCellName(cellAddress, cellName) {
        this.invokeExcelCall("setCellName", { address: cellAddress, cellName: cellName });
    }
    calculate() {
        this.invokeExcelCall("calculateSheet");
    }
    getCellByName(cellName, callback) {
        this.invokeExcelCall("getCellByName", { cellName: cellName }, callback);
    }
    protect(password) {
        this.invokeExcelCall("protectSheet", { password: password ? password : null });
    }
    toObject() {
        return {
            addEventListener: this.addEventListener.bind(this),
            dispatchEvent: this.dispatchEvent.bind(this),
            removeEventListener: this.removeEventListener.bind(this),
            name: this.name,
            activate: this.activate.bind(this),
            activateCell: this.activateCell.bind(this),
            addButton: this.addButton.bind(this),
            calculate: this.calculate.bind(this),
            clearAllCellContents: this.clearAllCellContents.bind(this),
            clearAllCellFormats: this.clearAllCellFormats.bind(this),
            clearAllCells: this.clearAllCells.bind(this),
            clearRange: this.clearRange.bind(this),
            clearRangeContents: this.clearRangeContents.bind(this),
            clearRangeFormats: this.clearRangeFormats.bind(this),
            formatRange: this.formatRange.bind(this),
            getCellByName: this.getCellByName.bind(this),
            getCells: this.getCells.bind(this),
            getColumn: this.getColumn.bind(this),
            getRow: this.getRow.bind(this),
            protect: this.protect.bind(this),
            setCellName: this.setCellName.bind(this),
            setCells: this.setCells.bind(this),
            setFilter: this.setFilter.bind(this)
        };
    }
}
exports.ExcelWorksheet = ExcelWorksheet;
//# sourceMappingURL=ExcelWorksheet.js.map