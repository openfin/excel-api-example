"use strict";
var __extends = (this && this.__extends) || function (d, b) {
    for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];
    function __() { this.constructor = d; }
    d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
};
var RpcDispatcher_1 = require('./RpcDispatcher');
var ExcelWorksheet = (function (_super) {
    __extends(ExcelWorksheet, _super);
    function ExcelWorksheet(name, workbook) {
        _super.call(this);
        this.name = name;
        this.workbook = workbook;
    }
    ExcelWorksheet.prototype.getDefaultMessage = function () {
        return {
            workbook: this.workbook.name,
            worksheet: this.name
        };
    };
    ExcelWorksheet.prototype.setCells = function (values, offset) {
        if (!offset)
            offset = "A1";
        this.invokeRemote("setCells", { offset: offset, values: values });
    };
    ExcelWorksheet.prototype.getCells = function (start, offsetWidth, offsetHeight, callback) {
        this.invokeRemote("getCells", { start: start, offsetWidth: offsetWidth, offsetHeight: offsetHeight }, callback);
    };
    ExcelWorksheet.prototype.getRow = function (start, width, callback) {
        this.invokeRemote("getCellsRow", { start: start, offsetWidth: width }, callback);
    };
    ExcelWorksheet.prototype.getColumn = function (start, offsetHeight, callback) {
        this.invokeRemote("getCellsColumn", { start: start, offsetHeight: offsetHeight }, callback);
    };
    ExcelWorksheet.prototype.activate = function () {
        this.invokeRemote("activateSheet");
    };
    ExcelWorksheet.prototype.activateCell = function (cellAddress) {
        this.invokeRemote("activateCell", { address: cellAddress });
    };
    ExcelWorksheet.prototype.addButton = function (name, caption, cellAddress) {
        this.invokeRemote("addButton", { address: cellAddress, buttonName: name, buttonCaption: caption });
    };
    ExcelWorksheet.prototype.setFilter = function (start, offsetWidth, offsetHeight, field, criteria1, op, criteria2, visibleDropDown) {
        this.invokeRemote("setFilter", {
            start: start,
            offsetWidth: offsetWidth,
            offsetHeight: offsetHeight,
            field: field,
            criteria1: criteria1,
            op: op,
            criteria2: criteria2,
            visibleDropDown: visibleDropDown
        });
    };
    ExcelWorksheet.prototype.formatRange = function (rangeCode, format, callback) {
        this.invokeRemote("formatRange", { rangeCode: rangeCode, format: format }, callback);
    };
    ExcelWorksheet.prototype.clearRange = function (rangeCode, callback) {
        this.invokeRemote("clearRange", { rangeCode: rangeCode }, callback);
    };
    ExcelWorksheet.prototype.clearRangeContents = function (rangeCode, callback) {
        this.invokeRemote("clearRangeContents", { rangeCode: rangeCode }, callback);
    };
    ExcelWorksheet.prototype.clearRangeFormats = function (rangeCode, callback) {
        this.invokeRemote("clearRangeFormats", { rangeCode: rangeCode }, callback);
    };
    ExcelWorksheet.prototype.clearAllCells = function (callback) {
        this.invokeRemote("clearAllCells", null, callback);
    };
    ExcelWorksheet.prototype.clearAllCellContents = function (callback) {
        this.invokeRemote("clearAllCellContents", null, callback);
    };
    ExcelWorksheet.prototype.clearAllCellFormats = function (callback) {
        this.invokeRemote("clearAllCellFormats", null, callback);
    };
    ExcelWorksheet.prototype.setCellName = function (cellAddress, cellName) {
        this.invokeRemote("setCellName", { address: cellAddress, cellName: cellName });
    };
    ExcelWorksheet.prototype.calculate = function () {
        this.invokeRemote("calculateSheet");
    };
    ExcelWorksheet.prototype.getCellByName = function (cellName, callback) {
        this.invokeRemote("getCellByName", { cellName: cellName }, callback);
    };
    ExcelWorksheet.prototype.protect = function (password) {
        this.invokeRemote("protectSheet", { password: password ? password : null });
    };
    ExcelWorksheet.prototype.toObject = function () {
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
    };
    return ExcelWorksheet;
}(RpcDispatcher_1.RpcDispatcher));
exports.ExcelWorksheet = ExcelWorksheet;
//# sourceMappingURL=ExcelWorksheet.js.map