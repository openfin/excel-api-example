"use strict";
var __extends = (this && this.__extends) || function (d, b) {
    for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];
    function __() { this.constructor = d; }
    d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
};
var RpcDispatcher_1 = require('./RpcDispatcher');
var ExcelWorkbook = (function (_super) {
    __extends(ExcelWorkbook, _super);
    function ExcelWorkbook(application, name) {
        _super.call(this);
        this.application = application;
        this.name = name;
    }
    ExcelWorkbook.prototype.getDefaultMessage = function () {
        return {
            workbook: this.name
        };
    };
    ExcelWorkbook.prototype.getWorksheets = function (callback) {
        this.invokeRemote("getWorksheets", null, callback);
    };
    ExcelWorkbook.prototype.getWorksheetByName = function (name) {
        return this.application.getWorksheetByName(this.name, name);
    };
    ExcelWorkbook.prototype.addWorksheet = function (callback) {
        this.invokeRemote("addSheet", null, callback);
    };
    ExcelWorkbook.prototype.activate = function () {
        this.invokeRemote("activateWorkbook");
    };
    ExcelWorkbook.prototype.save = function () {
        this.invokeRemote("saveWorkbook");
    };
    ExcelWorkbook.prototype.close = function () {
        this.invokeRemote("closeWorkbook");
    };
    ExcelWorkbook.prototype.toObject = function () {
        var _this = this;
        return {
            addEventListener: this.addEventListener.bind(this),
            dispatchEvent: this.dispatchEvent.bind(this),
            removeEventListener: this.removeEventListener.bind(this),
            name: this.name,
            activate: this.activate.bind(this),
            addWorksheet: this.addWorksheet.bind(this),
            close: this.close.bind(this),
            getWorksheetByName: function (name) { return _this.getWorksheetByName(name).toObject(); },
            getWorksheets: this.getWorksheets.bind(this),
            save: this.save.bind(this)
        };
    };
    return ExcelWorkbook;
}(RpcDispatcher_1.RpcDispatcher));
exports.ExcelWorkbook = ExcelWorkbook;
//# sourceMappingURL=ExcelWorkbook.js.map