"use strict";
const RpcDispatcher_1 = require('./RpcDispatcher');
class ExcelWorkbook extends RpcDispatcher_1.RpcDispatcher {
    constructor(application, name) {
        super();
        this.connectionUuid = application.connectionUuid;
        this.application = application;
        this.name = name;
    }
    getDefaultMessage() {
        return {
            workbook: this.name
        };
    }
    getWorksheets(callback) {
        this.invokeExcelCall("getWorksheets", null, callback);
    }
    getWorksheetByName(name) {
        return this.application.getWorksheetByName(this.name, name);
    }
    addWorksheet(callback) {
        this.invokeExcelCall("addSheet", null, callback);
    }
    activate() {
        this.invokeExcelCall("activateWorkbook");
    }
    save() {
        this.invokeExcelCall("saveWorkbook");
    }
    close() {
        this.invokeExcelCall("closeWorkbook");
    }
    toObject() {
        return {
            addEventListener: this.addEventListener.bind(this),
            dispatchEvent: this.dispatchEvent.bind(this),
            removeEventListener: this.removeEventListener.bind(this),
            name: this.name,
            activate: this.activate.bind(this),
            addWorksheet: this.addWorksheet.bind(this),
            close: this.close.bind(this),
            getWorksheetByName: name => this.getWorksheetByName(name).toObject(),
            getWorksheets: this.getWorksheets.bind(this),
            save: this.save.bind(this)
        };
    }
}
exports.ExcelWorkbook = ExcelWorkbook;
//# sourceMappingURL=ExcelWorkbook.js.map