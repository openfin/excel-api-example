"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.RpcDispatcher = void 0;
const EventEmitter_1 = require("./EventEmitter");
class RpcDispatcher extends EventEmitter_1.EventEmitter {
    getDefaultMessage() {
        return {};
    }
    invokeExcelCall(functionName, data, callback) {
        return this.invokeRemoteCall('excelCall', functionName, data, callback);
    }
    invokeServiceCall(functionName, data, callback) {
        return this.invokeRemoteCall('excelServiceCall', functionName, data, callback);
    }
    invokeRemoteCall(topic, functionName, data, callback) {
        var message = this.getDefaultMessage();
        var args = data || {};
        var invoker = this;
        Object.assign(message, {
            messageId: RpcDispatcher.messageId,
            target: {
                connectionUuid: this.connectionUuid,
                workbookName: invoker.workbookName || (invoker.workbook && invoker.workbook.workbookName) || args.workbookName || args.workbook,
                worksheetName: invoker.worksheetName || args.worksheetName || args.worksheet,
                rangeCode: args.rangeCode
            },
            action: functionName,
            data: data
        });
        var executor;
        var promise = new Promise((resolve, reject) => {
            executor = {
                resolve,
                reject
            };
        });
        // Legacy Callback-style API
        promise = this.applyCallbackToPromise(promise, callback);
        var currentMessageId = RpcDispatcher.messageId;
        RpcDispatcher.messageId++;
        if (this.connectionUuid !== undefined) {
            RpcDispatcher.promiseExecutors[currentMessageId] = executor;
            fin.desktop.InterApplicationBus.send(this.connectionUuid, topic, message, ack => {
                // TODO: log once we support configurable logging.
            }, nak => {
                delete RpcDispatcher.promiseExecutors[currentMessageId];
                executor.reject(new Error(nak));
            });
        }
        else {
            executor.reject(new Error('The target UUID of the remote call is undefined.'));
        }
        return promise;
    }
    applyCallbackToPromise(promise, callback) {
        if (callback) {
            promise
                .then(result => {
                callback(result);
                return result;
            }).catch(err => {
                console.error(err);
            });
            promise = undefined;
        }
        return promise;
    }
}
exports.RpcDispatcher = RpcDispatcher;
RpcDispatcher.messageId = 1;
RpcDispatcher.promiseExecutors = {};
//# sourceMappingURL=RpcDispatcher.js.map