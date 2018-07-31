"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
/**
 * @abstract
 * @class Top level class that communicates with the Excel application
 */
class RpcDispatcher {
    constructor() {
        /**
         * @private
         * @description Holds event listeners
         */
        this.listeners = {};
    }
    /**
     * @public
     * @function addEventListener Adds event listener to listen to events coming from Excel application
     * @param type The type of the event to listen to
     * @param listener The method to execute when the event has been fired
     */
    addEventListener(type, listener) {
        if (this.hasEventListener(type, listener)) {
            return;
        }
        if (!this.listeners[type]) {
            this.listeners[type] = [];
        }
        this.listeners[type].push(listener);
    }
    /**
     * @public
     * @function removeEventListener Removes the event from the local store
     * @param type The type of the event to listen to
     * @param listener The method to execute when the event has been fired
     */
    removeEventListener(type, listener) {
        if (!this.hasEventListener(type, listener)) {
            return;
        }
        var callbacksOfType = this.listeners[type];
        callbacksOfType.splice(callbacksOfType.indexOf(listener), 1);
    }
    /**
     * @private
     * @function hasEventListener Check whether an event listener has been registered
     * @param type The type of the event
     * @param listener The method to execute when the event has been fired
     */
    hasEventListener(type, listener) {
        if (!this.listeners[type]) {
            return false;
        }
        if (!listener) {
            return true;
        }
        return (this.listeners[type].indexOf(listener) >= 0);
    }
    /**
     * @public
     * @function dispatchEvent Sends event over to the correct entity e.g. Workbook, worksheet
     * @param evtOrTypeArg Pass either an event or event type as a string
     * @param data The data to be passed to the receiving entity
     */
    dispatchEvent(evtOrTypeArg, data) {
        var event;
        if (typeof evtOrTypeArg == "string") {
            event = Object.assign({
                target: this.toObject(),
                type: evtOrTypeArg,
                defaultPrevented: false
            }, data);
        }
        else {
            event = evtOrTypeArg;
        }
        var callbacks = this.listeners[event.type] || [];
        callbacks.forEach(callback => callback(event));
        return event.defaultPrevented;
    }
    /**
     * @private
     * @function getDefaultMessage Get the default message when invoking a remote call
     * @returns {object} Returns an empty object to be populated
     */
    getDefaultMessage() {
        return {};
    }
    /**
     * @protected
     * @function invokeExcelCall Invokes a call in excel application via RPC
     * @param functionName The name of the function to invoke
     * @param data Any data to be sent over as part of the invocation
     */
    invokeExcelCall(functionName, data) {
        return this.invokeRemoteCall('excelCall', functionName, data);
    }
    /**
     * @protected
     * @function invokeServiceCall Invokes a call in the excel service process via RPC
     * @param functionName The name of the function to invoke
     * @param data Any data to be sent over as part of the invocation
     */
    invokeServiceCall(functionName, data) {
        return this.invokeRemoteCall('excelServiceCall', functionName, data);
    }
    /**
     * @private
     * @function invokeRemoteCall Invokes a remote procedure call
     * @param topic Topic to send on
     * @param functionName The name of the function to invoke
     * @param data The data to be sent over as part of the invocation
     * @param callback Callback to be applied to the promise
     */
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
            fin.desktop.InterApplicationBus.send(this.connectionUuid, topic, message, ack => {
                RpcDispatcher.promiseExecutors[currentMessageId] = executor;
            }, nak => {
                executor.reject(new Error(nak));
            });
        }
        else {
            executor.reject(new Error('The target UUID of the remote call is undefined.'));
        }
        return promise;
    }
    /**
     * @protected
     * @function applyCallbackToPromise Applies a callback to the promise
     * @param promise The promise to be acted on
     * @param callback THe callback to be applied to the promise
     * @returns {Promise<any>} A promise with the callback applied
     */
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
/**
 * @protected
 * @static
 * @description The message id of the action being sent
 */
RpcDispatcher.messageId = 1;
/**
 * @protected
 * @static
 * @description Promises to be executed
 */
RpcDispatcher.promiseExecutors = {};
exports.RpcDispatcher = RpcDispatcher;
//# sourceMappingURL=RpcDispatcher.js.map