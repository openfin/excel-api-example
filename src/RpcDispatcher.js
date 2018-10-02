"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
/**
 * @abstract
 * @class
 * @description Top level class that communicates with the Excel application
 */
class RpcDispatcher {
    constructor() {
        /**
         * @public
         * @description The connectionUuid of the excel application
         */
        this.connectionUuid = '';
        /**
         * @private
         * @description Holds event listeners
         */
        this.listeners = {};
    }
    /**
     * @public
     * @function addEventListener Adds event listener to listen to events coming
     * from Excel application
     * @param {string} type The type of the event to listen to
     * @param {EventListenerOrEventListenerObject} listener The method to execute when the event has been fired
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
     * @function removeEventListener
     * @description Removes the event from the local store
     * @param {string} type The type of the event to listen to
     * @param {EventListenerOrEventListenerObject} listener The method to execute when the event has been fired
     */
    removeEventListener(type, listener) {
        if (!this.hasEventListener(type, listener)) {
            return;
        }
        const callbacksOfType = this.listeners[type];
        callbacksOfType.splice(callbacksOfType.indexOf(listener), 1);
    }
    /**
     * @private
     * @function hasEventListener
     * @description Check whether an event listener has been
     * registered
     * @param {string} type The type of the event
     * @param {EventListenerOrEventListenerObject} listener The method to execute when the event has been fired
     * @returns {boolean} True or false depending on if the listener exists
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
     * @function dispatchEvent
     * @description Sends event over to the correct entity e.g.
     * Workbook, worksheet
     * @param {string|Event} evtOrTypeArg Pass either an event or event type as a string
     * @param {T} data The data to be passed to the receiving entity
     * @returns {boolean} Whether or not the events default behaviour has been prevented
     */
    dispatchEvent(evtOrTypeArg, data) {
        let event;
        if (typeof evtOrTypeArg === 'string') {
            event =
                Object.assign({
                    target: this.toObject(),
                    type: evtOrTypeArg,
                    defaultPrevented: false,
                }, data);
        }
        else {
            event = evtOrTypeArg;
        }
        const callbacks = this.listeners[event.type] || [];
        callbacks.forEach((callback) => {
            callback(event);
        });
        return event.defaultPrevented;
    }
    /**
     * @private
     * @function getDefaultMessage
     * @description Get the default message when invoking a remote
     * call
     * @returns {object} Returns an empty object to be populated
     */
    getDefaultMessage() {
        return {};
    }
    /**
     * @protected
     * @function invokeExcelCall
     * @description Invokes a call in excel application via RPC
     * @param {string} functionName The name of the function to invoke
     * @param {RemoteData?} data Any data to be sent over as part of the invocation
     * @returns {Promise<T>} A Promise with generic data depending on which function calls it
     */
    invokeExcelCall(functionName, data) {
        return this.invokeRemoteCall('excelCall', functionName, data);
    }
    /**
     * @protected
     * @function invokeServiceCall Invokes a call in the excel service process via
     * RPC
     * @param {string} functionName The name of the function to invoke
     * @param {ExcelData|null} data Any data to be sent over as part of the invocation
     * @returns {Promise<T>} A Promise with generic data depending on which function calls it
     */
    invokeServiceCall(functionName, data) {
        return this.invokeRemoteCall('excelServiceCall', functionName, data);
    }
    /**
     * @private
     * @function invokeRemoteCall
     * @description Invokes a remote procedure call
     * @param {string} topic Topic to send on
     * @param {string} functionName The name of the function to invoke
     * @param {RemoteData?} data The data to be sent over as part of the invocation
     * @returns {Promise<T>} A Promise with generic data depending on which function calls it
     */
    invokeRemoteCall(topic, functionName, data) {
        const message = this.getDefaultMessage();
        const args = data || {};
        const invoker = this;
        const workbook = (invoker.workbook) ||
            (invoker) || null;
        const worksheet = (invoker) || null;
        Object.assign(message, {
            messageId: RpcDispatcher.messageId,
            target: {
                connectionUuid: this.connectionUuid,
                workbookName: workbook.name ? workbook.name : null,
                worksheetName: worksheet.name ? worksheet.name : null,
                rangeCode: args.rangeCode
            },
            action: functionName,
            data
        });
        let executor;
        const promise = new Promise((resolve, reject) => {
            executor = { resolve, reject };
        });
        const currentMessageId = RpcDispatcher.messageId;
        RpcDispatcher.messageId++;
        if (this.connectionUuid !== undefined) {
            fin.desktop.InterApplicationBus.send(this.connectionUuid, topic, message, () => {
                RpcDispatcher.promiseExecutors[currentMessageId] = executor;
            }, (nak) => {
                executor.reject(new Error(nak));
            });
        }
        else {
            executor.reject(new Error('The target UUID of the remote call is undefined.'));
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