"use strict";
class RpcDispatcher {
    constructor() {
        this.listeners = {};
    }
    addEventListener(type, listener) {
        if (this.hasEventListener(type, listener)) {
            return;
        }
        if (!this.listeners[type]) {
            this.listeners[type] = [];
        }
        this.listeners[type].push(listener);
    }
    removeEventListener(type, listener) {
        if (!this.hasEventListener(type, listener)) {
            return;
        }
        var callbacksOfType = this.listeners[type];
        callbacksOfType.splice(callbacksOfType.indexOf(listener), 1);
    }
    hasEventListener(type, listener) {
        if (!this.listeners[type]) {
            return false;
        }
        if (!listener) {
            return true;
        }
        return (this.listeners[type].indexOf(listener) >= 0);
    }
    dispatchEvent(event) {
        event.target = this;
        if (!this.listeners[event.type]) {
            return false;
        }
        var callbacks = this.listeners[event.type];
        for (var i = 0; i < callbacks.length; i++) {
            callbacks[i](event);
        }
        return event.defaultPrevented;
    }
    getDefaultMessage() {
        return {};
    }
    invokeExcelCall(functionName, data, callback) {
        this.invokeRemoteCall('excelCall', functionName, data, callback);
    }
    invokeServiceCall(functionName, data, callback) {
        this.invokeRemoteCall('excelServiceCall', functionName, data, callback);
    }
    invokeRemoteCall(topic, functionName, data, callback) {
        var message = this.getDefaultMessage();
        message.messageId = RpcDispatcher.messageId;
        message.action = functionName;
        Object.assign(message, data);
        if (callback) {
            RpcDispatcher.callbacks[RpcDispatcher.messageId] = callback;
        }
        fin.desktop.InterApplicationBus.send(this.connectionUuid, topic, message);
        RpcDispatcher.messageId++;
    }
}
RpcDispatcher.messageId = 1;
RpcDispatcher.callbacks = {};
exports.RpcDispatcher = RpcDispatcher;
//# sourceMappingURL=RpcDispatcher.js.map