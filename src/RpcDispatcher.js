"use strict";
var RpcDispatcher = (function () {
    function RpcDispatcher() {
        this.listeners = {};
    }
    RpcDispatcher.prototype.addEventListener = function (type, listener) {
        if (this.hasEventListener(type, listener)) {
            return;
        }
        if (!this.listeners[type]) {
            this.listeners[type] = [];
        }
        this.listeners[type].push(listener);
    };
    RpcDispatcher.prototype.removeEventListener = function (type, listener) {
        if (!this.hasEventListener(type, listener)) {
            return;
        }
        var callbacksOfType = this.listeners[type];
        callbacksOfType.splice(callbacksOfType.indexOf(listener), 1);
    };
    RpcDispatcher.prototype.hasEventListener = function (type, listener) {
        if (!this.listeners[type]) {
            return false;
        }
        if (!listener) {
            return true;
        }
        return (this.listeners[type].indexOf(listener) >= 0);
    };
    RpcDispatcher.prototype.dispatchEvent = function (event) {
        event.target = this;
        if (!this.listeners[event.type]) {
            return false;
        }
        var callbacks = this.listeners[event.type];
        for (var i = 0; i < callbacks.length; i++) {
            callbacks[i](event);
        }
        return event.defaultPrevented;
    };
    RpcDispatcher.prototype.getDefaultMessage = function () {
        return {};
    };
    RpcDispatcher.prototype.invokeRemote = function (functionName, data, callback) {
        var message = this.getDefaultMessage();
        message.messageId = RpcDispatcher.messageId;
        message.action = functionName;
        if (data) {
            for (var key in data) {
                message[key] = data[key];
            }
        }
        if (callback) {
            RpcDispatcher.callbacks[RpcDispatcher.messageId] = callback;
        }
        fin.desktop.InterApplicationBus.publish("excelCall", message);
        RpcDispatcher.messageId++;
    };
    return RpcDispatcher;
}());
RpcDispatcher.messageId = 1;
RpcDispatcher.callbacks = {};
exports.RpcDispatcher = RpcDispatcher;
//# sourceMappingURL=RpcDispatcher.js.map