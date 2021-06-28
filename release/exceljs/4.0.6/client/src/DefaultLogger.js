"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.DefaultLogger = void 0;
class DefaultLogger {
    constructor(name) {
        this.name = name || "logger";
    }
    trace(message, ...args) {
        console.log(this.name + ": " + message, ...args);
    }
    debug(message, ...args) {
        console.log(this.name + ": " + message, ...args);
    }
    info(message, ...args) {
        console.info(this.name + ": " + message, ...args);
    }
    warn(message, ...args) {
        console.warn(this.name + ": " + message, ...args);
    }
    error(message, error, ...args) {
        console.error(this.name + ": " + message, error, ...args);
    }
    fatal(message, error, ...args) {
        console.error(this.name + ": " + message, error, ...args);
    }
}
exports.DefaultLogger = DefaultLogger;
//# sourceMappingURL=DefaultLogger.js.map