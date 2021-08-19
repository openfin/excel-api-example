"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.ExcelRtd = void 0;
const EventEmitter_1 = require("./EventEmitter");
class ExcelRtd extends EventEmitter_1.EventEmitter {
    constructor(providerName, logger, heartbeatIntervalInMilliseconds = 10000) {
        super();
        this.heartbeatIntervalInMilliseconds = heartbeatIntervalInMilliseconds;
        this.listeners = {};
        this.connectedTopics = {};
        this.connectedKey = 'connected';
        this.disconnectedKey = 'disconnected';
        this.loggerName = "ExcelRtd";
        this.initialized = false;
        this.disposed = false;
        var minimumDefaultHeartbeat = 10000;
        if (this.heartbeatIntervalInMilliseconds < minimumDefaultHeartbeat) {
            logger.warn(`heartbeatIntervalInMilliseconds cannot be less than ${minimumDefaultHeartbeat}. Setting heartbeatIntervalInMilliseconds to ${minimumDefaultHeartbeat}.`);
            this.heartbeatIntervalInMilliseconds = minimumDefaultHeartbeat;
        }
        this.providerName = providerName;
        this.logger = logger;
        logger.debug(this.loggerName + ": instance created for provider: " + providerName);
    }
    static create(providerName, logger, heartbeatIntervalInMilliseconds = 10000) {
        return __awaiter(this, void 0, void 0, function* () {
            logger.debug("ExcelRtd: create called to create provider: " + providerName);
            const instance = new ExcelRtd(providerName, logger, heartbeatIntervalInMilliseconds);
            yield instance.init();
            if (!instance.isInitialized) {
                return undefined;
            }
            return instance;
        });
    }
    init() {
        return __awaiter(this, void 0, void 0, function* () {
            if (this.isInitialized) {
                return;
            }
            this.logger.debug(this.loggerName + ": Initialise called for provider: " + this.providerName);
            try {
                // A channel is created to ensure it is a singleton so you don't have two apps pushing updates over each other or two windows within the same app
                this.provider = yield fin.InterApplicationBus.Channel.create(`excelRtd/${this.providerName}`);
            }
            catch (err) {
                this.logger.warn(this.loggerName + `: The excelRtd/${this.providerName} channel already exists. You can only have one instance of a connection for a provider to avoid confusion. It may be you have multiple instances or another window or application has created a provider with the same name.`, err);
                return;
            }
            this.logger.debug(this.loggerName + `: Subscribing to messages to this provider (${this.providerName}) from excel.`);
            yield fin.InterApplicationBus.subscribe({ uuid: '*' }, `excelRtd/pong/${this.providerName}`, this.onSubscribe.bind(this));
            yield fin.InterApplicationBus.subscribe({ uuid: '*' }, `excelRtd/ping-request/${this.providerName}`, this.ping.bind(this));
            yield fin.InterApplicationBus.subscribe({ uuid: '*' }, `excelRtd/unsubscribed/${this.providerName}`, this.onUnsubscribe.bind(this));
            yield this.ping();
            this.establishHeartbeat();
            this.logger.debug(this.loggerName + `: initialisation for provider (${this.providerName}) finished.`);
            this.initialized = true;
        });
    }
    get isDisposed() {
        return this.disposed;
    }
    get isInitialized() {
        return this.initialized;
    }
    setValue(topic, value) {
        this.logger.trace(this.loggerName + `: Publishing on rtdTopic: ${topic} and provider: ${this.providerName} value: ${JSON.stringify(value)}`);
        fin.InterApplicationBus.publish(`excelRtd/data/${this.providerName}/${topic}`, value);
    }
    dispose() {
        return __awaiter(this, void 0, void 0, function* () {
            if (!this.disposed) {
                this.logger.debug(this.loggerName + `: dispose called. Will send message to clear values for this provider (${this.providerName}).`);
                if (this.heartbeatToken) {
                    clearInterval(this.heartbeatToken);
                }
                this.clear();
                if (this.provider !== undefined) {
                    try {
                        yield this.provider.destroy();
                    }
                    catch (err) {
                        // without a catch the rest of the initialisation would be broken
                        this.logger.warn(this.loggerName + `: The excelRtd/${this.providerName} channel could not be destroyed during cleanup.`, err);
                    }
                }
                this.logger.debug(this.loggerName + `: UnSubscribing to messages to this provider (${this.providerName}) from excel.`);
                yield fin.InterApplicationBus.unsubscribe({ uuid: '*' }, `excelRtd/pong/${this.providerName}`, this.onSubscribe.bind(this));
                yield fin.InterApplicationBus.unsubscribe({ uuid: '*' }, `excelRtd/ping-request/${this.providerName}`, this.ping.bind(this));
                yield fin.InterApplicationBus.subscribe({ uuid: '*' }, `excelRtd/unsubscribed/${this.providerName}`, this.onUnsubscribe.bind(this));
                this.disposed = true;
                this.initialized = false;
            }
            else {
                this.logger.debug(this.loggerName + `: This provider (${this.providerName}) has already been disposed.`);
            }
        });
    }
    // Overriding
    addEventListener(type, listener) {
        this.logger.debug(this.loggerName + `: Event listener add requested for type ${type} received.`);
        if (super.hasEventListener(type, listener)) {
            this.logger.debug(this.loggerName + `: Event listener add requested for type ${type} received.`);
            return;
        }
        let connectedTopicIds = Object.keys(this.connectedTopics);
        let topics = this.connectedTopics;
        if (connectedTopicIds.length > 0) {
            // need to simulate async action as by default this method would return and then a listener would be called
            setTimeout(() => {
                connectedTopicIds.forEach(id => {
                    this.logger.debug(this.loggerName + `: Raising synthetic event as the event listener was added after the event for connected for rtdTopic: ${id}.`);
                    listener(topics[id]);
                });
            }, 0);
        }
        super.addEventListener(type, listener);
    }
    dispatchEvent(evtOrTypeArg, data) {
        var event;
        if (typeof evtOrTypeArg == "string" && data !== undefined) {
            this.logger.debug(this.loggerName + `: dispatch event called for type ${evtOrTypeArg} and data: ${JSON.stringify(data)}`);
            event = Object.assign({
                target: this.toObject(),
                type: evtOrTypeArg,
                defaultPrevented: false
            }, data);
            if (data.topic !== undefined) {
                if (evtOrTypeArg === this.connectedKey) {
                    this.connectedTopics[data.topic] = event;
                    this.logger.debug(this.loggerName + `: Saving connected event for rtdTopic: ${data.topic}.`);
                }
                else if (evtOrTypeArg === this.disconnectedKey) {
                    this.logger.debug(this.loggerName + `: Disconnected event for rtdTopic: ${data.topic} received.`);
                    if (this.connectedTopics[data.topic] !== undefined) {
                        // we have removed the topic so clear it from the connected list for late subscribers
                        this.logger.debug(this.loggerName + `: Clearing saved connected event for rtdTopic: ${data.topic}.`);
                        delete this.connectedTopics[data.topic];
                    }
                }
            }
            this.logger.debug(this.loggerName + `: Dispatching event.`);
            return super.dispatchEvent(event, data);
        }
        event = evtOrTypeArg;
        return super.dispatchEvent(event);
    }
    toObject() {
        return this;
    }
    ping(topic) {
        return __awaiter(this, void 0, void 0, function* () {
            if (topic !== undefined) {
                this.pingPath = `excelRtd/ping/${this.providerName}/${topic}`;
            }
            else {
                this.pingPath = `excelRtd/ping/${this.providerName}`;
            }
            this.logger.debug(this.loggerName + `: Publishing ping message for this provider (${this.providerName}) to excel on topic: ${this.pingPath}.`);
            yield fin.InterApplicationBus.publish(`${this.pingPath}`, true);
        });
    }
    establishHeartbeat() {
        this.heartbeatPath = `excelRtd/heartbeat/${this.providerName}`;
        this.heartbeatToken = setInterval(() => {
            this.logger.debug(`Heartbeating for ${this.heartbeatPath}.`);
            fin.InterApplicationBus.publish(`${this.heartbeatPath}`, this.heartbeatIntervalInMilliseconds);
        }, this.heartbeatIntervalInMilliseconds);
    }
    onSubscribe(topic) {
        this.logger.debug(this.loggerName + `: Subscription for rtdTopic ${topic} found. Dispatching connected event for rtdTopic.`);
        this.dispatchEvent(this.connectedKey, { topic });
    }
    onUnsubscribe(topic) {
        this.logger.debug(this.loggerName + `: Unsubscribe for rtdTopic ${topic}. Dispatching disconnected event for rtdTopic.`);
        this.dispatchEvent(this.disconnectedKey, { topic });
    }
    clear() {
        let path = `excelRtd/clear/${this.providerName}`;
        this.logger.debug(this.loggerName + `: Clear called. Publishing to excel on topic: ${path} `);
        fin.InterApplicationBus.publish(`excelRtd/clear/${this.providerName}`, true);
    }
}
exports.ExcelRtd = ExcelRtd;
//# sourceMappingURL=ExcelRtd.js.map