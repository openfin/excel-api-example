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
exports.ExcelRtd2 = exports.ExcelRtd = void 0;
const EventEmitter_1 = require("./EventEmitter");
class ExcelRtd extends EventEmitter_1.EventEmitter {
    constructor(providerName) {
        super();
        this.listeners = {};
        this.providerName = providerName;
    }
    static create(providerName) {
        return __awaiter(this, void 0, void 0, function* () {
            const instance = new ExcelRtd(providerName);
            yield instance.init();
            return instance;
        });
    }
    init() {
        return __awaiter(this, void 0, void 0, function* () {
            this.provider = yield fin.InterApplicationBus.Channel.create(`excelRtd/${this.providerName}`);
            this.provider.register('topicRegistered', topic => this.onTopicRegistered(topic));
            this.provider.register('topicUnregistered', topic => this.onTopicUnregistered(topic));
            fin.InterApplicationBus.subscribe({ uuid: '*' }, `excelRtd/${this.providerName}`, () => { });
            this.initialized = true;
        });
    }
    onTopicRegistered(topic) {
        this.dispatchEvent('connected', { topic });
    }
    onTopicUnregistered(topic) {
        this.dispatchEvent('disconnected', { topic });
    }
    setValue(topic, value) {
        //should be:
        //this.provider.publish(topic, value);
        fin.InterApplicationBus.publish(`excelRtd/${this.providerName}/${topic}`, value);
    }
    toObject() {
        return this;
    }
}
exports.ExcelRtd = ExcelRtd;
class ExcelRtd2 extends EventEmitter_1.EventEmitter {
    constructor(providerName) {
        super();
        this.listeners = {};
        this.providerName = providerName;
    }
    static create(providerName) {
        return __awaiter(this, void 0, void 0, function* () {
            const instance = new ExcelRtd2(providerName);
            yield instance.init();
            return instance;
        });
    }
    init() {
        return __awaiter(this, void 0, void 0, function* () {
            this.provider = yield fin.InterApplicationBus.Channel.create(`excelRtd/${this.providerName}`);
            fin.desktop.InterApplicationBus.addSubscribeListener((_, topic) => this.onSubscribe(topic));
            fin.desktop.InterApplicationBus.addUnsubscribeListener((_, topic) => this.onUnsubscribe(topic));
            yield fin.InterApplicationBus.subscribe({ uuid: '*' }, `excelRtd/pong/${this.providerName}`, rtdTopic => this.onSubscribe(`excelRtd/data/${this.providerName}/${rtdTopic}`));
            yield fin.InterApplicationBus.publish(`excelRtd/ping/${this.providerName}`, true);
        });
    }
    setValue(topic, value) {
        fin.InterApplicationBus.publish(`excelRtd/data/${this.providerName}/${topic}`, value);
    }
    onSubscribe(topic) {
        let match = topic.match(`excelRtd/data/${this.providerName}/(.+)`);
        if (match && match.length === 2) {
            let rtdTopic = match[1];
            this.dispatchEvent('connected', { topic: rtdTopic });
        }
    }
    onUnsubscribe(topic) {
        let match = topic.match(`excelRtd/data/${this.providerName}/(.+)`);
        if (match && match.length === 2) {
            let rtdTopic = match[1];
            this.dispatchEvent('disconnected', { topic: rtdTopic });
        }
    }
    toObject() {
        return this;
    }
}
exports.ExcelRtd2 = ExcelRtd2;
//# sourceMappingURL=ExcelRtd.js.map