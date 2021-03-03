import { EventEmitter } from './EventEmitter';
import { ILog } from './ILog';
export declare class ExcelRtd extends EventEmitter {
    providerName: string;
    provider: any;
    logger: ILog;
    listeners: {
        [eventType: string]: Function[];
    };
    connectedTopics: {};
    connectedKey: string;
    disconnectedKey: string;
    loggerName: string;
    private initialized;
    private disposed;
    static create(providerName: any, logger: ILog): Promise<ExcelRtd>;
    constructor(providerName: any, logger: ILog);
    init(): Promise<void>;
    get isDisposed(): boolean;
    get isInitialized(): boolean;
    setValue(topic: any, value: any): void;
    dispose(clearValues?: boolean): Promise<void>;
    addEventListener(type: string, listener: (data?: any) => any): void;
    dispatchEvent(evt: Event): boolean;
    dispatchEvent(typeArg: string, data?: any): boolean;
    toObject(): this;
    private ping;
    private onSubscribe;
    private onUnsubscribe;
    private clear;
}
