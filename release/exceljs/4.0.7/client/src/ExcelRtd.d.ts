import { EventEmitter } from './EventEmitter';
import { ILog } from './ILog';
export declare class ExcelRtd extends EventEmitter {
    private heartbeatIntervalInMilliseconds;
    providerName: string;
    provider: any;
    logger: ILog;
    listeners: {
        [eventType: string]: Function[];
    };
    pingPath: string;
    heartbeatPath: string;
    connectedTopics: {};
    connectedKey: string;
    disconnectedKey: string;
    loggerName: string;
    private initialized;
    private disposed;
    heartbeatToken: number;
    static create(providerName: any, logger: ILog, heartbeatIntervalInMilliseconds?: number): Promise<ExcelRtd>;
    constructor(providerName: any, logger: ILog, heartbeatIntervalInMilliseconds?: number);
    init(): Promise<void>;
    get isDisposed(): boolean;
    get isInitialized(): boolean;
    setValue(topic: any, value: any): void;
    dispose(): Promise<void>;
    addEventListener(type: string, listener: (data?: any) => any): void;
    dispatchEvent(evt: Event): boolean;
    dispatchEvent(typeArg: string, data?: any): boolean;
    toObject(): this;
    private ping;
    private establishHeartbeat;
    private onSubscribe;
    private onUnsubscribe;
    private clear;
}
