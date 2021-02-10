import { EventEmitter } from './EventEmitter';
export declare class ExcelRtd extends EventEmitter {
    providerName: string;
    provider: any;
    initialized: boolean;
    listeners: {
        [eventType: string]: Function[];
    };
    static create(providerName: any): Promise<ExcelRtd>;
    constructor(providerName: any);
    init(): Promise<void>;
    setValue(topic: any, value: any): void;
    onSubscribe(topic: string): void;
    onUnsubscribe(topic: any): void;
    toObject(): this;
}