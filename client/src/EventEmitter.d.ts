export declare abstract class EventEmitter implements EventTarget {
    listeners: {
        [eventType: string]: Function[];
    };
    addEventListener(type: string, listener: (data?: any) => any): void;
    removeEventListener(type: string, listener: (data?: any) => any): void;
    private hasEventListener;
    dispatchEvent(evt: Event): boolean;
    dispatchEvent(typeArg: string, data?: any): boolean;
    abstract toObject(): any;
}
