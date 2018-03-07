export declare abstract class RpcDispatcher implements EventTarget {
    protected static messageId: number;
    protected static callbacksP: {
        [messageId: number]: {
            resolve: Function;
            reject: Function;
        };
    };
    connectionUuid: string;
    listeners: {
        [eventType: string]: Function[];
    };
    addEventListener(type: string, listener: (data?: any) => any): void;
    removeEventListener(type: string, listener: (data?: any) => any): void;
    private hasEventListener(type, listener);
    dispatchEvent(evt: Event): boolean;
    dispatchEvent(typeArg: string, data?: any): boolean;
    getDefaultMessage(): any;
    protected invokeExcelCall(functionName: string, data?: any, callback?: Function): void;
    protected invokeServiceCall(functionName: string, data?: any, callback?: Function): void;
    private invokeRemoteCall(topic, functionName, data?, callback?);
    abstract toObject(): any;
}
