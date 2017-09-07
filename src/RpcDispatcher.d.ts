export declare abstract class RpcDispatcher implements EventTarget {
    protected static messageId: number;
    protected static callbacks: {
        [messageId: number]: Function;
    };
    connectionUuid: string;
    listeners: {
        [eventType: string]: Function[];
    };
    addEventListener(type: string, listener: (data?: any) => any): void;
    removeEventListener(type: string, listener: (data?: any) => any): void;
    private hasEventListener(type, listener);
    dispatchEvent(event: any): boolean;
    getDefaultMessage(): any;
    protected invokeExcelCall(functionName: string, data?: any, callback?: Function): void;
    protected invokeServiceCall(functionName: string, data?: any, callback?: Function): void;
    private invokeRemoteCall(topic, functionName, data?, callback?);
}
