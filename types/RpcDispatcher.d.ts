export declare abstract class RpcDispatcher implements EventTarget {
    protected static messageId: number;
    protected static callbacks: {
        [messageId: number]: Function;
    };
    listeners: {
        [eventType: string]: Function[];
    };
    addEventListener(type: string, listener: () => any): void;
    removeEventListener(type: string, listener: () => any): void;
    private hasEventListener(type, listener);
    dispatchEvent(event: any): boolean;
    getDefaultMessage(): any;
    protected invokeRemote(functionName: string, data?: any, callback?: Function): void;
}
