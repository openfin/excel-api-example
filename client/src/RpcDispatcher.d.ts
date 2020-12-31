import { EventEmitter } from './EventEmitter';
export declare abstract class RpcDispatcher extends EventEmitter {
    protected static messageId: number;
    protected static promiseExecutors: {
        [messageId: number]: {
            resolve: Function;
            reject: Function;
        };
    };
    connectionUuid: string;
    getDefaultMessage(): any;
    protected invokeExcelCall(functionName: string, data?: any, callback?: Function): Promise<any>;
    protected invokeServiceCall(functionName: string, data?: any, callback?: Function): Promise<any>;
    private invokeRemoteCall;
    protected applyCallbackToPromise(promise: Promise<any>, callback: Function): Promise<any>;
    abstract toObject(): any;
}
