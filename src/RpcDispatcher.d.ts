/**
 * @abstract
 * @class Top level class that communicates with the Excel application
 */
export declare abstract class RpcDispatcher implements EventTarget {
    /**
     * @protected
     * @static
     * @description The message id of the action being sent
     */
    protected static messageId: number;
    /**
     * @protected
     * @static
     * @description Promises to be executed
     */
    protected static promiseExecutors: {
        [messageId: number]: {
            resolve: Function;
            reject: Function;
        };
    };
    /**
     * @protected
     * @description The connectionUuid of the excel application
     */
    connectionUuid: string;
    /**
     * @private
     * @description Holds event listeners
     */
    private listeners;
    /**
     * @public
     * @function addEventListener Adds event listener to listen to events coming from Excel application
     * @param type The type of the event to listen to
     * @param listener The method to execute when the event has been fired
     */
    addEventListener(type: string, listener: (data?: any) => any): void;
    /**
     * @public
     * @function removeEventListener Removes the event from the local store
     * @param type The type of the event to listen to
     * @param listener The method to execute when the event has been fired
     */
    removeEventListener(type: string, listener: (data?: any) => any): void;
    /**
     * @private
     * @function hasEventListener Check whether an event listener has been registered
     * @param type The type of the event
     * @param listener The method to execute when the event has been fired
     */
    private hasEventListener(type, listener);
    /**
     * @public
     * @function dispatchEvent Sends event over to the correct entity e.g. Workbook, worksheet
     * @param evtOrTypeArg Pass either an event or event type as a string
     * @param data The data to be passed to the receiving entity
     */
    dispatchEvent(evtOrTypeArg: string | Event, data?: any): boolean;
    /**
     * @private
     * @function getDefaultMessage Get the default message when invoking a remote call
     * @returns {object} Returns an empty object to be populated
     */
    protected getDefaultMessage(): object;
    /**
     * @protected
     * @function invokeExcelCall Invokes a call in excel application via RPC
     * @param functionName The name of the function to invoke
     * @param data Any data to be sent over as part of the invocation
     */
    protected invokeExcelCall(functionName: string, data?: any): Promise<any>;
    /**
     * @protected
     * @function invokeServiceCall Invokes a call in the excel service process via RPC
     * @param functionName The name of the function to invoke
     * @param data Any data to be sent over as part of the invocation
     */
    protected invokeServiceCall(functionName: string, data?: any): Promise<any>;
    /**
     * @private
     * @function invokeRemoteCall Invokes a remote procedure call
     * @param topic Topic to send on
     * @param functionName The name of the function to invoke
     * @param data The data to be sent over as part of the invocation
     * @param callback Callback to be applied to the promise
     */
    private invokeRemoteCall(topic, functionName, data?, callback?);
    /**
     * @protected
     * @function applyCallbackToPromise Applies a callback to the promise
     * @param promise The promise to be acted on
     * @param callback THe callback to be applied to the promise
     * @returns {Promise<any>} A promise with the callback applied
     */
    protected applyCallbackToPromise(promise: Promise<any>, callback: Function): Promise<any>;
    abstract toObject(): any;
}
