import { ExcelService } from './ExcelApi';
import { Workbook } from './ExcelWorkbook';
import { CellAddress, GetCellsPayload, SetCellsPayload, Worksheet } from './ExcelWorksheet';
interface Executor {
    resolve: Function;
    reject: Function;
}
interface PromiseExecutors {
    [messageId: number]: Executor;
}
interface ExcelData {
    [key: string]: string | number;
}
declare type RemoteData = ExcelData | null | SetCellsPayload | CellAddress | GetCellsPayload;
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
    protected static promiseExecutors: PromiseExecutors;
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
     * @function addEventListener Adds event listener to listen to events coming
     * from Excel application
     * @param type The type of the event to listen to
     * @param listener The method to execute when the event has been fired
     */
    addEventListener(type: string, listener: EventListenerOrEventListenerObject): void;
    /**
     * @public
     * @function removeEventListener Removes the event from the local store
     * @param type The type of the event to listen to
     * @param listener The method to execute when the event has been fired
     */
    removeEventListener(type: string, listener: EventListenerOrEventListenerObject): void;
    /**
     * @private
     * @function hasEventListener Check whether an event listener has been
     * registered
     * @param type The type of the event
     * @param listener The method to execute when the event has been fired
     */
    private hasEventListener;
    /**
     * @public
     * @function dispatchEvent Sends event over to the correct entity e.g.
     * Workbook, worksheet
     * @param evtOrTypeArg Pass either an event or event type as a string
     * @param data The data to be passed to the receiving entity
     */
    dispatchEvent<T>(evtOrTypeArg: string | Event, data?: T): boolean;
    /**
     * @private
     * @function getDefaultMessage Get the default message when invoking a remote
     * call
     * @returns {object} Returns an empty object to be populated
     */
    protected getDefaultMessage(): object;
    /**
     * @protected
     * @function invokeExcelCall Invokes a call in excel application via RPC
     * @param functionName The name of the function to invoke
     * @param data Any data to be sent over as part of the invocation
     */
    protected invokeExcelCall<T>(functionName: string, data?: RemoteData): Promise<T>;
    /**
     * @protected
     * @function invokeServiceCall Invokes a call in the excel service process via
     * RPC
     * @param functionName The name of the function to invoke
     * @param data Any data to be sent over as part of the invocation
     */
    protected invokeServiceCall<T>(functionName: string, data?: null | ExcelData): Promise<T>;
    /**
     * @private
     * @function invokeRemoteCall Invokes a remote procedure call
     * @param topic Topic to send on
     * @param functionName The name of the function to invoke
     * @param data The data to be sent over as part of the invocation
     * @param callback Callback to be applied to the promise
     */
    private invokeRemoteCall;
    abstract toObject(): Worksheet | Workbook | ExcelService | {};
}
export {};
