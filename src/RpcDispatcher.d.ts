import { ExcelService } from './ExcelApi';
import { Workbook } from './ExcelWorkbook';
import { CellAddress, GetCellsPayload, SetCellsPayload, Worksheet } from './ExcelWorksheet';
/**
 * @description Executor for the command being executed
 */
interface Executor {
    resolve: Function;
    reject: Function;
}
/**
 * @description A store for all executors waiting to be executed
 */
interface PromiseExecutors {
    [messageId: number]: Executor;
}
interface ExcelData {
    [key: string]: string | number;
}
declare type RemoteData = ExcelData | null | SetCellsPayload | CellAddress | GetCellsPayload | {
    password: string;
};
/**
 * @abstract
 * @class
 * @description Top level class that communicates with the Excel application
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
     * @public
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
     * @param {string} type The type of the event to listen to
     * @param {EventListenerOrEventListenerObject} listener The method to execute when the event has been fired
     */
    addEventListener(type: string, listener: EventListenerOrEventListenerObject): void;
    /**
     * @public
     * @function removeEventListener
     * @description Removes the event from the local store
     * @param {string} type The type of the event to listen to
     * @param {EventListenerOrEventListenerObject} listener The method to execute when the event has been fired
     */
    removeEventListener(type: string, listener: EventListenerOrEventListenerObject): void;
    /**
     * @private
     * @function hasEventListener
     * @description Check whether an event listener has been
     * registered
     * @param {string} type The type of the event
     * @param {EventListenerOrEventListenerObject} listener The method to execute when the event has been fired
     * @returns {boolean} True or false depending on if the listener exists
     */
    private hasEventListener;
    /**
     * @public
     * @function dispatchEvent
     * @description Sends event over to the correct entity e.g.
     * Workbook, worksheet
     * @param {string|Event} evtOrTypeArg Pass either an event or event type as a string
     * @param {T} data The data to be passed to the receiving entity
     * @returns {boolean} Whether or not the events default behaviour has been prevented
     */
    dispatchEvent<T>(evtOrTypeArg: string | Event, data?: T): boolean;
    /**
     * @private
     * @function getDefaultMessage
     * @description Get the default message when invoking a remote
     * call
     * @returns {object} Returns an empty object to be populated
     */
    protected getDefaultMessage(): object;
    /**
     * @protected
     * @function invokeExcelCall
     * @description Invokes a call in excel application via RPC
     * @param {string} functionName The name of the function to invoke
     * @param {RemoteData?} data Any data to be sent over as part of the invocation
     * @returns {Promise<T>} A Promise with generic data depending on which function calls it
     */
    protected invokeExcelCall<T>(functionName: string, data?: RemoteData): Promise<T>;
    /**
     * @protected
     * @function invokeServiceCall Invokes a call in the excel service process via
     * RPC
     * @param {string} functionName The name of the function to invoke
     * @param {ExcelData|null} data Any data to be sent over as part of the invocation
     * @returns {Promise<T>} A Promise with generic data depending on which function calls it
     */
    protected invokeServiceCall<T>(functionName: string, data?: ExcelData | null): Promise<T>;
    /**
     * @private
     * @function invokeRemoteCall
     * @description Invokes a remote procedure call
     * @param {string} topic Topic to send on
     * @param {string} functionName The name of the function to invoke
     * @param {RemoteData?} data The data to be sent over as part of the invocation
     * @returns {Promise<T>} A Promise with generic data depending on which function calls it
     */
    private invokeRemoteCall;
    /**
     * @public
     * @abstract
     * @function toObject
     * @description This function gets called when we attach to the window we only
     * expose methods that the user needs
     * @returns {Worksheet|Workbook|ExcelService} Returns an object
     */
    abstract toObject(): Worksheet | Workbook | ExcelService | {} | undefined;
}
export {};
