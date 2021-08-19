import { ILog } from './ILog';
export declare class NoOpLogger implements ILog {
    name: string;
    constructor();
    trace(message: any, ...args: any[]): void;
    debug(message: any, ...args: any[]): void;
    info(message: any, ...args: any[]): void;
    warn(message: any, ...args: any[]): void;
    error(message: any, error: any, ...args: any[]): void;
    fatal(message: any, error: any, ...args: any[]): void;
}
