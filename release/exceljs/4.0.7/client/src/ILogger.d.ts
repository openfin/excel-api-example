export interface ILogger {
    name?: string;
    trace?: (message: any, ...args: any[]) => void;
    debug?: (message: any, ...args: any[]) => void;
    info?: (message: any, ...args: any[]) => void;
    warn?: (message: any, ...args: any[]) => void;
    error?: (message: any, error: any, ...args: any[]) => void;
    fatal?: (message: any, error: any, ...args: any[]) => void;
}
