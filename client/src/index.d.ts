import { ExcelService } from './ExcelApi';
import { ExcelApplication } from './ExcelApplication';
declare global {
    interface Window {
        fin: {
            desktop: {
                ExcelService: ExcelService;
                Excel: ExcelApplication;
            };
        };
    }
}
