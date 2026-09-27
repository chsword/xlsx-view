import type { SheetReference, Workbook } from "./workbook.js";
export interface SpreadsheetView {
    render(): void;
    setActiveSheet(nextSheet: SheetReference): void;
}
interface SpreadsheetViewOptions {
    container?: HTMLElement | null;
    workbook?: Workbook;
    sheet?: SheetReference;
}
export declare function createSpreadsheetView({ container, workbook, sheet }?: SpreadsheetViewOptions): SpreadsheetView;
export {};
