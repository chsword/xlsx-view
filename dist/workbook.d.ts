export type SheetReference = number | string;
export interface Cell {
    value: unknown;
    formula: string | null;
    style: Record<string, unknown>;
    comment: string | null;
}
export type CellInput = Partial<Cell> | string | number | boolean | null | undefined;
export interface Sheet {
    name: string;
    rows: Cell[][];
}
export interface SheetInput {
    name?: string;
    rows?: CellInput[][];
}
export interface WorkbookSnapshot {
    sheets: Sheet[];
}
export interface CellUpdate {
    value?: unknown;
    formula?: string | null;
    style?: Record<string, unknown>;
    comment?: string | null;
}
type AddSheetOperation = {
    type: "addSheet";
    name?: string;
    rows?: CellInput[][];
};
type SetCellOperation = {
    type: "setCell";
    sheet?: SheetReference;
    row: number;
    column: number;
    value: unknown;
};
type UpdateCellOperation = {
    type: "updateCell";
    sheet?: SheetReference;
    row: number;
    column: number;
    detail?: CellUpdate;
};
export type WorkbookOperation = AddSheetOperation | SetCellOperation | UpdateCellOperation;
export interface Workbook {
    listSheets(): string[];
    getSheet(sheetRef?: SheetReference): Sheet | undefined;
    addSheet(name?: string, rows?: CellInput[][]): Sheet;
    getCell(rowIndex: number, columnIndex: number): Cell | null;
    getCell(sheetRef: SheetReference, rowIndex: number, columnIndex: number): Cell | null;
    setCell(sheetRef: SheetReference, rowIndex: number, columnIndex: number, value: unknown): Cell;
    updateCell(sheetRef: SheetReference, rowIndex: number, columnIndex: number, detail?: CellUpdate): Cell;
    applyOperation(operation: WorkbookOperation): Sheet | Cell;
    toJSON(): WorkbookSnapshot;
}
export declare function createWorkbook(options?: {
    sheets?: SheetInput[];
}): Workbook;
export {};
