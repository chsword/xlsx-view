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
type RenameSheetOperation = {
    type: "renameSheet";
    sheet?: SheetReference;
    name: string;
};
type RemoveSheetOperation = {
    type: "removeSheet";
    sheet?: SheetReference;
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
type InsertRowOperation = {
    type: "insertRow";
    sheet?: SheetReference;
    row: number;
    count?: number;
};
type DeleteRowOperation = {
    type: "deleteRow";
    sheet?: SheetReference;
    row: number;
    count?: number;
};
type InsertColumnOperation = {
    type: "insertColumn";
    sheet?: SheetReference;
    column: number;
    count?: number;
};
type DeleteColumnOperation = {
    type: "deleteColumn";
    sheet?: SheetReference;
    column: number;
    count?: number;
};
export type WorkbookOperation = AddSheetOperation | RenameSheetOperation | RemoveSheetOperation | SetCellOperation | UpdateCellOperation | InsertRowOperation | DeleteRowOperation | InsertColumnOperation | DeleteColumnOperation;
export interface Workbook {
    listSheets(): string[];
    getSheet(sheetRef?: SheetReference): Sheet | undefined;
    addSheet(name?: string, rows?: CellInput[][]): Sheet;
    renameSheet(sheetRef: SheetReference, name: string): Sheet;
    removeSheet(sheetRef: SheetReference): Sheet;
    insertRow(sheetRef: SheetReference, rowIndex: number, count?: number): Sheet;
    deleteRow(sheetRef: SheetReference, rowIndex: number, count?: number): Sheet;
    insertColumn(sheetRef: SheetReference, columnIndex: number, count?: number): Sheet;
    deleteColumn(sheetRef: SheetReference, columnIndex: number, count?: number): Sheet;
    getCell(rowIndex: number, columnIndex: number): Cell | null;
    getCell(sheetRef: SheetReference, rowIndex: number, columnIndex: number): Cell | null;
    setCell(sheetRef: SheetReference, rowIndex: number, columnIndex: number, value: unknown): Cell;
    updateCell(sheetRef: SheetReference, rowIndex: number, columnIndex: number, detail?: CellUpdate): Cell;
    applyOperation(operation: WorkbookOperation): Sheet | Cell;
    applyOperations(operations: WorkbookOperation[]): Array<Sheet | Cell>;
    toJSON(): WorkbookSnapshot;
}
export declare function createWorkbook(options?: {
    sheets?: SheetInput[];
}): Workbook;
export {};
