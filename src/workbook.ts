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

function cloneRows(rows: CellInput[][] = [[]]): Cell[][] {
  return rows.map((row) => row.map((cell) => normalizeCell(cell)));
}

function normalizeCell(cell: CellInput): Cell {
  if (cell && typeof cell === "object" && !Array.isArray(cell)) {
    return {
      value: cell.value ?? "",
      formula: cell.formula ?? null,
      style: { ...(cell.style ?? {}) },
      comment: cell.comment ?? null
    };
  }

  return {
    value: cell ?? "",
    formula: null,
    style: {},
    comment: null
  };
}

function normalizeSheet(sheet: SheetInput | undefined, index: number): Sheet {
  return {
    name: sheet?.name || `Sheet${index + 1}`,
    rows: cloneRows(sheet?.rows?.length ? sheet.rows : [[]])
  };
}

function resolveSheet(workbook: WorkbookSnapshot, sheetRef: SheetReference = 0): Sheet | undefined {
  if (typeof sheetRef === "number") {
    return workbook.sheets[sheetRef];
  }

  return workbook.sheets.find((sheet) => sheet.name === sheetRef);
}

function ensureCell(sheet: Sheet, rowIndex: number, columnIndex: number): Cell {
  while (sheet.rows.length <= rowIndex) {
    sheet.rows.push([]);
  }

  const row = sheet.rows[rowIndex];

  while (row.length <= columnIndex) {
    row.push(normalizeCell(""));
  }

  return row[columnIndex];
}

function validateCoordinates(rowIndex: number, columnIndex: number): void {
  if (!Number.isInteger(rowIndex) || rowIndex < 0) {
    throw new Error(`Invalid row index: ${rowIndex}`);
  }

  if (!Number.isInteger(columnIndex) || columnIndex < 0) {
    throw new Error(`Invalid column index: ${columnIndex}`);
  }
}

export function createWorkbook(options: { sheets?: SheetInput[] } = {}): Workbook {
  const workbook: WorkbookSnapshot = {
    sheets: (options.sheets?.length ? options.sheets : [{ name: "Sheet1", rows: [[]] }]).map((sheet, index) =>
      normalizeSheet(sheet, index)
    )
  };

  return {
    listSheets() {
      return workbook.sheets.map((sheet) => sheet.name);
    },
    getSheet(sheetRef = 0) {
      return resolveSheet(workbook, sheetRef);
    },
    addSheet(name, rows = [[]]) {
      const sheet = normalizeSheet({ name, rows }, workbook.sheets.length);
      workbook.sheets.push(sheet);
      return sheet;
    },
    getCell(sheetRef: SheetReference | number, rowIndex?: number, columnIndex?: number) {
      let targetSheetRef: SheetReference = sheetRef;
      let targetRowIndex = rowIndex;
      let targetColumnIndex = columnIndex;

      if (targetColumnIndex === undefined) {
        targetColumnIndex = targetRowIndex;
        targetRowIndex = typeof sheetRef === "number" ? sheetRef : undefined;
        targetSheetRef = 0;
      }

      if (targetRowIndex === undefined || targetColumnIndex === undefined) {
        return null;
      }

      validateCoordinates(targetRowIndex, targetColumnIndex);
      return resolveSheet(workbook, targetSheetRef)?.rows?.[targetRowIndex]?.[targetColumnIndex] ?? null;
    },
    setCell(sheetRef, rowIndex, columnIndex, value) {
      const sheet = resolveSheet(workbook, sheetRef);

      if (!sheet) {
        throw new Error(`Unknown sheet: ${sheetRef}`);
      }

      validateCoordinates(rowIndex, columnIndex);
      const cell = ensureCell(sheet, rowIndex, columnIndex);
      cell.value = value ?? "";
      return cell;
    },
    updateCell(sheetRef, rowIndex, columnIndex, detail = {}) {
      const sheet = resolveSheet(workbook, sheetRef);

      if (!sheet) {
        throw new Error(`Unknown sheet: ${sheetRef}`);
      }

      validateCoordinates(rowIndex, columnIndex);
      const cell = ensureCell(sheet, rowIndex, columnIndex);
      if ("value" in detail) {
        cell.value = detail.value ?? "";
      }
      if ("formula" in detail) {
        cell.formula = detail.formula ?? null;
      }
      if ("comment" in detail) {
        cell.comment = detail.comment ?? null;
      }
      if ("style" in detail) {
        cell.style = { ...(cell.style ?? {}), ...(detail.style ?? {}) };
      }
      return cell;
    },
    applyOperation(operation) {
      switch (operation?.type) {
        case "addSheet":
          return this.addSheet(operation.name, operation.rows);
        case "setCell":
          return this.setCell(operation.sheet ?? 0, operation.row, operation.column, operation.value);
        case "updateCell":
          return this.updateCell(operation.sheet ?? 0, operation.row, operation.column, operation.detail);
        default:
          throw new Error(`Unsupported operation: ${(operation as { type?: string } | undefined)?.type}`);
      }
    },
    toJSON() {
      return {
        sheets: workbook.sheets.map((sheet) => ({
          name: sheet.name,
          rows: sheet.rows.map((row) => row.map((cell) => ({ ...cell, style: { ...cell.style } })))
        }))
      };
    }
  };
}
