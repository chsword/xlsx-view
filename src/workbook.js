function cloneRows(rows = [[]]) {
  return rows.map((row) => row.map((cell) => normalizeCell(cell)));
}

function normalizeCell(cell) {
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

function normalizeSheet(sheet, index) {
  return {
    name: sheet?.name || `Sheet${index + 1}`,
    rows: cloneRows(sheet?.rows?.length ? sheet.rows : [[]])
  };
}

function resolveSheet(workbook, sheetRef = 0) {
  if (typeof sheetRef === "number") {
    return workbook.sheets[sheetRef];
  }

  return workbook.sheets.find((sheet) => sheet.name === sheetRef);
}

function ensureCell(sheet, rowIndex, columnIndex) {
  while (sheet.rows.length <= rowIndex) {
    sheet.rows.push([]);
  }

  const row = sheet.rows[rowIndex];

  while (row.length <= columnIndex) {
    row.push(normalizeCell(""));
  }

  return row[columnIndex];
}

function validateCoordinates(rowIndex, columnIndex) {
  if (!Number.isInteger(rowIndex) || rowIndex < 0) {
    throw new Error(`Invalid row index: ${rowIndex}`);
  }

  if (!Number.isInteger(columnIndex) || columnIndex < 0) {
    throw new Error(`Invalid column index: ${columnIndex}`);
  }
}

export function createWorkbook(options = {}) {
  const workbook = {
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
    getCell(sheetRef, rowIndex, columnIndex) {
      if (columnIndex === undefined) {
        columnIndex = rowIndex;
        rowIndex = sheetRef;
        sheetRef = 0;
      }

      const sheet = resolveSheet(workbook, sheetRef);
      return sheet?.rows?.[rowIndex]?.[columnIndex] ?? null;
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
          throw new Error(`Unsupported operation: ${operation?.type}`);
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
