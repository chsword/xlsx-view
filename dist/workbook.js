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
function validateCount(count) {
    if (!Number.isInteger(count) || count < 1) {
        throw new Error(`Invalid count: ${count}`);
    }
}
export function createWorkbook(options = {}) {
    const workbook = {
        sheets: (options.sheets?.length ? options.sheets : [{ name: "Sheet1", rows: [[]] }]).map((sheet, index) => normalizeSheet(sheet, index))
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
        renameSheet(sheetRef, name) {
            if (!name || !name.trim()) {
                throw new Error("Sheet name is required");
            }
            const sheet = resolveSheet(workbook, sheetRef);
            if (!sheet) {
                throw new Error(`Unknown sheet: ${sheetRef}`);
            }
            sheet.name = name.trim();
            return sheet;
        },
        removeSheet(sheetRef) {
            const index = typeof sheetRef === "number"
                ? sheetRef
                : workbook.sheets.findIndex((sheet) => sheet.name === sheetRef);
            if (index < 0 || index >= workbook.sheets.length) {
                throw new Error(`Unknown sheet: ${sheetRef}`);
            }
            if (workbook.sheets.length === 1) {
                throw new Error("Workbook must contain at least one sheet");
            }
            return workbook.sheets.splice(index, 1)[0];
        },
        insertRow(sheetRef, rowIndex, count = 1) {
            const sheet = resolveSheet(workbook, sheetRef);
            if (!sheet) {
                throw new Error(`Unknown sheet: ${sheetRef}`);
            }
            validateCoordinates(rowIndex, 0);
            validateCount(count);
            const maxColumns = Math.max(1, ...sheet.rows.map((row) => row.length));
            const newRows = Array.from({ length: count }, () => Array.from({ length: maxColumns }, () => normalizeCell("")));
            if (rowIndex > sheet.rows.length) {
                while (sheet.rows.length < rowIndex) {
                    sheet.rows.push(Array.from({ length: maxColumns }, () => normalizeCell("")));
                }
            }
            sheet.rows.splice(rowIndex, 0, ...newRows);
            return sheet;
        },
        deleteRow(sheetRef, rowIndex, count = 1) {
            const sheet = resolveSheet(workbook, sheetRef);
            if (!sheet) {
                throw new Error(`Unknown sheet: ${sheetRef}`);
            }
            validateCoordinates(rowIndex, 0);
            validateCount(count);
            if (rowIndex >= sheet.rows.length) {
                return sheet;
            }
            sheet.rows.splice(rowIndex, count);
            if (sheet.rows.length === 0) {
                sheet.rows.push([]);
            }
            return sheet;
        },
        insertColumn(sheetRef, columnIndex, count = 1) {
            const sheet = resolveSheet(workbook, sheetRef);
            if (!sheet) {
                throw new Error(`Unknown sheet: ${sheetRef}`);
            }
            validateCoordinates(0, columnIndex);
            validateCount(count);
            const rowCount = Math.max(1, sheet.rows.length);
            while (sheet.rows.length < rowCount) {
                sheet.rows.push([]);
            }
            for (const row of sheet.rows) {
                while (row.length < columnIndex) {
                    row.push(normalizeCell(""));
                }
                const newCells = Array.from({ length: count }, () => normalizeCell(""));
                row.splice(columnIndex, 0, ...newCells);
            }
            return sheet;
        },
        deleteColumn(sheetRef, columnIndex, count = 1) {
            const sheet = resolveSheet(workbook, sheetRef);
            if (!sheet) {
                throw new Error(`Unknown sheet: ${sheetRef}`);
            }
            validateCoordinates(0, columnIndex);
            validateCount(count);
            for (const row of sheet.rows) {
                if (columnIndex < row.length) {
                    row.splice(columnIndex, count);
                }
            }
            return sheet;
        },
        getCell(sheetRef, rowIndex, columnIndex) {
            let targetSheetRef = sheetRef;
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
                case "renameSheet":
                    return this.renameSheet(operation.sheet ?? 0, operation.name);
                case "removeSheet":
                    return this.removeSheet(operation.sheet ?? 0);
                case "setCell":
                    return this.setCell(operation.sheet ?? 0, operation.row, operation.column, operation.value);
                case "updateCell":
                    return this.updateCell(operation.sheet ?? 0, operation.row, operation.column, operation.detail);
                case "insertRow":
                    return this.insertRow(operation.sheet ?? 0, operation.row, operation.count);
                case "deleteRow":
                    return this.deleteRow(operation.sheet ?? 0, operation.row, operation.count);
                case "insertColumn":
                    return this.insertColumn(operation.sheet ?? 0, operation.column, operation.count);
                case "deleteColumn":
                    return this.deleteColumn(operation.sheet ?? 0, operation.column, operation.count);
                default:
                    throw new Error(`Unsupported operation: ${operation?.type}`);
            }
        },
        applyOperations(operations) {
            return operations.map((operation) => this.applyOperation(operation));
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
