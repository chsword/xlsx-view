import type { Cell, SheetReference, Workbook } from "./workbook.js";

export interface SpreadsheetView {
  render(): void;
  setActiveSheet(nextSheet: SheetReference): void;
}

interface SpreadsheetViewOptions {
  container?: HTMLElement | null;
  workbook?: Workbook;
  sheet?: SheetReference;
}

function createInput(cell: Cell | undefined, label: string, onChange: (value: string) => void): HTMLInputElement {
  const input = document.createElement("input");
  input.type = "text";
  input.value = String(cell?.value ?? "");
  input.dataset.formula = cell?.formula ?? "";
  input.placeholder = "Cell";
  input.setAttribute("aria-label", label);
  input.addEventListener("input", (event) => {
    onChange((event.target as HTMLInputElement).value);
  });
  return input;
}

function getColumnLabel(columnIndex: number): string {
  let label = "";
  let current = columnIndex;

  do {
    label = String.fromCharCode(65 + (current % 26)) + label;
    current = Math.floor(current / 26) - 1;
  } while (current >= 0);

  return label;
}

export function createSpreadsheetView({ container, workbook, sheet = 0 }: SpreadsheetViewOptions = {}): SpreadsheetView {
  if (!container) {
    throw new Error("container is required");
  }

  if (!workbook) {
    throw new Error("workbook is required");
  }

  const resolvedContainer = container;
  const resolvedWorkbook = workbook;
  let activeSheet = sheet;

  function render(): void {
    const renderedSheet = activeSheet;
    const sheetModel = resolvedWorkbook.getSheet(activeSheet);

    if (!sheetModel) {
      throw new Error(`Unknown sheet: ${activeSheet}`);
    }

    resolvedContainer.innerHTML = "";

    const title = document.createElement("h2");
    title.textContent = sheetModel.name;
    resolvedContainer.appendChild(title);

    const table = document.createElement("table");
    table.style.borderCollapse = "collapse";
    table.style.border = "1px solid #999";
    table.setAttribute("aria-label", `${sheetModel.name} worksheet`);

    const headerRow = document.createElement("tr");
    const corner = document.createElement("th");
    corner.scope = "col";
    corner.textContent = "#";
    headerRow.appendChild(corner);

    const maxColumns = Math.max(5, ...sheetModel.rows.map((row) => row.length));
    for (let columnIndex = 0; columnIndex < maxColumns; columnIndex += 1) {
      const th = document.createElement("th");
      th.scope = "col";
      th.textContent = getColumnLabel(columnIndex);
      th.style.border = "1px solid #999";
      th.style.padding = "4px";
      headerRow.appendChild(th);
    }
    table.appendChild(headerRow);

    const rowCount = Math.max(5, sheetModel.rows.length);

    for (let rowIndex = 0; rowIndex < rowCount; rowIndex += 1) {
      const row = sheetModel.rows[rowIndex] ?? [];
      const tr = document.createElement("tr");
      const rowHeader = document.createElement("th");
      rowHeader.scope = "row";
      rowHeader.textContent = String(rowIndex + 1);
      rowHeader.style.border = "1px solid #999";
      rowHeader.style.padding = "4px";
      tr.appendChild(rowHeader);

      for (let columnIndex = 0; columnIndex < maxColumns; columnIndex += 1) {
        const cell = row[columnIndex];
        const td = document.createElement("td");
        td.style.border = "1px solid #999";
        td.style.padding = "4px";
        td.appendChild(
          createInput(cell, `${sheetModel.name} ${getColumnLabel(columnIndex)}${rowIndex + 1}`, (value) => {
            resolvedWorkbook.updateCell(renderedSheet, rowIndex, columnIndex, { value, formula: null });
          })
        );
        tr.appendChild(td);
      }

      table.appendChild(tr);
    }

    resolvedContainer.appendChild(table);
  }

  render();

  return {
    render,
    setActiveSheet(nextSheet) {
      if (!resolvedWorkbook.getSheet(nextSheet)) {
        throw new Error(`Unknown sheet: ${nextSheet}`);
      }

      activeSheet = nextSheet;
      render();
    }
  };
}
