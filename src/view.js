function createInput(cell, label, onChange) {
  const input = document.createElement("input");
  input.type = "text";
  input.value = cell?.value ?? "";
  input.dataset.formula = cell?.formula ?? "";
  input.placeholder = "Cell";
  input.setAttribute("aria-label", label);
  input.addEventListener("input", (event) => {
    onChange(event.target.value);
  });
  return input;
}

function getColumnLabel(columnIndex) {
  let label = "";
  let current = columnIndex;

  do {
    label = String.fromCharCode(65 + (current % 26)) + label;
    current = Math.floor(current / 26) - 1;
  } while (current >= 0);

  return label;
}

export function createSpreadsheetView({ container, workbook, sheet = 0 } = {}) {
  if (!container) {
    throw new Error("container is required");
  }

  if (!workbook) {
    throw new Error("workbook is required");
  }

  let activeSheet = sheet;

  function render() {
    const renderedSheet = activeSheet;
    const sheetModel = workbook.getSheet(activeSheet);

    if (!sheetModel) {
      throw new Error(`Unknown sheet: ${activeSheet}`);
    }

    container.innerHTML = "";

    const title = document.createElement("h2");
    title.textContent = sheetModel?.name ?? "Sheet";
    container.appendChild(title);

    const table = document.createElement("table");
    table.style.borderCollapse = "collapse";
    table.style.border = "1px solid #999";
    table.setAttribute("aria-label", `${sheetModel?.name ?? "Sheet"} worksheet`);

    const headerRow = document.createElement("tr");
    const corner = document.createElement("th");
    corner.scope = "col";
    corner.textContent = "#";
    headerRow.appendChild(corner);

    const maxColumns = Math.max(0, ...(sheetModel?.rows ?? [[]]).map((row) => row.length));
    for (let columnIndex = 0; columnIndex < maxColumns; columnIndex += 1) {
      const th = document.createElement("th");
      th.scope = "col";
      th.textContent = getColumnLabel(columnIndex);
      th.style.border = "1px solid #999";
      th.style.padding = "4px";
      headerRow.appendChild(th);
    }
    table.appendChild(headerRow);

    (sheetModel?.rows ?? [[]]).forEach((row, rowIndex) => {
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
          createInput(cell, `${sheetModel?.name ?? "Sheet"} ${getColumnLabel(columnIndex)}${rowIndex + 1}`, (value) => {
            workbook.setCell(renderedSheet, rowIndex, columnIndex, value);
          })
        );
        tr.appendChild(td);
      }
      table.appendChild(tr);
    });

    container.appendChild(table);
  }

  render();

  return {
    render,
    setActiveSheet(nextSheet) {
      if (!workbook.getSheet(nextSheet)) {
        throw new Error(`Unknown sheet: ${nextSheet}`);
      }

      activeSheet = nextSheet;
      render();
    }
  };
}
