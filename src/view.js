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
    const sheetModel = workbook.getSheet(activeSheet);

    container.innerHTML = "";

    const title = document.createElement("h2");
    title.textContent = sheetModel?.name ?? "Sheet";
    container.appendChild(title);

    const table = document.createElement("table");
    table.border = "1";
    table.cellPadding = "4";
    table.style.borderCollapse = "collapse";
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
      headerRow.appendChild(th);
    }
    table.appendChild(headerRow);

    (sheetModel?.rows ?? [[]]).forEach((row, rowIndex) => {
      const tr = document.createElement("tr");
      const rowHeader = document.createElement("th");
      rowHeader.scope = "row";
      rowHeader.textContent = String(rowIndex + 1);
      tr.appendChild(rowHeader);
      row.forEach((cell, columnIndex) => {
        const td = document.createElement("td");
        td.appendChild(
          createInput(cell, `${sheetModel?.name ?? "Sheet"} ${getColumnLabel(columnIndex)}${rowIndex + 1}`, (value) => {
            workbook.setCell(activeSheet, rowIndex, columnIndex, value);
          })
        );
        tr.appendChild(td);
      });
      table.appendChild(tr);
    });

    container.appendChild(table);
  }

  render();

  return {
    render,
    setActiveSheet(nextSheet) {
      activeSheet = nextSheet;
      render();
    }
  };
}
