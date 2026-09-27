function createInput(cell, onChange) {
  const input = document.createElement("input");
  input.type = "text";
  input.value = cell?.value ?? "";
  input.dataset.formula = cell?.formula ?? "";
  input.placeholder = "Cell";
  input.addEventListener("input", (event) => {
    onChange(event.target.value);
  });
  return input;
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

    (sheetModel?.rows ?? [[]]).forEach((row, rowIndex) => {
      const tr = document.createElement("tr");
      row.forEach((cell, columnIndex) => {
        const td = document.createElement("td");
        td.appendChild(
          createInput(cell, (value) => {
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
