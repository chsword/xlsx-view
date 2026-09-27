function applyInputStyle(input, cell) {
    const style = (cell?.style ?? {});
    input.style.width = "100%";
    input.style.boxSizing = "border-box";
    input.style.border = "0";
    input.style.outline = "none";
    input.style.background = "transparent";
    input.style.fontWeight = style.bold ? "700" : "400";
    input.style.fontStyle = style.italic ? "italic" : "normal";
    input.style.textDecoration = style.underline ? "underline" : "none";
    input.style.color = typeof style.textColor === "string" ? style.textColor : "#111";
    input.style.backgroundColor = typeof style.backgroundColor === "string" ? style.backgroundColor : "transparent";
    input.style.textAlign = typeof style.align === "string" ? style.align : "left";
}
function createInput(cell, label, selected, onChange, onSelect) {
    const input = document.createElement("input");
    input.type = "text";
    input.value = String(cell?.value ?? "");
    input.dataset.formula = cell?.formula ?? "";
    input.placeholder = "Cell";
    input.setAttribute("aria-label", label);
    applyInputStyle(input, cell);
    if (selected) {
        input.style.outline = "2px solid #3b82f6";
    }
    input.addEventListener("input", (event) => {
        onChange(event.target.value);
    });
    input.addEventListener("focus", onSelect);
    input.addEventListener("click", onSelect);
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
function getCellAddress(rowIndex, columnIndex) {
    return `${getColumnLabel(columnIndex)}${rowIndex + 1}`;
}
export function createSpreadsheetView({ container, workbook, sheet = 0 } = {}) {
    if (!container) {
        throw new Error("container is required");
    }
    if (!workbook) {
        throw new Error("workbook is required");
    }
    const resolvedContainer = container;
    const resolvedWorkbook = workbook;
    let activeSheet = sheet;
    let selectedRow = 0;
    let selectedColumn = 0;
    function render() {
        const renderedSheet = activeSheet;
        const sheetModel = resolvedWorkbook.getSheet(activeSheet);
        if (!sheetModel) {
            throw new Error(`Unknown sheet: ${activeSheet}`);
        }
        resolvedContainer.innerHTML = "";
        const heading = document.createElement("div");
        heading.style.display = "flex";
        heading.style.alignItems = "center";
        heading.style.justifyContent = "space-between";
        heading.style.gap = "12px";
        const title = document.createElement("h2");
        title.textContent = sheetModel.name;
        title.style.margin = "0";
        heading.appendChild(title);
        const sheetTabs = document.createElement("div");
        sheetTabs.style.display = "flex";
        sheetTabs.style.flexWrap = "wrap";
        sheetTabs.style.gap = "6px";
        resolvedWorkbook.listSheets().forEach((name, index) => {
            const tab = document.createElement("button");
            tab.type = "button";
            tab.textContent = name;
            tab.style.padding = "4px 8px";
            tab.style.cursor = "pointer";
            tab.style.border = "1px solid #999";
            tab.style.background = index === renderedSheet ? "#dbeafe" : "#fff";
            tab.addEventListener("click", () => {
                activeSheet = index;
                selectedRow = 0;
                selectedColumn = 0;
                render();
            });
            sheetTabs.appendChild(tab);
        });
        const addSheetButton = document.createElement("button");
        addSheetButton.type = "button";
        addSheetButton.textContent = "+";
        addSheetButton.style.padding = "4px 8px";
        addSheetButton.style.cursor = "pointer";
        addSheetButton.style.border = "1px solid #999";
        addSheetButton.style.background = "#fff";
        addSheetButton.addEventListener("click", () => {
            resolvedWorkbook.addSheet();
            activeSheet = resolvedWorkbook.listSheets().length - 1;
            selectedRow = 0;
            selectedColumn = 0;
            render();
        });
        sheetTabs.appendChild(addSheetButton);
        heading.appendChild(sheetTabs);
        resolvedContainer.appendChild(heading);
        const formulaRow = document.createElement("div");
        formulaRow.style.display = "flex";
        formulaRow.style.gap = "8px";
        formulaRow.style.margin = "8px 0";
        const nameBox = document.createElement("input");
        nameBox.type = "text";
        nameBox.value = getCellAddress(selectedRow, selectedColumn);
        nameBox.setAttribute("aria-label", "Selected cell");
        nameBox.style.width = "80px";
        nameBox.readOnly = true;
        formulaRow.appendChild(nameBox);
        const formulaInput = document.createElement("input");
        formulaInput.type = "text";
        formulaInput.placeholder = "Formula or value";
        formulaInput.style.flex = "1";
        formulaInput.setAttribute("aria-label", "Formula bar");
        const selectedCell = resolvedWorkbook.getCell(renderedSheet, selectedRow, selectedColumn);
        formulaInput.value = String(selectedCell?.formula ?? selectedCell?.value ?? "");
        formulaInput.addEventListener("input", (event) => {
            const value = event.target.value;
            if (value.startsWith("=")) {
                resolvedWorkbook.updateCell(renderedSheet, selectedRow, selectedColumn, {
                    value,
                    formula: value
                });
            }
            else {
                resolvedWorkbook.updateCell(renderedSheet, selectedRow, selectedColumn, {
                    value,
                    formula: null
                });
            }
            render();
        });
        formulaRow.appendChild(formulaInput);
        resolvedContainer.appendChild(formulaRow);
        const table = document.createElement("table");
        table.style.borderCollapse = "collapse";
        table.style.border = "1px solid #999";
        table.setAttribute("aria-label", `${sheetModel.name} worksheet`);
        const headerRow = document.createElement("tr");
        const corner = document.createElement("th");
        corner.scope = "col";
        corner.textContent = "#";
        headerRow.appendChild(corner);
        const maxColumns = Math.max(12, ...sheetModel.rows.map((row) => row.length));
        for (let columnIndex = 0; columnIndex < maxColumns; columnIndex += 1) {
            const th = document.createElement("th");
            th.scope = "col";
            th.textContent = getColumnLabel(columnIndex);
            th.style.border = "1px solid #999";
            th.style.padding = "4px";
            headerRow.appendChild(th);
        }
        table.appendChild(headerRow);
        const rowCount = Math.max(20, sheetModel.rows.length);
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
                td.appendChild(createInput(cell, `${sheetModel.name} ${getCellAddress(rowIndex, columnIndex)}`, selectedRow === rowIndex && selectedColumn === columnIndex, (value) => {
                    resolvedWorkbook.updateCell(renderedSheet, rowIndex, columnIndex, { value, formula: null });
                }, () => {
                    selectedRow = rowIndex;
                    selectedColumn = columnIndex;
                    render();
                }));
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
            selectedRow = 0;
            selectedColumn = 0;
            render();
        },
        getActiveSheet() {
            return activeSheet;
        }
    };
}
