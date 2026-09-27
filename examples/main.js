import { createSpreadsheetView, createWorkbook } from "../src/index.js";

const workbook = createWorkbook({
  sheets: [
    {
      name: "Sheet1",
      rows: [
        [
          { value: "A1", style: { bold: true } },
          { value: "B1" }
        ],
        [{ value: "A2" }, { value: "B2", comment: "Editable" }]
      ]
    }
  ]
});

const snapshot = document.querySelector("#snapshot");

createSpreadsheetView({
  container: document.querySelector("#app"),
  workbook
});

const renderSnapshot = () => {
  snapshot.textContent = JSON.stringify(workbook.toJSON(), null, 2);
};

document.querySelector("#app").addEventListener("input", renderSnapshot);
renderSnapshot();
