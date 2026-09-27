import { createSpreadsheetView, createWorkbook } from "./dist/index.js";

const workbook = createWorkbook({
  sheets: [
    {
      name: "Overview",
      rows: [
        [
          { value: "Item", style: { bold: true } },
          { value: "Value", style: { bold: true } }
        ],
        [{ value: "Library" }, { value: "TypeScript" }],
        [{ value: "Preview" }, { value: "GitHub Pages" }]
      ]
    }
  ]
});

const app = document.querySelector("#app");
const snapshot = document.querySelector("#snapshot");

createSpreadsheetView({
  container: app,
  workbook
});

const renderSnapshot = () => {
  snapshot.textContent = JSON.stringify(workbook.toJSON(), null, 2);
};

app.addEventListener("input", renderSnapshot);
renderSnapshot();
