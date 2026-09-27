import test from "node:test";
import assert from "node:assert/strict";

import { createWorkbook } from "../src/index.js";

test("createWorkbook exposes a default sheet and cell operations", () => {
  const workbook = createWorkbook();

  assert.deepEqual(workbook.listSheets(), ["Sheet1"]);

  workbook.setCell(0, 0, 0, "hello");
  workbook.updateCell(0, 0, 0, {
    formula: "=LOWER(\"HELLO\")",
    style: { bold: true },
    comment: "seed"
  });

  assert.deepEqual(workbook.getCell(0, 0, 0), {
    value: "hello",
    formula: "=LOWER(\"HELLO\")",
    style: { bold: true },
    comment: "seed"
  });
});

test("createWorkbook can apply agent-friendly operations by sheet name", () => {
  const workbook = createWorkbook({
    sheets: [{ name: "Budget", rows: [[{ value: "Q1" }]] }]
  });

  workbook.applyOperation({
    type: "setCell",
    sheet: "Budget",
    row: 1,
    column: 0,
    value: "1200"
  });

  workbook.applyOperation({
    type: "updateCell",
    sheet: "Budget",
    row: 1,
    column: 0,
    detail: { style: { numberFormat: "#,##0" } }
  });

  assert.equal(workbook.getCell("Budget", 1, 0).value, "1200");
  assert.deepEqual(workbook.toJSON().sheets[0].rows[1][0].style, {
    numberFormat: "#,##0"
  });
});
