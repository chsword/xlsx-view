import test from "node:test";
import assert from "node:assert/strict";

import { createSpreadsheetView, createWorkbook } from "../src/index.js";

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

class FakeElement {
  constructor(tagName) {
    this.tagName = tagName;
    this.children = [];
    this.attributes = {};
    this.dataset = {};
    this.eventListeners = {};
    this.style = {};
    this.value = "";
    this.textContent = "";
    this.type = "";
    this.placeholder = "";
    this.border = "";
    this.cellPadding = "";
    this.scope = "";
    this.parentNode = null;
  }

  appendChild(child) {
    child.parentNode = this;
    this.children.push(child);
    return child;
  }

  addEventListener(name, callback) {
    this.eventListeners[name] = callback;
  }

  dispatchEvent(name) {
    this.eventListeners[name]?.({ target: this });
  }

  setAttribute(name, value) {
    this.attributes[name] = value;
  }

  getAttribute(name) {
    return this.attributes[name];
  }

  set innerHTML(value) {
    this.children = [];
    this._innerHTML = value;
  }

  get innerHTML() {
    return this._innerHTML ?? "";
  }
}

function withFakeDocument(callback) {
  const previousDocument = global.document;
  global.document = {
    createElement(tagName) {
      return new FakeElement(tagName);
    }
  };

  try {
    callback();
  } finally {
    global.document = previousDocument;
  }
}

test("createSpreadsheetView validates required arguments", () => {
  const workbook = createWorkbook();

  withFakeDocument(() => {
    assert.throws(() => createSpreadsheetView({ workbook }), /container is required/);
    assert.throws(() => createSpreadsheetView({ container: new FakeElement("div") }), /workbook is required/);
  });
});

test("createSpreadsheetView renders headers, labels, and sheet switching", () => {
  const workbook = createWorkbook({
    sheets: [
      { name: "Alpha", rows: [[{ value: "A1" }, { value: "B1" }]] },
      { name: "Beta", rows: [[{ value: "X1" }]] }
    ]
  });

  withFakeDocument(() => {
    const container = new FakeElement("div");
    const view = createSpreadsheetView({ container, workbook });

    const title = container.children[0];
    const table = container.children[1];
    const headerRow = table.children[0];
    const firstDataRow = table.children[1];
    const firstInput = firstDataRow.children[1].children[0];

    assert.equal(title.textContent, "Alpha");
    assert.equal(table.getAttribute("aria-label"), "Alpha worksheet");
    assert.deepEqual(headerRow.children.map((child) => child.textContent), ["#", "A", "B"]);
    assert.equal(firstDataRow.children[0].textContent, "1");
    assert.equal(firstInput.getAttribute("aria-label"), "Alpha A1");

    firstInput.value = "changed";
    firstInput.dispatchEvent("input");
    assert.equal(workbook.getCell(0, 0, 0).value, "changed");

    view.setActiveSheet(1);
    assert.equal(container.children[0].textContent, "Beta");
    assert.equal(container.children[1].children[1].children[1].children[0].value, "X1");
  });
});
