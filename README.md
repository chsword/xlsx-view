# xlsx-view

一个轻量级的 xlsx 所见即所得组件库初始化版本，当前包含：

- 一个可在浏览器中直接使用的工作簿模型 API
- 一个最小可编辑的表格视图组件
- 一个静态 `examples` 演示页面
- 面向 AI Agent 的操作接口（`applyOperation`）

## 快速开始

```bash
npm test
python3 -m http.server 4173
```

然后访问 `http://localhost:4173/examples/` 查看静态示例。

## API

```js
import { createSpreadsheetView, createWorkbook } from "./src/index.js";

const workbook = createWorkbook({
  sheets: [
    {
      name: "Sheet1",
      rows: [[{ value: "A1" }, { value: "B1" }]]
    }
  ]
});

workbook.setCell(0, 1, 0, "A2");
workbook.updateCell(0, 1, 0, { style: { bold: true }, comment: "editable" });
workbook.applyOperation({
  type: "setCell",
  sheet: "Sheet1",
  row: 1,
  column: 1,
  value: "B2"
});

createSpreadsheetView({
  container: document.querySelector("#app"),
  workbook
});
```

## 当前范围

这是一个初始化版本，先提供统一的数据模型、基础视图与 Agent 操作入口，后续可以在这个骨架上继续补充更细粒度的 Open XML 能力、样式能力与真实 xlsx 读写能力。