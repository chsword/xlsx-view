# xlsx-view

一个轻量级的 xlsx 所见即所得 TypeScript 组件库初始化版本，当前包含：

- 一个可在浏览器中直接使用的工作簿模型 API
- 一个最小可编辑的表格视图组件
- 一个静态 `examples` 演示页面
- 一个仓库根目录的 GitHub Pages 预览页面
- 面向 AI Agent 的操作接口（`applyOperation`）

## 快速开始

```bash
npm run build
npm test
python3 -m http.server 4173
```

然后访问：

- `http://localhost:4173/` 查看 GitHub Pages 预览页
- `http://localhost:4173/examples/` 查看静态示例

## API

```ts
import { createSpreadsheetView, createWorkbook } from "./dist/index.js";

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

这是一个初始化版本，先提供统一的数据模型、基础视图、TypeScript 类型定义与 Agent 操作入口，后续可以在这个骨架上继续补充更细粒度的 Open XML 能力、样式能力与真实 xlsx 读写能力。

## GitHub Pages

仓库根目录提供了 `index.html` + `main.js` 预览入口，构建产物位于 `dist/`。开启 GitHub Pages 的 branch root 发布后，可以直接使用该页面作为在线预览入口。