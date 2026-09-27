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

在线预览地址：https://chsword.github.io/xlsx-view/

首次启用需要仓库管理员在 **Settings → Pages → Build and deployment → Source** 中选择 **GitHub Actions**。

`.github/workflows/pages.yml` 会在默认分支 `main` 更新时自动安装依赖、构建并测试，通过后发布预览站点。也可以在 **Actions → Deploy GitHub Pages → Run workflow** 中选择 `main` 手动发布；其他分支的手动运行会跳过部署。如果更改默认分支名称，需要同步更新工作流的 `push.branches`。

部署内容仅包含根目录的 `index.html`、`main.js`、构建产物 `dist/` 和 `examples/`，保留相对路径以支持 `/xlsx-view/` 项目站点。静态示例地址为 https://chsword.github.io/xlsx-view/examples/ 。首次部署成功后即可访问；无需将仓库根目录配置为 branch root 发布。