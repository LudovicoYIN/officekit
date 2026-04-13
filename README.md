# officekit

> **Node.js/Bun Office 文档操作工具集**

一个全面的 Office 文档处理工具包，支持 Word (.docx)、Excel (.xlsx) 和 PowerPoint (.pptx) 格式，提供 CLI 接口和程序化 API。

[![npm version](https://img.shields.io/npm/v/officekit)](https://www.npmjs.com/package/officekit)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

[English](./README.md) | [中文](./README.zh-CN.md)

## 特性

### 核心功能

| 功能 | Word | Excel | PowerPoint |
|------|:----:|:-----:|:----------:|
| 文档创建/编辑 | ✅ | ✅ | ✅ |
| 样式管理 | ✅ | ✅ | ✅ |
| 节点查询 (CSS 选择器) | ✅ | ✅ | ✅ |
| 公式支持 | - | ✅ (50+ 函数) | - |
| 图表 | ✅ | ✅ | ✅ |
| 母版/布局 | - | - | ✅ |
| 动画/变体切换 | - | - | ✅ |
| 预览服务器 | ✅ | ✅ | ✅ |

### Word 文档
- 段落、表格、样式管理
- 页眉/页脚、目录 (TOC)
- 图片、超链接、批注
- 文档保护、水印

### Excel 表格
- 单元格操作，150+ 公式函数支持
- 命名区域（支持链式解析）
- 数据验证、条件格式
- 透视表、迷你图
- CSV/TSV 导入

### PowerPoint 演示
- 幻灯片增删改查
- 形状对齐、分布、效果
- 动画、变体切换
- 母版视图、主题管理
- 表格、图表、媒体支持

## 安装

```bash
# npm
npm install -g officekit

# bun
bun add -g officekit
```

## 快速开始

### CLI 用法

```bash
# 创建文档
officekit create demo.docx
officekit create spreadsheet.xlsx
officekit create presentation.pptx

# 查看内容
officekit view demo.docx
officekit view spreadsheet.xlsx --sheet Sheet1

# 查询结构
officekit query demo.docx /body/p
officekit query spreadsheet.xlsx /Sheet1/A1:B10

# 添加内容
officekit add demo.docx /body --type paragraph --prop "text=Hello World"

# 设置属性
officekit set demo.docx /body/p[1] --prop "bold=true"

# 移动/交换节点
officekit move demo.docx /p[1] /to /p[3]
officekit swap demo.docx /p[1] /p[2]

# 批量操作
officekit batch demo.docx '[{"op":"add","path":"/body","type":"paragraph","props":{"text":"New"}}]'

# 验证文档
officekit validate demo.docx

# 检查问题
officekit check presentation.pptx

# 实时预览
officekit watch demo.docx
officekit watch spreadsheet.xlsx
officekit watch presentation.pptx
officekit unwatch demo.docx

# 驻留会话
officekit open demo.docx
officekit set demo.docx /body/p[1] --prop "text=Updated in resident mode"
officekit close demo.docx

# 原始 XML / 模板 / 部件
officekit raw-set demo.docx /document --xpath "//w:body" --action append --xml "<w:p><w:r><w:t>Hello</w:t></w:r></w:p>"
officekit add-part demo.pptx /slide[1] --type chart --prop "title=Quarterly"
officekit merge template.xlsx output.xlsx --data '{"name":"Alice"}'

# 获取帮助
officekit help
officekit help create
```

### API 用法

```typescript
import { createWordDocument, getWordNode, setWordNode } from "@officekit/word";
import { createExcelWorkbook, queryExcelNodes, setExcelCell } from "@officekit/excel";
import { createPresentation, addSlide, getSlide } from "@officekit/ppt";

// Word
await createWordDocument("output.docx");
await setWordNode("document.docx", "/body/p[1]", { props: { text: "Hello" } });

// Excel
await createExcelWorkbook("output.xlsx");
await setExcelCell("spreadsheet.xlsx", "/Sheet1/A1", { value: 42, formula: "=SUM(B1:B10)" });

// PowerPoint
await createPresentation("output.pptx");
await addSlide("presentation.pptx");
const slide = await getSlide("presentation.pptx", 1);
```

## 命令参考

| 命令 | 说明 |
|------|------|
| `create <file>` | 创建新文档 |
| `view <file>` | 查看文档内容 (text/outline/annotated/stats/html) |
| `get <file> <path>` | 获取指定节点 |
| `query <file> <path>` | 使用 CSS 选择器查询 |
| `set <file> <path>` | 设置节点属性 |
| `add <file> <path>` | 添加新节点 |
| `remove <file> <path>` | 删除节点 |
| `move <file> <from> <to>` | 移动节点 |
| `swap <file> <path1> <path2>` | 交换两个节点 |
| `copy <file> --from <source>` | 复制节点 |
| `batch <file> <operations>` | 批量操作 |
| `raw <file>` | 查看原始 XML |
| `raw-set <file> <part>` | 通过 XPath 修改原始 XML |
| `add-part <file> <path>` | 添加图表、页眉、页脚等文档部件 |
| `merge <template> <output>` | 将 JSON 数据合并到模板文档 |
| `validate <file>` | OpenXML 架构验证 |
| `check <file>` | 检查布局/结构问题 |
| `watch <file>` | 启动实时预览服务器 |
| `unwatch <file>` | 停止活动中的实时预览会话 |
| `open <file>` | 打开驻留会话以复用文档状态 |
| `close <file>` | 关闭驻留会话并持久化更改 |
| `about` | 显示版本信息 |
| `contracts` | 显示功能合同摘要 |

## 预览服务器

`watch` 命令启动本地服务器，提供交互式预览：

```bash
officekit watch demo.docx    # Word 预览
officekit watch data.xlsx    # Excel 预览
officekit watch slides.pptx  # PowerPoint 预览
```

预览服务器支持：
- 实时刷新
- 形状/表格渲染
- 图表显示
- SSE 推送更新
- `unwatch` 显式结束活跃预览会话

## 驻留会话

`open` / `close` 提供了驻留式工作流，可以在多次命令之间复用当前文档状态：

```bash
officekit open report.docx
officekit add report.docx /body --type paragraph --prop "text=Cached edit"
officekit set report.docx /body/p[1] --prop "bold=true"
officekit close report.docx
```

这条链路适合低延迟的多步编辑，也和实时预览/批处理形成互补。

## 路径语法

使用类 CSS 选择器定位文档节点：

```
Word:  /body/p[1]/r[2]       # 第1段第2个run
Excel: /Sheet1/A1:B10        # 范围
       /Sheet1/$A$1          # 绝对引用
PPT:   /slide[1]/shape[2]    # 第1页第2个形状
```

选择器支持：
- `:contains(text)` - 文本包含
- `:has(selector)` - 子节点匹配
- `:eq(n)` - 索引匹配

## 包结构

```
officekit/
├── packages/
│   ├── cli/           # CLI 入口
│   ├── core/          # 核心框架 (routing, document-store)
│   ├── word/          # Word 文档适配器
│   ├── excel/         # Excel 文档适配器
│   ├── ppt/           # PowerPoint 文档适配器
│   ├── preview/       # HTML 预览服务器
│   ├── docs/          # 文档加载器
│   ├── skills/        # Agent 技能注册
│   └── install/       # 安装助手
└── README.md
```

## 开发

```bash
# 安装依赖
npm install

# 构建
npm run build

# 测试
npm test

# CLI 调试
node packages/cli/src/index.js view test.docx
```

## 许可证

MIT
