# 📊 Excel 分析器 - 金融建模专用工具

基于 **SheetJS (xlsx)** 库的强大 Excel 分析工具，专为金融建模和数据分析设计。

## 🚀 功能特点

### ✅ 已实现
- 📖 **读取 Excel 文件** - 支持 `.xlsx`, `.xlsm`, `.xltx`, `.xltm` 格式
- 🧮 **公式提取** - 完整提取所有公式及其计算结果
- 📊 **数据统计** - 自动统计单元格类型（文本、数字、公式、日期）
- 👀 **数据预览** - 表格形式美观展示数据
- 🔗 **跨表引用分析** - 识别工作表之间的公式引用关系
- 💾 **JSON 导出** - 将 Excel 数据导出为 JSON 格式
- 📑 **多工作表支持** - 分析整个工作簿的所有工作表

### 🎯 特别适合
- 💼 金融建模分析
- 📈 财务报表审计
- 🔍 数据质量检查
- 🧪 模型验证测试

## 📦 安装

```bash
# 克隆或进入项目目录
cd /root/excel-analyzer

# 依赖已安装（SheetJS xlsx）
npm install
```

## 💡 使用方法

### 基础用法

```bash
# 分析整个工作簿（默认）
node analyze-excel.js your-file.xlsx

# 分析所有工作表（详细模式）
node analyze-excel.js your-file.xlsx --all
```

### 指定工作表分析

```bash
# 分析特定工作表
node analyze-excel.js financial-model.xlsx --sheet "损益表"

# 分析并预览更多行
node analyze-excel.js financial-model.xlsx --sheet "资产负债表" --preview 20
```

### 提取公式

```bash
# 提取所有公式和计算逻辑
node analyze-excel.js financial-model.xlsx --sheet "成本分析" --formulas
```

### 导出数据

```bash
# 导出为 JSON 文件
node analyze-excel.js financial-model.xlsx --sheet "收入预测" --export revenue.json
```

## 📝 示例

### 演示文件

项目包含一个演示金融建模文件 `financial-model-demo.xlsx`，包含：

1. **收入预测表** - 多年收入预测和增长率
2. **成本分析表** - 成本结构和毛利率计算（含跨表引用）
3. **现金流预测表** - 现金流量预测模型

运行演示：

```bash
# 创建演示文件
node create-demo.js

# 分析演示文件
node analyze-excel.js financial-model-demo.xlsx --all

# 查看成本分析的公式
node analyze-excel.js financial-model-demo.xlsx --sheet "成本分析" --formulas
```

### 输出示例

```
📊 工作簿信息:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
文件名: financial-model-demo.xlsx
文件大小: 24.13 KB
工作表数量: 3
工作表列表: 收入预测, 成本分析, 现金流预测

📄 分析工作表: "成本分析"
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
范围: A1:E12
行数: 12
列数: 5
总单元格: 60

单元格类型统计:
  📝 文本: 32
  🔢 数字: 28
  🧮 公式: 12
  📅 日期: 0
  ⬜ 空白: 0

🧮 公式列表 (共 12 个):
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
1. B9: SUM(B4:B7)
   结果: 2100000

2. B11: 收入预测!B8-B9
   结果: 200000

3. B12: B11/收入预测!B8
   结果: 0.087
```

## 🛠️ 技术栈

- **Node.js** - JavaScript 运行时
- **SheetJS (xlsx)** - Excel 文件处理库
  - 版本: 0.20.3
  - 来源: https://cdn.sheetjs.com/xlsx-0.20.3/xlsx-0.20.3.tgz

## 📚 API 文档

### ExcelAnalyzer 类

```javascript
const ExcelAnalyzer = require('./analyze-excel.js');

// 创建分析器实例
const analyzer = new ExcelAnalyzer('your-file.xlsx');

// 加载文件
analyzer.load();

// 获取工作簿信息
const info = analyzer.getInfo();

// 分析指定工作表
const stats = analyzer.analyzeSheet('Sheet1');

// 获取工作表数据（JSON）
const data = analyzer.getSheetData('Sheet1');

// 预览数据
analyzer.previewSheet('Sheet1', 10);

// 提取公式
const formulas = analyzer.extractFormulas('Sheet1');

// 导出 JSON
analyzer.exportToJSON('Sheet1', 'output.json');

// 分析所有工作表
analyzer.analyzeAll();
```

## 🎓 高级用法

### 在您的代码中使用

```javascript
const ExcelAnalyzer = require('./analyze-excel.js');

// 加载并分析
const analyzer = new ExcelAnalyzer('financial-model.xlsx');
analyzer.load();

// 获取特定工作表数据
const incomeData = analyzer.getSheetData('损益表');

// 处理数据
incomeData.forEach(row => {
  console.log(`收入: ${row['收入']}, 成本: ${row['成本']}`);
});

// 提取所有公式以理解计算逻辑
const formulas = analyzer.extractFormulas('损益表');
formulas.forEach(f => {
  console.log(`${f.cell}: ${f.formula} = ${f.value}`);
});
```

### 批量处理多个文件

```javascript
const fs = require('fs');
const ExcelAnalyzer = require('./analyze-excel.js');

const files = fs.readdirSync('.').filter(f => f.endsWith('.xlsx'));

files.forEach(file => {
  const analyzer = new ExcelAnalyzer(file);
  analyzer.load();
  analyzer.analyzeAll();
});
```

## 🔍 关键特性说明

### 公式读取
- ✅ 读取单元格内的完整公式
- ✅ 保留公式引用关系（如 `=A1+B1`）
- ✅ 支持跨工作表引用（如 `=Sheet2!A1`）
- ✅ 显示公式计算结果

### 数据格式
- ✅ 识别数字、文本、日期、空白单元格
- ✅ 保留数字格式（百分比、货币等）
- ✅ 支持自定义数字格式

### 性能
- ⚡ 高效处理大型 Excel 文件
- 📊 支持分页读取（避免内存溢出）
- 🚀 快速公式提取和分析

## 📖 命令行帮助

```bash
node analyze-excel.js --help
```

## ⚠️ 注意事项

1. **文件路径**：请使用相对路径或绝对路径
2. **中文支持**：完全支持中文工作表名和单元格内容
3. **大文件**：处理超大文件时建议使用 `--preview` 限制预览行数
4. **公式计算**：仅读取公式和已有结果，不重新计算

## 🤝 使用场景

### 1. 金融建模审计
```bash
node analyze-excel.js DCF-model.xlsx --sheet "估值模型" --formulas
```
快速理解 DCF 模型的计算逻辑和公式依赖关系。

### 2. 数据提取
```bash
node analyze-excel.js sales-data.xlsx --sheet "Q1销售" --export q1-sales.json
```
将 Excel 数据转换为 JSON，便于程序处理。

### 3. 模型验证
```bash
node analyze-excel.js budget-2024.xlsx --all
```
全面检查预算模型的所有工作表结构。

## 🚀 可用的 Slash 命令

本项目支持 **28 个全局 slash 命令**（来自 `~/.claude/commands/`），无需审批即可使用。

### 🧠 AI 思考与分析
- `/think [hard|harder|ultra]` - 启用扩展思考模式
- `/ultrathink` - 最大深度推理（20k+ tokens）
- `/explain-code [文件路径]` - 深度代码解释
- `/code-review [文件路径]` - 综合代码审查
- `/security-check` - 安全漏洞分析
- `/debug-help [错误]` - 调试辅助

### 🔍 研究与搜索
- `/deep-search <域名> <目标>` - 深度网站探索
- `/research-deep <主题>` - 多源研究（15-20次搜索）
- `/search-tree <问题>` - MCTS树搜索
- `/reflect-search <目标>` - 反思驱动搜索

### 🛠️ 代码生成与重构
- `/refactor [文件路径]` - 重构建议
- `/optimize [文件路径]` - 性能优化
- `/test-gen [文件路径]` - 生成测试
- `/doc-gen [文件路径]` - 生成文档

### 🎨 Artifact 创建
- `/artifact-react [应用名]` - 创建 React 应用
- `/artifact-mermaid [类型]` - 创建 Mermaid 图表
- `/artifact-excel-analyzer` - Excel 分析工具
- `/quick-prototype [描述]` - 快速原型

### 📊 文件分析
- `/file-analyze [文件路径]` - 分析 PDF、Excel、Word、图片

### 🚀 Git 工作流
- `/push` - 验证式 git push 与自动暂存
- `/pull` - 带 stash 管理的 pull
- `/quick-commit [消息]` - 自动生成提交
- `/checkpoint` - 创建 git 检查点

### ⚙️ 系统管理
- `/status` - 显示 Claude Code 配置
- `/fswatch` - 文件监视工具
- `/playwright-helper` - Playwright MCP 指南

详细使用方法请查看 `~/.claude/commands/README.md`。

---

## 🔗 与 Claude Code 集成

本工具已深度集成到 **Claude Code** 的全局配置中！

### 在 Claude Code 对话中使用

#### 方法 1: 通过 `/file-analyze` 命令（推荐）

```
/file-analyze your-file.xlsx "提取所有公式"
```

Claude 会自动识别这是Excel文件，并为你提供三种分析选项：
1. **CLI快速分析** - 使用本工具（excel-analyzer）
2. **Web可视化分析** - 创建React应用
3. **AI深度分析** - 使用Anthropic API

#### 方法 2: 使用快捷wrapper

```bash
# 在任何地方使用
~/.claude/bin/quick-excel your-file.xlsx --formulas
quick-excel your-file.xlsx --all
```

#### 方法 3: 直接调用（开发者）

```bash
node /root/excel-analyzer/analyze-excel.js your-file.xlsx
```

### Claude Code 集成优势

✅ **智能路由**: Claude自动判断最佳分析方式
✅ **多工具协作**: 可与Web可视化工具配合
✅ **对话式**: 在Claude对话中自然调用
✅ **零配置**: 已预配置，开箱即用

### 配合使用示例

```
User: "分析这个财务模型的公式"
Claude:
  → 使用 excel-analyzer 提取所有公式
  → 显示公式列表和依赖关系
  → 建议创建可视化工具查看数据
```

### 相关 Claude Code 命令

- `/file-analyze` - 通用文件分析入口（包含Excel）
- `/artifact-excel-analyzer` - 创建Web版Excel分析器
- `/quick-prototype` - 快速创建Excel相关原型

---

## 📄 许可证

ISC

## 👨‍💻 作者

Created by Claude Code (Anthropic)

---

## 🎓 了解更多

- **Claude Code文档**: https://docs.claude.com/en/docs/claude-code
- **SheetJS文档**: https://docs.sheetjs.com/
- **全局配置**: `~/.claude/CONFIGURATION_SUMMARY.md`

**提示**: 如果您有任何问题或需要添加新功能，请随时修改代码！SheetJS 库功能非常强大，支持更多高级特性。
