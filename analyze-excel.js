#!/usr/bin/env node

const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

/**
 * Excel 分析器 - 专为金融建模设计
 * 可以读取复杂的 Excel 文件并提取关键信息
 */

class ExcelAnalyzer {
  constructor(filePath) {
    this.filePath = filePath;
    this.workbook = null;
  }

  /**
   * 加载 Excel 文件
   */
  load() {
    if (!fs.existsSync(this.filePath)) {
      throw new Error(`文件不存在: ${this.filePath}`);
    }

    console.log(`\n📁 正在加载: ${this.filePath}\n`);
    this.workbook = XLSX.readFile(this.filePath, {
      cellFormula: true,  // 读取公式
      cellStyles: true,   // 读取样式
      cellNF: true,       // 读取数字格式
      cellDates: true,    // 读取日期
      sheetStubs: true,   // 读取空单元格
    });

    console.log('✅ 文件加载成功!\n');
    return this;
  }

  /**
   * 获取工作簿基本信息
   */
  getInfo() {
    const info = {
      fileName: path.basename(this.filePath),
      fileSize: (fs.statSync(this.filePath).size / 1024).toFixed(2) + ' KB',
      sheetCount: this.workbook.SheetNames.length,
      sheetNames: this.workbook.SheetNames,
    };

    console.log('📊 工作簿信息:');
    console.log('━'.repeat(50));
    console.log(`文件名: ${info.fileName}`);
    console.log(`文件大小: ${info.fileSize}`);
    console.log(`工作表数量: ${info.sheetCount}`);
    console.log(`工作表列表: ${info.sheetNames.join(', ')}\n`);

    return info;
  }

  /**
   * 分析指定工作表
   */
  analyzeSheet(sheetName) {
    if (!sheetName) {
      sheetName = this.workbook.SheetNames[0];
    }

    const sheet = this.workbook.Sheets[sheetName];
    if (!sheet) {
      throw new Error(`工作表 "${sheetName}" 不存在`);
    }

    console.log(`\n📄 分析工作表: "${sheetName}"`);
    console.log('━'.repeat(50));

    // 获取范围
    const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1');
    const stats = {
      sheetName,
      range: sheet['!ref'],
      rows: range.e.r - range.s.r + 1,
      cols: range.e.c - range.s.c + 1,
      formulas: 0,
      numbers: 0,
      text: 0,
      dates: 0,
      empty: 0,
    };

    // 统计单元格类型
    for (let R = range.s.r; R <= range.e.r; R++) {
      for (let C = range.s.c; C <= range.e.c; C++) {
        const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
        const cell = sheet[cellAddress];

        if (!cell) {
          stats.empty++;
          continue;
        }

        if (cell.f) stats.formulas++;

        switch (cell.t) {
          case 'n': stats.numbers++; break;
          case 's': stats.text++; break;
          case 'd': stats.dates++; break;
        }
      }
    }

    console.log(`范围: ${stats.range}`);
    console.log(`行数: ${stats.rows}`);
    console.log(`列数: ${stats.cols}`);
    console.log(`总单元格: ${stats.rows * stats.cols}`);
    console.log(`\n单元格类型统计:`);
    console.log(`  📝 文本: ${stats.text}`);
    console.log(`  🔢 数字: ${stats.numbers}`);
    console.log(`  🧮 公式: ${stats.formulas}`);
    console.log(`  📅 日期: ${stats.dates}`);
    console.log(`  ⬜ 空白: ${stats.empty}\n`);

    return stats;
  }

  /**
   * 获取工作表数据（JSON 格式）
   */
  getSheetData(sheetName, options = {}) {
    if (!sheetName) {
      sheetName = this.workbook.SheetNames[0];
    }

    const sheet = this.workbook.Sheets[sheetName];
    if (!sheet) {
      throw new Error(`工作表 "${sheetName}" 不存在`);
    }

    // 转换为 JSON（第一行作为表头）
    const data = XLSX.utils.sheet_to_json(sheet, {
      header: options.header || undefined,
      raw: options.raw !== undefined ? options.raw : false,
      defval: options.defval || '',
    });

    return data;
  }

  /**
   * 显示工作表数据预览
   */
  previewSheet(sheetName, rows = 10) {
    if (!sheetName) {
      sheetName = this.workbook.SheetNames[0];
    }

    const data = this.getSheetData(sheetName);

    console.log(`\n👀 数据预览 (前 ${Math.min(rows, data.length)} 行):`);
    console.log('━'.repeat(50));

    if (data.length === 0) {
      console.log('(空工作表)');
      return;
    }

    // 显示前 N 行
    const preview = data.slice(0, rows);
    console.table(preview);

    if (data.length > rows) {
      console.log(`\n... 还有 ${data.length - rows} 行数据\n`);
    }

    return preview;
  }

  /**
   * 提取所有公式
   */
  extractFormulas(sheetName) {
    if (!sheetName) {
      sheetName = this.workbook.SheetNames[0];
    }

    const sheet = this.workbook.Sheets[sheetName];
    const formulas = [];

    for (let cell in sheet) {
      if (cell[0] === '!') continue; // 跳过元数据

      if (sheet[cell].f) {
        formulas.push({
          cell,
          formula: sheet[cell].f,
          value: sheet[cell].v,
        });
      }
    }

    console.log(`\n🧮 公式列表 (共 ${formulas.length} 个):`);
    console.log('━'.repeat(50));

    formulas.forEach((item, index) => {
      console.log(`${index + 1}. ${item.cell}: ${item.formula}`);
      console.log(`   结果: ${item.value}\n`);
    });

    return formulas;
  }

  /**
   * 导出为 JSON 文件
   */
  exportToJSON(sheetName, outputPath) {
    const data = this.getSheetData(sheetName);
    const jsonPath = outputPath || `./${sheetName || 'sheet'}.json`;

    fs.writeFileSync(jsonPath, JSON.stringify(data, null, 2), 'utf8');
    console.log(`\n💾 已导出到: ${jsonPath}\n`);

    return jsonPath;
  }

  /**
   * 全面分析（所有工作表）
   */
  analyzeAll() {
    this.getInfo();

    this.workbook.SheetNames.forEach((sheetName, index) => {
      console.log(`\n${'='.repeat(60)}`);
      console.log(`工作表 ${index + 1}/${this.workbook.SheetNames.length}`);
      console.log('='.repeat(60));

      this.analyzeSheet(sheetName);
      this.previewSheet(sheetName, 5);
    });
  }
}

// ============================================
// 命令行接口
// ============================================

function printUsage() {
  console.log(`
📊 Excel 分析器 - 使用说明

用法:
  node analyze-excel.js <Excel文件路径> [选项]

选项:
  --sheet <名称>     指定要分析的工作表
  --preview <行数>   预览指定行数的数据 (默认: 10)
  --formulas        提取所有公式
  --export <路径>    导出为 JSON 文件
  --all             分析所有工作表

示例:
  # 分析整个工作簿
  node analyze-excel.js financial-model.xlsx

  # 分析特定工作表
  node analyze-excel.js financial-model.xlsx --sheet "损益表"

  # 提取公式
  node analyze-excel.js financial-model.xlsx --sheet "损益表" --formulas

  # 导出为 JSON
  node analyze-excel.js financial-model.xlsx --sheet "损益表" --export output.json

  # 完整分析
  node analyze-excel.js financial-model.xlsx --all
`);
}

// 主程序
function main() {
  const args = process.argv.slice(2);

  if (args.length === 0 || args.includes('--help') || args.includes('-h')) {
    printUsage();
    process.exit(0);
  }

  const filePath = args[0];

  try {
    const analyzer = new ExcelAnalyzer(filePath);
    analyzer.load();

    // 解析选项
    const sheetName = args.includes('--sheet')
      ? args[args.indexOf('--sheet') + 1]
      : null;

    const previewRows = args.includes('--preview')
      ? parseInt(args[args.indexOf('--preview') + 1])
      : 10;

    const exportPath = args.includes('--export')
      ? args[args.indexOf('--export') + 1]
      : null;

    if (args.includes('--all')) {
      analyzer.analyzeAll();
    } else {
      analyzer.getInfo();
      analyzer.analyzeSheet(sheetName);
      analyzer.previewSheet(sheetName, previewRows);

      if (args.includes('--formulas')) {
        analyzer.extractFormulas(sheetName);
      }

      if (exportPath) {
        analyzer.exportToJSON(sheetName, exportPath);
      }
    }

    console.log('✅ 分析完成!\n');

  } catch (error) {
    console.error(`\n❌ 错误: ${error.message}\n`);
    process.exit(1);
  }
}

// 运行
if (require.main === module) {
  main();
}

module.exports = ExcelAnalyzer;
