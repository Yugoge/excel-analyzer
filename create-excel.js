#!/usr/bin/env node

const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

/**
 * Excel 创建器 - 支持创建包含公式的金融建模 Excel 文件
 */

class ExcelBuilder {
  constructor(fileName = 'output.xlsx') {
    this.fileName = fileName;
    this.workbook = XLSX.utils.book_new();
  }

  /**
   * 添加工作表（从二维数组）
   */
  addSheet(sheetName, data, options = {}) {
    const sheet = XLSX.utils.aoa_to_sheet(data);

    // 设置列宽
    if (options.colWidths) {
      sheet['!cols'] = options.colWidths.map(wch => ({ wch }));
    }

    // 设置行高
    if (options.rowHeights) {
      sheet['!rows'] = options.rowHeights.map(hpx => ({ hpx }));
    }

    XLSX.utils.book_append_sheet(this.workbook, sheet, sheetName);
    return sheet;
  }

  /**
   * 添加公式到指定单元格
   * @param {string} sheetName - 工作表名称
   * @param {string} cell - 单元格地址 (如 'A1')
   * @param {string} formula - 公式 (如 'SUM(A1:A10)')
   * @param {number} value - 公式计算结果（可选）
   * @param {string} format - 数字格式（可选，如 '0.00%' 表示百分比）
   */
  addFormula(sheetName, cell, formula, value = 0, format = null) {
    const sheet = this.workbook.Sheets[sheetName];
    if (!sheet) {
      throw new Error(`工作表 "${sheetName}" 不存在`);
    }

    const cellObj = { t: 'n', f: formula };
    if (value !== null && value !== undefined) {
      cellObj.v = value;
    }
    if (format) {
      cellObj.z = format;
    }

    sheet[cell] = cellObj;
  }

  /**
   * 批量添加公式
   */
  addFormulas(sheetName, formulas) {
    formulas.forEach(({ cell, formula, value, format }) => {
      this.addFormula(sheetName, cell, formula, value, format);
    });
  }

  /**
   * 设置单元格值
   */
  setCellValue(sheetName, cell, value, format = null) {
    const sheet = this.workbook.Sheets[sheetName];
    if (!sheet) {
      throw new Error(`工作表 "${sheetName}" 不存在`);
    }

    let cellType = 'n';
    if (typeof value === 'string') cellType = 's';
    else if (typeof value === 'boolean') cellType = 'b';
    else if (value instanceof Date) cellType = 'd';

    const cellObj = { t: cellType, v: value };
    if (format) {
      cellObj.z = format;
    }

    sheet[cell] = cellObj;
  }

  /**
   * 保存工作簿
   */
  save(customFileName = null) {
    const fileName = customFileName || this.fileName;
    XLSX.writeFile(this.workbook, fileName);
    console.log(`\n✅ Excel 文件已创建: ${fileName}\n`);
    return fileName;
  }

  /**
   * 获取工作表对象
   */
  getSheet(sheetName) {
    return this.workbook.Sheets[sheetName];
  }
}

// ==========================================
// 预设模板
// ==========================================

/**
 * 创建 DCF 估值模型
 */
function createDCFModel(outputFile = 'dcf-valuation.xlsx') {
  console.log('📊 创建 DCF 估值模型...\n');

  const builder = new ExcelBuilder(outputFile);

  // 假设与参数
  const assumptions = [
    ['DCF 估值模型 - 假设与参数', '', '', '', ''],
    ['', '', '', '', ''],
    ['财务假设', '数值', '单位', '说明', ''],
    ['初始收入', 10000000, '元', '基准年收入', ''],
    ['收入增长率', 0.15, '%', '年增长率', ''],
    ['EBITDA利润率', 0.25, '%', '息税折旧摊销前利润率', ''],
    ['税率', 0.25, '%', '企业所得税率', ''],
    ['折旧率', 0.05, '%', '固定资产折旧率', ''],
    ['资本支出率', 0.03, '%', '占收入比例', ''],
    ['营运资金变动率', 0.02, '%', '占收入变动比例', ''],
    ['', '', '', '', ''],
    ['估值参数', '', '', '', ''],
    ['预测期', 5, '年', '详细预测年数', ''],
    ['WACC', 0.10, '%', '加权平均资本成本', ''],
    ['永续增长率', 0.03, '%', '终值增长率', ''],
  ];

  builder.addSheet('假设', assumptions, {
    colWidths: [25, 15, 10, 30, 5],
  });

  // 设置百分比格式
  builder.addFormula('假设', 'B5', 'B5', 0.15, '0.00%');
  builder.addFormula('假设', 'B6', 'B6', 0.25, '0.00%');
  builder.addFormula('假设', 'B7', 'B7', 0.25, '0.00%');
  builder.addFormula('假设', 'B8', 'B8', 0.05, '0.00%');
  builder.addFormula('假设', 'B9', 'B9', 0.03, '0.00%');
  builder.addFormula('假设', 'B10', 'B10', 0.02, '0.00%');
  builder.addFormula('假设', 'B14', 'B14', 0.10, '0.00%');
  builder.addFormula('假设', 'B15', 'B15', 0.03, '0.00%');

  // 现金流预测
  const fcfData = [
    ['自由现金流预测', '', '', '', '', '', ''],
    ['', '', '', '', '', '', ''],
    ['项目', '年份0', '年份1', '年份2', '年份3', '年份4', '年份5'],
    ['收入', 10000000, '', '', '', '', ''],
    ['收入增长率', '', '', '', '', '', ''],
    ['EBITDA', '', '', '', '', '', ''],
    ['折旧', '', '', '', '', '', ''],
    ['EBIT', '', '', '', '', '', ''],
    ['税后EBIT', '', '', '', '', '', ''],
    ['加回: 折旧', '', '', '', '', '', ''],
    ['减: 资本支出', '', '', '', '', '', ''],
    ['减: 营运资金变动', '', '', '', '', '', ''],
    ['自由现金流 (FCF)', '', '', '', '', '', ''],
  ];

  builder.addSheet('FCF预测', fcfData, {
    colWidths: [20, 12, 12, 12, 12, 12, 12],
  });

  // 添加收入预测公式 (年份1-5)
  for (let year = 1; year <= 5; year++) {
    const col = String.fromCharCode(66 + year); // C, D, E, F, G
    const prevCol = String.fromCharCode(65 + year); // B, C, D, E, F

    // 收入 = 上年收入 * (1 + 增长率)
    builder.addFormula('FCF预测', `${col}4`, `${prevCol}4*(1+假设!$B$5)`, 0);

    // 收入增长率
    builder.addFormula('FCF预测', `${col}5`, `假设!$B$5`, 0.15, '0.00%');

    // EBITDA = 收入 * EBITDA利润率
    builder.addFormula('FCF预测', `${col}6`, `${col}4*假设!$B$6`, 0);

    // 折旧 = 收入 * 折旧率
    builder.addFormula('FCF预测', `${col}7`, `${col}4*假设!$B$8`, 0);

    // EBIT = EBITDA - 折旧
    builder.addFormula('FCF预测', `${col}8`, `${col}6-${col}7`, 0);

    // 税后EBIT = EBIT * (1 - 税率)
    builder.addFormula('FCF预测', `${col}9`, `${col}8*(1-假设!$B$7)`, 0);

    // 加回折旧
    builder.addFormula('FCF预测', `${col}10`, `${col}7`, 0);

    // 资本支出 = 收入 * 资本支出率
    builder.addFormula('FCF预测', `${col}11`, `${col}4*假设!$B$9`, 0);

    // 营运资金变动 = (当年收入 - 上年收入) * 营运资金变动率
    builder.addFormula('FCF预测', `${col}12`, `(${col}4-${prevCol}4)*假设!$B$10`, 0);

    // FCF = 税后EBIT + 折旧 - 资本支出 - 营运资金变动
    builder.addFormula('FCF预测', `${col}13`, `${col}9+${col}10-${col}11-${col}12`, 0);
  }

  // 估值汇总
  const valuationData = [
    ['DCF 估值汇总', '', '', ''],
    ['', '', '', ''],
    ['项目', '年份', '现金流', '现值'],
    ['年份1 FCF', 1, '', ''],
    ['年份2 FCF', 2, '', ''],
    ['年份3 FCF', 3, '', ''],
    ['年份4 FCF', 4, '', ''],
    ['年份5 FCF', 5, '', ''],
    ['', '', '', ''],
    ['预测期现值合计', '', '', ''],
    ['', '', '', ''],
    ['终值计算', '', '', ''],
    ['最后一年 FCF', '', '', ''],
    ['终值', '', '', ''],
    ['终值现值', '', '', ''],
    ['', '', '', ''],
    ['企业价值', '', '', ''],
  ];

  builder.addSheet('估值', valuationData, {
    colWidths: [20, 10, 15, 15],
  });

  // 添加估值公式
  for (let i = 0; i < 5; i++) {
    const row = 4 + i;
    const col = String.fromCharCode(67 + i); // C, D, E, F, G

    // 现金流引用
    builder.addFormula('估值', `C${row}`, `FCF预测!${col}13`, 0);

    // 现值 = FCF / (1 + WACC)^年份
    builder.addFormula('估值', `D${row}`, `C${row}/((1+假设!$B$14)^B${row})`, 0);
  }

  // 预测期现值合计
  builder.addFormula('估值', 'D10', 'SUM(D4:D8)', 0);

  // 最后一年 FCF
  builder.addFormula('估值', 'C13', 'FCF预测!G13', 0);

  // 终值 = 最后一年FCF * (1 + 永续增长率) / (WACC - 永续增长率)
  builder.addFormula('估值', 'C14', 'C13*(1+假设!$B$15)/(假设!$B$14-假设!$B$15)', 0);

  // 终值现值 = 终值 / (1 + WACC)^5
  builder.addFormula('估值', 'C15', 'C14/((1+假设!$B$14)^5)', 0);

  // 企业价值 = 预测期现值 + 终值现值
  builder.addFormula('估值', 'C17', 'D10+C15', 0);

  builder.save();
  console.log('包含以下工作表:');
  console.log('  1. 假设 - 财务假设与估值参数');
  console.log('  2. FCF预测 - 5年自由现金流预测（含完整公式）');
  console.log('  3. 估值 - DCF估值汇总（自动计算企业价值）\n');
}

/**
 * 创建三表模型（损益表、资产负债表、现金流量表）
 */
function createThreeStatementModel(outputFile = 'three-statement-model.xlsx') {
  console.log('📊 创建三表财务模型...\n');

  const builder = new ExcelBuilder(outputFile);

  // 损益表
  const incomeData = [
    ['损益表', '', '', '', '', ''],
    ['单位: 万元', '', '', '', '', ''],
    ['项目', '2023A', '2024E', '2025E', '2026E', '2027E'],
    ['营业收入', 10000, 11500, 13225, 15209, 17490],
    ['营业成本', -6000, -6900, -7935, -9125, -10494],
    ['毛利润', '', '', '', '', ''],
    ['毛利率', '', '', '', '', ''],
    ['', '', '', '', '', ''],
    ['销售费用', -1500, -1725, -1984, -2281, -2624],
    ['管理费用', -800, -920, -1058, -1217, -1399],
    ['研发费用', -500, -575, -661, -760, -874],
    ['营业利润', '', '', '', '', ''],
    ['营业利润率', '', '', '', '', ''],
    ['', '', '', '', '', ''],
    ['利息费用', -200, -200, -200, -200, -200],
    ['税前利润', '', '', '', '', ''],
    ['所得税', '', '', '', '', ''],
    ['净利润', '', '', '', '', ''],
    ['净利率', '', '', '', '', ''],
  ];

  builder.addSheet('损益表', incomeData, {
    colWidths: [18, 12, 12, 12, 12, 12],
  });

  // 添加损益表公式
  const cols = ['B', 'C', 'D', 'E', 'F'];
  cols.forEach(col => {
    // 毛利润 = 收入 + 成本（成本为负数）
    builder.addFormula('损益表', `${col}6`, `${col}4+${col}5`, 0);
    // 毛利率 = 毛利润 / 收入
    builder.addFormula('损益表', `${col}7`, `${col}6/${col}4`, 0, '0.00%');
    // 营业利润 = 毛利润 - 三项费用
    builder.addFormula('损益表', `${col}12`, `${col}6+${col}9+${col}10+${col}11`, 0);
    // 营业利润率
    builder.addFormula('损益表', `${col}13`, `${col}12/${col}4`, 0, '0.00%');
    // 税前利润 = 营业利润 + 利息费用
    builder.addFormula('损益表', `${col}16`, `${col}12+${col}15`, 0);
    // 所得税 = 税前利润 * 25%
    builder.addFormula('损益表', `${col}17`, `${col}16*0.25`, 0);
    // 净利润 = 税前利润 - 所得税
    builder.addFormula('损益表', `${col}18`, `${col}16-${col}17`, 0);
    // 净利率
    builder.addFormula('损益表', `${col}19`, `${col}18/${col}4`, 0, '0.00%');
  });

  // 资产负债表
  const balanceData = [
    ['资产负债表', '', '', '', '', ''],
    ['单位: 万元', '', '', '', '', ''],
    ['项目', '2023A', '2024E', '2025E', '2026E', '2027E'],
    ['资产', '', '', '', '', ''],
    ['流动资产', '', '', '', '', ''],
    ['  货币资金', 3000, 3500, 4200, 5100, 6200],
    ['  应收账款', 2000, 2300, 2645, 3042, 3498],
    ['  存货', 1500, 1725, 1984, 2281, 2624],
    ['  其他流动资产', 500, 575, 661, 760, 874],
    ['流动资产合计', '', '', '', '', ''],
    ['', '', '', '', '', ''],
    ['非流动资产', '', '', '', '', ''],
    ['  固定资产', 8000, 8500, 9000, 9500, 10000],
    ['  无形资产', 2000, 2100, 2200, 2300, 2400],
    ['非流动资产合计', '', '', '', '', ''],
    ['', '', '', '', '', ''],
    ['资产总计', '', '', '', '', ''],
    ['', '', '', '', '', ''],
    ['负债和所有者权益', '', '', '', '', ''],
    ['流动负债', 4000, 4400, 4840, 5324, 5856],
    ['长期借款', 5000, 5000, 5000, 5000, 5000],
    ['负债合计', '', '', '', '', ''],
    ['', '', '', '', '', ''],
    ['所有者权益', 8000, '', '', '', ''],
    ['资本公积', 2000, 2000, 2000, 2000, 2000],
    ['留存收益', 2000, '', '', '', ''],
    ['所有者权益合计', '', '', '', '', ''],
    ['', '', '', '', '', ''],
    ['负债和权益总计', '', '', '', '', ''],
  ];

  builder.addSheet('资产负债表', balanceData, {
    colWidths: [18, 12, 12, 12, 12, 12],
  });

  // 资产负债表公式
  cols.forEach(col => {
    // 流动资产合计
    builder.addFormula('资产负债表', `${col}10`, `SUM(${col}6:${col}9)`, 0);
    // 非流动资产合计
    builder.addFormula('资产负债表', `${col}15`, `SUM(${col}13:${col}14)`, 0);
    // 资产总计
    builder.addFormula('资产负债表', `${col}17`, `${col}10+${col}15`, 0);
    // 负债合计
    builder.addFormula('资产负债表', `${col}22`, `${col}20+${col}21`, 0);
    // 所有者权益合计
    builder.addFormula('资产负债表', `${col}27`, `${col}24+${col}25+${col}26`, 0);
    // 负债和权益总计
    builder.addFormula('资产负债表', `${col}29`, `${col}22+${col}27`, 0);
  });

  // 留存收益 = 上期留存收益 + 本期净利润（简化处理）
  builder.addFormula('资产负债表', 'C26', 'B26+损益表!C18', 0);
  builder.addFormula('资产负债表', 'D26', 'C26+损益表!D18', 0);
  builder.addFormula('资产负债表', 'E26', 'D26+损益表!E18', 0);
  builder.addFormula('资产负债表', 'F26', 'E26+损益表!F18', 0);

  builder.save();
  console.log('包含以下工作表:');
  console.log('  1. 损益表 - 完整的利润表（含所有财务比率公式）');
  console.log('  2. 资产负债表 - 资产负债表（与损益表联动）\n');
}

/**
 * 创建敏感性分析表
 */
function createSensitivityAnalysis(outputFile = 'sensitivity-analysis.xlsx') {
  console.log('📊 创建敏感性分析表...\n');

  const builder = new ExcelBuilder(outputFile);

  // 基础参数
  const baseData = [
    ['敏感性分析 - 基础参数', '', '', ''],
    ['', '', '', ''],
    ['参数', '基准值', '单位', '说明'],
    ['初始投资', 10000000, '元', '项目初始投资额'],
    ['年收入', 3000000, '元', '每年稳定收入'],
    ['年成本', 1500000, '元', '每年运营成本'],
    ['折现率', 0.10, '%', 'WACC'],
    ['项目年限', 10, '年', '项目运营年限'],
    ['', '', '', ''],
    ['计算结果', '', '', ''],
    ['NPV', '', '元', '净现值'],
  ];

  builder.addSheet('基础参数', baseData, {
    colWidths: [18, 15, 10, 30],
  });

  builder.addFormula('基础参数', 'B7', 'B7', 0.10, '0.00%');

  // NPV = -初始投资 + Σ(年收入-年成本)/(1+折现率)^年份
  builder.addFormula('基础参数', 'B11',
    '-B4+SUMPRODUCT((B5-B6)/((1+B7)^ROW(1:10)),1/ROW(1:10)^0)', 0);

  // 双因素敏感性分析
  const sensitivityData = [
    ['NPV 敏感性分析 (收入 vs 折现率)', '', '', '', '', '', '', ''],
    ['', '', '', '', '', '', '', ''],
    ['年收入 ↓ / 折现率 →', '8%', '9%', '10%', '11%', '12%', '13%', '14%'],
  ];

  // 收入变化：80% - 120%
  const revenueMultipliers = [0.8, 0.9, 1.0, 1.1, 1.2];
  revenueMultipliers.forEach((mult, i) => {
    const revenue = 3000000 * mult;
    sensitivityData.push([revenue, '', '', '', '', '', '', '']);
  });

  builder.addSheet('敏感性分析', sensitivityData, {
    colWidths: [25, 12, 12, 12, 12, 12, 12, 12],
  });

  // 添加敏感性分析公式
  const discountRates = [0.08, 0.09, 0.10, 0.11, 0.12, 0.13, 0.14];
  const rows = [4, 5, 6, 7, 8]; // 对应不同收入水平

  rows.forEach((row, i) => {
    discountRates.forEach((rate, j) => {
      const col = String.fromCharCode(66 + j); // B, C, D, E, F, G, H
      // NPV计算
      builder.addFormula('敏感性分析', `${col}${row}`,
        `-基础参数!$B$4+SUMPRODUCT((A${row}-基础参数!$B$6)/((1+${rate})^ROW(1:10)),1/ROW(1:10)^0)`, 0);
    });
  });

  builder.save();
  console.log('包含以下工作表:');
  console.log('  1. 基础参数 - NPV计算基础参数');
  console.log('  2. 敏感性分析 - 双因素敏感性分析矩阵\n');
}

// ==========================================
// 命令行接口
// ==========================================

function printUsage() {
  console.log(`
📊 Excel 创建器 - 使用说明

用法:
  create-excel <模板类型> [输出文件名]

模板类型:
  dcf              创建 DCF 估值模型
  three-statement  创建三表财务模型（损益表、资产负债表）
  sensitivity      创建敏感性分析表
  demo             创建简单演示文件

示例:
  # 创建 DCF 估值模型
  create-excel dcf my-dcf-model.xlsx

  # 创建三表模型
  create-excel three-statement financial-statements.xlsx

  # 创建敏感性分析
  create-excel sensitivity analysis.xlsx

  # 创建演示文件
  create-excel demo

特点:
  ✅ 包含完整的 Excel 公式
  ✅ 自动计算（打开即可使用）
  ✅ 工作表之间自动引用
  ✅ 专业的金融建模结构
  ✅ 可自定义参数后立即看到结果
`);
}

function main() {
  const args = process.argv.slice(2);

  if (args.length === 0 || args.includes('--help') || args.includes('-h')) {
    printUsage();
    process.exit(0);
  }

  const template = args[0];
  const outputFile = args[1];

  try {
    switch (template) {
      case 'dcf':
        createDCFModel(outputFile);
        break;
      case 'three-statement':
      case '3s':
        createThreeStatementModel(outputFile);
        break;
      case 'sensitivity':
      case 'sens':
        createSensitivityAnalysis(outputFile);
        break;
      case 'demo':
        require('./create-demo.js');
        break;
      default:
        console.error(`\n❌ 未知模板类型: ${template}\n`);
        printUsage();
        process.exit(1);
    }

    console.log('✅ 创建完成!\n');
    console.log('💡 提示: 用 Excel 或其他表格软件打开文件，所有公式都已配置好！\n');

  } catch (error) {
    console.error(`\n❌ 错误: ${error.message}\n`);
    process.exit(1);
  }
}

if (require.main === module) {
  main();
}

module.exports = ExcelBuilder;
