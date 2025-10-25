#!/usr/bin/env node

const XLSX = require('xlsx');

/**
 * 创建一个简单的金融建模演示 Excel 文件
 */

console.log('📊 创建演示金融建模 Excel 文件...\n');

// 创建工作簿
const wb = XLSX.utils.book_new();

// ============================================
// 工作表 1: 收入预测
// ============================================
const revenueData = [
  ['收入预测表', '', '', '', '', ''],
  ['', '', '', '', '', ''],
  ['项目', '2023', '2024', '2025', '2026', '增长率'],
  ['产品A销售额', 1000000, 1200000, 1440000, 1728000, '20%'],
  ['产品B销售额', 800000, 920000, 1058000, 1216700, '15%'],
  ['服务收入', 500000, 625000, 781250, 976563, '25%'],
  ['', '', '', '', '', ''],
  ['总收入', 2300000, 2745000, 3279250, 3921263, ''],
];

const revenueSheet = XLSX.utils.aoa_to_sheet(revenueData);

// 添加公式
revenueSheet['B8'] = { t: 'n', f: 'SUM(B4:B6)', v: 2300000 };
revenueSheet['C8'] = { t: 'n', f: 'SUM(C4:C6)', v: 2745000 };
revenueSheet['D8'] = { t: 'n', f: 'SUM(D4:D6)', v: 3279250 };
revenueSheet['E8'] = { t: 'n', f: 'SUM(E4:E6)', v: 3921263 };

// 设置列宽
revenueSheet['!cols'] = [
  { wch: 20 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 10 }
];

XLSX.utils.book_append_sheet(wb, revenueSheet, '收入预测');

// ============================================
// 工作表 2: 成本分析
// ============================================
const costData = [
  ['成本分析表', '', '', '', ''],
  ['', '', '', '', ''],
  ['项目', '2023', '2024', '2025', '2026'],
  ['原材料成本', 600000, 720000, 864000, 1036800],
  ['人工成本', 800000, 880000, 968000, 1064800],
  ['运营成本', 400000, 440000, 484000, 532400],
  ['营销费用', 300000, 360000, 432000, 518400],
  ['', '', '', '', ''],
  ['总成本', 2100000, 2400000, 2748000, 3152400],
  ['', '', '', '', ''],
  ['毛利润', 200000, 345000, 531250, 768863],
  ['毛利率', '8.7%', '12.6%', '16.2%', '19.6%'],
];

const costSheet = XLSX.utils.aoa_to_sheet(costData);

// 添加公式
costSheet['B9'] = { t: 'n', f: 'SUM(B4:B7)', v: 2100000 };
costSheet['C9'] = { t: 'n', f: 'SUM(C4:C7)', v: 2400000 };
costSheet['D9'] = { t: 'n', f: 'SUM(D4:D7)', v: 2748000 };
costSheet['E9'] = { t: 'n', f: 'SUM(E4:E7)', v: 3152400 };

// 毛利润 = 收入 - 成本（引用另一个工作表）
costSheet['B11'] = { t: 'n', f: '收入预测!B8-B9', v: 200000 };
costSheet['C11'] = { t: 'n', f: '收入预测!C8-C9', v: 345000 };
costSheet['D11'] = { t: 'n', f: '收入预测!D8-D9', v: 531250 };
costSheet['E11'] = { t: 'n', f: '收入预测!E8-E9', v: 768863 };

// 毛利率 = 毛利润 / 收入
costSheet['B12'] = { t: 'n', f: 'B11/收入预测!B8', v: 0.087, z: '0.0%' };
costSheet['C12'] = { t: 'n', f: 'C11/收入预测!C8', v: 0.126, z: '0.0%' };
costSheet['D12'] = { t: 'n', f: 'D11/收入预测!D8', v: 0.162, z: '0.0%' };
costSheet['E12'] = { t: 'n', f: 'E11/收入预测!E8', v: 0.196, z: '0.0%' };

costSheet['!cols'] = [
  { wch: 20 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 12 }
];

XLSX.utils.book_append_sheet(wb, costSheet, '成本分析');

// ============================================
// 工作表 3: 现金流预测
// ============================================
const cashFlowData = [
  ['现金流预测表', '', '', '', ''],
  ['', '', '', '', ''],
  ['项目', '2023', '2024', '2025', '2026'],
  ['期初现金', 500000, 600000, 945000, 1476250],
  ['经营现金流入', 2300000, 2745000, 3279250, 3921263],
  ['经营现金流出', -2100000, -2400000, -2748000, -3152400],
  ['投资支出', -200000, -150000, -180000, -216000],
  ['融资活动', 100000, 150000, 180000, 0],
  ['', '', '', '', ''],
  ['期末现金', 600000, 945000, 1476250, 2029113],
  ['现金净变化', 100000, 345000, 531250, 552863],
];

const cashFlowSheet = XLSX.utils.aoa_to_sheet(cashFlowData);

// 添加公式
cashFlowSheet['B10'] = { t: 'n', f: 'SUM(B4:B8)', v: 600000 };
cashFlowSheet['C10'] = { t: 'n', f: 'SUM(C4:C8)', v: 945000 };
cashFlowSheet['D10'] = { t: 'n', f: 'SUM(D4:D8)', v: 1476250 };
cashFlowSheet['E10'] = { t: 'n', f: 'SUM(E4:E8)', v: 2029113 };

cashFlowSheet['B11'] = { t: 'n', f: 'B10-B4', v: 100000 };
cashFlowSheet['C11'] = { t: 'n', f: 'C10-C4', v: 345000 };
cashFlowSheet['D11'] = { t: 'n', f: 'D10-D4', v: 531250 };
cashFlowSheet['E11'] = { t: 'n', f: 'E10-E4', v: 552863 };

cashFlowSheet['!cols'] = [
  { wch: 20 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 12 }
];

XLSX.utils.book_append_sheet(wb, cashFlowSheet, '现金流预测');

// 保存文件
const fileName = 'financial-model-demo.xlsx';
XLSX.writeFile(wb, fileName);

console.log(`✅ 演示文件创建成功: ${fileName}`);
console.log('\n包含以下工作表:');
console.log('  1. 收入预测 - 多年收入预测数据');
console.log('  2. 成本分析 - 成本结构和毛利率计算');
console.log('  3. 现金流预测 - 现金流量表\n');
console.log('现在可以使用以下命令分析:\n');
console.log(`  node analyze-excel.js ${fileName}`);
console.log(`  node analyze-excel.js ${fileName} --all`);
console.log(`  node analyze-excel.js ${fileName} --sheet "成本分析" --formulas\n`);
