/**
 * 日规划 Markdown 转 Excel 转换器
 * 
 * 格式规范（基于成品文件分析）：
 * - 第1行：标题，楷体20号粗体，合并A1:H1，行高51
 * - 第2行：核心目标，楷体12号粗体，合并A2:H2，行高40
 * - 第3行：空行
 * - 第4行：表头，楷体16号粗体，行高20.4
 * - 第5行起：数据行，楷体12号，行高100
 *   - 日期和星期列：红色粗体
 *   - 其他列：普通黑色
 * - 所有单元格：左对齐，垂直居中，自动换行，细边框
 */

import XLSX from 'xlsx-js-style';
import { saveAs } from 'file-saver';

// 解析结果接口
export interface ParseResult {
  studentName: string;
  dateRange: string;
  coreTarget: string;
  columns: string[];
  dataRows: string[][];
  errors: string[];
  warnings: string[];
}

// 从第一行智能提取学生名字和日期
function extractStudentNameAndDate(firstLine: string): { studentName: string; dateRange: string } {
  // 移除 # 和 ** 符号
  let line = firstLine.replace(/#/g, '').replace(/\*\*/g, '').trim();
  
  // 提取日期：查找 M.D - M.D 或 M月D日 - M月D日 的模式
  const datePatterns = [
    /(\d{1,2}\.\d{1,2}\s*[-–—]\s*\d{1,2}\.\d{1,2})/,
    /(\d{1,2}月\d{1,2}日?\s*[-–—]\s*\d{1,2}月\d{1,2}日?)/,
    /(\d{4}年\d{1,2}月\d{1,2}日\s*[-–—]\s*\d{4}年\d{1,2}月\d{1,2}日)/
  ];
  
  let dateRange = '';
  let studentName = '';
  
  for (const pattern of datePatterns) {
    const dateMatch = line.match(pattern);
    if (dateMatch) {
      dateRange = dateMatch[1].trim();
      // 移除日期，剩下的就是学生名字
      studentName = line.substring(0, dateMatch.index).trim();
      break;
    }
  }
  
  // 清理学生名字
  if (studentName) {
    // 移除"日规划"、"执行表"等词和括号
    studentName = studentName.replace(/日规划|执行表|周规划|规划|计划|[（()）]/g, '').trim();
    // 只保留中文字符
    const chineseMatch = studentName.match(/[\u4e00-\u9fff]+/);
    if (chineseMatch) {
      studentName = chineseMatch[0];
    }
  }
  
  if (!studentName) {
    // 如果找不到名字，尝试提取中文名字
    const chineseMatch = line.match(/[\u4e00-\u9fff]+/);
    if (chineseMatch) {
      studentName = chineseMatch[0].replace(/日规划|执行表|周规划|规划|计划/g, '');
    } else {
      studentName = '学生';
    }
  }
  
  return { studentName, dateRange };
}

// 从文本行中智能提取核心目标
function extractCoreTarget(lines: string[]): string {
  let targetText = '';
  
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    // 匹配"本周核心目标"或"核心目标"
    if (line.includes('核心目标') || line.includes('本周目标')) {
      // 找到了目标行，提取其后的内容
      targetText = line.replace(/\*\*/g, '').replace(/本周核心目标[：:]/g, '').replace(/核心目标[：:]/g, '').replace(/本周目标[：:]/g, '').trim();
      // 如果这一行太短，可能内容在下一行
      if (targetText.length < 20 && i + 1 < lines.length) {
        // 检查下一行是否是表格
        if (!lines[i + 1].includes('|')) {
          targetText += ' ' + lines[i + 1].replace(/\*\*/g, '').trim();
        }
      }
      break;
    }
  }
  
  // 清理文本
  targetText = targetText.split(/\s+/).join(' ');
  return targetText;
}

// 解析 Markdown 表格
function parseMarkdownTable(markdownText: string): { columns: string[]; dataRows: string[][] } {
  const lines = markdownText.trim().split('\n');
  
  // 寻找并提取表头 - 查找包含"日期"或"星期"的行
  const headerLine = lines.find(line => line.includes('|') && (line.includes('日期') || line.includes('星期')));
  if (!headerLine) {
    throw new Error('在输入文件中未找到有效的Markdown表格表头（需要包含"日期"或"星期"列）');
  }
  
  const columns = headerLine.split('|').slice(1, -1).map(col => col.trim());
  
  // 提取数据行
  const dataRows: string[][] = [];
  let tableStarted = false;
  
  for (const line of lines) {
    // 检测分隔行（包含 --- 或 ---- 的行）
    if (line.includes('---') && line.includes('|')) {
      tableStarted = true;
      continue;
    }
    if (!tableStarted || !line.includes('|')) {
      continue;
    }
    
    const rowData = line.split('|').slice(1, -1).map(col => col.trim());
    // 检查是否是有效数据行（不是空行，且列数匹配）
    if (rowData.length === columns.length && rowData.some(cell => cell.length > 0)) {
      dataRows.push(rowData);
    }
  }
  
  if (dataRows.length === 0) {
    throw new Error('未从表格中提取到任何数据行');
  }
  
  return { columns, dataRows };
}

// 生成输出文件名
function generateOutputFilename(studentName: string, dateRange: string): string {
  if (!studentName) {
    studentName = '学生';
  }
  
  if (!dateRange) {
    return `${studentName}日规划执行表.xlsx`;
  }
  
  // 清理日期范围中的特殊字符
  const cleanDateRange = dateRange.replace(/[–—]/g, '-').replace(/\s+/g, '');
  
  return `${studentName}日规划执行表_${cleanDateRange}.xlsx`;
}

// 解析 Markdown 内容
export function parseMarkdown(markdownContent: string): ParseResult {
  const errors: string[] = [];
  const warnings: string[] = [];
  
  if (!markdownContent.trim()) {
    return {
      studentName: '',
      dateRange: '',
      coreTarget: '',
      columns: [],
      dataRows: [],
      errors: ['输入内容为空'],
      warnings: []
    };
  }
  
  const lines = markdownContent.trim().split('\n');
  
  // 提取学生名字和日期
  let studentName = '';
  let dateRange = '';
  
  if (lines.length >= 1) {
    const result = extractStudentNameAndDate(lines[0]);
    studentName = result.studentName;
    dateRange = result.dateRange;
    
    if (!studentName) {
      warnings.push('未能从标题中提取学生名字，将使用默认名称');
      studentName = '学生';
    }
    if (!dateRange) {
      warnings.push('未能从标题中提取日期范围');
    }
  }
  
  // 提取核心目标
  const coreTarget = extractCoreTarget(lines);
  if (!coreTarget) {
    warnings.push('未能提取核心目标信息');
  }
  
  // 解析表格
  let columns: string[] = [];
  let dataRows: string[][] = [];
  
  try {
    const tableResult = parseMarkdownTable(markdownContent);
    columns = tableResult.columns;
    dataRows = tableResult.dataRows;
  } catch (e) {
    errors.push(e instanceof Error ? e.message : '表格解析失败');
  }
  
  return {
    studentName,
    dateRange,
    coreTarget,
    columns,
    dataRows,
    errors,
    warnings
  };
}

// 边框样式
const thinBorder = {
  top: { style: 'thin', color: { rgb: '000000' } },
  bottom: { style: 'thin', color: { rgb: '000000' } },
  left: { style: 'thin', color: { rgb: '000000' } },
  right: { style: 'thin', color: { rgb: '000000' } },
};

// 标题样式 - 楷体20号粗体
const titleStyle = {
  font: {
    name: '楷体',
    sz: 20,
    bold: true,
    color: { rgb: '000000' },
  },
  alignment: {
    horizontal: 'left',
    vertical: 'center',
    wrapText: true,
  },
};

// 核心目标样式 - 楷体12号粗体
const coreTargetStyle = {
  font: {
    name: '楷体',
    sz: 12,
    bold: true,
    color: { rgb: '000000' },
  },
  alignment: {
    horizontal: 'left',
    vertical: 'center',
    wrapText: true,
  },
};

// 表头样式 - 楷体16号粗体
const headerStyle = {
  font: {
    name: '楷体',
    sz: 16,
    bold: true,
    color: { rgb: '000000' },
  },
  alignment: {
    horizontal: 'left',
    vertical: 'center',
    wrapText: true,
  },
  border: thinBorder,
};

// 日期/星期样式 - 楷体12号红色粗体
const dateStyle = {
  font: {
    name: '楷体',
    sz: 12,
    bold: true,
    color: { rgb: 'FF0000' },
  },
  alignment: {
    horizontal: 'left',
    vertical: 'center',
    wrapText: true,
  },
  border: thinBorder,
};

// 普通数据样式 - 楷体12号
const normalStyle = {
  font: {
    name: '楷体',
    sz: 12,
    bold: false,
    color: { rgb: '000000' },
  },
  alignment: {
    horizontal: 'left',
    vertical: 'center',
    wrapText: true,
  },
  border: thinBorder,
};

// 创建并下载 Excel 文件
export function downloadExcel(parseResult: ParseResult): string {
  const { studentName, dateRange, coreTarget, columns, dataRows } = parseResult;
  
  // 创建工作簿
  const wb = XLSX.utils.book_new();
  
  // 计算列数
  const colCount = columns.length;
  
  // 准备数据
  const wsData: any[][] = [];
  
  // 第1行：标题
  const titleText = dateRange 
    ? `${studentName}日规划（${dateRange}）执行表`
    : `${studentName}日规划执行表`;
  
  const titleRow: any[] = [{ v: titleText, s: titleStyle }];
  for (let i = 1; i < colCount; i++) {
    titleRow.push({ v: '', s: titleStyle });
  }
  wsData.push(titleRow);
  
  // 第2行：核心目标
  const targetText = `本周核心目标：${coreTarget || ''}`;
  const targetRow: any[] = [{ v: targetText, s: coreTargetStyle }];
  for (let i = 1; i < colCount; i++) {
    targetRow.push({ v: '', s: coreTargetStyle });
  }
  wsData.push(targetRow);
  
  // 第3行：空行
  const emptyRow: any[] = [];
  for (let i = 0; i < colCount; i++) {
    emptyRow.push({ v: '', s: {} });
  }
  wsData.push(emptyRow);
  
  // 第4行：表头
  const headerRow: any[] = columns.map(col => ({ v: col, s: headerStyle }));
  wsData.push(headerRow);
  
  // 数据行（从第5行开始）
  for (const row of dataRows) {
    const dataRow: any[] = row.map((cell, index) => {
      // 前两列（日期和星期）使用红色粗体
      if (index < 2) {
        return { v: cell || '', s: dateStyle };
      }
      // 处理单元格内容中的<br>标签，转换为换行符
      const processedCell = (cell || '').replace(/<br\s*\/?>/gi, '\n');
      return { v: processedCell, s: normalStyle };
    });
    wsData.push(dataRow);
  }
  
  // 创建工作表
  const ws = XLSX.utils.aoa_to_sheet(wsData);
  
  // 设置列宽（基于成品文件）
  const colWidths: { wch: number }[] = [];
  for (let i = 0; i < colCount; i++) {
    if (i === 0) {
      colWidths.push({ wch: 8 }); // 日期列
    } else if (i === 1) {
      colWidths.push({ wch: 8.43 }); // 星期列（默认宽度）
    } else if (i === 2 || i === 4) {
      colWidths.push({ wch: 40 }); // 语文、英语列
    } else if (i === 6) {
      colWidths.push({ wch: 35 }); // 化学列
    } else {
      colWidths.push({ wch: 30 }); // 其他学科列
    }
  }
  ws['!cols'] = colWidths;
  
  // 设置行高（基于成品文件）
  const rowHeights: { hpt: number }[] = [];
  rowHeights.push({ hpt: 51 });  // 第1行：标题
  rowHeights.push({ hpt: 40 });  // 第2行：核心目标
  rowHeights.push({ hpt: 15 });  // 第3行：空行
  rowHeights.push({ hpt: 20.4 }); // 第4行：表头
  // 数据行
  for (let i = 0; i < dataRows.length; i++) {
    rowHeights.push({ hpt: 100 }); // 每天的任务内容
  }
  ws['!rows'] = rowHeights;
  
  // 设置合并单元格
  ws['!merges'] = [
    // 标题合并 A1:最后一列
    { s: { r: 0, c: 0 }, e: { r: 0, c: colCount - 1 } },
    // 核心目标合并 A2:最后一列
    { s: { r: 1, c: 0 }, e: { r: 1, c: colCount - 1 } },
  ];
  
  // 添加工作表到工作簿
  XLSX.utils.book_append_sheet(wb, ws, '日规划执行表');
  
  // 生成文件名
  const filename = generateOutputFilename(studentName, dateRange);
  
  // 生成并下载文件
  const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  const blob = new Blob([wbout], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  saveAs(blob, filename);
  
  return filename;
}

// 示例 Markdown 内容
export const exampleMarkdown = `**徐冰鑫日规划（1.13 - 1.18）执行表**
**本周核心目标：** 语文提升古诗与文言文能力，数学紧跟期末复习解决问题，英语强化多方面训练，物理掌握带电粒子运动题型，化学完成期末复习解决重点问题，生物巩固除遗传外知识。

| 日期 | 星期 | 语文 | 数学 | 英语 | 物理 | 化学 | 生物 |
| ---- | ---- | ---- | ---- | ---- | ---- | ---- | ---- |
| 1.13 | 周一 | 1. 花15-20分钟积累文言文虚实词，完成2道翻译题，记录重点字词和翻译思路<br>2. 用20-30分钟做2-3道诗歌鉴赏主观题，参考群里语文知识体系 | 1. 按学校期末复习进度学习<br>2. 遇问题多问学长学姐，找专项题目训练 | 1. 早中晚各10分钟背诵30个单词<br>2. 完成1套BC阅读，分析错题原因，若BC正确率低做1套AB阅读<br>3. 加做1道语法填空<br>4. 练2句中译英，向学长学姐要资料 | 1. 按学校进度学习物理<br>2. 花20-30分钟做1道带电粒子运动大题和1道选择题，分析解题思路 | 1. 进行期末复习<br>2. 做半套选择题<br>3. 遇问题多问学长学姐，若氧化还原配平有问题看网课链接 | 1. 早读读生物课本<br>2. 做2道除遗传外大题，约30分钟<br>3. 做半套选择题 |
| 1.14 | 周二 | 1. 花15-20分钟积累文言文虚实词，完成2道翻译题<br>2. 用20-30分钟做2-3道诗歌鉴赏主观题 | 1. 按学校期末复习进度学习<br>2. 遇问题多问学长学姐 | 1. 早中晚各10分钟背诵30个单词<br>2. 完成1套BC阅读<br>3. 加做1道语法填空 | 1. 按学校进度学习物理<br>2. 做1道带电粒子运动大题 | 1. 进行期末复习<br>2. 做半套选择题 | 1. 早读读生物课本<br>2. 做2道除遗传外大题 |
| 1.15 | 周三 | 1. 积累文言文虚实词<br>2. 做诗歌鉴赏主观题 | 1. 按学校进度学习<br>2. 专项训练 | 1. 背诵30个单词<br>2. 完成阅读训练 | 1. 按学校进度学习<br>2. 做练习题 | 1. 期末复习<br>2. 做选择题 | 1. 读生物课本<br>2. 做大题 |
| 1.16 | 周四 | 1. 花15-20分钟积累文言文虚实词，完成2道翻译题，记录重点字词和翻译思路<br>2. 用20-30分钟做2-3道诗歌鉴赏主观题，参考群里语文知识体系 | 1. 按学校期末复习进度学习<br>2. 遇问题多问学长学姐，找专项题目训练 | 1. 早中晚各10分钟背诵30个单词<br>2. 完成1套BC阅读，分析错题原因，若BC正确率低做1套AB阅读<br>3. 加做1道语法填空<br>4. 练2句中译英，向学长学姐要资料 | 1. 按学校进度学习物理<br>2. 花20-30分钟做1道带电粒子运动大题和1道选择题，分析解题思路 | 1. 进行期末复习<br>2. 做半套选择题<br>3. 遇问题多问学长学姐，若氧化还原配平有问题看网课链接 | 1. 早读读生物课本<br>2. 做2道除遗传外大题，约30分钟<br>3. 做半套选择题 |
| 1.17 | 周五 | 1. 花15-20分钟积累文言文虚实词，完成2道翻译题，记录重点字词和翻译思路<br>2. 用20-30分钟做2-3道诗歌鉴赏主观题，参考群里语文知识体系 | 1. 按学校期末复习进度学习<br>2. 遇问题多问学长学姐，找专项题目训练 | 1. 早中晚各10分钟背诵30个单词<br>2. 完成1套BC阅读，分析错题原因，若BC正确率低做1套AB阅读<br>3. 加做1道语法填空<br>4. 练2句中译英，向学长学姐要资料 | 1. 按学校进度学习物理<br>2. 花20-30分钟做1道带电粒子运动大题和1道选择题，分析解题思路 | 1. 进行期末复习<br>2. 做半套选择题<br>3. 遇问题多问学长学姐，若氧化还原配平有问题看网课链接 | 1. 早读读生物课本<br>2. 做2道除遗传外大题，约30分钟<br>3. 做半套选择题 |
| 1.18 | 周六 | 1. 快速回顾本周诗歌鉴赏主观题答题情况，整理答题思路和技巧 | - | - | - | - | 1. 抽30分钟左右扫荡生物课本 |`;
