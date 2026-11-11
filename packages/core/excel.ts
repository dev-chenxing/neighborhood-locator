import path from "node:path";
import ExcelJS from "exceljs";
import { 镇街字段, 非镇街字段, 地址列, 其他表头字段 } from "./config";

function assertXlsxPath(filePath: string) {
  if (path.extname(filePath).toLowerCase() !== ".xlsx") {
    throw new Error(`Only .xlsx files are supported: ${filePath}`);
  }
}

export function getWorksheetByIndex(workbook: ExcelJS.Workbook, sheetIndex = 0) {
  const worksheet = workbook.worksheets[sheetIndex];
  if (!worksheet) {
    throw new Error(`Sheet index out of range: ${sheetIndex}`);
  }
  return worksheet;
}

export async function loadWorkbook(filePath: string) {
  assertXlsxPath(filePath);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);
  return workbook;
}

export async function saveWorkbook(workbook: ExcelJS.Workbook, filePath: string) {
  assertXlsxPath(filePath);
  await workbook.xlsx.writeFile(filePath);
}

/**
 * 获取单元格的显示文本。
 * 支持富文本、超链接和公式等单元格值。
 */
export function getCellText(cell: ExcelJS.Cell): string {
  return cell.text.trim();
}

type 表头匹配结果 = {
  字段: string;
  优先级: number;
};

type 表头候选 = 表头匹配结果 & {
  列索引: number;
};

function 匹配表头(文本: string, 字段列表: string[]): 表头匹配结果 | null {
  for (const [优先级, 字段] of 字段列表.entries()) {
    if (文本.includes(字段)) {
      return { 字段, 优先级 };
    }
  }
  return null;
}

function 挑选最优表头(候选列表: 表头候选[]) {
  return 候选列表.reduce<表头候选 | null>((当前最优, 候选) => {
    if (
      !当前最优 ||
      候选.优先级 < 当前最优.优先级 ||
      (候选.优先级 === 当前最优.优先级 && 候选.列索引 < 当前最优.列索引)
    ) {
      return 候选;
    }
    return 当前最优;
  }, null);
}

export function detectHeaderRow(worksheet: ExcelJS.Worksheet) {
  const 最大检查行数 = 4; // 检查前4行
  for (let 行号 = 1; 行号 <= Math.min(最大检查行数, worksheet.rowCount); 行号++) {
    const 当前行 = worksheet.getRow(行号);
    // 获取当前行的所有单元格文本值
    const 单元格值列表 = Array.from({ length: worksheet.columnCount }, (_, 索引) =>
      getCellText(当前行.getCell(索引 + 1)),
    );

    const 镇街候选列表: 表头候选[] = [];
    const 地址候选列表: 表头候选[] = [];
    let 已匹配表头类别数 = 0;

    // 统计命中的表头字段数量
    单元格值列表.forEach((单元格值, 列索引) => {
      if (!单元格值) return; // 跳过空单元格

      // 检查是否是镇街字段
      const 非镇街匹配 = 匹配表头(单元格值, 非镇街字段);
      const 镇街匹配 = 匹配表头(单元格值, 镇街字段);
      if (!非镇街匹配 && 镇街匹配) {
        镇街候选列表.push({ 列索引: 列索引 + 1, ...镇街匹配 });
      }

      // 检查是否是地址列字段
      const 地址匹配 = 匹配表头(单元格值, 地址列);
      if (地址匹配) {
        地址候选列表.push({ 列索引: 列索引 + 1, ...地址匹配 });
      }

      // 检查是否是其他表头字段
      if (匹配表头(单元格值, 其他表头字段)) {
        已匹配表头类别数++;
      }
    });

    // 镇街列和地址列分别只计为一个表头类别，避免重复候选改变判定阈值。
    if (镇街候选列表.length > 0) 已匹配表头类别数++;
    if (地址候选列表.length > 0) 已匹配表头类别数++;

    // 如果命中2个及以上表头字段，且找到了镇街字段，则认为是表头行
    const 镇街候选 = 挑选最优表头(镇街候选列表);
    const 地址候选 = 挑选最优表头(地址候选列表);

    if (已匹配表头类别数 >= 2 && 镇街候选) {
      return {
        headerRowNum: 行号,
        streetColIndex: 镇街候选.列索引,
        addressColIndex: 地址候选?.列索引 ?? -1,
      }; // 返回表头行号和镇街列索引
    }
  }
  return { headerRowNum: -1, streetColIndex: -1, addressColIndex: -1 }; // 未找到表头行
}
