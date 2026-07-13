import path from "node:path";
import ExcelJS from "exceljs";
import { 镇街字段, 地址列, 其他表头字段 } from "./config";

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

function findMatch(text: string, fields: string[]) {
  let bestMatch: string | null = null;
  for (const field of fields) {
    if (text.includes(field)) {
      if (!bestMatch || field.length > bestMatch.length) {
        bestMatch = field;
      }
    }
  }
  return bestMatch;
}

export function detectHeaderRow(worksheet: ExcelJS.Worksheet) {
  for (let rowNum = 1; rowNum <= Math.min(3, worksheet.rowCount); rowNum++) {
    const row = worksheet.getRow(rowNum);
    // 获取当前行的所有单元格文本值
    const cellValues = (row.values as ExcelJS.CellValue[]).map((val) =>
      val !== undefined ? String(val).trim() : "",
    );

    let streetColIndex = -1;
    let addressColIndex = -1;
    let matchedHeaderCount = 0;

    // 统计命中的表头字段数量
    cellValues.forEach((cellValue, colIndex) => {
      if (!cellValue) return; // 跳过空单元格

      // 检查是否是镇街字段
      if (findMatch(cellValue, 镇街字段)) {
        if (streetColIndex === -1) {
          streetColIndex = colIndex;
          matchedHeaderCount++;
        }
      }

      // 检查是否是地址列字段
      if (findMatch(cellValue, 地址列)) {
        if (addressColIndex === -1) {
          addressColIndex = colIndex;
          matchedHeaderCount++;
        }
      }

      // 检查是否是其他表头字段
      if (findMatch(cellValue, 其他表头字段)) {
        matchedHeaderCount++;
      }
    });

    // 如果命中2个及以上表头字段，且找到了镇街字段，则认为是表头行
    if (matchedHeaderCount >= 2 && streetColIndex !== -1) {
      return { headerRowNum: rowNum, streetColIndex, addressColIndex }; // 返回表头行号和镇街列索引
    }
  }
  return { headerRowNum: -1, streetColIndex: -1, addressColIndex: -1 }; // 未找到表头行
}
