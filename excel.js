const path = require("path");
const ExcelJS = require("exceljs");

const 镇街字段 = ["所属镇街", "所属街道", "行政区划", "街道", "镇街", "所属街"];
const 地址列 = ["注册地址", "经营场所", "机构地址", "场所地址", "住所"];
const 居委列名 = "所属社区";

// 表头字段，命中2个及以上则认为是表头行
const 其他表头字段 = [
  "序号",
  "企业名称",
  "单位名称",
  "经营者名称",
  "机构名称",
  "主体名称",
  "统一社会信用代码",
  "法定代表人",
  "法人姓名",
  "负责人姓名",
  "法定代表人联系电话",
  "电话号码",
];

function createWorkbook() {
  return new ExcelJS.Workbook();
}

function assertXlsxPath(filePath) {
  if (path.extname(filePath).toLowerCase() !== ".xlsx") {
    throw new Error(`Only .xlsx files are supported: ${filePath}`);
  }
}

function getWorksheetByIndex(workbook, sheetIndex = 0) {
  const worksheet = workbook.worksheets[sheetIndex];
  if (!worksheet) {
    throw new Error(`Sheet index out of range: ${sheetIndex}`);
  }
  return worksheet;
}

function worksheetToObjects(worksheet) {
  const headerRow = worksheet.getRow(1);
  const headers = headerRow.values
    .slice(1)
    .map((value) => (value == null ? "" : String(value).trim()));

  const rows = [];
  for (let rowNumber = 2; rowNumber <= worksheet.rowCount; rowNumber += 1) {
    const row = worksheet.getRow(rowNumber);
    const record = {};
    let hasValue = false;

    headers.forEach((header, index) => {
      if (!header) return;
      const cellValue = row.getCell(index + 1).value;
      const normalized =
        cellValue == null
          ? ""
          : typeof cellValue === "object" &&
              cellValue !== null &&
              "text" in cellValue
            ? String(cellValue.text)
            : String(cellValue);

      record[header] = normalized;
      if (normalized !== "") hasValue = true;
    });

    if (hasValue) rows.push(record);
  }

  return { headers, rows };
}

function worksheetToMatrix(worksheet) {
  const rows = [];
  worksheet.eachRow({ includeEmpty: true }, (row) => {
    rows.push(
      row.values.slice(1).map((value) => {
        if (value == null) return "";
        if (typeof value === "object" && value !== null && "text" in value) {
          return String(value.text);
        }
        return String(value);
      }),
    );
  });
  return rows;
}

function writeObjectsToWorksheet(worksheet, headers, rows) {
  worksheet.spliceRows(1, worksheet.rowCount || 1);
  worksheet.addRow(headers);
  for (const row of rows) {
    worksheet.addRow(headers.map((header) => row[header] ?? ""));
  }
}

function writeMatrixToWorksheet(worksheet, rows) {
  for (const row of rows) {
    worksheet.addRow(row);
  }
}

async function loadWorkbook(filePath) {
  assertXlsxPath(filePath);
  const workbook = createWorkbook();
  await workbook.xlsx.readFile(filePath);
  return workbook;
}

async function saveWorkbook(workbook, filePath) {
  assertXlsxPath(filePath);
  await workbook.xlsx.writeFile(filePath);
}

function findMatch(text, fields) {
  let bestMatch = null;
  for (const field of fields) {
    if (text.includes(field)) {
      if (!bestMatch || field.length > bestMatch.length) {
        bestMatch = field;
      }
    }
  }
  return bestMatch;
}

function detectHeaderRow(worksheet) {
  for (let rowNum = 1; rowNum <= Math.min(3, worksheet.rowCount); rowNum++) {
    const row = worksheet.getRow(rowNum);
    // 获取当前行的所有单元格文本值
    const cellValues = row.values.map((val) =>
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

module.exports = {
  assertXlsxPath,
  createWorkbook,
  getWorksheetByIndex,
  loadWorkbook,
  saveWorkbook,
  worksheetToMatrix,
  worksheetToObjects,
  writeMatrixToWorksheet,
  writeObjectsToWorksheet,
  镇街字段,
  地址列,
  居委列名,
  其他表头字段,
  detectHeaderRow,
};
