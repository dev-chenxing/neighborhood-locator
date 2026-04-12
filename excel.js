const path = require("path");
const ExcelJS = require("exceljs");

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
          : typeof cellValue === "object" && cellValue !== null && "text" in cellValue
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
    rows.push(row.values.slice(1).map((value) => {
      if (value == null) return "";
      if (typeof value === "object" && value !== null && "text" in value) {
        return String(value.text);
      }
      return String(value);
    }));
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
};