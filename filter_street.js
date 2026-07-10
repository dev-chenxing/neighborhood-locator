const path = require("path");
const { detectHeaderRow, loadWorkbook, saveWorkbook } = require("./excel");

async function main() {
  const excelFile = process.argv[2];
  const district = process.argv[3];
  const filterStreet = process.argv[4];

  // 拆分路径，仅修改文件名部分
  const fileDir = path.dirname(excelFile);
  const fileBasename = path.basename(excelFile, path.extname(excelFile));
  const newBasename = fileBasename.replace(district, filterStreet);
  const newExcelFile = path.join(fileDir, `${newBasename}.xlsx`);

  // 加载原工作簿
  let workbook;
  try {
    workbook = await loadWorkbook(excelFile);
  } catch (error) {
    throw new Error(`无法加载Excel文件 ${excelFile}: ${error.message}`);
  }

  let successfulSheets = 0; // 成功处理的工作表计数
  for (const sheet of workbook.worksheets) {
    // 检测表头行
    const { headerRowNum, streetColIndex } = detectHeaderRow(sheet);
    if (streetColIndex === -1) {
      continue; // 未找到镇街列，跳过当前工作表
    }
    if (headerRowNum === -1) {
      continue; // 未找到表头行，跳过当前工作表
    }

    // 倒序删除不符合镇街条件的行
    for (let rowNum = sheet.rowCount; rowNum > headerRowNum; rowNum--) {
      const row = sheet.getRow(rowNum);
      const 镇街列值 = String(row.getCell(streetColIndex).value || "").trim();
      // 不匹配的行删除
      if (!镇街列值.includes(filterStreet)) {
        sheet.spliceRows(rowNum, 1);
      }
    }
    successfulSheets++; // 成功处理的工作表计数加1
  }

  // 全文件无有效表头时给出提示，不保存新文件
  if (successfulSheets === 0) {
    console.warn(`⚠️ ${excelFile} 全部工作表无有效表头，未生成新文件`);
    return;
  } else {
    // 保存为新文件
    await saveWorkbook(workbook, newExcelFile);
    // 输出新文件路径供后续脚本使用
    if (process.argv[5] === "-o") console.log(newExcelFile);
  }
}

main().catch((error) => {
  console.error(error instanceof Error ? error.message : String(error));
  process.exitCode = 1;
});
