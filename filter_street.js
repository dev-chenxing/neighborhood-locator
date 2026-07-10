const path = require("path");
const { getWorksheetByIndex, loadWorkbook, saveWorkbook } = require("./excel");

const 镇街字段 = ["所属镇街", "所属街道", "行政区划", "街道", "镇街", "所属街"];

// 表头字段，命中2个及以上则认为是表头行
const 其他表头字段 = [
  "序号",
  "企业名称",
  "单位名称",
  "经营者名称",
  "机构名称",
  "主体名称",
  "注册地址",
  "经营场所",
  "机构地址",
  "场所地址",
  "住所",
  "统一社会信用代码",
  "法定代表人",
  "法人姓名",
  "负责人姓名",
  "法定代表人联系电话",
  "电话号码",
];

function detectHeaderRow(worksheet) {
  for (let rowNum = 1; rowNum <= Math.min(3, worksheet.rowCount); rowNum++) {
    const row = worksheet.getRow(rowNum);
    // 获取当前行的所有单元格文本值
    const cellValues = row.values.map((val) =>
      val !== undefined ? String(val).trim() : "",
    );

    let streetColIndex = -1;
    let matchedHeaderCount = 0;

    // 统计命中的表头字段数量
    cellValues.forEach((cellValue, colIndex) => {
      if (!cellValue) return; // 跳过空单元格
      console.log(`检查表头行 ${rowNum} 列 ${colIndex}: ${cellValue}`);

      // 检查是否是镇街字段
      if (镇街字段.includes(cellValue)) {
        if (streetColIndex === -1) {
          streetColIndex = colIndex;
          matchedHeaderCount++;
        }
      }

      // 检查是否是其他表头字段
      if (其他表头字段.includes(cellValue)) {
        matchedHeaderCount++;
      }
    });

    // 如果命中2个及以上表头字段，且找到了镇街字段，则认为是表头行
    if (matchedHeaderCount >= 2 && streetColIndex !== -1) {
      return { rowNum, streetColIndex }; // 返回表头行号和镇街列索引
    }
  }
  return { rowNum: -1, streetColIndex: -1 }; // 未找到表头行
}

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
  const workbook = await loadWorkbook(excelFile);
  const worksheet = getWorksheetByIndex(workbook, 0);

  // 检测表头行
  const { rowNum: headerRowNum, streetColIndex } = detectHeaderRow(worksheet);
  if (streetColIndex === -1) {
    throw new Error(`${excelFile} 未找到镇街字段列，请检查表头`);
  }
  if (headerRowNum === -1) {
    throw new Error(`${excelFile} 未找到表头行，请检查表头`);
  }

  // 倒序删除不符合镇街条件的行
  for (let rowNum = worksheet.rowCount; rowNum > headerRowNum; rowNum--) {
    const row = worksheet.getRow(rowNum);
    const 镇街列值 = String(row.getCell(streetColIndex).value || "").trim();
    // 不匹配的行删除
    if (!镇街列值.includes(filterStreet)) {
      worksheet.spliceRows(rowNum, 1);
    }
  }

  // 保存为新文件
  await saveWorkbook(workbook, newExcelFile);

  // 输出新文件路径供后续脚本使用
  if (process.argv[5] === "-o") console.log(newExcelFile);
}

main().catch((error) => {
  console.error(error instanceof Error ? error.message : String(error));
  process.exitCode = 1;
});
