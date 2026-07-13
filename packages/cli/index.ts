import { styleText } from "node:util";
import yargs from "yargs";
import { hideBin } from "yargs/helpers";
import { 居委列名 } from "../core/config";
import { detectHeaderRow, getWorksheetByIndex, loadWorkbook, saveWorkbook } from "../core/excel";
import { 匹配所属社区 } from "../core/resolver";
import type { LogLevel } from "../core/types";

const args = yargs(hideBin(process.argv))
  .usage("用法: $0 <企业名单.xlsx> [地址列名] [居委列名] [选项]")
  .demandCommand(1, "缺少必填参数：请输入Excel文件路径")
  .option("create-copy", {
    alias: "c",
    describe: "是否创建副本文件，默认：否",
    type: "boolean",
    default: false,
  })
  .option("delete-unknown", {
    alias: "d",
    describe: "是否删除无法匹配居委的行，默认：否",
    type: "boolean",
    default: false,
  })
  .option("log-level", {
    alias: "l",
    describe: "日志等级，INFO：输出分居委信息，ERROR：仅输出错误信息",
    default: "INFO",
    choices: ["INFO", "ERROR"],
  })
  .option("sheet-index", {
    alias: "s",
    describe: "指定处理的工作表索引，默认：全部工作表（-1）",
    type: "number",
    default: -1,
  })
  .help()
  .parseSync();

const logNeighborhoodResult = (
  address: string,
  neighborhoodName: string | undefined,
  logLevel: LogLevel,
) => {
  if (neighborhoodName) {
    if (logLevel == "INFO")
      console.log(
        styleText("blue", `${neighborhoodName}`) + " " + styleText("green", `${address}`),
      );
  } else {
    console.log(styleText("red", "未知") + " " + styleText("green", `${address}`));
  }
};

async function main() {
  const [excel_file] = args._;
  if (excel_file === undefined || typeof excel_file !== "string") {
    throw new Error("缺少必填参数：请输入Excel文件路径");
  }

  const { logLevel, sheetIndex, createCopy, deleteUnknown } = args;

  const workbook = await loadWorkbook(excel_file);

  // 获取指定工作表
  let sheets = [];
  if (sheetIndex === -1) {
    sheets = workbook.worksheets;
  } else {
    const worksheet = getWorksheetByIndex(workbook, sheetIndex);
    if (!worksheet) {
      throw new Error(`工作表索引 ${sheetIndex} 不存在`);
    }
    sheets = [worksheet];
  }

  let totalFinished = 0;
  let totalCount = 0;

  for (let sheetIndex = 0; sheetIndex < sheets.length; sheetIndex++) {
    const sheet = sheets[sheetIndex];
    if (sheet === undefined) {
      throw new Error(`工作表索引 ${sheetIndex} 不存在`);
    }
    console.log(
      "正在处理工作表" +
        styleText("cyan", `${sheetIndex}`) +
        "：" +
        styleText("cyan", `${sheet.name}`) +
        `（${sheet.rowCount}行，${sheet.columnCount}列）`,
    );
    // 获取表头行和镇街列索引
    const { headerRowNum, streetColIndex, addressColIndex } = detectHeaderRow(sheet);
    console.log(
      `表头行号：` +
        styleText("cyan", `${headerRowNum}`) +
        `，镇街列索引：` +
        styleText("cyan", `${streetColIndex}`) +
        `，地址列索引：` +
        styleText("cyan", `${addressColIndex}`),
    );

    if (headerRowNum === -1 || streetColIndex === -1 || addressColIndex === -1) {
      continue; // 未找到表头行或镇街列，跳过当前工作表
    }

    // 在最后一列添加居委列
    const headerRow = sheet.getRow(headerRowNum);
    const 居委列 = headerRow.cellCount + 1; // 居委列索引为最后一列的下一个索引
    headerRow.getCell(居委列).value = 居委列名;

    // 遍历数据行，匹配所属社区，并将结果写入居委列
    for (let rowNum = headerRowNum + 1; rowNum <= sheet.rowCount; rowNum++) {
      const row = sheet.getRow(rowNum);
      const address = String(row.getCell(addressColIndex).value).trim();
      const neighborhood = address ? 匹配所属社区(address) : "";

      logNeighborhoodResult(address, neighborhood, logLevel !== "ERROR" ? "INFO" : "ERROR");

      if (neighborhood) totalFinished++;
      totalCount++;

      // 将匹配结果写入居委列
      row.getCell(居委列).value = neighborhood;
    }

    // 删除无法匹配居委的行
    if (deleteUnknown) {
      for (let rowNum = sheet.rowCount; rowNum > headerRowNum; rowNum--) {
        const row = sheet.getRow(rowNum);
        const neighborhood = row.getCell(居委列).value;
        if (!neighborhood) {
          sheet.spliceRows(rowNum, 1);
        }
      }
    }
  }

  console.log(
    "\n已分居委" +
      styleText("cyan", `${totalFinished}`) +
      "条，共" +
      styleText("cyan", `${totalCount}`) +
      "条，完成率" +
      styleText(
        "cyan",
        `${totalCount === 0 ? 0 : ((totalFinished / totalCount) * 100).toFixed(2)}%`,
      ),
  );

  if (createCopy) {
    const newExcelFile = excel_file.replace(/\.xlsx$/i, " - 副本.xlsx");
    await saveWorkbook(workbook, newExcelFile);
  } else {
    await saveWorkbook(workbook, excel_file);
  }
}

main().catch((error) => {
  console.error(error instanceof Error ? error.stack : String(error));
  process.exitCode = 1;
});
