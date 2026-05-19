const richConsle = require("rich-console");
const yargs = require("yargs");
const { hideBin } = require("yargs/helpers");
const { 匹配所属社区 } = require("./resolver");
const {
  createWorkbook,
  getWorksheetByIndex,
  loadWorkbook,
  saveWorkbook,
  worksheetToMatrix,
  worksheetToObjects,
  writeMatrixToWorksheet,
  writeObjectsToWorksheet,
} = require("./excel");

const args = yargs(hideBin(process.argv))
  .demandCommand(1, "Missing arguments: <input.xlsx>")
  .demandCommand(1, "Missing arguments: <address_column>")
  .option("create-copy", {
    alias: "c",
    describe: "Specify whether to create a copy or not",
    type: "boolean",
    default: true,
  })
  .option("delete-unknown", {
    alias: "d",
    describe: "Specify whether to delete unknown address or not",
    type: "boolean",
    default: false,
  })
  .option("log-level", {
    alias: "l",
    describe: "Set the logging level for console",
    default: "INFO",
    choices: ["INFO", "ERROR"],
  })
  .option("reorder", {
    alias: "r",
    describe: "Specify whether to group and reorder or not",
    type: "boolean",
    default: true,
  })
  .option("sheet-index", {
    alias: "s",
    describe: "Specify the index of sheet to process",
    type: "number",
    default: 0,
  })

  .parse();

const logNeighborhoodResult = (address, neighborhoodName, logLevel) => {
  if (neighborhoodName) {
    if (logLevel == "INFO")
      richConsle.log(`<blue>${neighborhoodName}</blue> <green>${address}</green>`);
  } else {
    richConsle.log(`<red>未知</red> <green>${address}</green>`);
  }
};

async function main() {
  const [excel_file, address_column, neighborhood_col_name] = args._;
  const { logLevel, sheetIndex, createCopy, reorder, deleteUnknown } = args;
  void logLevel;

  const workbook = await loadWorkbook(excel_file);
  const worksheet = getWorksheetByIndex(workbook, sheetIndex);
  const sheet_name = worksheet.name;
  const { headers, rows } = worksheetToObjects(worksheet);

  const new_worksheet_json = [];
  let finished_count = 0;
  const total_count = rows.length;

  for (const row of rows) {
    const address = row[address_column];
    if (!address) richConsle.log(`<red>${JSON.stringify(row)}</red>`);
    const neighborhood = 匹配所属社区(String(address || ""));
    logNeighborhoodResult(String(address || ""), neighborhood, logLevel);
    if (neighborhood) {
      finished_count = finished_count + 1;
    }
    if (!deleteUnknown || neighborhood) {
      row[neighborhood_col_name] = neighborhood;
      new_worksheet_json.push(row);
    }
  }

  richConsle.log(`
  已分居委<cyan>${finished_count}</cyan>条，共<cyan>${total_count}</cyan>条，完成率<cyan>${(total_count ===
  0
    ? 0
    : (finished_count / total_count) * 100
  ).toFixed(2)}%</cyan>
  `);

  if (reorder) {
    new_worksheet_json.sort((row1, row2) => (row1[address_column] < row2[address_column] ? 1 : -1));
    new_worksheet_json.sort((row1, row2) =>
      row1[neighborhood_col_name] < row2[neighborhood_col_name] ? 1 : -1,
    );
  }

  const outputHeaders = headers.includes(neighborhood_col_name)
    ? headers
    : [...headers, neighborhood_col_name];

  if (createCopy) {
    const outputWorkbook = createWorkbook();
    const outputWorksheet = outputWorkbook.addWorksheet(sheet_name);
    writeObjectsToWorksheet(outputWorksheet, outputHeaders, new_worksheet_json);
    const new_excel_file = excel_file.replace(/\.xlsx$/i, " - 副本.xlsx");
    await saveWorkbook(outputWorkbook, new_excel_file);
  } else {
    const outputWorkbook = createWorkbook();
    workbook.worksheets.forEach((sourceSheet, index) => {
      const outputWorksheet = outputWorkbook.addWorksheet(sourceSheet.name);
      if (index === sheetIndex) {
        writeObjectsToWorksheet(outputWorksheet, outputHeaders, new_worksheet_json);
        return;
      }
      writeMatrixToWorksheet(outputWorksheet, worksheetToMatrix(sourceSheet));
    });
    await saveWorkbook(outputWorkbook, excel_file);
  }
}

main().catch((error) => {
  console.error(error instanceof Error ? error.message : String(error));
  process.exitCode = 1;
});
