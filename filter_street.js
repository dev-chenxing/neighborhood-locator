const {
  createWorkbook,
  getWorksheetByIndex,
  loadWorkbook,
  saveWorkbook,
  worksheetToObjects,
  writeObjectsToWorksheet,
} = require("./excel");

async function main() {
  const excel_file = process.argv[2];
  const district = process.argv[3];
  const filter_street = process.argv[4];
  const new_excel_file = excel_file.replace(district, filter_street);

  const workbook = await loadWorkbook(excel_file);
  const worksheet = getWorksheetByIndex(workbook, 0);
  const { rows } = worksheetToObjects(worksheet);

  const new_worksheet_json = [];

  for (const row of rows) {
    if (String(row["镇街"] || "").includes(filter_street)) {
      new_worksheet_json.push({
        统一社会信用代码: row["统一社会信用代码"] ?? "",
        主体名称: row["主体名称"] ?? "",
        住所: row["住所"] ?? "",
      });
    }
  }

  const outputWorkbook = createWorkbook();
  const outputWorksheet = outputWorkbook.addWorksheet("主体名单");
  writeObjectsToWorksheet(
    outputWorksheet,
    ["统一社会信用代码", "主体名称", "住所"],
    new_worksheet_json,
  );
  await saveWorkbook(outputWorkbook, new_excel_file);

  if (process.argv[5] === "-o") console.log(new_excel_file);
}

main().catch((error) => {
  console.error(error instanceof Error ? error.message : String(error));
  process.exitCode = 1;
});
