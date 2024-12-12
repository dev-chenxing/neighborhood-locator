const XLSX = require("xlsx");

const excel_file = process.argv[2];
const district = process.argv[3];
const filter_street = process.argv[4];
const new_excel_file = excel_file.replace(district, filter_street);

const workbook = XLSX.readFile(excel_file);
const sheet_name = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheet_name];
const worksheet_json = XLSX.utils.sheet_to_json(worksheet);

let new_worksheet_json = [];

for (const row of worksheet_json) {
  if (row["镇街"].includes(filter_street)) {
    new_worksheet_json.push({
      统一社会信用代码: row["统一社会信用代码"],
      主体名称: row["主体名称"],
      住所: row["住所"],
    });
  }
}

let new_workbook = XLSX.utils.book_new();
const new_worksheet = XLSX.utils.json_to_sheet(new_worksheet_json);
XLSX.utils.book_append_sheet(new_workbook, new_worksheet, "主体名单");
XLSX.writeFile(new_workbook, new_excel_file);

if (process.argv[5] == "-o") console.log(new_excel_file);
