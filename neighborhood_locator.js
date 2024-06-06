const XLSX = require("xlsx");
const fs = require("fs");
const richConsle = require("rich-console");

let excel_file = "";
let address_column = "";
let print_unknown_only = false;

if (process.argv.length >= 4) {
  if (process.argv.length > 4 && process.argv[4] == "-e")
    print_unknown_only = true;
  excel_file = process.argv[2];
  address_column = process.argv[3];
} else if (process.argv.length == 3) {
  console.error("Missing arguments: <column_name>");
  process.exit(1);
} else if (process.argv.length == 2) {
  console.error("Missing arguments: <input.xlsx> <column_name>");
  process.exit(1);
}

neighborhood_col_name = "所属居委";
neighborhood_column = [];

const workbook = XLSX.readFile(excel_file);
const sheet_name = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheet_name];
const worksheet_json = XLSX.utils.sheet_to_json(worksheet);

let rawdata = fs.readFileSync("streets.json");
let streets = JSON.parse(rawdata);

const get_address_number = (street, address) => {
  const regexp = new RegExp(`${street}(\\d+)号`);
  let number = address.match(regexp);
  if (number) return Number(number[1]);
  else {
    number = address.match(new RegExp(`${street}(.+)号`))[1];
    return Number(number.match(/\d+/g)[0]);
  }
};

const find_neighborhood = (address, streets) => {
  let neighborhood_name = null;
  for (const street in streets) {
    const neighborhoods = streets[street];
    if (address.includes(street)) {
      for (const neighborhood of neighborhoods) {
        if (neighborhood["all"]) {
          neighborhood_name = neighborhood["name"];
          break;
        } else {
          const address_number = get_address_number(street, address);
          if (neighborhood["oddity"] == address_number % 2) {
            if (neighborhood["start"]) {
              if (
                neighborhood["start"] <= address_number &&
                address_number <= neighborhood["end"]
              ) {
                neighborhood_name = neighborhood["name"];
                break;
              }
            } else {
              neighborhood_name = neighborhood["name"];
              break;
            }
          }
        }
      }
      break;
    }
  }
  if (neighborhood_name) {
    if (!print_unknown_only)
      richConsle.log(
        `<blue>${neighborhood_name}</blue> <green>${address}</green>`
      );
  } else {
    richConsle.log(`<red>未知</red> <green>${address}</green>`);
  }
  return neighborhood_name;
};

for (const row of worksheet_json) {
  const address = row[address_column];
  row[neighborhood_col_name] = find_neighborhood(address, streets);
}

worksheet_json.sort((row1, row2) =>
  row1[address_column] < row2[address_column] ? 1 : -1
);

worksheet_json.sort((row1, row2) =>
  row1[neighborhood_col_name] < row2[neighborhood_col_name] ? 1 : -1
);

workbook.Sheets[sheet_name] = XLSX.utils.json_to_sheet(worksheet_json);
XLSX.writeFile(workbook, excel_file);
