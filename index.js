const XLSX = require("xlsx");
const fs = require("fs");

let excel_file = "";
let address_column = "";

if (process.argv.length >= 4) {
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
const worksheet = workbook.Sheets[workbook.SheetNames[0]];
const addresses = XLSX.utils
  .sheet_to_json(worksheet)
  .map((row) => row[address_column]);

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
          if (neighborhood["oddity"] == address_number % 2)
            if (
              neighborhood["start"] <=
              address_number <=
              neighborhood["end"]
            ) {
              neighborhood_name = neighborhood["name"];
              break;
            }
        }
      }
      break;
    }
  }
  return neighborhood_name;
};

for (const address of addresses) {
  console.log(address);
  const neighborhood_name = find_neighborhood(address, streets);
  console.log(neighborhood_name);
}
