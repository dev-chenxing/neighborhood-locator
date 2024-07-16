const XLSX = require("xlsx");
const fs = require("fs");
const richConsle = require("rich-console");
const yargs = require("yargs");
const { hideBin } = require("yargs/helpers");

const args = yargs(hideBin(process.argv))
    .demandCommand(1, "Missing arguments: <input.xlsx>")
    .demandCommand(1, "Missing arguments: <address_column>")
    .option("create-copy", {
        alias: "c",
        describe: "Specify whether to create a copy or not",
        type: "boolean",
        default: false,
    })
    .option("delete-unknown", {
        alias: "d",
        describe: "Specify whether to delete unknown address or not",
        type: "boolean",
        default: true,
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

const [excel_file, address_column, neighborhood_col_name] = args._;
const { logLevel, sheetIndex, createCopy, reorder, deleteUnknown } = args;

const workbook = XLSX.readFile(excel_file);
const sheet_name = workbook.SheetNames[sheetIndex];
const worksheet = workbook.Sheets[sheet_name];
const worksheet_json = XLSX.utils.sheet_to_json(worksheet);

let rawdata = fs.readFileSync("streets.json");
let streets = JSON.parse(rawdata);

const get_address_number = (street, address) => {
    const regexp = new RegExp(`${street}(\\d+)号`);
    let number = address.match(regexp);
    if (number) return Number(number[1]);
    else {
        let match_result = address.match(new RegExp(`${street}(.+)号`));
        if (match_result) {
            number = match_result[1];
            match_result = number.match(/\d+/g);
        } else {
            match_result = address.match(/\d+/g);
        }
        if (match_result) {
            return Number(match_result[0]);
        } else {
            console.log(street, address);
            return 0;
        }
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
                            if (neighborhood["start"] <= address_number && address_number <= neighborhood["end"]) {
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
        if (logLevel == "INFO") richConsle.log(`<blue>${neighborhood_name}</blue> <green>${address}</green>`);
    } else {
        richConsle.log(`<red>未知</red> <green>${address}</green>`);
    }
    return neighborhood_name;
};

let new_worksheet_json = [];
let finished_count = 0;
const total_count = worksheet_json.length;

for (const row of worksheet_json) {
    const address = row[address_column];
    const neighborhood = find_neighborhood(address, streets);
    if (neighborhood) {
        finished_count = finished_count + 1;
    }
    if (!deleteUnknown || neighborhood) {
        row[neighborhood_col_name] = neighborhood;
        new_worksheet_json.push(row);
    }
}

richConsle.log(`
  已分居委<cyan>${finished_count}</cyan>条，共<cyan>${total_count}</cyan>条，完成率<cyan>${((finished_count / total_count) * 100).toFixed(2)}%</cyan>
  `);

if (reorder) {
    new_worksheet_json.sort((row1, row2) => (row1[address_column] < row2[address_column] ? 1 : -1));
    new_worksheet_json.sort((row1, row2) => (row1[neighborhood_col_name] < row2[neighborhood_col_name] ? 1 : -1));
}

if (createCopy) {
    let new_workbook = XLSX.utils.book_new();
    const new_worksheet = XLSX.utils.json_to_sheet(new_worksheet_json);
    XLSX.utils.book_append_sheet(new_workbook, new_worksheet, sheet_name);
    const new_excel_file = excel_file.replace(".xls", " - 副本.xls");
    XLSX.writeFile(new_workbook, new_excel_file);
} else {
    workbook.Sheets[sheet_name] = XLSX.utils.json_to_sheet(new_worksheet_json);
    XLSX.writeFile(workbook, excel_file);
}
