const fs = require("fs");
const XLSX = require("xlsx");
const util = require("util");

function compareLogsWithExcel() {
  try {
    // Read JSON log file
    const jsonData = JSON.parse(
      fs.readFileSync("./logs/json/success_log.json", "utf8")
    );

    // Read Excel file
    const workbook = XLSX.readFile("./memeysel.xlsx");
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
    const excelData = XLSX.utils.sheet_to_json(worksheet);

    // Array to store matching results
    const matches = [];

    // Iterate through Excel rows and check for matches in JSON
    excelData.forEach((row) => {
      const nikMatch = jsonData.find((logEntry) => logEntry.NIK === row.NIK);
      if (nikMatch) {
        matches.push({
          NIK: row.NIK,
          No: row.No,
        });
      }
    });

    return matches;
  } catch (error) {
    console.error("Error comparing logs:", error);
    return [];
  }
}

module.exports = {
  compareLogsWithExcel,
};

// Execute if run directly
if (require.main === module) {
  const matches = compareLogsWithExcel();
  console.log(
    "Matching records:",
    util.inspect(matches, { maxArrayLength: null, depth: null })
  );
  console.log("Total matching records:", matches.length);
}
