const robot = require("robotjs");
const sleep = require("sleep");
const Excel = require("exceljs");
const workbook = new Excel.Workbook();
// Which workbook
const workbookName = process.argv[2];
// Which worksheet
const worksheetName = process.argv[3];
// How many line need key in this time
const linesNeedKeyIn = process.argv[4];

const seRow = [2, 16];
const seCol = [2, 14];
var ws;

async function getData() {
  // await workbook.xlsx.readFile('ogfn.xlsx')
  await workbook.xlsx.readFile(`${workbookName}.xlsx`);

  ws = await workbook.getWorksheet(`${worksheetName}`);
  // console.log(typeof ws);
  // Giving 6s preparing time, In this period, I need locate the first cell that I want to input
  sleep.sleep(6);
  // The keyin Function
  cool();
}
// First, I need get data
getData();
// console.log(linesNeedKeyIn);
function cool() {
  // r means ROW, Parsing Row Scope is from 2nd to ${linesNeedKeyIn}
  for (r = 2; r < parseInt(linesNeedKeyIn) + 2; r++) {
    // c means Column, Parsing Column Scope is from 2nd to 14th
    for (c = 2; c < 15; c++) {
      let currentRow = ws.getRow(r);
      // console.log(currentRow)
      let data = currentRow.getCell(c).value;
      console.log(data);

      // The 14th Column is For Description
      if (c !== 14) {
        if (data) {
          // robot.typeStringDelayed(data,200)
          robot.typeString(data);
        } else {
          robot.keyTap("delete");
        }
        robot.keyTap("tab");
      } else {
        function dataParser(n) {
          if (currentRow.getCell(n).value) {
            return currentRow.getCell(n).text;
          } else {
            return " ";
          }
        }
        for (n = 15; n < 19; n++) {
          data += dataParser(n);
        }
        // robot.typeStringDelayed(data,200)
        robot.typeString(data);
        robot.keyTap("tab");
        robot.keyTap("tab");
        robot.keyTap("tab");
        robot.keyTap("tab");
      }
    }
  }
}
