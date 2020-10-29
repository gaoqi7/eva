const robot = require("robotjs");
const sleep = require("sleep");
const Excel = require("exceljs");
const workbook = new Excel.Workbook();
const workbookName = process.argv[2];
const worksheetName = process.argv[3];
const linesNeedKeyIn = process.argv[4];

const seRow = [2, 16];
const seCol = [2, 14];
var ws;

async function getData() {
  // await workbook.xlsx.readFile('ogfn.xlsx')
  await workbook.xlsx.readFile(`${workbookName}.xlsx`);

  ws = await workbook.getWorksheet(`${worksheetName}`);
  console.log(typeof ws);
  sleep.sleep(6);
  cool();
}

getData();
console.log(linesNeedKeyIn);
console.log(typeof linesNeedKeyIn);
function cool() {
  for (r = 2; r < parseInt(linesNeedKeyIn) + 2; r++) {
    for (c = 2; c < 15; c++) {
      let currentRow = ws.getRow(r);
      // console.log(currentRow)
      let data = currentRow.getCell(c).value;
      console.log(data);
      if (c !== 14) {
        if (data) {
          // robot.typeStringDelayed(data,200)
          robot.typeString(data);
        } else {
          robot.keyTap("delete");
        }
        robot.keyTap("tab");
      } else {
        // let data15 = currentRow.getCell(15).value;
        // let data16 = currentRow.getCell(16).value;
        // let data17 = currentRow.getCell(17).value;
        // let data18 = currentRow.getCell(18).value;
        // data = data + data15 + data16 + data17 + data18;
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
