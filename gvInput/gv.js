const robot = require("robotjs");
const sleep = require("sleep");
const Excel = require("exceljs");
const workbook = new Excel.Workbook();

const linesNeedKeyIn = process.argv[2];

const seRow = [2, 16];
const seCol = [2, 14];
var ws;

async function getData() {
  // await workbook.xlsx.readFile('ogfn.xlsx')
  await workbook.xlsx.readFile("yyyymmdd.xlsx");

  ws = await workbook.getWorksheet("ymd");
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
