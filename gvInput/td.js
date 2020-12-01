const Excel = require("exceljs");
const workbook = new Excel.Workbook();

async function dataTransfer() {
  let cMonth;
  let cData = [];
  async function dataExtrate() {
    await workbook.xlsx.readFile("td.xlsx");
    const ws = await workbook.getWorksheet(1);
    for (i = 5; i < 8; i++) {
      let subContainer = [];
      subContainer.push(ws.getCell(`B${i}`).value);
      subContainer.push(ws.getCell(`J${i}`).value);
      subContainer.push(ws.getCell(`L${i}`).value.result);
      cData.push(subContainer);
    }
    console.log(cData);
    cMonth = ws.getCell("L2").value.slice(0, 3);
    console.log(cMonth);
  }
  await dataExtrate();
  async function dataRelease() {
    await workbook.xlsx.readFile("ymd.xlsx");
    const ws = await workbook.getWorksheet("td");
    for (i = 0; i < 3; i++) {
      //Month
      ws.getCell(`O${2 * i + 2}`).value = cMonth;
      ws.getCell(`O${2 * i + 3}`).value = cMonth;
      // Doc No.
      ws.getCell(`M${2 * i + 2}`).value = cData[i][0];
      //Amount
      ws.getCell(`J${2 * i + 2}`).value = cData[i][2];
      ws.getCell(`J${2 * i + 3}`).value = cData[i][2];
      //Period From
      ws.getCell(`K${2 * i + 2}`).value = cData[i][1]
        .slice(0, 10)
        .replace(/\//g, "");
      ws.getCell(`K${2 * i + 3}`).value = cData[i][1]
        .slice(0, 10)
        .replace(/\//g, "");
      //Period To
      ws.getCell(`L${2 * i + 2}`).value = cData[i][1]
        .slice(11)
        .replace(/\//g, "");
      ws.getCell(`L${2 * i + 3}`).value = cData[i][1]
        .slice(11)
        .replace(/\//g, "");
    }
    return workbook.xlsx.writeFile("ymd.xlsx");
  }
  await dataRelease();
}
dataTransfer();
