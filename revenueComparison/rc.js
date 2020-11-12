const Excel = require("exceljs");
const workbook = new Excel.Workbook();
async function getExcelReady() {
  await workbook.xlsx.readFile("收入比較-202009.xlsx");
  workbook.eachSheet(function (worksheet, sheetId) {
    // total Row
    let ttlR = worksheet.rowCount;
    for (i = 2; i < worksheet.rowCount; i++) {
      // worksheet.getCell(`C${i}`).richText = worksheet.getCell(`B${i}`).richText;
      console.log("did i change", worksheet.getCell(`C${i}`).value);
      // if (i === 2) {
      //   let a = worksheet.getCell(`C${i}`).richText.replace(/A/g, "B");
      //   console.log(a);
      //   worksheet.getCell(`C${i}`).richText = a;
      // }
    }
    // return workbook.xlsx.writeFile("收入比較-202009.xlsx");
  });
}
getExcelReady();
