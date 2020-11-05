const Excel = require("exceljs");
const workbook = new Excel.Workbook();
async function getExcelReady() {
  await workbook.xlsx.readFile("收入比較-202010.xlsx");
  workbook.eachSheet(function (worksheet, sheetId) {
    // total Row
    let ttlR = worksheet.rowCount;
    for (i = 2; i < worksheet.rowCount; i++) {
      worksheet.getCell(`C${i}`).value = worksheet.getCell(`B${i}`).value;
      console.log("did i change", worksheet.getCell(`C${i}`).value);
      if (i === 2) {
        let a = worksheet.getCell(`C${i}`).text.replace(/A/g, "B");
        console.log(a);
        worksheet.getCell(`C${i}`).value = a;
      }
    }
    return workbook.xlsx.writeFile("收入比較-202010.xlsx");
  });
}
getExcelReady();
