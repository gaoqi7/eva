const Excel = require("exceljs");
const workbook = new Excel.Workbook();

async function getExcelReady() {
  await workbook.xlsx.readFile("收入比較-202010.xlsx");
  workbook.eachSheet(async function (worksheet, sheetId) {
    // total Row
    let ttlR = worksheet.rowCount;

    console.log("total row ", ttlR);
    for (i = 2; i < ttlR; i++) {
      let d = new Date();
      b_month = d.getMonth();
      if (i === 2) {
        let crt = [];
        crt.push({ text: `${b_month}` });
        crt.push(worksheet.getCell(`C2`).value.richText[1]);
        crt.push({
          font: {
            size: 16,
            color: { theme: 1 },
            name: "Times New Roman",
            family: 1,
          },
          text: "(B)",
        });
        worksheet.getCell(`C2`).value.richText = crt;

        let rt = [];
        rt.push({ text: `${b_month + 1}` });
        rt.push(worksheet.getCell(`B2`).value.richText[1]);
        rt.push({
          font: {
            size: 16,
            color: { theme: 1 },
            name: "Times New Roman",
            family: 1,
          },
          text: "(A)",
        });

        worksheet.getCell(`B2`).value.richText = rt;
      } else {
        worksheet.getCell(`C${i}`).value = worksheet.getCell(`B${i}`).value;
      }
    }
  });
  return workbook.xlsx.writeFile("收入比較-202010.xlsx");
}
getExcelReady();
