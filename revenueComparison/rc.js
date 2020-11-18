const Excel = require("exceljs");
const workbook = new Excel.Workbook();
async function getExcelReady() {
  await workbook.xlsx.readFile("收入比較-202009.xlsx");
  workbook.eachSheet(function (worksheet, sheetId) {
    // total Row
    let ttlR = worksheet.rowCount;

    console.log("total row ", ttlR);
    for (i = 2; i < ttlR - 1; i++) {
      // console.log("rich text ", worksheet.getCell(`C${i}`).value.richText);
      if (i === 2) {
        worksheet.getCell("C2").value.richText[
          worksheet.getCell(`C2`).value.richText.length - 1
        ].text = "(B)";

        let b_month = parseInt(
          worksheet.getCell(`B2`).value.richText[0].text[0]
        );
        console.log("bbbbbbbbbbbbbbbbbbbbbbbbbbbbbb", 1 + b_month);
        console.log(worksheet.getCell(`B2`).value.richText[0].text[0]);
        if (worksheet.getCell(`B2`).value.richText[0].text[0].length === 1) {
        }
        // console.log(`10${fff.slice(-3)}`);
        // fff = `${1 + b_month}${fff.slice(-(fff.length - 1))}`;

        // console.log("hahahah", 1 + b_month);
        // console.log("B22222222222222", worksheet.getCell(`B2`).value);
      } else {
        // console.log("B's text", worksheet.getCell(`B${i}`).text);

        worksheet.getCell(`C${i}`).value = worksheet.getCell(`B${i}`).value;
        // console.log("did i change", worksheet.getCell(`C${i}`).text);
      }
    }
  });
  return workbook.xlsx.writeFile("收入比較-202009.xlsx");
}
getExcelReady();
