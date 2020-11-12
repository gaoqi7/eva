// Only xlsx file can be read and write. This limition is adopted by exceljs.
const Excel = require("exceljs");

const workbook = new Excel.Workbook();

function Ap(
  seq,
  reqDate,
  evaPayDay,
  amount,
  localPayDay,
  localPayType,
  chequeNum,
  vendorCode,
  vendorName
) {
  this.seq = seq;
  this.reqDate = reqDate;
  this.evaPayDay = evaPayDay;
  this.amount = amount;
  this.localPayDay = localPayDay;
  this.localPayType = localPayType;
  this.chequeNum = chequeNum;
  this.vendorCode = vendorCode;
  this.vendorName = vendorName;
}

const apList = [];

async function parseAP(path) {
  const seqList = [];
  await workbook.xlsx.readFile(path);
  const worksheet = workbook.getWorksheet(1);
  const seqNoCol = worksheet.getColumn(1);
  const r011 = worksheet.getColumn("R");
  const requestDate = worksheet.getCell(`D8`).value;
  // Reference Point should be sequence number
  r011.eachCell((cell, rowNumber) => {
    if (cell.value === "DATE:" && rowNumber < 133) {
      worksheet.spliceRows(rowNumber - 1, 10);
    }
  });
  seqNoCol.eachCell((cell, rowNumber) => {
    if (cell.value && cell.value.startsWith("0") && rowNumber < 133) {
      seqList.push([cell.value, rowNumber]);
    }
  });
  console.log(seqList);

  seqList.forEach((el) => {
    const seq = el[0];
    const reqDate = requestDate;
    const evaPayDay = worksheet.getCell(`N${el[1]}`).value;
    const amount = worksheet.getCell(`R${el[1]}`).value;
    const localPayType = worksheet.getCell(`D${el[1] + 5}`).value;
    const localPayDay = worksheet.getCell(`J${el[1] + 5}`).value;
    const chequeNum = worksheet.getCell(`M${el[1] + 5}`).value;
    const vendorCode = worksheet.getCell(`C${el[1] + 7}`).value;
    const vendorName = worksheet.getCell(`E${el[1] + 7}`).value;
    const i = new Ap(
      seq,
      reqDate,
      evaPayDay,
      amount,
      localPayDay,
      localPayType,
      chequeNum,
      vendorCode,
      vendorName
    );
    apList.push(i);
  });
  console.log(apList[2].amount);
}

parseAP("./excelFile/7ap-cv.xlsx");
