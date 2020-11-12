const { PDFDocument, rgb, StandardFonts } = require("pdf-lib");
//Use ExcelJS to Modify ACH.xlsx File
const Excel = require("exceljs");
const tlFileName = process.argv[2];
const workbook = new Excel.Workbook();
//Use fs-extra package to make Copy and delete Copy
//ExcelJS not support read and write a file at same time. I have to create a copy which is ACH_bk.xlsx, read and modify that copy, and then write into the original file.
const fse = require("fs-extra");
//Parse the Payment Check List file
const pdf = require("pdf-extraction");
//Actually, I can just use one dependent(fs-extra) to deal with file system operation like copy paste, read and write.
const fs = require("fs");
const { match } = require("assert");
//tl.pdf !!! what a great name!!!
// Parse pdf file
let dataBuffer = fs.readFileSync(`${tlFileName}.pdf`);
pdf(dataBuffer).then((data) => {
  const rawArr = data.text.split("\n");
  console.log(rawArr);
  // use array to store the [['10/07/2020',36.56],[transactioin date,transaction amount]...]
  const transactionInfo = [];
  rawArr.forEach((el, i) => {
    // get info around SETTLEMENT line. previous line get transactioin day and next line get transaction amount
    // the info around SETTLEMENT is about Debit ACH
    if (el.trim() === "SETTLEMENT") {
      let tDay = rawArr[i - 1].split(" ")[2];
      let tAmount = rawArr[i + 1].split(" ")[0];
      transactionInfo.push([tDay, tAmount]);
      // Line start with account number is about the EFT Debit
      // tDay and tAmount is in the same line
    } else if (el.startsWith("211367350") && el.length > 45) {
      // must use trim(), because of some line ending with empty space
      let tInfoArr = rawArr[i].trim().split(" ");
      let tDay = tInfoArr[2];
      let tAmount = tInfoArr[tInfoArr.length - 4];
      transactionInfo.push([tDay, tAmount]);
    }
  });

  console.log(
    "The result of Bank Transaction List Parsing is ",
    transactionInfo
  );
  // After pasing the transactioin list and collect all the data I need.
  // Now it is time to transfer the collected information to excel
  async function modifyExcel() {
    fse.copySync(
      `\\\\10.101.1.240\\ftafs\\SUP\\FIN\\HELEN\\Important\\ACH.xlsx`,
      `\\\\10.101.1.240\\ftafs\\Scan\\H36360\\ach_bk\\ACH_bk.xlsx`
    );
    await workbook.xlsx.readFile(
      `\\\\10.101.1.240\\ftafs\\Scan\\H36360\\ach_bk\\ACH_bk.xlsx`
    );
    const ws = workbook.getWorksheet("ap");
    // Extract the transaction amount(amountArr) from the transaction list
    // Extract the transaction Day(tDay) from the transaction list and reformate the date
    let amountArr = [];
    let tDay = transactionInfo[0][0].split("/").join("");
    let td1 = tDay.slice(4);
    let td2 = tDay.slice(0, 4);
    tDay = `${td1}${td2}`;
    transactionInfo.forEach((el) => {
      let a = parseInt(parseFloat(el[1].replace(/,/g, "")) * 100) / 100;
      amountArr.push(a);
    });
    let totalRow = ws.rowCount;
    // Search the matching payday from localPayDay column
    const localPDCol = ws.getColumn("F");
    // all variable start with xy means location in Excel
    const xyAmountMatchPool = [];
    // This functioin is used for filter out all the UNPAID AMOUNT in ACH file. Store in unpaidAmountList
    // And, list out all the unpaid amount ROW NUMBER. store in xyAmountMatchPool
    function amountMatchPoolCreate() {
      let unpaidAmountList = [];
      localPDCol.eachCell((cell, rowNumber) => {
        // As long as there is no content in Transaction Day Column, The amount number will be added in unpaidAmountList array
        if (ws.getCell(`E${rowNumber}`).value === null) {
          unpaidAmountList.push(ws.getCell(`G${rowNumber}`).value);
          xyAmountMatchPool.push(rowNumber);
        }
      });

      return unpaidAmountList;
    }
    // Feed this function an array including all the unpaid amount(amountList)
    //and the Matching target from the amount number from transaction list which is amountArr
    function amountSumMatch(arr, target) {
      function powerset(arr) {
        var ps = [[]];
        for (var i = 0; i < arr.length; i++) {
          for (var j = 0, len = ps.length; j < len; j++) {
            ps.push(ps[j].concat(arr[i]));
          }
        }
        return ps;
      }

      function sum(arr) {
        var total = 0;
        for (var i = 0; i < arr.length; i++) total += arr[i];
        return total;
      }

      function findSum(numbers, targetSum) {
        var numberSets = powerset(numbers);
        for (var i = 0; i < numberSets.length; i++) {
          var numberSet = numberSets[i];
          if (sum(numberSet) == targetSum) return numberSet;
        }
      }

      return findSum(arr, target);
    }
    // just get unpaidAmountList
    const aMP = amountMatchPoolCreate();
    console.log(tDay);
    console.log(aMP);
    console.log(amountArr);
    amountArr.forEach((el) => {
      console.log(`for${el}`);
      let matchedAmountArr = amountSumMatch(aMP, el);
      console.log(matchedAmountArr);
      // Sometime , Because of Auto Deducted payment type, the creation of payment application is later than the actual transaction
      // The result is there will be no matching payment application  for one bank transaction.
      // So, I still need verity after tl.pdf modified.
      if (matchedAmountArr) {
        matchedAmountArr.forEach((el) => {
          let r = xyAmountMatchPool[aMP.indexOf(el)];
          console.log(
            "the Row Number need edit is ",
            xyAmountMatchPool[aMP.indexOf(el)]
          );
          ws.getCell(`E${r}`).value = tDay;
          workbook.xlsx.writeFile(
            `\\\\10.101.1.240\\ftafs\\SUP\\FIN\\HELEN\\Important\\ACH.xlsx`
          );
          //   fse.removeSync("Ach_bk.xlsx");
          //***************** */
          // Modify Transaction List
          //**************** */
          async function modifyTL(x0, y0) {
            const existingPdfBytes = fse.readFileSync("tl.pdf");
            const pdfDoc = await PDFDocument.load(existingPdfBytes);
            const helveticaFont = await pdfDoc.embedFont(
              StandardFonts.Helvetica
            );
            const pages = pdfDoc.getPages();
            const modifyPage = pages[0];
            const { width, height } = firstPage.getSize();
            // modifyPage.drawText(ws)
          }
        });
      }
    });
  }
  modifyExcel();
});
