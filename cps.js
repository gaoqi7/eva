const fs = require("fs");
// The package below extraction the text from pdf.
const pdf = require("pdf-extraction");
const pdfFiles = fs.readdirSync("./pdf").map((el) => `./pdf/${el}`);
const totalPdfFiles = pdfFiles.length
console.log(totalPdfFiles)
//The Package Below can be use to CREATE a new pdf which shows the cps result in this project
// const {jsPDF} = require("jspdf")
//Use the package below for merging all pdf files. Make it easy to print
const merge = require('easy-pdf-merge');

const { PDFDocument, rgb, StandardFonts } = require('pdf-lib');

var cd = [];
let unCashOutCheckAmount = 0
console.log(pdfFiles);
// use this function to accumulate day's amount change for each account
function accountChangeSum(arr) {
  const result = arr.reduce((r, a) => a.map((b, i) => (r[i] || 0) + b), []);;
  cd = result.map(el=>(el/100).toString())
  console.log(cd)
}

// Get the Start Amount
function getOriginalAmount(){
  let oAS = []
  let dataBuffer = fs.readFileSync('./pdf/1.pdf')
  pdf(dataBuffer).then(data=>{
    let bss = data.text
        .split("\n")
        .filter((el) => el.endsWith("USD"))
        .map((el) => el.replace(/,/g, ""));
        console.log(bss)
     bss.map(ele=>{
       const oA = parseInt(parseFloat(ele.trim().split("  ")[3].split(' ')[0])*100)
      //  console.log(oA) 
       oAS.push(oA)
     })
     cd.push(oAS)
     console.log("this is original cd")
     console.log(cd)
  })
}




//Deal with the raw data and get the data I need for CPS, which is Debit amount and Credit amount
async function getDCSummary() {

  await getOriginalAmount()
  for (const pdfFile of pdfFiles) {
    let dataBuffer = fs.readFileSync(pdfFile);
    await pdf(dataBuffer).then((data) => {
      const bankStatementSummary = data.text
        .split("\n")
        .filter((el) => el.endsWith("USD"))
        .map((el) => el.replace(/,/g, ""));
      const bss_trim = bankStatementSummary.map((el) =>
        el
          .split("  ")
          .slice(1, 3)
          .map((ele) => (ele = parseFloat(ele)))
      );
      const bss_sum = bss_trim.map(
        // *100
        (el) => parseInt(-100 * el[0] + 100 * el[1])
      );
      cd.push(bss_sum);
    });
  }
  await accountChangeSum(cd);
  // await createResultPDF()
  await mergePDF();
}



// function createResultPDF(){

//   const doc = new jsPDF({
//     orientation: "landscape",
//     unit: "in",
//     format: [11, 8.5]
//   });
  
//   // doc.text(cd[0], 9.5, 2.875);
//   doc.text(cd[1], 9.5, 2.875);
//   doc.text(cd[2], 9.5, 3.125);
//   doc.text(cd[3], 9.5, 3.75);
//   doc.text(cd[4], 9.5, 4);
//   doc.text(cd[5], 9.5, 4.63);
//   doc.save("./pdf/forReview.pdf");


// }

function mergePDF(){

const pdfPath =fs.readdirSync('./pdf').map(el=>`./pdf/${el}`)
merge(pdfPath, './pdf/forPrint.pdf', function (err) {
    if (err) {
        return console.log(err)
    }
    console.log('Success')
    modifyLastPage()
});
}

getDCSummary();

async function modifyLastPage(){
  const existingPdfBytes = await fs.readFileSync('./pdf/forPrint.pdf')
  // console.log(existingPdfBytes)
  const pdfDoc = await PDFDocument.load(existingPdfBytes)

  // Embed the Helvetica font
const helveticaFont = await pdfDoc.embedFont(StandardFonts.Helvetica)
  const pages = pdfDoc.getPages()
  const firstPage = pages[totalPdfFiles-1]

  // Get the width and height of the first page
const { width, height } = firstPage.getSize()
console.log(width)
console.log(height)
// Draw a string of text diagonally across the first page
firstPage.drawText(cd[0], {
  x: 760,
  y: 458,
  size: 12,
  font: helveticaFont,
  color: rgb(0.95, 0.1, 0.1),
})
firstPage.drawText(cd[1], {
  x: 760,
  y: 403,
  size: 12,
  font: helveticaFont,
  color: rgb(0.95, 0.1, 0.1),
})
firstPage.drawText(cd[2], {
  x: 760,
  y: 388,
  size: 12,
  font: helveticaFont,
  color: rgb(0.95, 0.1, 0.1),
})
firstPage.drawText(cd[3], {
  x: 760,
  y: 335,
  size: 12,
  font: helveticaFont,
  color: rgb(0.95, 0.1, 0.1),
})
firstPage.drawText(cd[4], {
  x: 760,
  y: 320,
  size: 12,
  font: helveticaFont,
  color: rgb(0.95, 0.1, 0.1),
})
firstPage.drawText(cd[5], {
  x: 760,
  y: 265,
  size: 12,
  font: helveticaFont,
  color: rgb(0.95, 0.1, 0.1),
})


const pdfBytes = await pdfDoc.save()
await fs.writeFileSync('./pdf/forPrint.pdf', pdfBytes)


}

// This part is used for parse the r710 report. 
function r710(){
  let r710Buffer = fs.readFileSync('./r710/710.pdf')
  pdf(r710Buffer).then(data => {
    // data.text.split('\n')
    const textArr = data.text.split('\n')
    // console.log(textArr[textArr.length-2])
    unCashOutCheckAmount =  parseFloat(textArr[textArr.length-2].replace(',',''))
    console.log(unCashOutCheckAmount)
  })
}

// r710()