const fs = require("fs");
const pdf = require("pdf-extraction");
const pdfFiles = fs.readdirSync("./pdf").map((el) => `./pdf/${el}`);
const {jsPDF} = require("jspdf")

const merge = require('easy-pdf-merge');



var cd = [];
console.log(pdfFiles);

function accountChangeSum(arr) {
  const result = arr.reduce((r, a) => a.map((b, i) => (r[i] || 0) + b), []);;
  cd = result.map(el=>(el/100).toString())
  console.log(cd)
  // console.log(result.map(el=>(el/100).toString()));
}

function getOriginalAmount(){
  let oAS = []
  let dataBuffer = fs.readFileSync('./pdf/1.pdf')
  pdf(dataBuffer).then(data=>{
    let bss = data.text
        .split("\n")
        .filter((el) => el.endsWith("USD"))
        .map((el) => el.replace(/,/g, ""));
      // console.log(bss)
     bss.map(ele=>{
       const oA = parseInt(parseFloat(ele.trim().split("  ")[3].split(' ')[0])*100)
      //  console.log(oA) 
       oAS.push(oA)
     })
     cd.push(oAS)
    //  console.log(cd)
  })
}





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
  await createResultPDF()
  await mergePDF()
}

getDCSummary();

function createResultPDF(){

  const doc = new jsPDF({
    orientation: "landscape",
    unit: "in",
    format: [11, 8.5]
  });
  
  // doc.text(cd[0], 9.5, 2.875);
  doc.text(cd[1], 9.5, 2.875);
  doc.text(cd[2], 9.5, 3.125);
  doc.text(cd[3], 9.5, 3.75);
  doc.text(cd[4], 9.5, 4);
  doc.text(cd[5], 9.5, 4.63);
  doc.save("./pdf/forReview.pdf");


}

function mergePDF(){

const pdfPath = fs.readdirSync('./pdf').map(el=>`./pdf/${el}`)
merge(pdfPath, './pdf/forPrint.pdf', function (err) {
    if (err) {
        return console.log(err)
    }
    console.log('Success')
});
}