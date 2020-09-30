const fs = require("fs");
const pdf = require("pdf-extraction");
const pdfFiles = fs.readdirSync("./pdf").map((el) => `./pdf/${el}`);
const cd = [];
console.log(pdfFiles);

function accountChangeSum() {
  const result = cd.reduce((r, a) => a.map((b, i) => (r[i] || 0) + b), []);;
  result.map(el=>el/100)
  console.log(result.map(el=>el/100));
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
}

getDCSummary();
// getOriginalAmount()