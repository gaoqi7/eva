const fs = require("fs");
const pdf = require("pdf-extraction");
const pdfFiles = fs.readdirSync("./pdf").map((el) => `./pdf/${el}`);
const cd = [];
console.log(pdfFiles);

function accountChangeSum(arr) {
  const result = cd.reduce((r, a) => a.map((b, i) => (r[i] || 0) + b), []);
  console.log(result);
}

async function getDCSummary() {
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
        (el) => parseInt(-100 * el[0] + 100 * el[1]) / 100
      );
      cd.push(bss_sum);
    });
  }
  await accountChangeSum(cd);
}

getDCSummary();
