const fse = require("fs-extra");
const pdf = require("pdf-extraction");

const afterFilter = [];
const xyEmployee = [];
const xyManualCheck = [];
let dataBuffer = fse.readFileSync("pp.pdf");

Array.prototype.chunk = function (chunk_size) {
  if (!this.length) {
    return [];
  }

  return [this.slice(0, chunk_size)].concat(
    this.slice(chunk_size).chunk(chunk_size)
  );
};
pdf(dataBuffer).then(function (data) {
  console.log(data.text);
  let rawArr = data.text.split("\n");
  let counter = 0;
  let tc;
  let container = [[]];
  // console.log(rawArr);
  let temp = [];
  rawArr.forEach((el, i) => {
    let v;
    if (el.startsWith("Employee:")) {
      container[counter] = [];
      container[counter].push(el);
      tc = counter;
      if (!isNaN(rawArr[i - 1]) && counter !== 0) {
        v = rawArr[i - 1];
        if (container[counter - 1].indexOf(v) === -1) {
          container[counter - 1].push(v);
        }
      }

      // console.log(tc);
      counter++;
    }
    if (
      (el.startsWith("FEDMEDCARE-ER") || el.startsWith("CASUI-ER")) &&
      !isNaN(rawArr[i + 1])
    ) {
      // console.log(rawArr[i + 1]);
      container[tc].push(rawArr[i + 1]);
    }

    // console.log(container[tc]);
  });
  console.log(container);
});
