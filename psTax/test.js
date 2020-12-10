const fse = require("fs-extra");
const pdf = require("pdf-extraction");
let dataBuffer = fse.readFileSync("pp.pdf");
pdf(dataBuffer).then(function (data) {
  console.log(data.text);
  let rawArr = data.text.split("\n");
  let counter = 0;
  let tc;
  let container = [[]];
  // console.log(rawArr);
  // Three Cases
  // 1. Employee ... FEDMEDCARE-ER {Amount} Employee
  // 2. Employee ... FEDMEDCARE-ER {Amount} ... CASUI-ER {Amount} Employee
  // 3. Employee ... FEDMEDCARE-ER {Amount} ... CASUI-ER ...(Because of new page) {Amount} Employee
  rawArr.forEach((el, i) => {
    // V as Verify; Verify if it is case 3 - can't get amount because of header on new page.
    let v;

    if (el.startsWith("Employee:")) {
      container[counter] = [];
      container[counter].push(el);
      //tc as temporary counter, used for collect employee's tax
      tc = counter;
      // isNaN means is Not a Number
      // Deal with the case 3.
      if (!isNaN(rawArr[i - 1]) && counter !== 0) {
        v = rawArr[i - 1];
        if (container[counter - 1].indexOf(v) === -1) {
          container[counter - 1].push(v);
        }
      }
      // counter plus one for next employee; for current employee, use tc.
      counter++;
    }
    // Sometime, new employee will have CASUI-ER. 
    if (
      (el.startsWith("FEDMEDCARE-ER") || el.startsWith("CASUI-ER")) &&
      !isNaN(rawArr[i + 1])
    ) {
      container[tc].push(rawArr[i + 1]);
    }
  })
  console.log(container);
});
