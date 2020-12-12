const address = {
  Affemann_AlexeiE: "AF58",
  BailesJr_CharlesV: "AA58",
  Baron_Cody: "AJ58",
  Baxes_CoryM: "AQ58",
  Chang_WeiYu: "O55",
  Chen_Minyi: "H56",
  Chiu_ChihYing: "L54",
  Chiu_YenWen: "T57",
  Chou_YuLin: "AR58",
  Chu_Ge: "D53",
  Davis_RonaldL: "AH58",
  DeSantosDelgadoJr_Humberto: "AO58",
  Gao_Qi: "K54",
  George_AlexanderN: "AP58",
  Gerkus_Michael: "V57",
  Hisato_JohnK: "G56",
  Hsieh_FuChung: "R57",
  Hu_Kevin: "AL58",
  Kamoshita_TeHuaM: "M54",
  Lee_HaoChieh: "F56",
  Lee_JamesJ: "AE58",
  Lombart_MichaelD: "AD58",
  Lopina_BrianA: "AB58",
  McIlwain_DanaJ: "AM58",
  Nagahamulla_DewniY: "AT58",
  Roberts_HaydenE: "AU58",
  Rosensteel_WilliamJ: "W57",
  Share_PhilipA: "AC58",
  Shih_YinCheng: "J54",
  Smith_SarahA: "AG58",
  ThompsonJR_GeorgeT: "AN58",
  Yang_ChiMing: "U57",
  Yoon_Junga: "AI58",
  Zheng_Hongming: "AV58",
};
//===================================================

const fse = require("fs-extra");
const pdf = require("pdf-extraction");
let dataBuffer = fse.readFileSync("pp1.pdf");
let container = [[]];
function getData() {
  pdf(dataBuffer).then(function (data) {
    //   console.log(data.text);
    let rawArr = data.text.split("\n");
    let counter = 0;
    let tc;
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
        container[counter].push(el.slice(9));
        //tc as temporary counter, used for collect employee's tax
        tc = counter;
        // isNaN means is Not a Number
        // Deal with the case 3.
        if (!isNaN(rawArr[i - 1]) && counter !== 0) {
          v = rawArr[i - 1];
          if (
            container[counter - 1].indexOf(v) === -1 &&
            container[counter - 1].length < 3
          ) {
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
    });
    console.log(container);

    container.map((el) => {
      el[0] = el[0].replace(",", "_").replace("-", "");
      if (el.length === 2) {
        el[1] = parseFloat(el[1]);
      } else {
        el[1] =
          (100 * (parseFloat(el[1]) || 0) + 100 * parseFloat(el[2])) / 100;
        el.pop();
      }
    });
    console.log(container);
  });
}
//=========================================
const Excel = require("exceljs");
const workbook = new Excel.Workbook();

async function inputTax() {
  await workbook.xlsx.readFile(`20201215.xlsx`);
  const worksheet = await workbook.getWorksheet(`FIN`);
  container.forEach((el) => {
    worksheet.getCell(address[el[0]]).value = el[1];
  });
  return workbook.xlsx.writeFile(`20201215.xlsx`);
}

// console.log(container[0][0]);
// console.log(address[container[0][0]]);
async function transferData() {
  await getData();
  await inputTax();
}

transferData();
