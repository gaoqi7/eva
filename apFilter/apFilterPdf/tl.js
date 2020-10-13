//Use ExcelJS to Modify ACH.xlsx File
const Excel = require('exceljs')
const workbook = new Excel.Workbook()
//Use fs-extra package to make Copy and delete Copy
//ExcelJS not support read and write a file at same time. I have to create a copy which is ACH_bk.xlsx, read and modify that copy, and then write into the original file.
const fse = require('fs-extra')
//Parse the Payment Check List file
const pdf =  require('pdf-extraction')
//Actually, I can just use one dependent(fs-extra) to deal with file system operation like copy paste, read and write.
const fs = require('fs')
//PDF >> Array of Text >> cherry pick what I need
// I store all the information from one payment checklist in ONE big Array and 
// Then cut the single array in small pieces. each piece match only one payment info
//Care the file name, tend to use process.argv
//1.pdf !!! what a great name!!!    
let dataBuffer = fs.readFileSync('tl_both.pdf')
pdf(dataBuffer).then(data=>{
    const rawArr = data.text.split('\n')
    console.log(rawArr)
const transactionInfo = []
rawArr.forEach((el,i)=>{
    if (el.trim() === 'SETTLEMENT'){
        let tDay = rawArr[i-1].split(' ')[2]
        let tAmount = rawArr[i+1].split(' ')[0]
        transactionInfo.push([tDay,tAmount])
    } else if (el.startsWith('211367350') && el.length>45){
        // must use trim()
        let tInfoArr = rawArr[i].trim().split(' ')
        let tDay = tInfoArr[2]
        let tAmount = tInfoArr[tInfoArr.length - 4]
        transactionInfo.push([tDay,tAmount])
    }
})
console.log(transactionInfo)
// // console.log(afterFilter)
// // Copied from google. works well
// Array.prototype.chunk = function (chunk_size) {
//     if ( !this.length ) {
//         return [];
//     }

//     return [ this.slice( 0, chunk_size ) ].concat(this.slice(chunk_size).chunk(chunk_size));
// };



// // 對數據再處理，保證格式正確。
// const dataReady = []
// afterFilter.chunk(4).forEach(el=>{
//     let payType = el[1][1]
//     if (payType !== 'C'){
//         let apNO = `${apReqDate}-${apReqDep}-${el[0].split(' ')[0]}`
//         let amount = el[0].split(' ')[1].replace(',','') 
//         let remitDay =el[0].split(' ')[2].slice(-10).split('/').join('')
//         let localPayDay = el[1].slice(-10).split('/').join('')
//         let venderInfo = `${el[2]}${el[3]}`
//         // let venderCode = venderInfo.substring(0,8)
//         let venderName = venderInfo.slice(8)
//         dataReady.push([remitDay,apNO,venderName,localPayDay,amount])
//     }
// })

// console.log(dataReady)

// //Below is the output

// async function modifyExcle(){
//     fse.copySync('ACH.xlsx','ACH_bk.xlsx')
//     await workbook.xlsx.readFile('ACH_bk.xlsx');
// const ws = workbook.getWorksheet(1)
// let startRow = ws.rowCount+1
// dataReady.forEach(el=>{
// ws.getCell(`B${startRow}`).value=el[0]
// ws.getCell(`C${startRow}`).value=el[1]
// ws.getCell(`D${startRow}`).value=el[2]
// ws.getCell(`F${startRow}`).value=el[3]
// ws.getCell(`G${startRow}`).value=parseInt(parseFloat(el[4])*100)/100
// startRow++
// })
// await workbook.xlsx.writeFile('ACH.xlsx')
// await fse.removeSync('ACH_bk.xlsx')
// }


// modifyExcle()
})