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
const afterFilter = []
//Care the file name, tend to use process.argv
//1.pdf !!! what a great name!!!    
let dataBuffer = fs.readFileSync('1.pdf')
pdf(dataBuffer).then(data=>{
    const rawArr = data.text.split('\n')
    console.log(rawArr)
// For Request Date
const apReqDate = rawArr[12].split('/').join('')
console.log(apReqDate)
// For Request Department
const apReqDep = rawArr[13]

rawArr.forEach((el,i)=>{
// line ***start with '00' *** includes the AP number information
// line start with "N" and the third charactor is ":"
   const isLocalPayInfo = el.startsWith('N') && el[2]===':'
   //There is one Payment Application Called 4601, HHHHHHHH
if(el.startsWith('00') || isLocalPayInfo || el.startsWith('46')){
    afterFilter.push(el)
}
// The venderInfo is under the modifier line
if(el.startsWith('Modifier')){
    afterFilter.push(rawArr[i+1])
    afterFilter.push(rawArr[i+2])
}
})
console.log(afterFilter)
// Copied from google. works well
Array.prototype.chunk = function (chunk_size) {
    if ( !this.length ) {
        return [];
    }

    return [ this.slice( 0, chunk_size ) ].concat(this.slice(chunk_size).chunk(chunk_size));
};



// 對數據再處理，保證格式正確。
console.log(afterFilter.chunk(4))
const dataReady = []
afterFilter.chunk(4).forEach((el,i)=>{
    let payType = el[1][1]
    console.log("this is ",i)
    if (payType !== 'C'){
        let apNO = `${apReqDate}-${apReqDep}-${el[0].split(' ')[0]}`
        let amount = el[0].split(' ')[1].replace(',','') 
        let remitDay =el[0].split(' ')[2].slice(-10).split('/').join('')
        let localPayDay
        // There is one bug in financial system. 
        // Financial System Can't get local pay day from uploaded excel file
        // Attention: I use == not === to decide if the AP is reimburstment
        if(el[0].split(' ')[0]==4601){
            localPayDay = remitDay - 1
        }else{
            localPayDay = el[1].slice(-10).split('/').join('')
        }



        let venderInfo = `${el[2]}${el[3]}`
        // let venderCode = venderInfo.substring(0,8)
        let venderName
        //There is one kind of vendor called Employee ID
        if(el[2].length === 6){
            venderName = venderInfo.slice(6)
        }else{
            venderName = venderInfo.slice(8)
        }
        dataReady.push([remitDay,apNO,venderName,localPayDay,amount])
    }
})

console.log(dataReady)

//Below is the output

async function modifyExcle(){
    fse.copySync('ACH.xlsx','ACH_bk.xlsx')
    await workbook.xlsx.readFile('ACH_bk.xlsx');
const ws = workbook.getWorksheet(1)
let startRow = ws.rowCount+1
dataReady.forEach(el=>{
ws.getCell(`B${startRow}`).value=el[0]
ws.getCell(`C${startRow}`).value=el[1]
ws.getCell(`D${startRow}`).value=el[2]
ws.getCell(`F${startRow}`).value=el[3]
ws.getCell(`G${startRow}`).value=parseInt(parseFloat(el[4])*100)/100
startRow++
})
await workbook.xlsx.writeFile('ACH.xlsx')
await fse.removeSync('ACH_bk.xlsx')
}


modifyExcle()
})