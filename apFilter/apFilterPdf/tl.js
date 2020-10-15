
const { PDFDocument, rgb, StandardFonts } = require('pdf-lib');
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
const { match } = require('assert')
//PDF >> Array of Text >> cherry pick what I need
// I store all the information from one payment checklist in ONE big Array and 
// Then cut the single array in small pieces. each piece match only one payment info
//Care the file name, tend to use process.argv
//1.pdf !!! what a great name!!!    
let dataBuffer = fs.readFileSync('tl.pdf')
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

async function modifyExcel(){

    fse.copySync('ACH.xlsx','ACH_bk.xlsx')
    await workbook.xlsx.readFile('ACH_bk.xlsx')
    const ws = workbook.getWorksheet(1)
    let amountArr=[] 
    let tDay = transactionInfo[0][0].split('/').join('')
    let td1 = tDay.slice(4)
    let td2 = tDay.slice(0,4)
    tDay = `${td1}${td2}`
    transactionInfo.forEach(el=>{
        let a = parseInt(parseFloat(el[1].replace(/,/g,''))*100)/100
        amountArr.push(a)
    })
    let totalRow = ws.rowCount

    const localPDCol = ws.getColumn('F')
    const xyAmountMatchPool = []
    function amountMatchPoolCreate(date){
        let amountList = []

        localPDCol.eachCell((cell,rowNumber)=>{
            if( ws.getCell(`E${rowNumber}`).value === null){
                amountList.push(ws.getCell(`G${rowNumber}`).value)
                xyAmountMatchPool.push(rowNumber)
            }
        })
        
        return amountList
    }

    function amountSumMatch(arr, target){

        function powerset(arr) {
            var ps = [[]];
            for (var i=0; i < arr.length; i++) {
                for (var j = 0, len = ps.length; j < len; j++) {
                    ps.push(ps[j].concat(arr[i]));
                }
            }
            return ps;
        }
        
        function sum(arr) {
            var total = 0;
            for (var i = 0; i < arr.length; i++)
                total += arr[i];
            return total
        }
        
        function findSum(numbers, targetSum) {
            var numberSets = powerset(numbers);
            for (var i=0; i < numberSets.length; i++) {
                var numberSet = numberSets[i]; 
                if (sum(numberSet) == targetSum)
                    return numberSet;
            }
        }

       return findSum(arr,target)


    }
    
    const aMP = amountMatchPoolCreate(tDay)
    console.log(tDay)
    console.log(aMP)
    console.log(amountArr)
    amountArr.forEach(el=>{
        console.log(`for${el}`)
let matchedAmountArr =amountSumMatch(aMP,el)
console.log(matchedAmountArr)
console.log(typeof matchedAmountArr)
if (matchedAmountArr){
    matchedAmountArr.forEach(el=>{
        let r = xyAmountMatchPool[aMP.indexOf(el)]
        console.log("the Row Number need edit is " , xyAmountMatchPool[aMP.indexOf(el)])
        ws.getCell(`E${r}`).value = tDay
        workbook.xlsx.writeFile('ACH.xlsx')
        fse.removeSync('Ach_bk.xlsx')
//***************** */
// Modify Transaction List
//**************** */




    })

}
    })
    


}
modifyExcel()


})