const fse = require('fs-extra')
const Excel = require('exceljs')
const wb = new Excel.Workbook()
const eeList = []
async function modify(){
    fse.copySync('punchInOut.xlsx','punchInOutBk.xlsx')
    await wb.xlsx.readFile('punchInOutBk.xlsx')
    const ws = wb.getWorksheet(1)
    const eeID = ws.getColumn('D')

    eeID.eachCell((cell,rowNumber)=>{
        //if Not exist in eeList
        if(eeList.indexOf(cell.value) === -1) {
            eeList.push(cell.value)
            eeList.push(rowNumber)
        } 
    })
    
    console.log(eeList)
}

modify()