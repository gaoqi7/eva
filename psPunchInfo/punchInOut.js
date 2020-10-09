const fse = require('fs-extra')
const Excel = require('exceljs')
const wb = new Excel.Workbook()


Array.prototype.chunk = function (chunk_size) {
    if ( !this.length ) {
        return [];
    }

    return [ this.slice( 0, chunk_size ) ].concat(this.slice(chunk_size).chunk(chunk_size));
};

const eeList = []

async function modify(){
    fse.copySync('punchInOut.xlsx','punchInOutBk.xlsx')
    await wb.xlsx.readFile('punchInOutBk.xlsx')
    const ws = wb.getWorksheet(1)
    //Get Column D
    const eeID = ws.getColumn('D')

    eeID.eachCell((cell,rowNumber)=>{
        //if Not exist in eeList
        if(eeList.indexOf(cell.value) === -1) {
            eeList.push(cell.value)
            eeList.push(rowNumber)
        } 
    })
    // Remove the first two item which is [Employee,1]
    eeList.splice(0,2)
    console.log(eeList.chunk(2))
    let eeChunk = eeList.chunk(2)
    for (i=0;i<eeChunk.length-1;i++){
        for (j=eeChunk[i][1]+1;j<eeChunk[i+1][1];j++){
            console.log(j)
            ws.getCell(`A${j}`).value = null
            ws.getCell(`B${j}`).value = null
        }

    }
    // The previous two for loop can't finish to the end of the xlsx file.
    // I have to use the third for loop. Thanks the rowCount !
    let totalRow = ws.rowCount
    for (t = eeChunk[eeChunk.length-1][1]+1;t<=ws.rowCount;t++){
            ws.getCell(`A${t}`).value = null
            ws.getCell(`B${t}`).value = null
    }

    return wb.xlsx.writeFile('punchInOut.xlsx')

}

modify()