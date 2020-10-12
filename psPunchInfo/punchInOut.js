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
    // console.log(eeList.chunk(2))
    let eeChunk = eeList.chunk(2)
    eeChunk.push(['NNNNNN',ws.rowCount+1])
    console.log(eeChunk)
    
    // function dayCount(sp,ep){
        
    // }
    for (i=0;i<eeChunk.length-1;i++){
        let a = [] // used for collect the punch date info
        for (j=eeChunk[i][1]+1;j<eeChunk[i+1][1];j++){
            // Remove Duplicate content in Column A & B
            ws.getCell(`A${j}`).value = null
            ws.getCell(`B${j}`).value = null
            //collect the punch date
            if (a.indexOf(ws.getCell(`F${j-1}`).value) === -1){
                a.push(ws.getCell(`F${j-1}`).value)
                a.push(j-1)
            }
        }
        console.log(a)
        let b = a.chunk(2)
        console.log(b)
        for(t = 0;t<b.length-1;t++){
            if(ws.getCell(`H${b[t][1]}`).value.trim() === 'Punch In' && ws.getCell(`H${b[t+1][1]-1}`).value.trim().endsWith('Out')){
                ws.getCell(`E${b[t][1]}`).value = 1
                ws.getCell(`J${b[t][1]}`).value={formula:`I${b[t+1][1]-1}-I${b[t][1]}`}
                ws.getCell(`J${b[t][1]}`).numFmt = 'hh:mm'

            }else{
                ws.getCell(`E${b[t][1]}`).value = '?????'
            }
        }
    }

    wb.xlsx.writeFile('punchInOut.xlsx')
    fse.removeSync('punchInOutBk.xlsx')


}

modify()