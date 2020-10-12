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
    
        
    //This loop is looping the person. 
    for (i=0;i<eeChunk.length-1;i++){
        let a = [] // used for collect the punch date info
        // this loop is for everyday's punch info for one person
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
        // Why? because I need know the ending day of one person's last punch record
        // How? Next person's starter day is the last person's ending record.
        a.push(eeChunk[i+1][0])
        a.push(eeChunk[i+1][1])
        console.log(a)
        //Chunk the array. bounding the punch date and row number
        let b = a.chunk(2)
        console.log(b)
        //Looping each day which has punch info recorded.
        for(t = 0;t<b.length-1;t++){
            console.log(t)
            // if the daily records start with Punch In, and ending with Punch Out or Overtime Out. I will consider this employee have a good punch action.
            // Again, I still use the next day's start row to calculate the previous person's ending record row number.
            if(ws.getCell(`H${b[t][1]}`).value.trim() === 'Punch In' && ws.getCell(`H${b[t+1][1]-1}`).value.trim().endsWith('Out')){
                // Day Count
                ws.getCell(`E${b[t][1]}`).value = 1
                // Use Formula to calculate the working time. But the result is the floating number
                ws.getCell(`J${b[t][1]}`).value={formula:`I${b[t+1][1]-1}-I${b[t][1]}`}
                //Format the floating number to time format. 
                ws.getCell(`J${b[t][1]}`).numFmt = 'hh:mm'

            }else{
                // use ?????? to indicate the bad punch action.
                ws.getCell(`E${b[t][1]}`).value = '?????'
            }
        }
    }

    wb.xlsx.writeFile('punchInOut.xlsx')
    fse.removeSync('punchInOutBk.xlsx')


}

modify()