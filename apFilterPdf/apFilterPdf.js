const Excel = require('exceljs')
const workbook = new Excel.Workbook()


const pdf =  require('pdf-extraction')
const fs = require('fs')
const afterFilter = []
let dataBuffer = fs.readFileSync('1.pdf')
pdf(dataBuffer).then(data=>{
console.log(data.text)
const rawArr = data.text.split('\n')
// For Request Date
const apReqDate = rawArr[12]

// For Request Department
const apReqDep = rawArr[13]

rawArr.forEach((el,i)=>{
   const isLocalPayInfo = el.startsWith('N') && el[2]===':'
if(el.startsWith('00') || isLocalPayInfo){
    afterFilter.push(el)
}
if(el.startsWith('Modifier')){
    afterFilter.push(rawArr[i+1])
    afterFilter.push(rawArr[i+2])
}
})
// console.log(afterFilter)

Array.prototype.chunk = function (chunk_size) {
    if ( !this.length ) {
        return [];
    }

    return [ this.slice( 0, chunk_size ) ].concat(this.slice(chunk_size).chunk(chunk_size));
};

async function modifyExcle(){

    await workbook.xlsx.readFile('ACH.xlsx');
const ws = workbook.getWorksheet(1)
ws.getCell('A3').value="GOOD"
return workbook.xlsx.writeFile('ACH-1.xlsx')
// console.log(ws.getCell('B1').value)



}


console.log(afterFilter.chunk(4))

// afterFilter.chunk(4).forEach(el=>{

// el[0].split(' ')
// })
modifyExcle()
// const worksheet = workbook.getWorksheet(1)

})