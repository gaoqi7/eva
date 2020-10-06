const Excel = require('exceljs')
const workbook = createAndFillWorkbook();
await workbook.xlsx.writeFile('ACH&WIRE PAYMENT-bk.xlsx');
const worksheet = workbook.getWorksheet(1)

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

console.log(afterFilter.chunk(4))

afterFilter.chunk(4).forEach(el=>{

el[0].split(' ')
})


})