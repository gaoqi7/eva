const robot = require('robotjs')
const sleep = require('sleep')
const Excel = require('exceljs')
const workbook = new Excel.Workbook()

const linesNeedKeyIn = process.argv[2]

const seRow = [2,16]
const seCol = [2,14]
var ws

async function getData(){
  await workbook.xlsx.readFile('ogfn.xlsx')
  ws =await  workbook.getWorksheet('ofgn')
console.log(typeof ws)
cool()
}

getData()

// getData()
// sleep.sleep(5)
function cool(){


for (r = 2; r<linesNeedKeyIn+2; r++){
  for(c = 2; c<15; c++){
    
    function ha(){
    console.log(r)
    console.log(c)
      let cRow = ws.getRow(r)
      let b = cRow.getCell(c).value
      // let a =  ws.getCell(`B2`).value
      // console.log(a)
    console.log("b",b)
  }
    
    ha()




    // async function getData(){
      
      //   await workbook.xlsx.readFile('ogfn.xlsx')
//   const ws = workbook.getWorksheet('ogfn')
// console.log("haha")
// console.log(typeof ws)
// let currentRow = ws.getRow(r)
// // console.log(currentRow)
// let data = currentRow.getCell(c).value
// console.log(data)
// if(c !== 14){
// if(data){

//   robot.typeStringDelayed(data,80)
// }
//   robot.keyTap('tab')
// }else{    
//   robot.typeStringDelayed(data,80)
//   robot.keyTap('tab')
//   robot.keyTap('tab')
//   robot.keyTap('tab')
//   robot.keyTap('tab')

// }
// // currentRow.getCell(1).value = 'Done'
// }
// getData()
     
  }  
}
}