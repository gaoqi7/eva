//先parse PDF，然後找text文本的規律，然後filter想要的數據
//整理數據，至少要easy access
// 針對表格 進行修改
const pdf = require('pdf-extraction')
const fse = require('fs-extra')

const Excel = require('exceljs') 
const workbook = new Excel.Workbook()


const aiListAfterFilter = []
function parsePdf(){
    let dataBuffer = fse.readFileSync('coa.pdf')
    pdf(dataBuffer).then(data=>{
        // console.log(data.text)
        let rawText = data.text
        let aiList = rawText.split('\n')
        aiList.forEach((el,i)=>{
            // Two condition: 1. start with number. A little bit shame to express like below. I should use regex.
            // condition 2. all the account code are only 6 charactors
            // let condition1 = el.trim().startsWith('1') || el.trim().startsWith('2') ||el.trim().startsWith('3') ||el.trim().startsWith('4') ||el.trim().startsWith('5') ||el.trim().startsWith('6') ||el.trim().startsWith('7') ||el.trim().startsWith('8') ||el.trim().startsWith('9')
            let condition1 = el.trim().match(/^\d/)
            let condition2 = el.length === 6
          if (condition1 && condition2){
              aiListAfterFilter.push([el,aiList[i-1].slice(-4)])
          } 
        })
    }).then(()=>modifyExcel())
    
}

parsePdf()

function acQuery(aCode){
let a 
let arr = aiListAfterFilter.filter(el=>el[0]===aCode)
if (arr[0]){
a = arr[0][1]
}else{
    a = 'Sorry'
}
return a
}


async function modifyExcel(){
    fse.copySync('all.xlsx','allBk.xlsx')
    await workbook.xlsx.readFile('allBk.xlsx')
    const ws = workbook.getWorksheet('ALL')
    let allRow = ws.rowCount-1
    console.log(allRow)
    for (i=8;i<allRow;i++){
        //Because missing the last value, I waste two hours to debug!!!
        ws.getCell(`B${i}`).value = acQuery(ws.getCell(`A${i}.trim()`).value)
    }

    workbook.xlsx.writeFile('all.xlsx')
    fse.removeSync('allBk.xlsx')
}
