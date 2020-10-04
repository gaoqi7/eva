//Only xlsx file can be read and write. This limition is adopted by exceljs.
const Excel = require('exceljs')

const workbook = new Excel.Workbook()
const requestDate = []
async function readAP (path) {
  await workbook.xlsx.readFile(path)
  const worksheet = workbook.getWorksheet(1)
  const seqNoCol = worksheet.getColumn(1)
  // Reference Point should be sequence number
  seqNoCol.eachCell((cell, rowNumber) => {
    if (cell.value === '0001') {
      const xyRequestDate = `D${rowNumber - 4}`
      console.log(rowNumber)
      requestDate.push(worksheet.getCell(xyRequestDate).value)
    }
  })
  console.log(requestDate)
}

readAP('./excelFile/7ap-cv.xlsx')
