const { getSheet, getSheets, getRows } = require('./excelHandler')

const fileName = 'data/test.xlsx'

console.log('one', getSheet(fileName))
console.log('many', getSheets(fileName))
console.log('rows', getRows(fileName))
