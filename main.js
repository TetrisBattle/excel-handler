const {
	getRows,
	getSheet,
	getSheets,
	editSheet,
	newSheet,
	newFile,
} = require('./excelHandler')

const fileName = 'data/test.xlsx'

const newData = getSheet(fileName)
newData.push({ id: 30, name: 'test', age: 20 })

// console.log(getRows(fileName))
console.log(getSheet(fileName))
// console.log(getSheets(fileName))

// writeSheet(fileName, sheet)
// newSheet(fileName, sheet)
newFile('Code', newData)
