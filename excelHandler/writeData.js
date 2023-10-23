const XLSX = require('xlsx')

function editSheet(fileName, editedSheet, sheetIndex = 0) {
	const workbook = XLSX.readFile(fileName)
	const originalSheet = workbook.Sheets[workbook.SheetNames[sheetIndex]]
	XLSX.utils.sheet_add_json(originalSheet, editedSheet)
	XLSX.writeFile(workbook, fileName)
}

function newSheet(fileName, newSheet, sheetName) {
	const workbook = XLSX.readFile(fileName)
	const sheet = XLSX.utils.json_to_sheet(newSheet)
	XLSX.utils.book_append_sheet(workbook, sheet, sheetName)
	XLSX.writeFile(workbook, fileName)
}

function newFile(newFileName, newSheet, sheetName) {
	const newExcel = XLSX.utils.book_new()
	const sheet = XLSX.utils.json_to_sheet(newSheet)
	XLSX.utils.book_append_sheet(newExcel, sheet, sheetName)
	XLSX.writeFile(newExcel, `data/${newFileName}.xlsx`)
}

module.exports = {
	editSheet,
	newSheet,
	newFile,
}
