const XLSX = require('xlsx')

function getSheet(fileName, sheetIndex = 0) {
	const workbook = XLSX.readFile(fileName)
	const worksheet = workbook.Sheets[workbook.SheetNames[sheetIndex]]
	return XLSX.utils.sheet_to_json(worksheet)
}

function getSheets(fileName) {
	const workbook = XLSX.readFile(fileName)
	const worksheets = {}
	workbook.SheetNames.forEach((sheetName) => {
		worksheets[sheetName] = XLSX.utils.sheet_to_json(
			workbook.Sheets[sheetName]
		)
	})
	return worksheets
}

function getRows(fileName, sheetIndex = 0) {
	const workbook = XLSX.readFile(fileName)
	const sheet = workbook.Sheets[workbook.SheetNames[sheetIndex]]
	const range = sheet['!ref']
	const [start, end] = range.split(':')
	const [startCol, startRow] = start.match(/[A-Z]+|[0-9]+/g)
	const [endCol, endRow] = end.match(/[A-Z]+|[0-9]+/g)
	const colCount = endCol.charCodeAt(0) - startCol.charCodeAt(0) + 1
	const rowCount = endRow - startRow + 1
	const data = []
	for (let i = 0; i < rowCount; i++) {
		const row = []
		for (let j = 0; j < colCount; j++) {
			const cell =
				sheet[
					String.fromCharCode(startCol.charCodeAt(0) + j) +
						(parseInt(startRow) + i)
				]
			row.push(cell ? cell.v : undefined)
		}
		data.push(row)
	}
	return data
}

module.exports = {
	getSheet,
	getSheets,
	getRows,
}
