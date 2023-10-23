const { getSheet, getSheets, getRows } = require('./readData')
const { editSheet, newSheet, newFile } = require('./writeData')

module.exports = {
	getSheet,
	getSheets,
	getRows,

	editSheet,
	newSheet,
	newFile,
}
