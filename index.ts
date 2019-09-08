function createTotals() {
  const output = getOutputRange()
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0]

  for (let i = output.getColumn(); i <= output.getLastColumn(); i++) {
    sheet.getRange(output.getLastRow() + 1, i)
      .setFormula(`=SUBTOTAL(9, ${sheet.getRange(output.getRow(), i, output.getHeight(), 1).getA1Notation()})`)
      .setNumberFormat('#,###')
  }
}

function createAverages() {
  const output = getOutputRange()
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0]

  for (let i = output.getColumn(); i <= output.getLastColumn(); i++) {
    sheet.getRange(output.getLastRow() + 2, i)
      .setFormula(`=SUBTOTAL(1, ${sheet.getRange(output.getRow(), i, output.getHeight(), 1).getA1Notation()})`)
      .setNumberFormat('#,###')
  }

  sheet.getRange(output.getLastRow() + 2, 1).setValue('Average')
}

function findFirstAvailableCellInSheet(sheet: GoogleAppsScript.Spreadsheet.Sheet): GoogleAppsScript.Spreadsheet.Range {
  let row = 0
  do {
    row++
  } while (sheet.getRange(row, 1).getValue() !== '' || sheet.getRange(row + 1, 1).getValue() !== '')

  return sheet.getRange(row > 1 ? row + 1 : row, 1)
}

function createChart() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[1]
  const dateRange = sheet.getRange('A2:A6')
  const range1 = sheet.getRange('D2:D6')
  const range2 = sheet.getRange('D9:D13')
  const chartBuilder = sheet.newChart()
    .addRange(dateRange)
    .addRange(range1)
    .addRange(range2)
    .setChartType(Charts.ChartType.LINE)
    .setOption('title', 'My Line chart')
    .setPosition(5, 5, 0, 0)
    .setNumHeaders(1)
  sheet.insertChart(chartBuilder.build())
}

function columnToLetter(column: number): string {
  let temp: number, letter = ''
  while (column > 0) {
    temp = (column - 1) % 26
    letter = String.fromCharCode(temp + 65) + letter
    column = (column - temp - 1) / 26
  }
  return letter
}

function explodeRangeInColumns(range: GoogleAppsScript.Spreadsheet.Range): Array<GoogleAppsScript.Spreadsheet.Range> {
  const res = []

  for (let i = range.getColumn(); i <= range.getLastColumn(); i++) {
    res.push(range.getSheet().getRange(range.getRow(), i, range.getHeight(), 1))
  }

  return res
}

function createDimension(dimension: string) {
  const dataRange = getDataRange(dimension)

  const uniqueValues = dataRange.getValues().slice(1).map(([value]) => value).filter((value, index, values) => values.indexOf(value) === index)

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  const sheet = spreadsheet.getSheetByName(dimension) || spreadsheet.insertSheet()
  sheet.setName(dimension)
  sheet.clear()
  sheet.getCharts().forEach(chart => sheet.removeChart(chart))

  const dataSheetName = spreadsheet.getSheets()[0].getName()
  const outputRanges = explodeRangeInColumns(getOutputRange())

  outputRanges.forEach((outputRange, i) => {
    const startCell = findFirstAvailableCellInSheet(sheet)

    startCell.setValue(outputRange.getValue())
    const dataCell = sheet.getRange(startCell.getRow() + 1, 1)
    dataCell.setFormula(`UNIQUE('${dataSheetName}'!${getDataRange('Date').getA1Notation()})`)

    uniqueValues.forEach((uniqueValue, i) => {
      const cell = sheet.getRange(dataCell.getRow(), dataCell.getColumn() + i + 1)
      cell.setFormula(`=QUERY('${dataSheetName}'!${getMainDataRange().getA1Notation()}, "SELECT ${columnToLetter(outputRange.getColumn())} WHERE ${columnToLetter(dataRange.getColumn())} = '${uniqueValue}' LABEL ${columnToLetter(outputRange.getColumn())} '${uniqueValue}'")`)
    })

    const nextStartCell = findFirstAvailableCellInSheet(sheet)
    const chartBuilder = sheet.newChart()
      .addRange(sheet.getRange(dataCell.getRow(), dataCell.getColumn(), nextStartCell.getRow() - dataCell.getRow() - 2, uniqueValues.length + 1))
      .setChartType(Charts.ChartType.LINE)
      .setOption('title', outputRange.getValue())
      .setOption('width', 500)
      .setOption('height', 300)
      .setPosition(1, uniqueValues.length + 3, 0, 320 * i)
      .setNumHeaders(1)
    sheet.insertChart(chartBuilder.build())
  })
}

function main(...dimensions: Array<string>) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0]
  if (sheet.getFilter()) {
    sheet.getFilter().remove()
  }

  const inputRange = getInputRange()
  inputRange.createFilter()

  createTotals()
  createAverages()

  dimensions.forEach(dimension => {
    createDimension(dimension)
  })
}