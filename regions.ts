function searchRange(content: string, inRange?: GoogleAppsScript.Spreadsheet.Range): GoogleAppsScript.Spreadsheet.Range {
  const textFinder = (inRange ? inRange : SpreadsheetApp.getActiveSpreadsheet().getSheets()[0]).createTextFinder(content)
  textFinder.matchCase(true)
  textFinder.matchEntireCell(true)
  return textFinder.findNext()
}

function getTopLeftCell(): GoogleAppsScript.Spreadsheet.Range {
  return searchRange('Date')
}

function getMainDataWidth(): number {
  let cell = getTopLeftCell()
  let count = 0
  while (cell.getValue() !== '') {
    count++
    cell = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0].getRange(cell.getRow(), cell.getColumn() + 1)
  }
  return count
}

function getMainDataHeight(): number {
  let cell = getTopLeftCell()
  let count = 0
  while (cell.getValue() !== 'Total' && cell.getValue() !== '') {
    count++
    cell = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0].getRange(cell.getRow() + 1, cell.getColumn())
  }
  return count
}

function getMainDataRange(): GoogleAppsScript.Spreadsheet.Range {
  const topLeftCell = getTopLeftCell()
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0]
  return sheet.getRange(topLeftCell.getRow(), topLeftCell.getColumn(), getMainDataHeight(), getMainDataWidth())
}

function getDataRange(headerValue: string): GoogleAppsScript.Spreadsheet.Range {
  const headerRange = searchRange(headerValue, getMainDataRange())
  return SpreadsheetApp.getActiveSpreadsheet().getSheets()[0].getRange(headerRange.getRow(), headerRange.getColumn(), getMainDataHeight(), 1)
}

function getInputRange(): GoogleAppsScript.Spreadsheet.Range {
  const totalCell = searchRange('Total')
  let i = totalCell.getColumn() + 1
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0]
  while (i < getMainDataWidth() && sheet.getRange(totalCell.getRow(), i).getValue() === '') {
    i++
  }

  return sheet.getRange(getTopLeftCell().getRow(), 2, getMainDataHeight(), i - 2)
}

function getOutputRange() {
  const totalCell = searchRange('Total')
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0]
  let i = totalCell.getColumn() + 1
  while (i < getMainDataWidth() && sheet.getRange(totalCell.getRow(), i).getValue() === '') {
    i++
  }

  const firstOutputColumn = i
  while (i < getMainDataWidth() && sheet.getRange(totalCell.getRow(), i).getValue() !== '') {
    i++
  }

  return sheet.getRange(getTopLeftCell().getRow(), firstOutputColumn, getMainDataHeight(), i - firstOutputColumn + 1)
}