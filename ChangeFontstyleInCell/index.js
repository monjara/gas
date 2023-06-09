function ChangeFontStyleInCell() {
  const sheet = SpreadsheetApp.getActiveSheet()
  const range = sheet.getDataRange()
  const rowlength = range.getNumRows()
  const columnlength = range.getNumColumns()

  for (let row = 1; row <= rowlength; row++) {
    for (let column = 1; column <= columnlength; column++) {
      const cell = range.getCell(row, column)
      const isBlank = cell.getValue() === ''
      if (isBlank) {
        continue
      }
      _rebuildCell(cell)
    }
  }

  function _rebuildCell(cell) {
    const builder = SpreadsheetApp.newRichTextValue()
      .setText(cell.getValue())

    const runs = cell.getRichTextValue()?.getRuns()
    if (!runs) {
      return
    }
    for (const run of runs) {
      const isBold = run.getTextStyle().isBold()
      if (isBold) {
        const start = run.getStartIndex()
        const end = run.getEndIndex()
        const newStyle = _makeForegroundRed()
        builder.setTextStyle(start, end, newStyle)
      }
    }

    const newValue = builder.build()
    cell.setRichTextValue(newValue)
  }

  function _makeForegroundRed() {
    return SpreadsheetApp.newTextStyle()
      .setBold(true)
      .setForegroundColor('#ff0000')
      .build()
  }
}

