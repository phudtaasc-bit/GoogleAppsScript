/*************************************************
 * 99_UI.gs
 * UI helpers dùng chung cho FS Framework V2
 *************************************************/

const FS_UI = {
  COLOR: {
    TIME: '#D9E1F2',
    OPERATING: '#D9EAF7',
    COST: '#FFF2CC',
    VAT: '#E2F0D9',
    FUNDING: '#FCE4D6',
    DEBT: '#EADCF8',
    CASH: '#D9EAD3',
    KPI: '#B7DEE8',
    DARK_BLUE: '#1F4E78',
    WHITE: '#FFFFFF',
    BORDER: '#666666'
  },

  base(sheet, lastRow, lastCol) {
    sheet.getRange(1, 1, lastRow, lastCol)
      .setFontFamily('Arial')
      .setFontSize(10)
      .setVerticalAlignment('middle')
      .setWrap(true);
  },

  titleRow(sheet, row, startCol, numCols, title, color) {
    const range = sheet.getRange(row, startCol, 1, numCols);
    range.merge();
    range
      .setValue(title)
      .setBackground(color)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setBorder(true, true, true, true, true, true, this.COLOR.BORDER, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  },

  header(sheet, row, startCol, numCols, color) {
    sheet.getRange(row, startCol, 1, numCols)
      .setBackground(color)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setWrap(true)
      .setBorder(true, true, true, true, true, true, this.COLOR.BORDER, SpreadsheetApp.BorderStyle.SOLID);
  },

  block(sheet, startRow, startCol, numRows, numCols, color) {
    if (numRows <= 0 || numCols <= 0) return;
    sheet.getRange(startRow, startCol, numRows, numCols)
      .setBackground(color)
      .setBorder(true, true, true, true, true, true, this.COLOR.BORDER, SpreadsheetApp.BorderStyle.SOLID);
  },

  kpiColumn(sheet, startRow, col, numRows, color) {
    if (numRows <= 0) return;
    sheet.getRange(startRow, col, numRows, 1)
      .setBackground(color)
      .setFontWeight('bold')
      .setBorder(true, true, true, true, false, false, this.COLOR.BORDER, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  },

  numberFormat(sheet, startRow, startCol, numRows, numCols, format) {
    if (numRows <= 0 || numCols <= 0) return;
    sheet.getRange(startRow, startCol, numRows, numCols).setNumberFormat(format);
  },

  freeze(sheet, rows, cols) {
    sheet.setFrozenRows(rows);
    sheet.setFrozenColumns(cols);
  },

  autoWidth(sheet, startCol, numCols) {
    sheet.autoResizeColumns(startCol, numCols);
  },

  clearBanding(sheet) {
    const bandings = sheet.getBandings();
    bandings.forEach(b => b.remove());
  }
};