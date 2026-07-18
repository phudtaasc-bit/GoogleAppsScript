const FS00H_CFG = Object.freeze({
  SUMMARY: '00. Tổng hợp',
  HEADER_BG: '#9e9e9e',
  SECTION_BG: '#f4b400',
  PROJECT_BG: '#fce8e6',
  EQUITY_BG: '#e2f0d9',
  RED: '#ff0000',
  BORDER: '#000000'
});

/**
 * Hàm chính gọi từ menu riêng.
 * Chạy phân tích theo sản phẩm, sau đó định dạng lại Mục III Sheet 00.
 */
function FS00H_phanTichVaDinhDangTheoSanPham() {
  if (typeof FS00I_phanTichHieuQuaTheoSanPham !== 'function') {
    throw new Error('Thiếu hàm FS00I_phanTichHieuQuaTheoSanPham.');
  }

  const results = FS00I_phanTichHieuQuaTheoSanPham();
  FS00H_dinhDangMucIII_();
  SpreadsheetApp.getActive().toast(
    'Đã phân tích và định dạng bảng hiệu quả từng sản phẩm.',
    'FS sản phẩm',
    6
  );
  return results;
}

function FS00H_dinhDangMucIII_() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(FS00H_CFG.SUMMARY);
  if (!sheet) throw new Error('Không tìm thấy sheet "00. Tổng hợp".');

  const headerRow = FS00H_findHeaderRow_(sheet);
  const titleRow = headerRow - 1;
  const noteCol = FS00H_findHeaderCol_(sheet, headerRow, ['Ghi chú']);
  if (!noteCol) throw new Error('Không tìm thấy cột Ghi chú tại Mục III.');

  const lastRow = FS00H_findLastRow_(sheet, headerRow);
  const rowMap = FS00H_rowMap_(sheet, headerRow, lastRow);
  const fullRange = sheet.getRange(titleRow, 1, lastRow - titleRow + 1, noteCol);

  fullRange
    .setFontFamily('Times New Roman')
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true, FS00H_CFG.BORDER, SpreadsheetApp.BorderStyle.SOLID);

  sheet.getRange(titleRow, 1, 1, noteCol)
    .setBackground(FS00H_CFG.SECTION_BG)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setFontSize(12);

  sheet.getRange(headerRow, 1, 1, noteCol)
    .setBackground(FS00H_CFG.HEADER_BG)
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setWrap(true);

  if (lastRow > headerRow) {
    sheet.getRange(headerRow + 1, 1, lastRow - headerRow, 1).setHorizontalAlignment('center');
    sheet.getRange(headerRow + 1, 3, lastRow - headerRow, 1).setHorizontalAlignment('center');
    sheet.getRange(headerRow + 1, 4, lastRow - headerRow, noteCol - 4)
      .setHorizontalAlignment('center')
      .setNumberFormat('#,##0.0');
    sheet.getRange(headerRow + 1, noteCol, lastRow - headerRow, 1)
      .setHorizontalAlignment('center')
      .setWrap(true);
  }

  FS00H_boldRows_(sheet, rowMap, [
    'Tổng doanh thu có VAT',
    'Tổng chi phí có VAT',
    'Lợi nhuận sau thuế',
    'NPV dự án',
    'IRR dự án',
    'Thời gian hoàn vốn dự án',
    'NPV vốn CSH',
    'IRR vốn CSH',
    'Thời gian hoàn vốn - Vốn CSH',
    'Đỉnh dư nợ vay',
    'Tổng lãi vay',
    'Tổng Thuế TNDN',
    'Tổng VAT phải nộp'
  ], noteCol);

  FS00H_styleRows_(sheet, rowMap, [
    'NPV dự án',
    'IRR dự án',
    'Thời gian hoàn vốn dự án'
  ], noteCol, FS00H_CFG.PROJECT_BG, FS00H_CFG.RED);

  FS00H_styleRows_(sheet, rowMap, [
    'NPV vốn CSH',
    'IRR vốn CSH',
    'Thời gian hoàn vốn - Vốn CSH',
    'Thời gian hoàn vốn vốn CSH'
  ], noteCol, FS00H_CFG.EQUITY_BG, FS00H_CFG.RED);

  FS00H_numberFormatRows_(sheet, rowMap, ['IRR dự án', 'IRR vốn CSH'], 4, noteCol - 4, '0.00%');
  FS00H_numberFormatRows_(sheet, rowMap, [
    'Thời gian hoàn vốn dự án',
    'Thời gian hoàn vốn - Vốn CSH',
    'Thời gian hoàn vốn vốn CSH'
  ], 4, noteCol - 4, '0.00');

  FS00H_fontColorRows_(sheet, rowMap, [
    'NPV dự án',
    'IRR dự án',
    'Thời gian hoàn vốn dự án',
    'NPV vốn CSH',
    'IRR vốn CSH',
    'Thời gian hoàn vốn - Vốn CSH',
    'Thời gian hoàn vốn vốn CSH'
  ], noteCol, FS00H_CFG.RED);

  sheet.setColumnWidth(1, 45);
  sheet.setColumnWidth(2, 285);
  sheet.setColumnWidth(3, 110);
  sheet.setColumnWidth(4, 110);
  if (noteCol > 5) sheet.setColumnWidths(5, noteCol - 5, 125);
  sheet.setColumnWidth(noteCol, 210);

  sheet.setRowHeight(titleRow, 24);
  sheet.setRowHeight(headerRow, 28);

  // Tuyệt đối không cố định hàng/cột tại Sheet 00.
  sheet.setFrozenRows(0);
  sheet.setFrozenColumns(0);
}

function FS00H_findHeaderRow_(sheet) {
  const values = sheet.getRange(1, 1, sheet.getLastRow(), Math.min(12, sheet.getLastColumn())).getDisplayValues();
  for (let r = 0; r < values.length; r++) {
    const keys = values[r].map(FS00H_key_);
    if (keys.includes(FS00H_key_('Nội dung')) &&
        keys.includes(FS00H_key_('Đơn vị')) &&
        (keys.includes(FS00H_key_('Toàn dự án')) || keys.includes(FS00H_key_('Giá trị')))) {
      const hasSection = values.slice(0, r + 1).some(row => String(row[0] || '').trim().indexOf('III.') === 0);
      if (hasSection) return r + 1;
    }
  }
  throw new Error('Không xác định được dòng tiêu đề Mục III.');
}

function FS00H_findHeaderCol_(sheet, headerRow, aliases) {
  const headers = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const targets = aliases.map(FS00H_key_);
  for (let i = 0; i < headers.length; i++) {
    if (targets.includes(FS00H_key_(headers[i]))) return i + 1;
  }
  return 0;
}

function FS00H_findLastRow_(sheet, headerRow) {
  const values = sheet.getRange(headerRow + 1, 2, sheet.getLastRow() - headerRow, 1).getDisplayValues();
  let last = headerRow;
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0] || '').trim()) last = headerRow + i + 1;
  }
  return last;
}

function FS00H_rowMap_(sheet, headerRow, lastRow) {
  const values = sheet.getRange(headerRow + 1, 2, lastRow - headerRow, 1).getDisplayValues();
  const out = {};
  values.forEach((row, index) => {
    const key = FS00H_key_(row[0]);
    if (key && out[key] == null) out[key] = headerRow + index + 1;
  });
  return out;
}

function FS00H_boldRows_(sheet, rowMap, aliases, noteCol) {
  aliases.forEach(alias => {
    const row = rowMap[FS00H_key_(alias)];
    if (row) sheet.getRange(row, 1, 1, noteCol).setFontWeight('bold');
  });
}

function FS00H_styleRows_(sheet, rowMap, aliases, noteCol, background, fontColor) {
  aliases.forEach(alias => {
    const row = rowMap[FS00H_key_(alias)];
    if (row) {
      sheet.getRange(row, 1, 1, noteCol)
        .setBackground(background)
        .setFontColor(fontColor)
        .setFontWeight('bold');
    }
  });
}

function FS00H_numberFormatRows_(sheet, rowMap, aliases, firstCol, count, format) {
  aliases.forEach(alias => {
    const row = rowMap[FS00H_key_(alias)];
    if (row) sheet.getRange(row, firstCol, 1, count).setNumberFormat(format);
  });
}

function FS00H_fontColorRows_(sheet, rowMap, aliases, col, color) {
  aliases.forEach(alias => {
    const row = rowMap[FS00H_key_(alias)];
    if (row) sheet.getRange(row, col).setFontColor(color).setFontWeight('bold');
  });
}

function FS00H_key_(value) {
  return String(value == null ? '' : value)
    .trim()
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd')
    .replace(/[^a-z0-9]/g, '');
}
