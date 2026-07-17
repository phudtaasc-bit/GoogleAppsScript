function FS_SP_chonVaTinhMotSanPham_V2() {
  const originalLayout = FS_SP_chuanHoaCotSanPham_;
  FS_SP_chuanHoaCotSanPham_ = FS_SP_chuanHoaCotSanPhamV2_;
  try {
    FS_SP_chonVaTinhMotSanPham();
  } finally {
    FS_SP_chuanHoaCotSanPham_ = originalLayout;
  }
}

function FS_SP_chuanHoaCotSanPhamV2_(ss, products) {
  const sheet = ss.getSheetByName(FSSP_CFG.SUMMARY);
  if (!sheet) throw new Error('Không tìm thấy Sheet 00. Tổng hợp.');

  const rows = FS_SP_docDongChiTieu_(sheet);
  const firstProductCol = 5;
  const noteCol = firstProductCol + products.length;
  const lastRow = rows.vatPayable;

  if (sheet.getMaxColumns() < noteCol + 1) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), noteCol + 1 - sheet.getMaxColumns());
  }

  const headers = sheet.getRange(24, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  let currentNoteCol = headers.findIndex(value => FS_SP_key_(value) === 'ghichu') + 1;
  if (!currentNoteCol) currentNoteCol = 5;

  const tempCol = Math.max(sheet.getLastColumn(), noteCol) + 1;
  if (sheet.getMaxColumns() < tempCol) sheet.insertColumnsAfter(sheet.getMaxColumns(), tempCol - sheet.getMaxColumns());

  const noteHeight = lastRow - 23;
  sheet.getRange(24, currentNoteCol, noteHeight, 1).copyTo(
    sheet.getRange(24, tempCol, noteHeight, 1),
    SpreadsheetApp.CopyPasteType.PASTE_NORMAL,
    false
  );

  const oldProductValues = {};
  for (let col = 5; col <= sheet.getLastColumn(); col++) {
    if (col === currentNoteCol || col === tempCol) continue;
    const header = String(sheet.getRange(24, col).getDisplayValue() || '').trim();
    if (!header || FS_SP_key_(header) === 'ghichu') continue;
    oldProductValues[FS_SP_key_(header)] = sheet
      .getRange(rows.totalRevenue, col, lastRow - rows.totalRevenue + 1, 1)
      .getValues();
  }

  const clearWidth = tempCol - 5;
  if (clearWidth > 0) {
    sheet.getRange(24, 5, noteHeight, clearWidth)
      .breakApart()
      .clearContent()
      .clearFormat();
  }

  products.forEach((product, index) => {
    const col = firstProductCol + index;
    const label = product.name + (product.group ? ' - ' + product.group : '');

    sheet.getRange(24, 4, noteHeight, 1).copyTo(
      sheet.getRange(24, col, noteHeight, 1),
      SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
      false
    );

    sheet.getRange(24, col)
      .setValue(label)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setWrap(true);

    const old = oldProductValues[FS_SP_key_(label)];
    const dataHeight = lastRow - rows.totalRevenue + 1;
    if (old && old.length === dataHeight) {
      sheet.getRange(rows.totalRevenue, col, dataHeight, 1).setValues(old);
    } else {
      sheet.getRange(rows.totalRevenue, col, dataHeight, 1).setValue(0);
    }

    FS_SP_dinhDangCotKetQua_(sheet, rows, col);
    sheet.setColumnWidth(col, 125);
  });

  sheet.getRange(24, tempCol, noteHeight, 1).copyTo(
    sheet.getRange(24, noteCol, noteHeight, 1),
    SpreadsheetApp.CopyPasteType.PASTE_NORMAL,
    false
  );
  sheet.getRange(24, noteCol)
    .setValue('Ghi chú')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);
  sheet.setColumnWidth(noteCol, 170);

  sheet.getRange(24, tempCol, noteHeight, 1).clear();

  sheet.getRange(23, 1, 1, sheet.getMaxColumns()).breakApart();
  sheet.getRange(23, 1, 1, noteCol)
    .merge()
    .setValue('III. ĐÁNH GIÁ HIỆU QUẢ ĐẦU TƯ');

  sheet.getRange(23, 1, lastRow - 22, noteCol)
    .setFontFamily('Times New Roman')
    .setVerticalAlignment('middle')
    .setWrap(true)
    .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
}
