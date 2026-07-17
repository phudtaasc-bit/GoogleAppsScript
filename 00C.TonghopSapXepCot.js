function FS_lapSheet00_TheoDanhMucVaSanPham_GhiChuCuoi() {
  const result = FS_lapSheet00_TheoDanhMucVaSanPham();
  FS00C_sapXepCotSanPhamVaGhiChu_();
  return result;
}

function FS00C_sapXepCotSanPhamVaGhiChu_() {
  const ss = SpreadsheetApp.getActive();
  const tech = ss.getSheetByName('01A. Kỹ thuật');
  const summary = ss.getSheetByName('00. Tổng hợp');
  if (!tech || !summary) throw new Error('Thiếu Sheet 01A hoặc Sheet 00 để sắp xếp cột sản phẩm.');

  const products = FS00P_docSanPham_(tech);
  if (!products.length) return;

  const rows = FS00P_docDongChiTieu_(summary);
  const startRow = 23;
  const endRow = rows.vatPayable;
  const rowCount = endRow - startRow + 1;

  const noteSourceCol = 5;                // E hiện tại
  const oldFirstProductCol = 6;           // F hiện tại
  const oldLastProductCol = oldFirstProductCol + products.length - 1;
  const newFirstProductCol = 5;           // E mới
  const newNoteCol = newFirstProductCol + products.length;
  const tempCol = newNoteCol + 1;

  if (summary.getMaxColumns() < tempCol) {
    summary.insertColumnsAfter(summary.getMaxColumns(), tempCol - summary.getMaxColumns());
  }

  // Lưu nguyên cột Ghi chú vào cột tạm, gồm cả công thức và định dạng.
  summary.getRange(startRow, noteSourceCol, rowCount, 1).copyTo(
    summary.getRange(startRow, tempCol, rowCount, 1),
    SpreadsheetApp.CopyPasteType.PASTE_NORMAL,
    false
  );

  // Dịch các cột sản phẩm sang trái, bắt đầu từ cột E.
  summary.getRange(startRow, oldFirstProductCol, rowCount, products.length).copyTo(
    summary.getRange(startRow, newFirstProductCol, rowCount, products.length),
    SpreadsheetApp.CopyPasteType.PASTE_NORMAL,
    false
  );

  // Đưa Ghi chú ra cột cuối.
  summary.getRange(startRow, tempCol, rowCount, 1).copyTo(
    summary.getRange(startRow, newNoteCol, rowCount, 1),
    SpreadsheetApp.CopyPasteType.PASTE_NORMAL,
    false
  );

  summary.getRange(24, newNoteCol).setValue('Ghi chú');
  summary.setColumnWidth(newNoteCol, 175);
  for (let col = newFirstProductCol; col < newNoteCol; col++) summary.setColumnWidth(col, 125);

  // Xóa dữ liệu tạm và vùng sản phẩm cũ còn dư phía sau cột Ghi chú.
  summary.getRange(startRow, tempCol, rowCount, 1).clearContent().clearFormat();
  if (oldLastProductCol > newNoteCol) {
    summary.getRange(startRow, newNoteCol + 1, rowCount, oldLastProductCol - newNoteCol)
      .clearContent()
      .clearFormat();
  }

  summary.getRange(23, 1, rowCount, newNoteCol)
    .setFontFamily('Times New Roman')
    .setVerticalAlignment('middle')
    .setWrap(true)
    .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);

  summary.getRange(23, 1, 1, newNoteCol).breakApart().merge()
    .setValue('III. ĐÁNH GIÁ HIỆU QUẢ ĐẦU TƯ');

  SpreadsheetApp.flush();
}