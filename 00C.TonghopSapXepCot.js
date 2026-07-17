/**
 * 00C.TonghopSapXepCot.js
 *
 * Chuẩn bị phần cột tại Mục III Sheet 00:
 * A: TT
 * B: Nội dung
 * C: Đơn vị
 * D: Toàn dự án
 * E...: từng sản phẩm
 * Cột cuối: Ghi chú
 *
 * Hàm chạy lặp không chèn cột vô hạn.
 */

function FS00C_chuanBiCotSanPham_(summary, products) {
  if (!summary) throw new Error('Không tìm thấy Sheet 00. Tổng hợp.');
  if (!products || !products.length) throw new Error('Danh sách sản phẩm rỗng.');

  const headerRow = FS00C_findHeaderRow_(summary);
  const firstProductColumn = 5; // E
  const productCount = products.length;
  const desiredNoteColumn = firstProductColumn + productCount;

  FS00C_ensureColumnCount_(summary, desiredNoteColumn);

  const currentNoteColumn = FS00C_findNoteColumn_(summary, headerRow);
  if (currentNoteColumn && currentNoteColumn !== desiredNoteColumn) {
    FS00C_moveNoteColumn_(summary, headerRow, currentNoteColumn, desiredNoteColumn);
  }

  // Xóa nội dung các cột E đến trước cột ghi chú trong phạm vi Mục III,
  // nhưng không xóa định dạng để tránh làm biến dạng mẫu.
  const lastSectionRow = FS00C_lastSectionRow_(summary, headerRow);
  if (lastSectionRow > headerRow) {
    summary.getRange(
      headerRow,
      firstProductColumn,
      lastSectionRow - headerRow + 1,
      productCount
    ).clearContent();
  }

  summary.getRange(headerRow, 4).setValue('Toàn dự án');

  products.forEach((product, i) => {
    const title = product.name + (product.group ? ' - ' + product.group : '');
    summary.getRange(headerRow, firstProductColumn + i).setValue(title);
  });

  summary.getRange(headerRow, desiredNoteColumn).setValue('Ghi chú');

  FS00C_copyHeaderFormat_(summary, headerRow, 4, firstProductColumn, productCount);
  FS00C_copyBodyFormat_(summary, headerRow, lastSectionRow, 4, firstProductColumn, productCount);

  // Sheet 00 không cố định hàng/cột để người dùng cuộn tự do.
  summary.setFrozenRows(0);
  summary.setFrozenColumns(0);

  summary.setColumnWidth(4, 110);
  summary.setColumnWidths(firstProductColumn, productCount, 125);
  summary.setColumnWidth(desiredNoteColumn, 210);

  return {
    headerRow,
    firstProductColumn,
    productCount,
    noteColumn: desiredNoteColumn,
    lastSectionRow
  };
}

function FS00C_findHeaderRow_(summary) {
  const lastRow = summary.getLastRow();
  const values = summary.getRange(1, 1, lastRow, Math.min(8, summary.getLastColumn())).getDisplayValues();

  for (let r = 0; r < values.length; r++) {
    const rowText = values[r].map(FS00C_key_).join('|');
    if (
      rowText.indexOf(FS00C_key_('Nội dung')) >= 0 &&
      rowText.indexOf(FS00C_key_('Đơn vị')) >= 0 &&
      (
        rowText.indexOf(FS00C_key_('Giá trị')) >= 0 ||
        rowText.indexOf(FS00C_key_('Toàn dự án')) >= 0
      )
    ) {
      // Chỉ nhận dòng tiêu đề thuộc phần III.
      const sectionAbove = values.slice(0, r + 1)
        .some(row => String(row[0] || '').indexOf('III.') === 0);
      if (sectionAbove) return r + 1;
    }
  }

  throw new Error('Không xác định được dòng tiêu đề Mục III trên Sheet 00.');
}

function FS00C_findNoteColumn_(summary, headerRow) {
  const lastCol = summary.getLastColumn();
  const headers = summary.getRange(headerRow, 1, 1, lastCol).getDisplayValues()[0];
  for (let c = 0; c < headers.length; c++) {
    if (FS00C_key_(headers[c]) === FS00C_key_('Ghi chú')) return c + 1;
  }
  return 0;
}

function FS00C_moveNoteColumn_(summary, headerRow, fromCol, toCol) {
  const lastRow = FS00C_lastSectionRow_(summary, headerRow);
  if (lastRow < headerRow) return;

  const rowCount = lastRow - headerRow + 1;
  const values = summary.getRange(headerRow, fromCol, rowCount, 1).getValues();
  const notes = summary.getRange(headerRow, fromCol, rowCount, 1).getNotes();
  const backgrounds = summary.getRange(headerRow, fromCol, rowCount, 1).getBackgrounds();
  const fontColors = summary.getRange(headerRow, fromCol, rowCount, 1).getFontColors();
  const fontWeights = summary.getRange(headerRow, fromCol, rowCount, 1).getFontWeights();
  const horizontalAlignments = summary.getRange(headerRow, fromCol, rowCount, 1).getHorizontalAlignments();
  const verticalAlignments = summary.getRange(headerRow, fromCol, rowCount, 1).getVerticalAlignments();
  const numberFormats = summary.getRange(headerRow, fromCol, rowCount, 1).getNumberFormats();

  summary.getRange(headerRow, toCol, rowCount, 1)
    .setValues(values)
    .setNotes(notes)
    .setBackgrounds(backgrounds)
    .setFontColors(fontColors)
    .setFontWeights(fontWeights)
    .setHorizontalAlignments(horizontalAlignments)
    .setVerticalAlignments(verticalAlignments)
    .setNumberFormats(numberFormats);

  if (fromCol !== toCol) {
    summary.getRange(headerRow, fromCol, rowCount, 1).clearContent().clearNote();
  }
}

function FS00C_copyHeaderFormat_(summary, headerRow, sourceCol, firstTargetCol, count) {
  const source = summary.getRange(headerRow, sourceCol, 1, 1);
  const target = summary.getRange(headerRow, firstTargetCol, 1, count);
  source.copyTo(target, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
}

function FS00C_copyBodyFormat_(summary, headerRow, lastRow, sourceCol, firstTargetCol, count) {
  if (lastRow <= headerRow) return;
  const rowCount = lastRow - headerRow;
  for (let i = 0; i < count; i++) {
    summary.getRange(headerRow + 1, sourceCol, rowCount, 1)
      .copyTo(
        summary.getRange(headerRow + 1, firstTargetCol + i, rowCount, 1),
        SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
        false
      );
  }
}

function FS00C_lastSectionRow_(summary, headerRow) {
  const lastRow = summary.getLastRow();
  const values = summary.getRange(headerRow + 1, 2, lastRow - headerRow, 1).getDisplayValues();

  let last = headerRow;
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0] || '').trim()) last = headerRow + 1 + i;
  }
  return last;
}

function FS00C_ensureColumnCount_(summary, requiredLastColumn) {
  const maxColumns = summary.getMaxColumns();
  if (maxColumns < requiredLastColumn) {
    summary.insertColumnsAfter(maxColumns, requiredLastColumn - maxColumns);
  }
}

function FS00C_key_(value) {
  return String(value == null ? '' : value)
    .trim()
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd')
    .replace(/[^a-z0-9]/g, '');
}
