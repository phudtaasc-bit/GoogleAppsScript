const FS = {
  INPUT: '01. Đầu vào',
  TECH: '01A. Kỹ thuật',
  TECH_LEGACY: '01. Kỹ thuật',
  MENU: 'FS - CẬP NHẬT MÔ HÌNH'
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu(FS.MENU)
    .addItem('1. Tạo lại sheet 01A. Kỹ thuật', 'FS_taoKyThuatTuDauVao')
    .addSeparator()
    .addItem('2. Lập sheet 02 - Doanh thu', 'FS_lapSheet02')
    .addItem('3. Lập sheet 03 - Chi phí & Vốn', 'FS_lapSheet03')
    .addItem('4. Lập sheet 03A - Lợi nhuận & Thuế', 'FS_lapSheet03A')
    .addItem('5. Lập sheet 04 - Dòng tiền & Tài trợ', 'FS_lapSheet04')
    .addItem('6. Lập sheet 04A - Tổng hợp dòng tiền', 'FS_lapSheet04A')
    .addItem('7. Lập sheet 00 - Tổng hợp', 'FS_lapSheet00')
    .addItem('8. Lập sheet 99 - Kiểm tra', 'FS_lapSheet99')
    .addSeparator()
    .addItem('9. Chạy toàn bộ mô hình', 'FS_chayToanBoMoHinh')
    .addToUi();
}

function FS_chayToanBoMoHinh() {
  FS_taoKyThuatTuDauVao();
  FS_runIfExists_('FS_lapSheet02');
  FS_runIfExists_('FS_lapSheet03');
  FS_runIfExists_('FS_lapSheet03A');
  FS_runIfExists_('FS_lapSheet04');
  FS_runIfExists_('FS_lapSheet04A');
  FS_runIfExists_('FS_lapSheet00');
  FS_runIfExists_('FS_lapSheet99');
}

function FS_runIfExists_(functionName) {
  const fn = globalThis[functionName];
  if (typeof fn === 'function') fn();
}

function FS_taoKyThuatTuDauVao() {
  const ss = SpreadsheetApp.getActive();
  const input = ss.getSheetByName(FS.INPUT);
  let tech = ss.getSheetByName(FS.TECH);
  const legacy = ss.getSheetByName(FS.TECH_LEGACY);

  if (!input) throw new Error('Không tìm thấy sheet "01. Đầu vào".');

  if (!tech && legacy) {
    legacy.setName(FS.TECH);
    tech = legacy;
  }
  if (!tech) tech = ss.insertSheet(FS.TECH);

  tech.clear();
  tech.clearFormats();

  let row = 1;

  row = FS_writeThongTinChung_(input, tech, row);
  row += 2;

  row = FS_writeTable_(input, tech, row, {
    title: 'CHI_PHI_CHUNG',
    section: 'C. CHI PHÍ CHUNG',
    header: 'Khoản mục',
    headersOut: ['Khoản mục', 'Trước VAT', 'VAT đầu vào', 'Sau VAT', 'Ghi chú', 'Tỷ lệ']
  });
  row += 2;

  row = FS_writeSanPham_(input, tech, row);
  row += 2;

  row = FS_writeTable_(input, tech, row, {
    title: 'KE_HOACH_BAN_THU_TIEN',
    section: 'E. KẾ HOẠCH BÁN HÀNG',
    header: 'Nhóm',
    headersOut: ['Nhóm', 'Loại sản phẩm', 'Số đợt', 'Đợt', 'Tháng bắt đầu', 'Thời gian', 'Tỷ lệ', 'Ghi chú']
  });
  row += 2;

  FS_writeTable_(input, tech, row, {
    title: 'TIEN_DO_CHI_PHI',
    section: 'F. TIẾN ĐỘ CHI PHÍ',
    header: 'Khoản mục',
    headersOut: ['Khoản mục', 'Tháng bắt đầu', 'Thời gian', 'Tỷ lệ', 'Loại']
  });

  FS_formatTech_(tech);
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert('Đã tạo lại sheet "01A. Kỹ thuật" theo cấu trúc chuẩn.');
}

function FS_writeThongTinChung_(input, tech, startRow) {
  const fields = [
    { label: 'Tên dự án', section: 'A. THÔNG TIN CHUNG', type: 'text' },
    { label: 'Ngày bắt đầu dự án', section: 'A. THÔNG TIN CHUNG', type: 'date' },
    { label: 'Số tháng mô hình', section: 'A. THÔNG TIN CHUNG', type: 'integer' },
    { label: 'Đơn vị tiền', section: 'A. THÔNG TIN CHUNG', type: 'text' },
    { label: 'Tỷ suất chiết khấu', section: 'A. THÔNG TIN CHUNG', type: 'percent' },
    { label: 'Tỷ lệ tăng giá/năm', section: 'A. THÔNG TIN CHUNG', type: 'percent' },
    { label: 'Tỷ lệ trượt chi phí/năm', section: 'A. THÔNG TIN CHUNG', type: 'percent' },
    { label: 'Diện tích đất', section: 'B. QUY HOẠCH', type: 'number' },
    { label: 'Bắt đầu xây dựng', section: 'B. QUY HOẠCH', type: 'integer' },
    { label: 'Thời gian xây dựng', section: 'B. QUY HOẠCH', type: 'integer' },
    { label: 'Tỷ lệ vốn vay', section: 'B. QUY HOẠCH', type: 'percent' },
    { label: 'Lãi suất vay năm', section: 'B. QUY HOẠCH', type: 'percent' },
    { label: 'Tháng bắt đầu trả gốc', section: 'B. QUY HOẠCH', type: 'integer' },
    { label: 'Thời gian trả gốc', section: 'B. QUY HOẠCH', type: 'integer' }
  ];

  const out = [['THONG_TIN_CHUNG', 'GIÁ TRỊ', 'Ô NGUỒN', 'KIỂU DỮ LIỆU']];

  fields.forEach(field => {
    const cell = FS_findValueCellInSection_(input, field.section, field.label);
    out.push([
      field.label,
      cell ? cell.getValue() : '',
      cell ? cell.getA1Notation() : '',
      field.type
    ]);
  });

  tech.getRange(startRow, 1, out.length, 4).setValues(out);
  return startRow + out.length;
}

function FS_writeTable_(input, tech, startRow, cfg) {
  const table = FS_getTable_(input, cfg.section, cfg.header);
  tech.getRange(startRow, 1).setValue(cfg.title);

  if (!table) {
    tech.getRange(startRow + 1, 1).setValue('Không tìm thấy dữ liệu');
    return startRow + 2;
  }

  const out = [cfg.headersOut];
  table.rows.forEach(sourceRow => {
    out.push(cfg.headersOut.map(header => {
      const idx = FS_findHeaderIndex_(table.headers, header);
      return idx >= 0 ? sourceRow[idx] : '';
    }));
  });

  tech.getRange(startRow + 1, 1, out.length, out[0].length).setValues(out);
  return startRow + 1 + out.length;
}

function FS_writeSanPham_(input, tech, startRow) {
  const table = FS_getTable_(input, 'D. CHI TIẾT SẢN PHẨM', 'Loại sản phẩm');
  tech.getRange(startRow, 1).setValue('SAN_PHAM');

  if (!table) {
    tech.getRange(startRow + 1, 1).setValue('Không tìm thấy dữ liệu');
    return startRow + 2;
  }

  const headersOut = [
    'Mã SP',
    'Loại SP',
    'Nhóm',
    'DTKD',
    'Giá bán',
    'Giá thuê',
    'CPXD',
    'VAT đầu ra',
    'Thuế TNDN',
    'Lấp đầy',
    'CPVH/doanh thu',
    'Chi phí bảo trì/doanh thu',
    'Thời gian thuê',
    'Diện tích đất'
  ];

  const out = [headersOut];

  table.rows.forEach(sourceRow => {
    const code = String(FS_getByHeaderAny_(sourceRow, table.headers, ['Mã SP', 'Mã sản phẩm']) || '').trim().toUpperCase();
    const productName = FS_getByHeaderAny_(sourceRow, table.headers, ['Loại sản phẩm', 'Loại SP', 'Sản phẩm']);
    const group = FS_getByHeaderAny_(sourceRow, table.headers, ['Nhóm', 'Nhóm sản phẩm']);

    if (!code && !productName) return;

    out.push([
      code,
      productName,
      group,
      FS_getByHeaderAny_(sourceRow, table.headers, ['DTKD', 'Diện tích kinh doanh']),
      FS_getByHeaderAny_(sourceRow, table.headers, [
        'Giá bán trước thuế/m2', 'Giá bán trước thuế /m2', 'Giá bán trước thuế/m²',
        'Giá bán/m2', 'Giá bán /m2', 'Giá bán/m²', 'Giá bán'
      ]),
      FS_getByHeaderAny_(sourceRow, table.headers, [
        'Giá thuê/m2/tháng', 'Giá thuê /m2/tháng', 'Giá thuê/m2/th', 'Giá thuê'
      ]),
      FS_getByHeaderAny_(sourceRow, table.headers, [
        'CPXD/m2', 'CPXD /m2', 'Chi phí XD/m2', 'Suất CPXD', 'CPXD'
      ]),
      FS_getByHeaderAny_(sourceRow, table.headers, ['VAT đầu ra', 'VAT']),
      FS_getByHeaderAny_(sourceRow, table.headers, ['Thuế TNDN', 'TNDN']),
      FS_getByHeaderAny_(sourceRow, table.headers, ['Lấp đầy', 'Lấp đầy thuê', 'Tỷ lệ lấp đầy']),
      FS_getByHeaderAny_(sourceRow, table.headers, [
        'CPVH/doanh thu', 'CPVH / doanh thu', 'Chi phí vận hành/doanh thu', 'CPVH'
      ]),
      FS_getByHeaderAny_(sourceRow, table.headers, [
        'Chi phí bảo trì/doanh thu', 'Chi phí bảo trì / doanh thu', 'CPBT/doanh thu', 'CPBT'
      ]),
      FS_getByHeaderAny_(sourceRow, table.headers, [
        'Thời gian thuê (năm)', 'Thời gian thuê', 'Số năm thuê'
      ]),
      FS_getByHeaderAny_(sourceRow, table.headers, ['Diện tích đất', 'DT đất'])
    ]);
  });

  const dataRows = out.slice(1);
  const invalidCodes = dataRows
    .map((row, index) => ({ row: index + 1, code: String(row[0] || '').trim() }))
    .filter(item => !item.code);

  if (invalidCodes.length) {
    throw new Error(
      'D. CHI TIẾT SẢN PHẨM còn thiếu Mã SP tại ' + invalidCodes.length +
      ' dòng dữ liệu. Mã SP là khóa bắt buộc, không tự suy diễn từ tên sản phẩm.'
    );
  }

  const duplicatedCodes = FS_findDuplicateValues_(dataRows.map(row => String(row[0]).trim().toUpperCase()));
  if (duplicatedCodes.length) {
    throw new Error('Mã SP bị trùng trong D. CHI TIẾT SẢN PHẨM: ' + duplicatedCodes.join(', '));
  }

  tech.getRange(startRow + 1, 1, out.length, headersOut.length).setValues(out);
  return startRow + 1 + out.length;
}

function FS_findDuplicateValues_(values) {
  const counts = {};
  values.forEach(value => {
    if (!value) return;
    counts[value] = (counts[value] || 0) + 1;
  });
  return Object.keys(counts).filter(value => counts[value] > 1);
}

function FS_getTable_(sheet, sectionText, headerText) {
  const sectionRow = FS_findRowContains_(sheet, sectionText);
  if (!sectionRow) return null;

  const headerRow = FS_findHeaderRowAfter_(sheet, headerText, sectionRow);
  if (!headerRow) return null;

  const lastCol = FS_getLastColInRow_(sheet, headerRow);
  const headers = sheet.getRange(headerRow, 1, 1, lastCol).getDisplayValues()[0];
  const rows = [];
  let blankCount = 0;

  for (let r = headerRow + 1; r <= sheet.getLastRow(); r++) {
    const display = sheet.getRange(r, 1, 1, lastCol).getDisplayValues()[0];
    const values = sheet.getRange(r, 1, 1, lastCol).getValues()[0];
    const hasData = display.some(value => String(value).trim() !== '');
    const nextSection = display.some(value => /^[A-Z]\./.test(String(value).trim()));

    if (nextSection) break;

    if (!hasData) {
      blankCount++;
      if (blankCount >= 3) break;
      continue;
    }

    blankCount = 0;
    rows.push(values);
  }

  return { headers, rows };
}

function FS_findValueCellInSection_(sheet, sectionText, label) {
  const data = sheet.getDataRange().getDisplayValues();
  const sectionNorm = FS_norm_(sectionText);
  const labelNorm = FS_norm_(label);
  let sectionRow = -1;

  for (let r = 0; r < data.length; r++) {
    if (data[r].some(value => FS_norm_(value).includes(sectionNorm))) {
      sectionRow = r;
      break;
    }
  }

  if (sectionRow < 0) return null;

  const endRow = Math.min(sectionRow + 15, data.length - 1);
  for (let r = sectionRow + 1; r <= endRow; r++) {
    for (let c = 0; c < data[r].length - 1; c++) {
      const text = FS_norm_(data[r][c]);
      if (!text) continue;

      const isMatch =
        text === labelNorm ||
        text === FS_norm_(label + '/WACC') ||
        text.startsWith(labelNorm + '/');

      if (isMatch) return sheet.getRange(r + 1, c + 2);
    }
  }

  return null;
}

function FS_findRowContains_(sheet, text) {
  const data = sheet.getDataRange().getDisplayValues();
  const target = FS_norm_(text);

  for (let r = 0; r < data.length; r++) {
    if (data[r].some(value => FS_norm_(value).includes(target))) return r + 1;
  }
  return null;
}

function FS_findHeaderRowAfter_(sheet, headerText, startRow) {
  const target = FS_headerKey_(headerText);
  const maxRow = Math.min(startRow + 25, sheet.getLastRow());

  for (let r = startRow; r <= maxRow; r++) {
    const row = sheet.getRange(r, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
    if (row.some(value => FS_headerKey_(value) === target)) return r;
  }
  return null;
}

function FS_getLastColInRow_(sheet, row) {
  const values = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  let last = 1;
  values.forEach((value, index) => {
    if (String(value).trim() !== '') last = index + 1;
  });
  return last;
}

function FS_findHeaderIndex_(headers, name) {
  const target = FS_headerKey_(name);
  return headers.findIndex(header => FS_headerKey_(header) === target);
}

function FS_getByHeaderAny_(row, headers, names) {
  for (const name of names) {
    const idx = FS_findHeaderIndex_(headers, name);
    if (idx >= 0) return row[idx];
  }
  return '';
}

function FS_formatTech_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) return;

  sheet.getRange(1, 1, lastRow, lastCol)
    .setFontFamily('Arial')
    .setFontSize(10)
    .setVerticalAlignment('middle');

  const rowChiPhi = FS_findRowContains_(sheet, 'CHI_PHI_CHUNG');
  const rowSanPham = FS_findRowContains_(sheet, 'SAN_PHAM');
  const rowKeHoach = FS_findRowContains_(sheet, 'KE_HOACH_BAN_THU_TIEN');
  const rowTienDoCP = FS_findRowContains_(sheet, 'TIEN_DO_CHI_PHI');

  [1, rowChiPhi, rowSanPham, rowKeHoach, rowTienDoCP].forEach(row => {
    if (row) sheet.getRange(row, 1, 1, lastCol).setFontWeight('bold');
  });

  if (rowChiPhi) {
    sheet.getRange(rowChiPhi + 1, 1, 1, 6).setFontWeight('bold');
    const start = rowChiPhi + 2;
    const rows = rowSanPham ? rowSanPham - rowChiPhi - 4 : 0;
    if (rows > 0) {
      sheet.getRange(start, 2, rows, 1).setNumberFormat('#,##0');
      sheet.getRange(start, 3, rows, 1).setNumberFormat('0.00%');
      sheet.getRange(start, 4, rows, 1).setNumberFormat('#,##0');
      sheet.getRange(start, 6, rows, 1).setNumberFormat('0.00%');
    }
  }

  if (rowSanPham) {
    sheet.getRange(rowSanPham + 1, 1, 1, 14).setFontWeight('bold');
    const start = rowSanPham + 2;
    const rows = rowKeHoach ? rowKeHoach - rowSanPham - 4 : 0;
    if (rows > 0) {
      sheet.getRange(start, 4, rows, 1).setNumberFormat('#,##0');
      sheet.getRange(start, 5, rows, 3).setNumberFormat('#,##0');
      sheet.getRange(start, 8, rows, 5).setNumberFormat('0.00%');
      sheet.getRange(start, 13, rows, 1).setNumberFormat('0');
      sheet.getRange(start, 14, rows, 1).setNumberFormat('#,##0');
    }
  }

  if (rowKeHoach) {
    sheet.getRange(rowKeHoach + 1, 1, 1, 8).setFontWeight('bold');
    const start = rowKeHoach + 2;
    const rows = rowTienDoCP ? rowTienDoCP - rowKeHoach - 4 : 0;
    if (rows > 0) {
      sheet.getRange(start, 3, rows, 4).setNumberFormat('0');
      sheet.getRange(start, 7, rows, 1).setNumberFormat('0.00%');
    }
  }

  if (rowTienDoCP) {
    sheet.getRange(rowTienDoCP + 1, 1, 1, 5).setFontWeight('bold');
    const start = rowTienDoCP + 2;
    const rows = Math.max(0, lastRow - start + 1);
    if (rows > 0) {
      sheet.getRange(start, 2, rows, 2).setNumberFormat('0');
      sheet.getRange(start, 4, rows, 1).setNumberFormat('0.00%');
    }
  }

  const dateRow = FS_findRowContains_(sheet, 'Ngày bắt đầu dự án');
  if (dateRow) sheet.getRange(dateRow, 2).setNumberFormat('dd/mm/yyyy');

  ['Tỷ suất chiết khấu', 'Tỷ lệ tăng giá/năm', 'Tỷ lệ trượt chi phí/năm', 'Tỷ lệ vốn vay', 'Lãi suất vay năm']
    .forEach(label => {
      const row = FS_findRowContains_(sheet, label);
      if (row) sheet.getRange(row, 2).setNumberFormat('0.00%');
    });

  ['Số tháng mô hình', 'Bắt đầu xây dựng', 'Thời gian xây dựng', 'Tháng bắt đầu trả gốc', 'Thời gian trả gốc']
    .forEach(label => {
      const row = FS_findRowContains_(sheet, label);
      if (row) sheet.getRange(row, 2).setNumberFormat('0');
    });

  const areaRow = FS_findRowContains_(sheet, 'Diện tích đất');
  if (areaRow) sheet.getRange(areaRow, 2).setNumberFormat('#,##0');

  sheet.autoResizeColumns(1, lastCol);
}

function FS_norm_(value) {
  return String(value || '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd')
    .replace(/\s+/g, ' ')
    .trim();
}

function FS_headerKey_(value) {
  return FS_norm_(value)
    .replace(/²/g, '2')
    .replace(/\^2/g, '2')
    .replace(/m\s*2/g, 'm2')
    .replace(/m\s*²/g, 'm2')
    .replace(/[^a-z0-9]/g, '');
}
