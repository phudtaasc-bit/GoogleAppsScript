const FS = {
  INPUT: '01. Đầu vào',
  TECH: '01. Kỹ thuật',
  MENU: 'FS - CẬP NHẬT MÔ HÌNH'
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu(FS.MENU)
    .addItem('1. Tạo lại sheet Kỹ thuật từ Đầu vào', 'FS_taoKyThuatTuDauVao')
    .addSeparator()
    .addItem('2. Lập sheet 03 - Chi phí & vốn vay', 'FS_lapSheet03')
    .addItem('3. Lập sheet 02 - Doanh thu', 'FS_lapSheet02')
    .addItem('4. Lập sheet 04 - Dòng tiền', 'FS_lapSheet04')
    .addItem('5. Lập sheet 04A - TH dòng tiền', 'FS_lapSheet04A')
    .addItem('6. Lập sheet 00 - Tổng hợp', 'FS_lapSheet00')
    .addSeparator()
    .addItem('9. Chạy toàn bộ mô hình', 'FS_chayToanBoMoHinh')
    .addToUi();
}

function FS_chayToanBoMoHinh() {
  FS_taoKyThuatTuDauVao();
  FS_lapSheet03();
  FS_lapSheet02();
  FS_lapSheet04();
  FS_lapSheet04A();
  FS_lapSheet00();
}

function FS_taoKyThuatTuDauVao() {
  const ss = SpreadsheetApp.getActive();
  const input = ss.getSheetByName(FS.INPUT);
  let tech = ss.getSheetByName(FS.TECH);

  if (!input) throw new Error('Không tìm thấy sheet "01. Đầu vào".');
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

  row = FS_writeTable_(input, tech, row, {
    title: 'TIEN_DO_CHI_PHI',
    section: 'F. TIẾN ĐỘ CHI PHÍ',
    header: 'Khoản mục',
    headersOut: ['Khoản mục', 'Tháng bắt đầu', 'Thời gian', 'Tỷ lệ', 'Loại']
  });

  FS_formatTech_(tech);

  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert('Đã tạo lại sheet "01. Kỹ thuật".');
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

    { label: 'Diện tích đất', section: 'B. QUY HOẠCH', type: 'number2' },
    { label: 'Bắt đầu xây dựng', section: 'B. QUY HOẠCH', type: 'integer' },
    { label: 'Thời gian xây dựng', section: 'B. QUY HOẠCH', type: 'integer' },
    { label: 'Tỷ lệ vốn vay', section: 'B. QUY HOẠCH', type: 'percent' },
    { label: 'Lãi suất vay năm', section: 'B. QUY HOẠCH', type: 'percent' },
    { label: 'Tháng bắt đầu trả gốc', section: 'B. QUY HOẠCH', type: 'integer' },
    { label: 'Thời gian trả gốc', section: 'B. QUY HOẠCH', type: 'integer' }
  ];

  const out = [['THONG_TIN_CHUNG', 'GIÁ TRỊ', 'Ô NGUỒN', 'KIỂU DỮ LIỆU']];

  fields.forEach(f => {
    const cell = FS_findValueCellInSection_(input, f.section, f.label);
    out.push([f.label, cell ? cell.getValue() : '', cell ? cell.getA1Notation() : '', f.type]);
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

  table.rows.forEach(row => {
    out.push(cfg.headersOut.map(h => {
      const idx = FS_findHeaderIndex_(table.headers, h);
      return idx >= 0 ? row[idx] : '';
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

  const out = [[
    'Loại sản phẩm',
    'Hình thức',
    'DTKD',
    'Giá bán/m2',
    'Giá thuê/m2/tháng',
    'CPXD/m2',
    'VAT đầu ra',
    'Thuế TNDN',
    'Lấp đầy thuê',
    'CPVH thuê',
    'Ghi chú',
    'Diện tích đất'
  ]];

  table.rows.forEach(row => {
    const thueVH = FS_getByHeaderAny_(row, table.headers, [
      'Thuê/VH',
      'Thuê VH',
      'Thông tin thuê/VH'
    ]);

    const parsed = FS_parseThueVH_(thueVH);

    out.push([
      FS_getByHeaderAny_(row, table.headers, [
        'Loại sản phẩm',
        'Sản phẩm'
      ]),

      FS_getByHeaderAny_(row, table.headers, [
        'Hình thức'
      ]),

      FS_getByHeaderAny_(row, table.headers, [
        'DTKD',
        'Diện tích kinh doanh'
      ]),

      FS_getByHeaderAny_(row, table.headers, [
        'Giá bán trước thuế/m2',
        'Giá bán trước thuế /m2',
        'Giá bán trước thuế/m²',
        'Giá bán/m2',
        'Giá bán /m2',
        'Giá bán/m²',
        'Giá bán'
      ]),

      FS_getByHeaderAny_(row, table.headers, [
        'Giá thuê/m2/th',
        'Giá thuê/m2/tháng',
        'Giá thuê /m2/tháng',
        'Giá thuê'
      ]),

      FS_getByHeaderAny_(row, table.headers, [
        'CPXD/m2',
        'CPXD /m2',
        'Chi phí XD/m2',
        'Suất CPXD'
      ]),

      FS_getByHeaderAny_(row, table.headers, [
        'VAT đầu ra',
        'VAT'
      ]),

      FS_getByHeaderAny_(row, table.headers, [
        'Thuế TNDN',
        'TNDN'
      ]),

      parsed.lapDay || FS_getByHeaderAny_(row, table.headers, [
        'Lấp đầy thuê',
        'Tỷ lệ lấp đầy'
      ]),

      parsed.cpvh || FS_getByHeaderAny_(row, table.headers, [
        'CPVH thuê',
        'CPVH'
      ]),

      FS_getByHeaderAny_(row, table.headers, [
        'Ghi chú'
      ]),

      FS_getByHeaderAny_(row, table.headers, [
        'Diện tích đất',
        'DT đất'
      ])
    ]);
  });

  tech.getRange(startRow + 1, 1, out.length, out[0].length).setValues(out);
  return startRow + 1 + out.length;
}

function FS_parseThueVH_(value) {
  const s = String(value || '').toLowerCase();

  let lapDay = '';
  let cpvh = '';

  const m1 = s.match(/lấp\s*đầy\s*([0-9,.]+)\s*%/i);
  if (m1) lapDay = Number(m1[1].replace(',', '.')) / 100;

  const m2 = s.match(/cpvh\s*([0-9,.]+)\s*%/i);
  if (m2) cpvh = Number(m2[1].replace(',', '.')) / 100;

  return { lapDay, cpvh };
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

    const hasData = display.some(v => String(v).trim() !== '');
    const nextSection = display.some(v => /^[A-Z]\./.test(String(v).trim()));

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
    if (data[r].some(v => FS_norm_(v).includes(sectionNorm))) {
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
    if (data[r].some(v => FS_norm_(v).includes(target))) return r + 1;
  }

  return null;
}

function FS_findHeaderRowAfter_(sheet, headerText, startRow) {
  const target = FS_headerKey_(headerText);
  const maxRow = Math.min(startRow + 25, sheet.getLastRow());

  for (let r = startRow; r <= maxRow; r++) {
    const row = sheet.getRange(r, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
    if (row.some(v => FS_headerKey_(v) === target)) return r;
  }

  return null;
}

function FS_getLastColInRow_(sheet, row) {
  const values = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  let last = 1;

  values.forEach((v, i) => {
    if (String(v).trim() !== '') last = i + 1;
  });

  return last;
}

function FS_findHeaderIndex_(headers, name) {
  const target = FS_headerKey_(name);
  return headers.findIndex(h => FS_headerKey_(h) === target);
}

function FS_getByHeader_(row, headers, name) {
  const idx = FS_findHeaderIndex_(headers, name);
  return idx >= 0 ? row[idx] : '';
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

  [1, rowChiPhi, rowSanPham, rowKeHoach, rowTienDoCP].forEach(r => {
    if (r) sheet.getRange(r, 1, 1, Math.min(lastCol, 12)).setFontWeight('bold');
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
    sheet.getRange(rowSanPham + 1, 1, 1, 12).setFontWeight('bold');
    const start = rowSanPham + 2;
    const rows = rowKeHoach ? rowKeHoach - rowSanPham - 4 : 0;

    if (rows > 0) {
      sheet.getRange(start, 3, rows, 1).setNumberFormat('#,##0.00');
      sheet.getRange(start, 4, rows, 1).setNumberFormat('#,##0');
      sheet.getRange(start, 5, rows, 1).setNumberFormat('#,##0');
      sheet.getRange(start, 6, rows, 1).setNumberFormat('#,##0');
      sheet.getRange(start, 7, rows, 1).setNumberFormat('0.00%');
      sheet.getRange(start, 8, rows, 1).setNumberFormat('0.00%');
      sheet.getRange(start, 9, rows, 1).setNumberFormat('0.00%');
      sheet.getRange(start, 10, rows, 1).setNumberFormat('0.00%');
      sheet.getRange(start, 12, rows, 1).setNumberFormat('#,##0.00');
    }
  }

  if (rowKeHoach) {
    sheet.getRange(rowKeHoach + 1, 1, 1, 8).setFontWeight('bold');
    const start = rowKeHoach + 2;
    const rows = rowTienDoCP ? rowTienDoCP - rowKeHoach - 4 : 0;

    if (rows > 0) {
      sheet.getRange(start, 3, rows, 1).setNumberFormat('0');
      sheet.getRange(start, 4, rows, 1).setNumberFormat('0');
      sheet.getRange(start, 5, rows, 1).setNumberFormat('0');
      sheet.getRange(start, 6, rows, 1).setNumberFormat('0');
      sheet.getRange(start, 7, rows, 1).setNumberFormat('0.00%');
    }
  }

  if (rowTienDoCP) {
    sheet.getRange(rowTienDoCP + 1, 1, 1, 5).setFontWeight('bold');
    const start = rowTienDoCP + 2;
    const rows = Math.max(0, lastRow - start + 1);

    if (rows > 0) {
      sheet.getRange(start, 2, rows, 1).setNumberFormat('0');
      sheet.getRange(start, 3, rows, 1).setNumberFormat('0');
      sheet.getRange(start, 4, rows, 1).setNumberFormat('0.00%');
    }
  }

  const dateRow = FS_findRowContains_(sheet, 'Ngày bắt đầu dự án');
  if (dateRow) sheet.getRange(dateRow, 2).setNumberFormat('dd/mm/yyyy');

  ['Tỷ suất chiết khấu', 'Tỷ lệ tăng giá/năm', 'Tỷ lệ trượt chi phí/năm', 'Tỷ lệ vốn vay', 'Lãi suất vay năm']
    .forEach(label => {
      const r = FS_findRowContains_(sheet, label);
      if (r) sheet.getRange(r, 2).setNumberFormat('0.00%');
    });

  ['Số tháng mô hình', 'Bắt đầu xây dựng', 'Thời gian xây dựng', 'Tháng bắt đầu trả gốc', 'Thời gian trả gốc']
    .forEach(label => {
      const r = FS_findRowContains_(sheet, label);
      if (r) sheet.getRange(r, 2).setNumberFormat('0');
    });

  const areaRow = FS_findRowContains_(sheet, 'Diện tích đất');
  if (areaRow) sheet.getRange(areaRow, 2).setNumberFormat('#,##0.00');

  sheet.autoResizeColumns(1, Math.min(lastCol, 12));
}

function FS_norm_(v) {
  return String(v || '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd')
    .replace(/\s+/g, ' ')
    .trim();
}

function FS_headerKey_(v) {
  return FS_norm_(v)
    .replace(/²/g, '2')
    .replace(/\^2/g, '2')
    .replace(/m\s*2/g, 'm2')
    .replace(/m\s*²/g, 'm2')
    .replace(/[^a-z0-9]/g, '');
}