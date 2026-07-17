const FSSP_CFG = Object.freeze({
  TECH: '01A. Kỹ thuật',
  REVENUE: '02. Doanh thu',
  COST: '03. Chi phí & Vốn',
  PROFIT: '03A. Lợi nhuận & Thuế',
  CASH: '04. Dòng tiền & Tài trợ',
  CASH_SUMMARY: '04A. TH dòng tiền',
  SUMMARY: '00. Tổng hợp',
  BACKUP_PREFIX: '__FS_SP_BAK__',
  BACKUP_PROPERTY: 'FS_SP_BACKUPS_V1',
  STATUS_PROPERTY: 'FS_SP_STATUS_V1'
});

function FS_SP_chonVaTinhMotSanPham() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const tech = ss.getSheetByName(FSSP_CFG.TECH);
  if (!tech) throw new Error('Không tìm thấy Sheet 01A. Kỹ thuật.');

  const products = FS_SP_docSanPham_(tech);
  if (!products.length) throw new Error('Block SAN_PHAM không có sản phẩm hợp lệ.');

  const message = products.map(p => p.code + ' - ' + p.name + (p.group ? ' - ' + p.group : '')).join('\n');
  const response = ui.prompt(
    'Tổng hợp hiệu quả từng sản phẩm',
    'Nhập Mã SP cần tính:\n\n' + message,
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  const code = String(response.getResponseText() || '').trim().toUpperCase();
  const product = products.find(p => p.code === code);
  if (!product) throw new Error('Mã SP "' + code + '" không tồn tại trong block SAN_PHAM.');

  FS_SP_tinhMotSanPham_(product, products);
}

function FS_SP_tinhMotSanPham_(product, products) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(3000)) throw new Error('Mô hình đang được một tiến trình khác sử dụng.');

  const ss = SpreadsheetApp.getActive();
  const props = PropertiesService.getDocumentProperties();

  try {
    if (props.getProperty(FSSP_CFG.BACKUP_PROPERTY)) {
      throw new Error('Đang tồn tại bản sao khôi phục của lần chạy trước. Hãy chạy menu "Khôi phục mô hình gốc" trước.');
    }

    FS_SP_chuanHoaCotSanPham_(ss, products);
    SpreadsheetApp.flush();

    const sourceNames = [
      FSSP_CFG.REVENUE,
      FSSP_CFG.COST,
      FSSP_CFG.PROFIT,
      FSSP_CFG.CASH,
      FSSP_CFG.CASH_SUMMARY,
      FSSP_CFG.SUMMARY
    ].filter(name => ss.getSheetByName(name));

    const backups = FS_SP_taoBanSao_(ss, sourceNames);
    props.setProperty(FSSP_CFG.BACKUP_PROPERTY, JSON.stringify(backups));
    props.setProperty(FSSP_CFG.STATUS_PROPERTY, 'Đang tính: ' + product.code);

    FS_SP_apDungKichBan_(ss.getSheetByName(FSSP_CFG.REVENUE), product.code, 'REVENUE');
    FS_SP_apDungKichBan_(ss.getSheetByName(FSSP_CFG.COST), product.code, 'COST');
    SpreadsheetApp.flush();

    if (typeof FS_lapSheet03A !== 'function') throw new Error('Không tìm thấy hàm FS_lapSheet03A.');
    if (typeof FS_hoiTuTaiTro !== 'function') throw new Error('Không tìm thấy hàm FS_hoiTuTaiTro.');

    FS_lapSheet03A();
    FS_hoiTuTaiTro();
    if (typeof FS_lapSheet04A === 'function') FS_lapSheet04A();

    if (typeof FS_lapSheet00_TheoDanhMuc === 'function') {
      FS_lapSheet00_TheoDanhMuc();
    } else if (typeof FS_lapSheet00 === 'function') {
      FS_lapSheet00();
    } else {
      throw new Error('Không tìm thấy hàm lập Sheet 00.');
    }
    SpreadsheetApp.flush();

    const summaryScenario = ss.getSheetByName(FSSP_CFG.SUMMARY);
    const rows = FS_SP_docDongChiTieu_(summaryScenario);
    const scenarioValues = summaryScenario
      .getRange(rows.totalRevenue, 4, rows.vatPayable - rows.totalRevenue + 1, 1)
      .getValues();

    FS_SP_khoiPhucTuBanSao_(ss, backups);
    props.deleteProperty(FSSP_CFG.BACKUP_PROPERTY);
    props.deleteProperty(FSSP_CFG.STATUS_PROPERTY);

    const summaryBase = ss.getSheetByName(FSSP_CFG.SUMMARY);
    FS_SP_chuanHoaCotSanPham_(ss, products);
    const baseRows = FS_SP_docDongChiTieu_(summaryBase);
    const targetCol = 5 + products.findIndex(p => p.code === product.code);
    summaryBase
      .getRange(baseRows.totalRevenue, targetCol, scenarioValues.length, 1)
      .setValues(scenarioValues);

    FS_SP_dinhDangCotKetQua_(summaryBase, baseRows, targetCol);
    SpreadsheetApp.flush();

    SpreadsheetApp.getUi().alert(
      'Đã tổng hợp hiệu quả riêng cho ' + product.code + ' - ' + product.name + '.\n' +
      'Mô hình gốc đã được khôi phục.'
    );
  } catch (error) {
    const backupJson = props.getProperty(FSSP_CFG.BACKUP_PROPERTY);
    if (backupJson) {
      try {
        FS_SP_khoiPhucTuBanSao_(ss, JSON.parse(backupJson));
        props.deleteProperty(FSSP_CFG.BACKUP_PROPERTY);
        props.deleteProperty(FSSP_CFG.STATUS_PROPERTY);
      } catch (restoreError) {
        throw new Error(
          error.message + '\nKhông thể tự khôi phục: ' + restoreError.message +
          '\nHãy chạy menu "Khôi phục mô hình gốc".'
        );
      }
    }
    throw error;
  } finally {
    lock.releaseLock();
  }
}

function FS_SP_khoiPhucMoHinhGoc() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(3000)) throw new Error('Mô hình đang được một tiến trình khác sử dụng.');

  try {
    const ss = SpreadsheetApp.getActive();
    const props = PropertiesService.getDocumentProperties();
    const backupJson = props.getProperty(FSSP_CFG.BACKUP_PROPERTY);
    if (!backupJson) {
      SpreadsheetApp.getUi().alert('Không có bản sao cần khôi phục.');
      return;
    }

    FS_SP_khoiPhucTuBanSao_(ss, JSON.parse(backupJson));
    props.deleteProperty(FSSP_CFG.BACKUP_PROPERTY);
    props.deleteProperty(FSSP_CFG.STATUS_PROPERTY);
    SpreadsheetApp.flush();
    SpreadsheetApp.getUi().alert('Đã khôi phục mô hình gốc từ bản sao an toàn.');
  } finally {
    lock.releaseLock();
  }
}

function FS_SP_taoBanSao_(ss, sheetNames) {
  const timestamp = Date.now();
  const backups = [];

  sheetNames.forEach((name, index) => {
    const source = ss.getSheetByName(name);
    if (!source) return;
    const backupName = FSSP_CFG.BACKUP_PREFIX + index + '_' + timestamp;
    const backup = source.copyTo(ss).setName(backupName);
    backup.hideSheet();
    backups.push({ source: name, backup: backupName });
  });

  return backups;
}

function FS_SP_khoiPhucTuBanSao_(ss, backups) {
  backups.forEach(item => {
    const source = ss.getSheetByName(item.source);
    const backup = ss.getSheetByName(item.backup);
    if (!source || !backup) {
      throw new Error('Thiếu sheet nguồn hoặc sheet sao lưu: ' + item.source);
    }

    const rows = backup.getMaxRows();
    const cols = backup.getMaxColumns();
    if (source.getMaxRows() < rows) source.insertRowsAfter(source.getMaxRows(), rows - source.getMaxRows());
    if (source.getMaxColumns() < cols) source.insertColumnsAfter(source.getMaxColumns(), cols - source.getMaxColumns());

    source.clear();
    backup.getRange(1, 1, rows, cols).copyTo(
      source.getRange(1, 1, rows, cols),
      SpreadsheetApp.CopyPasteType.PASTE_NORMAL,
      false
    );
    source.setFrozenRows(backup.getFrozenRows());
    source.setFrozenColumns(backup.getFrozenColumns());
  });

  backups.forEach(item => {
    const backup = ss.getSheetByName(item.backup);
    if (backup) ss.deleteSheet(backup);
  });
}

function FS_SP_apDungKichBan_(sheet, selectedCode, type) {
  if (!sheet) throw new Error('Thiếu sheet kịch bản ' + type + '.');

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return;

  const headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  const normalized = headers.map(FS_SP_key_);
  const codeCol = normalized.indexOf('masp') + 1;
  if (!codeCol) throw new Error('Không tìm thấy cột Mã SP trên ' + sheet.getName() + '.');

  const targetCols = [];
  normalized.forEach((key, index) => {
    if (FS_SP_laCotCanVe0_(key, type)) targetCols.push(index + 1);
  });
  if (!targetCols.length) throw new Error('Không xác định được các cột kịch bản trên ' + sheet.getName() + '.');

  const codes = sheet.getRange(2, codeCol, lastRow - 1, 1).getDisplayValues();
  targetCols.forEach(col => {
    const values = sheet.getRange(2, col, lastRow - 1, 1).getValues();
    for (let row = 0; row < values.length; row++) {
      const code = String(codes[row][0] || '').trim().toUpperCase();
      if (code && code !== selectedCode) values[row][0] = 0;
    }
    sheet.getRange(2, col, lastRow - 1, 1).setValues(values);
  });
}

function FS_SP_laCotCanVe0_(key, type) {
  if (!key) return false;

  if (type === 'REVENUE') {
    return [
      'tiendothutien', 'doanhthubantruocvat', 'doanhthuthu etruocvat',
      'doanhthuthuetr uocvat', 'tongdoanhthutruocvat', 'vatdaura',
      'dongtienkhachhang', 'thuetndn', 'giavon', 'lnst'
    ].map(x => x.replace(/ /g, '')).includes(key);
  }

  const exact = [
    'dongtienkhachhang', 'vatdaura', 'xdtbtruocvat', 'gpmbtruocvat',
    'htkttruocvat', 'tiensddtruocvat', 'tienthuedattruocvat',
    'chiphibanhangtruocvat', 'chiphivanhanhtruocvat',
    'chiphibaotritruocvat', 'chiphiduphongtruocvat', 'vatdauvao',
    'tongchitruocvat', 'tongchisauvat'
  ];
  if (exact.includes(key)) return true;

  return key.indexOf('chiphi') === 0 ||
    key.indexOf('xdtb') === 0 || key.indexOf('gpmb') === 0 ||
    key.indexOf('htkt') === 0 || key.indexOf('tiensdd') === 0 ||
    key.indexOf('tienthuedat') === 0 || key.indexOf('tongchi') === 0;
}

function FS_SP_chuanHoaCotSanPham_(ss, products) {
  const sheet = ss.getSheetByName(FSSP_CFG.SUMMARY);
  if (!sheet) throw new Error('Không tìm thấy Sheet 00. Tổng hợp.');

  const rows = FS_SP_docDongChiTieu_(sheet);
  const firstProductCol = 5;
  const noteCol = firstProductCol + products.length;
  const lastRow = rows.vatPayable;
  const neededCols = noteCol;
  if (sheet.getMaxColumns() < neededCols) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), neededCols - sheet.getMaxColumns());
  }

  const headers = sheet.getRange(24, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  let currentNoteCol = headers.findIndex(value => FS_SP_key_(value) === 'ghichu') + 1;
  if (!currentNoteCol) currentNoteCol = 5;

  const tempCol = Math.max(sheet.getLastColumn(), noteCol) + 1;
  if (sheet.getMaxColumns() < tempCol) sheet.insertColumnAfter(sheet.getMaxColumns());

  sheet.getRange(24, currentNoteCol, lastRow - 23, 1).copyTo(
    sheet.getRange(24, tempCol, lastRow - 23, 1),
    SpreadsheetApp.CopyPasteType.PASTE_NORMAL,
    false
  );

  const oldProductValues = {};
  for (let col = 5; col <= sheet.getLastColumn(); col++) {
    const header = String(sheet.getRange(24, col).getDisplayValue() || '').trim();
    if (!header || FS_SP_key_(header) === 'ghichu') continue;
    oldProductValues[FS_SP_key_(header)] = sheet.getRange(rows.totalRevenue, col, lastRow - rows.totalRevenue + 1, 1).getValues();
  }

  sheet.getRange(24, 5, lastRow - 23, tempCol - 4).breakApart().clearContent().clearFormat();

  products.forEach((product, index) => {
    const col = firstProductCol + index;
    const label = product.name + (product.group ? ' - ' + product.group : '');
    sheet.getRange(24, 4, lastRow - 23, 1).copyTo(
      sheet.getRange(24, col, lastRow - 23, 1),
      SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
      false
    );
    sheet.getRange(24, col).setValue(label).setFontWeight('bold').setHorizontalAlignment('center').setWrap(true);
    const old = oldProductValues[FS_SP_key_(label)];
    if (old) sheet.getRange(rows.totalRevenue, col, old.length, 1).setValues(old);
    else sheet.getRange(rows.totalRevenue, col, lastRow - rows.totalRevenue + 1, 1).setValue(0);
    FS_SP_dinhDangCotKetQua_(sheet, rows, col);
    sheet.setColumnWidth(col, 125);
  });

  sheet.getRange(24, tempCol, lastRow - 23, 1).copyTo(
    sheet.getRange(24, noteCol, lastRow - 23, 1),
    SpreadsheetApp.CopyPasteType.PASTE_NORMAL,
    false
  );
  sheet.getRange(24, noteCol).setValue('Ghi chú').setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange(24, tempCol, lastRow - 23, 1).clear();

  sheet.getRange(23, 1, 1, sheet.getMaxColumns()).breakApart();
  sheet.getRange(23, 1, 1, noteCol).merge().setValue('III. ĐÁNH GIÁ HIỆU QUẢ ĐẦU TƯ');
  sheet.getRange(23, 1, lastRow - 22, noteCol)
    .setFontFamily('Times New Roman')
    .setVerticalAlignment('middle')
    .setWrap(true)
    .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
}

function FS_SP_dinhDangCotKetQua_(sheet, rows, col) {
  sheet.getRange(rows.totalRevenue, col, rows.pat - rows.totalRevenue + 1, 1).setNumberFormat('#,##0.0');
  sheet.getRange(rows.npvProject, col).setNumberFormat('#,##0.0');
  sheet.getRange(rows.irrProject, col).setNumberFormat('0.00%');
  sheet.getRange(rows.paybackProject, col).setNumberFormat('0.00');
  sheet.getRange(rows.npvEquity, col).setNumberFormat('#,##0.0');
  sheet.getRange(rows.irrEquity, col).setNumberFormat('0.00%');
  sheet.getRange(rows.paybackEquity, col).setNumberFormat('0.00');
  sheet.getRange(rows.peakDebt, col, rows.vatPayable - rows.peakDebt + 1, 1).setNumberFormat('#,##0.0');
}

function FS_SP_docSanPham_(sheet) {
  const finder = sheet.createTextFinder('SAN_PHAM').matchEntireCell(true).findNext();
  if (!finder) throw new Error('Không tìm thấy block SAN_PHAM tại 01A. Kỹ thuật.');

  const startRow = finder.getRow() + 2;
  const lastRow = sheet.getLastRow();
  const values = sheet.getRange(startRow, 1, lastRow - startRow + 1, 3).getDisplayValues();
  const products = [];
  const seen = {};

  for (const row of values) {
    const code = String(row[0] || '').trim().toUpperCase();
    const name = String(row[1] || '').trim();
    const group = String(row[2] || '').trim();
    if (!code && !name) break;
    if (!code || !name || seen[code]) continue;
    seen[code] = true;
    products.push({ code, name, group });
  }
  return products;
}

function FS_SP_docDongChiTieu_(sheet) {
  const aliases = {
    totalRevenue: ['Tổng doanh thu có VAT'],
    totalCost: ['Tổng chi phí có VAT'],
    coreInvestment: ['Tổng vốn đầu tư dự án'],
    selling: ['Chi phí bán hàng'],
    operating: ['Chi phí vận hành'],
    maintenance: ['Chi phí bảo trì'],
    pat: ['Lợi nhuận sau thuế'],
    npvProject: ['NPV dự án'],
    irrProject: ['IRR dự án'],
    paybackProject: ['Thời gian hoàn vốn dự án'],
    npvEquity: ['NPV vốn CSH'],
    irrEquity: ['IRR vốn CSH'],
    paybackEquity: ['Thời gian hoàn vốn - Vốn CSH', 'Thời gian hoàn vốn vốn CSH'],
    peakDebt: ['Đỉnh dư nợ vay'],
    interest: ['Tổng lãi vay'],
    cit: ['Tổng Thuế TNDN'],
    vatPayable: ['Tổng VAT phải nộp']
  };

  const values = sheet.getRange(1, 2, sheet.getLastRow(), 1).getDisplayValues().flat();
  const rows = {};
  Object.keys(aliases).forEach(key => {
    for (const alias of aliases[key]) {
      const target = FS_SP_key_(alias);
      const index = values.findIndex(value => FS_SP_key_(value) === target);
      if (index >= 0) {
        rows[key] = index + 1;
        break;
      }
    }
    if (!rows[key]) throw new Error('Không tìm thấy dòng "' + aliases[key][0] + '" trên Sheet 00.');
  });
  return rows;
}

function FS_SP_key_(value) {
  return String(value == null ? '' : value)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd')
    .replace(/Đ/g, 'D')
    .toLowerCase()
    .replace(/[^a-z0-9]/g, '');
}
