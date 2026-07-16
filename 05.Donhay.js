const FS05_CFG = Object.freeze({
  SHEET: '05. Độ nhạy',
  TECH: '01A. Kỹ thuật',
  SUMMARY: '00. Tổng hợp',
  CACHE: '_FS05_CACHE',
  FACTORS: [-0.10, -0.08, -0.05, -0.04, -0.02, 0, 0.02, 0.04, 0.05, 0.08, 0.10],
  MAX_RUNTIME_MS: 240000,
  COLOR_BASE: '#008000',
  COLOR_CENTER: '#FFFF00',
  COLOR_LOW: '#E6B8B7',
  COLOR_WHITE: '#FFFFFF',
  COLOR_HEADER: '#A6A6A6',
  COLOR_SECTION: '#FFC000',
  PROP_CURSOR: 'FS05_CURSOR',
  PROP_RUNNING: 'FS05_RUNNING'
});

function FS05_DoNhay_Fast() {
  return FS05_lapBangDoNhay();
}

function FS05_lapBangDoNhay() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  const tech = ss.getSheetByName(FS05_CFG.TECH);
  const summary = ss.getSheetByName(FS05_CFG.SUMMARY);
  if (!tech || !summary) {
    throw new Error('Cần lập đủ sheet "01A. Kỹ thuật" và "00. Tổng hợp" trước khi chạy độ nhạy.');
  }

  FS05_xoaTriggerTiepTuc_();
  PropertiesService.getDocumentProperties().deleteAllProperties();

  const sheet = FS05_getOrCreateSheet_(ss, FS05_CFG.SHEET);
  const cache = FS05_getOrCreateSheet_(ss, FS05_CFG.CACHE);
  cache.clear();
  cache.getRange(1, 1, 1, 6).setValues([['Nhóm', 'Dòng', 'Cột', 'NPV', 'IRR', 'Trạng thái']]);
  cache.hideSheet();

  FS05_taoKhung_(sheet);
  PropertiesService.getDocumentProperties().setProperties({
    [FS05_CFG.PROP_CURSOR]: '0',
    [FS05_CFG.PROP_RUNNING]: '1'
  });

  ui.alert(
    'Đã khởi tạo 363 kịch bản độ nhạy. Mô hình bắt đầu chạy theo lô và tự tiếp tục nếu vượt giới hạn thời gian.'
  );
  return FS05_tiepTucDoNhay();
}

function FS05_tiepTucDoNhay() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) return;

  const ss = SpreadsheetApp.getActive();
  const props = PropertiesService.getDocumentProperties();
  const startedAt = Date.now();
  let cursor = Number(props.getProperty(FS05_CFG.PROP_CURSOR) || 0);
  const total = 3 * FS05_CFG.FACTORS.length * FS05_CFG.FACTORS.length;

  try {
    if (props.getProperty(FS05_CFG.PROP_RUNNING) !== '1') return;

    const tech = ss.getSheetByName(FS05_CFG.TECH);
    const cache = ss.getSheetByName(FS05_CFG.CACHE);
    const sheet = ss.getSheetByName(FS05_CFG.SHEET);
    if (!tech || !cache || !sheet) throw new Error('Thiếu sheet phục vụ chạy độ nhạy.');

    const snapshot = FS05_docDauVao_(tech);

    try {
      while (cursor < total && Date.now() - startedAt < FS05_CFG.MAX_RUNTIME_MS) {
        const scenario = FS05_giaiMaKichBan_(cursor);
        FS05_apDungKichBan_(tech, snapshot, scenario);
        FS05_chayLaiMoHinh_();

        const kpi = FS05_docKpi_();
        cache.appendRow([
          scenario.group,
          scenario.rowIndex,
          scenario.colIndex,
          kpi.npv,
          kpi.irr,
          'OK'
        ]);

        cursor++;
        props.setProperty(FS05_CFG.PROP_CURSOR, String(cursor));
        FS05_capNhatTienDo_(sheet, cursor, total);
      }
    } finally {
      FS05_khoiPhucDauVao_(tech, snapshot);
      FS05_chayLaiMoHinh_();
    }

    if (cursor >= total) {
      FS05_ghiKetQua_();
      props.deleteProperty(FS05_CFG.PROP_CURSOR);
      props.deleteProperty(FS05_CFG.PROP_RUNNING);
      FS05_xoaTriggerTiepTuc_();
      ss.deleteSheet(cache);
      SpreadsheetApp.getUi().alert('Đã hoàn thành toàn bộ 363 kịch bản độ nhạy.');
    } else {
      FS05_taoTriggerTiepTuc_();
    }
  } catch (error) {
    props.deleteProperty(FS05_CFG.PROP_RUNNING);
    FS05_xoaTriggerTiepTuc_();
    throw error;
  } finally {
    lock.releaseLock();
  }
}

function FS05_giaiMaKichBan_(cursor) {
  const n = FS05_CFG.FACTORS.length;
  const perGroup = n * n;
  const groupIndex = Math.floor(cursor / perGroup);
  const within = cursor % perGroup;
  const rowIndex = Math.floor(within / n);
  const colIndex = within % n;
  const groups = ['SALE_COST', 'RATE_COST', 'RENT_COST'];

  return {
    group: groups[groupIndex],
    rowIndex,
    colIndex,
    costChange: FS05_CFG.FACTORS[rowIndex],
    horizontalChange: FS05_CFG.FACTORS[colIndex]
  };
}

function FS05_docDauVao_(tech) {
  const productBlock = FS05_timBlock_(tech, 'SAN_PHAM');
  const costBlock = FS05_timBlock_(tech, 'CHI_PHI_CHUNG');
  const loanRateCell = FS05_timOThongTin_(tech, 'Lãi suất vay năm');
  if (!loanRateCell) throw new Error('Không tìm thấy "Lãi suất vay năm" tại 01A. Kỹ thuật.');

  const productRows = tech.getRange(productBlock.startRow, 1, productBlock.rowCount, 14).getValues();
  const products = productRows.map((row, index) => ({
    row: productBlock.startRow + index,
    code: String(row[0] || '').trim(),
    group: FS05_chuanHoaNhom_(row[2]),
    salePrice: row[4],
    rentPrice: row[5]
  })).filter(item => item.code);

  const costRows = tech.getRange(costBlock.startRow, 1, costBlock.rowCount, 6).getValues();
  const costs = costRows.map((row, index) => ({
    row: costBlock.startRow + index,
    label: String(row[0] || '').trim(),
    value: row[1],
    formula: tech.getRange(costBlock.startRow + index, 2).getFormula()
  })).filter(item => FS05_laChiPhiVonDauTu_(item.label));

  return {
    products,
    costs,
    loanRate: loanRateCell.getValue(),
    loanRateFormula: loanRateCell.getFormula(),
    loanRateA1: loanRateCell.getA1Notation()
  };
}

function FS05_apDungKichBan_(tech, snapshot, scenario) {
  FS05_khoiPhucDauVao_(tech, snapshot);

  const costFactor = 1 + scenario.costChange;
  snapshot.costs.forEach(item => {
    tech.getRange(item.row, 2).setValue(FS05_so_(item.value) * costFactor);
  });

  if (scenario.group === 'SALE_COST') {
    const factor = 1 + scenario.horizontalChange;
    snapshot.products.filter(item => item.group === 'Bán').forEach(item => {
      tech.getRange(item.row, 5).setValue(FS05_so_(item.salePrice) * factor);
    });
  }

  if (scenario.group === 'RENT_COST') {
    const factor = 1 + scenario.horizontalChange;
    snapshot.products.filter(item => item.group === 'Cho thuê').forEach(item => {
      tech.getRange(item.row, 6).setValue(FS05_so_(item.rentPrice) * factor);
    });
  }

  if (scenario.group === 'RATE_COST') {
    const newRate = Math.max(0, FS05_tyLe_(snapshot.loanRate) + scenario.horizontalChange);
    tech.getRange(snapshot.loanRateA1).setValue(newRate);
  }

  SpreadsheetApp.flush();
}

function FS05_khoiPhucDauVao_(tech, snapshot) {
  snapshot.products.forEach(item => {
    tech.getRange(item.row, 5).setValue(item.salePrice);
    tech.getRange(item.row, 6).setValue(item.rentPrice);
  });

  snapshot.costs.forEach(item => {
    const cell = tech.getRange(item.row, 2);
    if (item.formula) cell.setFormula(item.formula);
    else cell.setValue(item.value);
  });

  const loanCell = tech.getRange(snapshot.loanRateA1);
  if (snapshot.loanRateFormula) loanCell.setFormula(snapshot.loanRateFormula);
  else loanCell.setValue(snapshot.loanRate);
  SpreadsheetApp.flush();
}

function FS05_chayLaiMoHinh_() {
  FS_lapSheet02();
  FS_lapSheet03();
  FS_lapSheet03A();
  FS_lapSheet00();
  SpreadsheetApp.flush();
}

function FS05_docKpi_() {
  const summary = SpreadsheetApp.getActive().getSheetByName(FS05_CFG.SUMMARY);
  const rows = {
    npvProject: FS05_timDongChiTieu_(summary, 'NPV dự án'),
    irrProject: FS05_timDongChiTieu_(summary, 'IRR dự án'),
    npvEquity: FS05_timDongChiTieu_(summary, 'NPV vốn CSH'),
    irrEquity: FS05_timDongChiTieu_(summary, 'IRR vốn CSH')
  };

  Object.keys(rows).forEach(key => {
    if (!rows[key]) throw new Error('Không tìm thấy chỉ tiêu ' + key + ' trên Sheet 00.');
  });

  return {
    project: {
      npv: FS05_so_(summary.getRange(rows.npvProject, 4).getValue()),
      irr: FS05_so_(summary.getRange(rows.irrProject, 4).getValue())
    },
    equity: {
      npv: FS05_so_(summary.getRange(rows.npvEquity, 4).getValue()),
      irr: FS05_so_(summary.getRange(rows.irrEquity, 4).getValue())
    },
    get npv() { return 0; },
    get irr() { return 0; }
  };
}

function FS05_ghiKetQua_() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(FS05_CFG.SHEET);
  const cache = ss.getSheetByName(FS05_CFG.CACHE);
  const summary = ss.getSheetByName(FS05_CFG.SUMMARY);
  const data = cache.getLastRow() > 1
    ? cache.getRange(2, 1, cache.getLastRow() - 1, 6).getValues()
    : [];

  const base = {
    npvProject: FS05_so_(summary.getRange(FS05_timDongChiTieu_(summary, 'NPV dự án'), 4).getValue()),
    irrProject: FS05_so_(summary.getRange(FS05_timDongChiTieu_(summary, 'IRR dự án'), 4).getValue()),
    npvEquity: FS05_so_(summary.getRange(FS05_timDongChiTieu_(summary, 'NPV vốn CSH'), 4).getValue()),
    irrEquity: FS05_so_(summary.getRange(FS05_timDongChiTieu_(summary, 'IRR vốn CSH'), 4).getValue())
  };

  const matrices = {};
  ['SALE_COST', 'RATE_COST', 'RENT_COST'].forEach(group => {
    matrices[group] = {
      npv: FS05_maTranRong_(),
      irr: FS05_maTranRong_()
    };
  });

  data.forEach(row => {
    const group = String(row[0]);
    const r = Number(row[1]);
    const c = Number(row[2]);
    if (!matrices[group]) return;
    matrices[group].npv[r][c] = FS05_so_(row[3]);
    matrices[group].irr[r][c] = FS05_so_(row[4]);
  });

  FS05_ghiBlock_(sheet, 4, 'NPV', 'Tăng/giảm giá bán', base.npvProject, matrices.SALE_COST.npv, '#,##0.0');
  FS05_ghiBlock_(sheet, 20, 'IRR', 'Tăng/giảm giá bán', base.irrProject, matrices.SALE_COST.irr, '0.0%');
  FS05_ghiBlock_(sheet, 36, 'NPV VỐN CSH', 'Tăng/giảm lãi suất vay', base.npvEquity, matrices.RATE_COST.npv, '#,##0.0');
  FS05_ghiBlock_(sheet, 52, 'IRR VỐN CSH', 'Tăng/giảm lãi suất vay', base.irrEquity, matrices.RATE_COST.irr, '0.0%');
  FS05_ghiBlock_(sheet, 68, 'NPV - GIÁ THUÊ/VỐN ĐẦU TƯ', 'Tăng/giảm giá thuê', base.npvProject, matrices.RENT_COST.npv, '#,##0.0');
  FS05_ghiBlock_(sheet, 84, 'IRR - GIÁ THUÊ/VỐN ĐẦU TƯ', 'Tăng/giảm giá thuê', base.irrProject, matrices.RENT_COST.irr, '0.0%');

  FS05_dinhDangToanSheet_(sheet);
  sheet.getRange('O1:P3').clearContent();
}

function FS05_docKpi_() {
  const summary = SpreadsheetApp.getActive().getSheetByName(FS05_CFG.SUMMARY);
  const scenario = FS05_giaiMaKichBan_(Number(PropertiesService.getDocumentProperties().getProperty(FS05_CFG.PROP_CURSOR) || 0));
  const isEquity = scenario.group === 'RATE_COST';
  const npvLabel = isEquity ? 'NPV vốn CSH' : 'NPV dự án';
  const irrLabel = isEquity ? 'IRR vốn CSH' : 'IRR dự án';
  const npvRow = FS05_timDongChiTieu_(summary, npvLabel);
  const irrRow = FS05_timDongChiTieu_(summary, irrLabel);
  if (!npvRow || !irrRow) throw new Error('Không tìm thấy NPV/IRR cần đọc trên Sheet 00.');
  return {
    npv: FS05_so_(summary.getRange(npvRow, 4).getValue()),
    irr: FS05_so_(summary.getRange(irrRow, 4).getValue())
  };
}

function FS05_taoKhung_(sheet) {
  sheet.getRange(1, 1, Math.max(sheet.getMaxRows(), 100), 16).breakApart();
  sheet.clear();
  sheet.clearFormats();
  sheet.getRange('A1:M1').merge().setValue('05. BẢNG PHÂN TÍCH ĐỘ NHẠY');
  sheet.getRange('A2:M2').merge().setValue('Màu hồng: kết quả thấp hơn Base. Màu trắng: kết quả lớn hơn hoặc bằng Base.');
  sheet.getRange('O1').setValue('Đang chạy');
  sheet.getRange('P1').setValue('0/363');
  FS05_dinhDangToanSheet_(sheet);
}

function FS05_ghiBlock_(sheet, startRow, title, horizontalTitle, base, matrix, numberFormat) {
  const n = FS05_CFG.FACTORS.length;
  const centerRow = startRow + 8;
  const centerCol = 8;

  sheet.getRange(startRow, 1, 1, 13)
    .setBackground(FS05_CFG.COLOR_SECTION)
    .setFontWeight('bold');
  sheet.getRange(startRow, 2).setValue(title).setFontWeight('bold').setFontSize(12);

  sheet.getRange(startRow + 1, 3, 1, n)
    .merge()
    .setValue(horizontalTitle)
    .setFontWeight('bold');

  sheet.getRange(startRow + 2, 2)
    .setValue(base)
    .setNumberFormat(numberFormat)
    .setBackground(FS05_CFG.COLOR_BASE)
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');

  sheet.getRange(startRow + 2, 3, 1, n)
    .setValues([FS05_CFG.FACTORS])
    .setNumberFormat('0.0%')
    .setBackground(FS05_CFG.COLOR_HEADER)
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');

  sheet.getRange(startRow + 3, 1, n, 1)
    .merge()
    .setValue('Tăng/giảm\nvốn đầu tư')
    .setFontWeight('bold')
    .setWrap(true);

  sheet.getRange(startRow + 3, 2, n, 1)
    .setValues(FS05_CFG.FACTORS.map(value => [value]))
    .setNumberFormat('0.0%')
    .setFontWeight('bold');

  const resultRange = sheet.getRange(startRow + 3, 3, n, n);
  resultRange.setValues(matrix).setNumberFormat(numberFormat);

  const backgrounds = matrix.map(row => row.map(value =>
    FS05_so_(value) < FS05_so_(base) ? FS05_CFG.COLOR_LOW : FS05_CFG.COLOR_WHITE
  ));
  resultRange.setBackgrounds(backgrounds);

  sheet.getRange(startRow + 2, 2, n + 1, n + 1)
    .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(startRow + 3, 1, n, n + 2)
    .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);

  sheet.getRange(centerRow, centerCol)
    .setBackground(FS05_CFG.COLOR_CENTER)
    .setFontWeight('bold')
    .setBorder(true, true, true, true, true, true, '#FF0000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}

function FS05_dinhDangToanSheet_(sheet) {
  const lastRow = 97;
  sheet.getRange(1, 1, lastRow, 13)
    .setFontFamily('Times New Roman')
    .setFontSize(10)
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('center');

  sheet.getRange('A1:M1')
    .setFontSize(14)
    .setFontWeight('bold')
    .setFontColor('#FFFFFF')
    .setBackground('#1F4E78');
  sheet.getRange('A2:M2').setFontStyle('italic').setBackground('#D9EAF7');

  sheet.setColumnWidth(1, 90);
  sheet.setColumnWidth(2, 90);
  for (let col = 3; col <= 13; col++) sheet.setColumnWidth(col, 82);
  sheet.setFrozenRows(2);
}

function FS05_capNhatTienDo_(sheet, current, total) {
  sheet.getRange('O1').setValue('Đang chạy');
  sheet.getRange('P1').setValue(current + '/' + total);
  SpreadsheetApp.flush();
}

function FS05_taoTriggerTiepTuc_() {
  FS05_xoaTriggerTiepTuc_();
  ScriptApp.newTrigger('FS05_tiepTucDoNhay').timeBased().after(60000).create();
}

function FS05_xoaTriggerTiepTuc_() {
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === 'FS05_tiepTucDoNhay') ScriptApp.deleteTrigger(trigger);
  });
}

function FS05_timBlock_(sheet, blockName) {
  const values = sheet.getDataRange().getDisplayValues();
  const target = FS05_key_(blockName);
  let blockRow = 0;
  for (let row = 0; row < values.length; row++) {
    if (values[row].some(value => FS05_key_(value) === target)) {
      blockRow = row + 1;
      break;
    }
  }
  if (!blockRow) throw new Error('Không tìm thấy block ' + blockName + ' tại 01A. Kỹ thuật.');

  const startRow = blockRow + 2;
  let endRow = startRow - 1;
  for (let row = startRow; row <= sheet.getLastRow(); row++) {
    const first = String(sheet.getRange(row, 1).getDisplayValue() || '').trim();
    if (row > startRow && /^[A-Z_]+$/.test(first)) break;
    if (!first) {
      if (endRow >= startRow) break;
      continue;
    }
    endRow = row;
  }
  return { startRow, rowCount: Math.max(0, endRow - startRow + 1) };
}

function FS05_timOThongTin_(sheet, label) {
  const target = FS05_key_(label);
  const values = sheet.getDataRange().getDisplayValues();
  for (let row = 0; row < values.length; row++) {
    if (FS05_key_(values[row][0]) === target) return sheet.getRange(row + 1, 2);
  }
  return null;
}

function FS05_timDongChiTieu_(sheet, label) {
  const target = FS05_key_(label);
  const values = sheet.getRange(1, 2, sheet.getLastRow(), 1).getDisplayValues();
  for (let row = 0; row < values.length; row++) {
    if (FS05_key_(values[row][0]) === target) return row + 1;
  }
  return 0;
}

function FS05_laChiPhiVonDauTu_(label) {
  const key = FS05_key_(label);
  if (!key || key.includes('tongmucdautu')) return false;
  return [
    'chiphixdtbkhac',
    'chiphixdtb',
    'chiphigpmb',
    'tiensdd',
    'tienthuedat',
    'chiphihtkt'
  ].some(token => key.includes(token));
}

function FS05_chuanHoaNhom_(value) {
  const key = FS05_key_(value);
  if (key === 'ban') return 'Bán';
  if (key === 'chothue' || key === 'thue') return 'Cho thuê';
  return String(value || '').trim();
}

function FS05_maTranRong_() {
  return FS05_CFG.FACTORS.map(() => FS05_CFG.FACTORS.map(() => 0));
}

function FS05_getOrCreateSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function FS05_tyLe_(value) {
  const number = FS05_so_(value);
  return Math.abs(number) > 1 ? number / 100 : number;
}

function FS05_so_(value) {
  if (typeof value === 'number') return isFinite(value) ? value : 0;
  const text = String(value == null ? '' : value).trim().replace(/\s/g, '');
  if (!text) return 0;
  const normalized = text.includes(',') && text.includes('.')
    ? text.replace(/\./g, '').replace(',', '.')
    : text.replace(/,/g, '');
  const number = Number(normalized);
  return isFinite(number) ? number : 0;
}

function FS05_key_(value) {
  return String(value == null ? '' : value)
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd')
    .replace(/[^a-z0-9]/g, '');
}
