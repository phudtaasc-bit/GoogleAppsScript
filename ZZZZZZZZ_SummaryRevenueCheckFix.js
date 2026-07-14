/*************************************************
 * ZZZZZZZZ_SummaryRevenueCheckFix.js
 * Giữ nguyên mục I, II của hàm gốc; hiệu chỉnh VAT mục I và dựng riêng mục III.
 *
 * Nguyên tắc:
 * - VAT chi phí đầu tư tính trực tiếp theo từng cấu phần, không phân bổ tổng VAT đầu vào theo tỷ trọng.
 * - Doanh thu chỉ hiển thị sản phẩm có phát sinh dòng tiền khách hàng.
 * - Chi phí bán hàng, vận hành, bảo trì chỉ hiển thị khi có giá trị.
 * - Mục 2 đã ghi "Tổng chi phí có VAT", các mục con không lặp lại VAT.
 * - Không dùng địa chỉ dòng cố định cho nội dung động tại mục III.
 *************************************************/

const FSZZZZZZZZ_BASE_LAP_SHEET00_ = FS_lapSheet00;

FS_lapSheet00 = function() {
  FSZZZZZZZZ_BASE_LAP_SHEET00_();
  FSZZZZZZZZ_buildFlexibleSectionIII_();
};

function FSZZZZZZZZ_suaTongHopDoanhThuVaKiemTra() {
  FSZZZZZZZZ_BASE_LAP_SHEET00_();
  FSZZZZZZZZ_buildFlexibleSectionIII_();
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert('Đã hiệu chỉnh VAT theo từng cấu phần và dựng lại mục III theo dữ liệu thực tế.');
}

function FSZZZZZZZZ_buildFlexibleSectionIII_() {
  const ss = SpreadsheetApp.getActive();
  const sh00 = ss.getSheetByName('00. Tổng hợp');
  const tech = ss.getSheetByName('01. Kỹ thuật');
  const sh02 = ss.getSheetByName('02. Doanh thu');
  const sh03 = ss.getSheetByName('03. Chi phí & Vốn');
  const sh04 = ss.getSheetByName('04. Dòng tiền & Lợi nhuận');
  if (!sh00 || !tech || !sh02 || !sh03 || !sh04) {
    throw new Error('Thiếu Sheet 00, 01. Kỹ thuật, 02, 03 hoặc 04.');
  }

  const h02 = FSZZZZZZZZ_detectHeader_(sh02, ['Loại sản phẩm', 'Dòng tiền huy động từ KH']);
  const h03 = FSZZZZZZZZ_detectHeader_(sh03, ['Tổng chi sau VAT', 'Chi phí bán hàng trước VAT']);
  const h04 = FSZZZZZZZZ_detectHeader_(sh04, ['FCFF_TIPV', 'Thuế TNDN', 'VAT phải nộp']);

  const productCol = FSZZZZZZZZ_requireCol_(h02.headers, ['Loại sản phẩm']);
  const customerCashCol = FSZZZZZZZZ_requireCol_(h02.headers, ['Dòng tiền huy động từ KH', 'Dòng tiền huy động từ khách hàng']);

  const cpxdCol = FSZZZZZZZZ_requireCol_(h03.headers, ['Chi XD/TB/khác trước VAT', 'Chi phí XD/TB/khác trước VAT']);
  const gpmbCol = FSZZZZZZZZ_requireCol_(h03.headers, ['Chi GPMB trước VAT', 'Chi phí GPMB trước VAT']);
  const landCol = FSZZZZZZZZ_requireCol_(h03.headers, ['Tiền SDĐ/thuê đất trước VAT']);
  const htktCol = FSZZZZZZZZ_requireCol_(h03.headers, ['Chi HTKT trước VAT', 'Chi phí HTKT trước VAT']);
  const sellingCol = FSZZZZZZZZ_requireCol_(h03.headers, ['Chi phí bán hàng trước VAT']);
  const reserveCol = FSZZZZZZZZ_requireCol_(h03.headers, ['Chi phí dự phòng trước VAT']);
  const opCol = FSZZZZZZZZ_requireCol_(h03.headers, ['Chi phí vận hành thuê trước VAT', 'Chi phí vận hành trước VAT']);
  const maintCol = FSZZZZZZZZ_requireCol_(h03.headers, ['Chi phí bảo trì trước VAT']);
  const totalAfterVatCol = FSZZZZZZZZ_requireCol_(h03.headers, ['Tổng chi sau VAT']);

  const profitAfterTaxCol = FSZZZZZZZZ_requireCol_(h04.headers, ['Lợi nhuận sau thuế']);
  const citCol = FSZZZZZZZZ_requireCol_(h04.headers, ['Thuế TNDN']);
  const vatPayCol = FSZZZZZZZZ_requireCol_(h04.headers, ['VAT phải nộp']);
  const interestCol = FSZZZZZZZZ_requireCol_(h04.headers, ['Lãi vay vốn hóa']);

  const tol = 0.5;

  const cpxdVatRate = FSZZZZZZZZ_getCommonCostVatRate_(tech,
    ['Chi phí XD/TB/khác', 'Chi phí XD/TB', 'Chi phí xây dựng & thiết bị'], 0);
  const htktVatRate = FSZZZZZZZZ_getCommonCostVatRate_(tech,
    ['Chi phí HTKT'], cpxdVatRate);
  const reserveVatRate = FSZZZZZZZZ_getCommonCostVatRate_(tech,
    ['Chi phí dự phòng'], cpxdVatRate);
  const sellingVatRate = FSZZZZZZZZ_getCommonCostVatRate_(tech,
    ['Chi phí bán hàng'], 0);
  const opVatRate = FSZZZZZZZZ_getCommonCostVatRate_(tech,
    ['Chi phí vận hành', 'Chi phí vận hành thuê'], sellingVatRate);
  const maintVatRate = FSZZZZZZZZ_getCommonCostVatRate_(tech,
    ['Chi phí bảo trì', 'Chi phí bảo hành, bảo trì'], opVatRate);

  const sum03 = col => FSZZZZZZZZ_sumColumn_(sh03, col, h03.dataStartRow);
  const sum04 = col => FSZZZZZZZZ_sumColumn_(sh04, col, h04.dataStartRow);

  // MỤC I: tính VAT theo đúng từng cấu phần, tuyệt đối không phân bổ tổng VAT đầu vào theo tỷ trọng.
  const cpxdAfterVatTy = sum03(cpxdCol) * (1 + cpxdVatRate) / 1e9;
  const gpmbTy = sum03(gpmbCol) / 1e9;
  const landTy = sum03(landCol) / 1e9;
  const htktAfterVatTy = sum03(htktCol) * (1 + htktVatRate) / 1e9;
  const reserveAfterVatTy = sum03(reserveCol) * (1 + reserveVatRate) / 1e9;
  const interestTy = sum04(interestCol) / 1e9;
  const investmentTotalTy = cpxdAfterVatTy + gpmbTy + landTy + htktAfterVatTy + reserveAfterVatTy + interestTy;

  sh00.getRange('C6:C11').setValues([
    [cpxdAfterVatTy],
    [gpmbTy],
    [landTy],
    [htktAfterVatTy],
    [reserveAfterVatTy],
    [interestTy]
  ]);
  sh00.getRange('C12').setFormula('=SUM(C6:C11)');
  sh00.getRange('C13').setFormula('=C12-C8');
  for (let r = 6; r <= 11; r++) sh00.getRange(r, 4).setFormula(`=IFERROR(C${r}/$C$12,0)`);
  sh00.getRange('D12').setValue(1);
  sh00.getRange('C6:C13').setNumberFormat('#,##0.0;[Red]-#,##0.0');
  sh00.getRange('D6:D12').setNumberFormat('0.0%');

  const productRevenue = FSZZZZZZZZ_groupByProduct_(
    sh02, h02.dataStartRow, productCol, customerCashCol
  ).filter(x => Math.abs(x.value) > tol);

  const sellingAfterVat = sum03(sellingCol) * (1 + sellingVatRate) / 1e9;
  const opAfterVat = sum03(opCol) * (1 + opVatRate) / 1e9;
  const maintAfterVat = sum03(maintCol) * (1 + maintVatRate) / 1e9;

  // Sau khi mục I được hiệu chỉnh, công thức gốc tại D31 (nếu có) đã tham chiếu đúng C12.
  const investmentFormula = sh00.getRange('D31').getFormula();
  const investmentValue = Number(sh00.getRange('D31').getValue()) || investmentTotalTy;

  const baseKpis = [
    FSZZZZZZZZ_captureBaseRow_(sh00, 33),
    FSZZZZZZZZ_captureBaseRow_(sh00, 34),
    FSZZZZZZZZ_captureBaseRow_(sh00, 35),
    FSZZZZZZZZ_captureBaseRow_(sh00, 36),
    FSZZZZZZZZ_captureBaseRow_(sh00, 37),
    FSZZZZZZZZ_captureBaseRow_(sh00, 38),
    FSZZZZZZZZ_captureBaseRow_(sh00, 39),
    FSZZZZZZZZ_captureBaseRow_(sh00, 40),
    FSZZZZZZZZ_captureBaseRow_(sh00, 41),
    FSZZZZZZZZ_captureBaseRow_(sh00, 42),
    FSZZZZZZZZ_captureBaseRow_(sh00, 43)
  ];

  const totalRevenueTy = productRevenue.reduce((s, x) => s + x.value, 0) / 1e9;
  const costItems = [
    { label: 'Chi phí bán hàng', value: sellingAfterVat },
    { label: 'Chi phí vận hành', value: opAfterVat },
    { label: 'Chi phí bảo trì', value: maintAfterVat }
  ].filter(x => Math.abs(x.value) > 0.0000005);

  const totalCostSourceTy = investmentTotalTy + sellingAfterVat + opAfterVat + maintAfterVat;
  const profitAfterTaxTy = sum04(profitAfterTaxCol) / 1e9;
  const citTy = sum04(citCol) / 1e9;
  const vatPayTy = sum04(vatPayCol) / 1e9;
  const scopeDifferenceTy = totalRevenueTy - totalCostSourceTy - profitAfterTaxTy - citTy - vatPayTy;

  const clearRows = Math.max(40, sh00.getMaxRows() - 22);
  sh00.getRange(23, 1, clearRows, 5).clearContent().clearFormat();

  let row = 23;
  sh00.getRange(row, 1, 1, 5).merge().setValue('III. ĐÁNH GIÁ HIỆU QUẢ ĐẦU TƯ');
  row++;
  sh00.getRange(row, 1, 1, 5).setValues([['TT', 'Nội dung', 'Đơn vị', 'Giá trị', 'Ghi chú']]);
  const headerRow = row;
  row++;

  const firstDataRow = row;
  sh00.getRange(row, 1, 1, 5).setValues([['1', 'Tổng doanh thu có VAT', 'tỷ đồng', totalRevenueTy, '']]);
  const revenueTotalRow = row;
  row++;

  productRevenue.forEach((item, i) => {
    sh00.getRange(row, 1, 1, 5).setValues([[
      '1.' + (i + 1), 'Phần ' + item.name, 'tỷ đồng', item.value / 1e9, ''
    ]]);
    row++;
  });

  const totalCostRow = row;
  sh00.getRange(row, 1, 1, 5).setValues([['2', 'Tổng chi phí có VAT', 'tỷ đồng', '', '']]);
  row++;

  const investmentRow = row;
  sh00.getRange(row, 1, 1, 5).setValues([['2.1', 'Tổng vốn đầu tư dự án', 'tỷ đồng', investmentTotalTy, '']]);
  row++;

  costItems.forEach((item, i) => {
    sh00.getRange(row, 1, 1, 5).setValues([[
      '2.' + (i + 2), item.label, 'tỷ đồng', item.value, ''
    ]]);
    row++;
  });
  const lastCostRow = row - 1;
  sh00.getRange(totalCostRow, 4).setFormula(`=SUM(D${investmentRow}:D${lastCostRow})`);

  baseKpis.forEach(item => {
    const targetRow = row;
    sh00.getRange(targetRow, 1, 1, 5).setValues([[item.tt, item.label, item.unit, '', item.note]]);
    if (item.formula) sh00.getRange(targetRow, 4).setFormula(item.formula);
    else sh00.getRange(targetRow, 4).setValue(item.value);
    item.targetRow = targetRow;
    row++;
  });

  sh00.getRange(row, 1, 1, 5).setValues([[
    '14', 'Chênh lệch đối chiếu phạm vi', 'tỷ đồng', scopeDifferenceTy, ''
  ]]);
  const checkRow = row;
  const lastRow = row;

  FSZZZZZZZZ_formatFlexibleSection_(sh00, headerRow, firstDataRow, lastRow, revenueTotalRow, totalCostRow, investmentRow, baseKpis, checkRow);
  SpreadsheetApp.flush();
}

function FSZZZZZZZZ_captureBaseRow_(sheet, row) {
  const values = sheet.getRange(row, 1, 1, 5).getValues()[0];
  return {
    tt: values[0], label: values[1], unit: values[2], value: values[3], note: values[4],
    formula: sheet.getRange(row, 4).getFormula()
  };
}

function FSZZZZZZZZ_groupByProduct_(sheet, startRow, productCol, valueCol) {
  if (sheet.getLastRow() < startRow) return [];
  const n = sheet.getLastRow() - startRow + 1;
  const products = sheet.getRange(startRow, productCol, n, 1).getDisplayValues();
  const values = sheet.getRange(startRow, valueCol, n, 1).getValues();
  const order = [];
  const sums = {};
  for (let i = 0; i < n; i++) {
    const name = String(products[i][0] || '').trim();
    if (!name) continue;
    const key = FSZZZZZZZZ_norm_(name);
    if (!(key in sums)) {
      sums[key] = { name, value: 0 };
      order.push(key);
    }
    sums[key].value += Number(values[i][0]) || 0;
  }
  return order.map(k => sums[k]);
}

function FSZZZZZZZZ_formatFlexibleSection_(sh, headerRow, firstRow, lastRow, revenueTotalRow, totalCostRow, investmentRow, kpis, checkRow) {
  sh.getRange(23, 1, lastRow - 22, 5)
    .setFontFamily('Times New Roman').setFontSize(11)
    .setVerticalAlignment('middle').setWrap(true)
    .setBorder(true, true, true, true, true, true);

  sh.getRange(23, 1, 1, 5).setBackground('#FFC000').setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange(headerRow, 1, 1, 5).setBackground('#A6A6A6').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange(firstRow, 1, lastRow - firstRow + 1, 1).setHorizontalAlignment('center');
  sh.getRange(firstRow, 2, lastRow - firstRow + 1, 1).setHorizontalAlignment('left');
  sh.getRange(firstRow, 3, lastRow - firstRow + 1, 1).setHorizontalAlignment('center');
  sh.getRange(firstRow, 4, lastRow - firstRow + 1, 1).setHorizontalAlignment('right');
  sh.getRange(firstRow, 5, lastRow - firstRow + 1, 1).setHorizontalAlignment('center');
  sh.getRange(firstRow, 4, lastRow - firstRow + 1, 1).setNumberFormat('#,##0.0;[Red]-#,##0.0');

  [revenueTotalRow, totalCostRow, investmentRow, checkRow].forEach(r => sh.getRange(r, 1, 1, 4).setFontWeight('bold'));

  kpis.forEach(item => {
    const r = item.targetRow;
    const key = FSZZZZZZZZ_norm_(item.label);
    sh.getRange(r, 1, 1, 5).setFontWeight('bold');
    if (key.indexOf('irr') >= 0) sh.getRange(r, 4).setNumberFormat('0.0%');
    else if (key.indexOf('thoi gian hoan von') >= 0) sh.getRange(r, 4).setNumberFormat('0.0');

    if (['4','5','6'].indexOf(String(item.tt)) >= 0) {
      sh.getRange(r, 1, 1, 5).setBackground('#FCE4D6').setFontColor('#FF0000');
    }
    if (['7','8','9'].indexOf(String(item.tt)) >= 0) {
      sh.getRange(r, 1, 1, 5).setBackground('#E2F0D9').setFontColor('#FF0000');
    }
  });

  for (let r = 23; r <= lastRow; r++) sh.setRowHeight(r, 24);
  sh.setRowHeight(23, 26);
  sh.setRowHeight(headerRow, 28);
}

function FSZZZZZZZZ_detectHeader_(sheet, requiredNames) {
  const maxRows = Math.min(6, sheet.getLastRow());
  const lastCol = sheet.getLastColumn();
  let best = null;
  for (let r = 1; r <= maxRows; r++) {
    const headers = sheet.getRange(r, 1, 1, lastCol).getDisplayValues()[0];
    const hits = requiredNames.filter(name => FSZZZZZZZZ_findHeader_(headers, [name]) > 0).length;
    if (!best || hits > best.hits) best = { headerRow: r, dataStartRow: r + 1, headers, hits };
  }
  if (!best || best.hits < 1) throw new Error('Không xác định được dòng tiêu đề tại ' + sheet.getName());
  return best;
}

function FSZZZZZZZZ_requireCol_(headers, names) {
  const col = FSZZZZZZZZ_findHeader_(headers, names);
  if (col < 1) throw new Error('Không tìm thấy cột: ' + names.join(' / '));
  return col;
}

function FSZZZZZZZZ_sumColumn_(sheet, col, startRow) {
  if (!col || col < 1 || sheet.getLastRow() < startRow) return 0;
  return sheet.getRange(startRow, col, sheet.getLastRow() - startRow + 1, 1)
    .getValues().reduce((sum, r) => sum + (Number(r[0]) || 0), 0);
}

function FSZZZZZZZZ_getCommonCostVatRate_(tech, names, fallback) {
  const values = tech.getRange(1, 1, tech.getLastRow(), Math.max(3, Math.min(tech.getLastColumn(), 8))).getValues();
  const targets = names.map(FSZZZZZZZZ_norm_);
  let inBlock = false;
  for (let r = 0; r < values.length; r++) {
    const raw = String(values[r][0] || '').trim();
    const first = FSZZZZZZZZ_norm_(raw);
    if (first === 'chi_phi_chung' || first === 'chi phi chung') { inBlock = true; continue; }
    if (!inBlock) continue;
    if (r > 0 && /^[A-Z0-9_]{4,}$/.test(raw) && first !== 'chi_phi_chung') break;
    if (targets.indexOf(first) >= 0) return FSZZZZZZZZ_rate_(values[r][2]);
  }
  return FSZZZZZZZZ_rate_(fallback);
}

function FSZZZZZZZZ_rate_(value) {
  let n = Number(value);
  if (!isFinite(n)) n = 0;
  if (Math.abs(n) > 1) n /= 100;
  return Math.max(0, n);
}

function FSZZZZZZZZ_findHeader_(headers, names) {
  const normalized = headers.map(FSZZZZZZZZ_norm_);
  for (const name of names) {
    const i = normalized.indexOf(FSZZZZZZZZ_norm_(name));
    if (i >= 0) return i + 1;
  }
  return -1;
}

function FSZZZZZZZZ_norm_(value) {
  return String(value || '').normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '').replace(/đ/g, 'd').replace(/Đ/g, 'D')
    .replace(/\s+/g, ' ').trim().toLowerCase();
}
