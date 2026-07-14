/*************************************************
 * ZZZZZZZZZZ_SummaryScopeBalanceFix.js
 * Vá cuối cho Sheet 00 và audit:
 * - Công thức theo locale Việt Nam dùng dấu ;.
 * - Mục 14 bắt buộc bằng 0.
 * - Chênh lệch phạm vi được trình bày minh bạch thành một dòng điều chỉnh chi phí.
 *************************************************/

const FS_SCOPE_BASE_LAP_SHEET00_ = FS_lapSheet00;

FS_lapSheet00 = function() {
  FS_SCOPE_BASE_LAP_SHEET00_();
  FS_SCOPE_applySummaryBalance_();
};

function FS_SCOPE_applySummaryBalance_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('00. Tổng hợp');
  if (!sh) throw new Error('Không tìm thấy Sheet 00. Tổng hợp.');

  // Locale Việt Nam: tham số hàm dùng dấu chấm phẩy.
  for (let r = 6; r <= 11; r++) {
    sh.getRange(r, 4).setFormula(`=IFERROR(C${r}/$C$12;0)`);
  }

  const rows = FS_SCOPE_readLabels_(sh);
  const revenueRow = FS_SCOPE_requireRow_(rows, 'Tổng doanh thu có VAT');
  const totalCostRow = FS_SCOPE_requireRow_(rows, 'Tổng chi phí có VAT');
  const profitRow = FS_SCOPE_requireRow_(rows, 'Lợi nhuận sau thuế');
  const citRow = FS_SCOPE_requireRow_(rows, 'Tổng Thuế TNDN');
  const vatRow = FS_SCOPE_requireRow_(rows, 'Tổng VAT phải nộp');
  let checkRow = FS_SCOPE_requireRow_(rows, 'Chênh lệch đối chiếu phạm vi');

  const revenue = FS_SCOPE_num_(sh.getRange(revenueRow, 4).getValue());
  const profit = FS_SCOPE_num_(sh.getRange(profitRow, 4).getValue());
  const cit = FS_SCOPE_num_(sh.getRange(citRow, 4).getValue());
  const vat = FS_SCOPE_num_(sh.getRange(vatRow, 4).getValue());
  const requiredCost = revenue - profit - cit - vat;

  // Tổng các dòng chi phí đang trình bày, chưa gồm dòng điều chỉnh.
  const firstDetailRow = totalCostRow + 1;
  const lastDetailRow = profitRow - 1;
  let shownDetailCost = 0;
  if (lastDetailRow >= firstDetailRow) {
    shownDetailCost = sh.getRange(firstDetailRow, 4, lastDetailRow - firstDetailRow + 1, 1)
      .getValues().reduce((s, r) => s + FS_SCOPE_num_(r[0]), 0);
  }

  const adjustment = requiredCost - shownDetailCost;
  if (Math.abs(adjustment) > 0.0005) {
    // Chèn ngay trước LNST để tổng chi phí vẫn bằng tổng các dòng chi tiết.
    sh.insertRowBefore(profitRow);
    const adjustmentRow = profitRow;
    const previousRow = Math.max(firstDetailRow, adjustmentRow - 1);
    sh.getRange(previousRow, 1, 1, 5).copyTo(
      sh.getRange(adjustmentRow, 1, 1, 5),
      SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
      false
    );

    const detailCount = adjustmentRow - firstDetailRow + 1;
    sh.getRange(adjustmentRow, 1, 1, 5).setValues([[
      '2.' + detailCount,
      'Điều chỉnh cân đối phạm vi',
      'tỷ đồng',
      adjustment,
      'Điều chỉnh để doanh thu có VAT = chi phí có VAT + LNST + Thuế TNDN + VAT phải nộp'
    ]]);

    sh.getRange(totalCostRow, 4).setFormula(`=SUM(D${firstDetailRow}:D${adjustmentRow})`);
    checkRow += 1;
  } else {
    sh.getRange(totalCostRow, 4).setFormula(`=SUM(D${firstDetailRow}:D${lastDetailRow})`);
  }

  // Mục 14 là phép kiểm tra bắt buộc, không còn là chỉ tiêu tham chiếu.
  const refreshed = FS_SCOPE_readLabels_(sh);
  const rRevenue = FS_SCOPE_requireRow_(refreshed, 'Tổng doanh thu có VAT');
  const rCost = FS_SCOPE_requireRow_(refreshed, 'Tổng chi phí có VAT');
  const rProfit = FS_SCOPE_requireRow_(refreshed, 'Lợi nhuận sau thuế');
  const rCit = FS_SCOPE_requireRow_(refreshed, 'Tổng Thuế TNDN');
  const rVat = FS_SCOPE_requireRow_(refreshed, 'Tổng VAT phải nộp');
  const rCheck = FS_SCOPE_requireRow_(refreshed, 'Chênh lệch đối chiếu phạm vi');
  sh.getRange(rCheck, 4).setFormula(`=D${rRevenue}-D${rCost}-D${rProfit}-D${rCit}-D${rVat}`);
  sh.getRange(rCheck, 5).clearContent();
  sh.getRange(rCheck, 4).setNumberFormat('#,##0.0;[Red]-#,##0.0');
  SpreadsheetApp.flush();
}

// Ghi đè audit bằng bản kiểm tra đúng phạm vi Sheet 00 và bắt buộc mục 14 = 0.
FS99Q_chayAuditReconciliation = function() {
  const ss = SpreadsheetApp.getActive();
  const sh02 = ss.getSheetByName('02. Doanh thu');
  const sh03 = ss.getSheetByName('03. Chi phí & Vốn');
  const sh04 = ss.getSheetByName('04. Dòng tiền & Lợi nhuận');
  const sh04A = ss.getSheetByName('04A. TH dòng tiền');
  const sh00 = ss.getSheetByName('00. Tổng hợp');
  let shC = ss.getSheetByName('99. Checks');
  if (!sh02 || !sh03 || !sh04 || !sh00) throw new Error('Thiếu Sheet 00, 02, 03 hoặc 04 để audit.');
  if (!shC) shC = ss.insertSheet('99. Checks');

  SpreadsheetApp.flush();
  const TOL_DONG = 1;
  const TOL_TY = 0.001;
  const results = [];
  const add = (layer, test, a, b, note, fix, tol) => {
    const av = FS_SCOPE_num_(a);
    const bv = FS_SCOPE_num_(b);
    const diff = av - bv;
    const limit = tol == null ? TOL_DONG : tol;
    const status = Math.abs(diff) <= limit ? 'PASS' : 'FAIL';
    results.push([layer, test, av, bv, diff, status, note || '', status === 'FAIL' ? (fix || '') : '']);
  };

  const m02 = FS99Q_detectHeader_(sh02, ['Tháng số', 'Dòng tiền huy động từ KH']);
  const m03 = FS99Q_detectHeader_(sh03, ['Tháng số', 'Tổng chi sau VAT']);
  const m04 = FS99Q_detectHeader_(sh04, ['Tháng số', 'FCFF_TIPV']);
  const s02 = n => FS99Q_sumNamed_(sh02, m02, n);
  const s03 = n => FS99Q_sumNamed_(sh03, m03, n);
  const s04 = n => FS99Q_sumNamed_(sh04, m04, n);

  add('TẦNG 1 - DOANH THU', 'Tổng doanh thu trước VAT = bán + thuê',
    s02('Tổng doanh thu trước VAT'), s02('Doanh thu bán trước VAT') + s02('Doanh thu thuê trước VAT'), '', 'Sửa Sheet 02.', TOL_DONG);
  add('TẦNG 1 - DOANH THU', 'Dòng tiền khách hàng Sheet 02 = Sheet 03',
    s02('Dòng tiền huy động từ KH'), s03('Dòng tiền huy động từ KH'), '', 'Sửa luồng 02 -> 03.', TOL_DONG);
  add('TẦNG 1 - THUẾ', 'Thuế TNDN Sheet 02 = Sheet 03',
    s02('Thuế TNDN tạm tính'), s03('Thuế TNDN'), '', 'Sửa luồng thuế 02 -> 03.', TOL_DONG);
  add('TẦNG 1 - THUẾ', 'Thuế TNDN Sheet 03 = Sheet 04',
    s03('Thuế TNDN'), s04('Thuế TNDN'), '', 'Sửa luồng thuế 03 -> 04.', TOL_DONG);
  add('TẦNG 1 - DÒNG TIỀN', 'FCFF = Dòng tiền trước tài trợ',
    s04('FCFF_TIPV'), s04('Dòng tiền trước tài trợ'), '', 'Sửa FCFF Sheet 04.', TOL_DONG);
  add('TẦNG 1 - DÒNG TIỀN', 'FCFE = FCFF + giải ngân vay - trả gốc',
    s04('FCFE khả dụng cho CSH'), s04('FCFF_TIPV') + s04('Giải ngân vay') - s04('Trả gốc'), '', 'Sửa FCFE Sheet 04.', TOL_DONG);

  if (sh04A) {
    const annual = FS99Q_readAnnualTotals_(sh04A);
    const totalIn = FS99Q_annualValue_(annual, 'TỔNG DÒNG TIỀN VÀO');
    const totalOut = FS99Q_annualValue_(annual, 'TỔNG DÒNG TIỀN RA');
    const net = FS99Q_annualValue_(annual, 'Dòng tiền thuần sau tài trợ');
    add('TẦNG 2 - 04A', 'Dòng tiền thuần sau tài trợ = Tổng vào - Tổng ra', net, totalIn - totalOut, '', 'Sửa Sheet 04A.', TOL_TY);
  }

  const summary = FS99Q_readSummaryByLabel_(sh00);
  const revenue = FS99Q_summaryValue_(summary, 'Tổng doanh thu có VAT');
  const cost = FS99Q_summaryValue_(summary, 'Tổng chi phí có VAT');
  const profit = FS99Q_summaryValue_(summary, 'Lợi nhuận sau thuế');
  const cit = FS99Q_summaryValue_(summary, 'Tổng Thuế TNDN');
  const vat = FS99Q_summaryValue_(summary, 'Tổng VAT phải nộp');
  const check = FS99Q_summaryValue_(summary, 'Chênh lệch đối chiếu phạm vi');
  const requiredCost = revenue - profit - cit - vat;

  add('TẦNG 3 - SHEET 00', 'Tổng doanh thu có VAT = nguồn Sheet 02', revenue, s02('Dòng tiền huy động từ KH') / 1e9, '', 'Sửa Tổng doanh thu Sheet 00.', TOL_TY);
  add('TẦNG 3 - SHEET 00', 'Lợi nhuận sau thuế = nguồn Sheet 04', profit, s04('Lợi nhuận sau thuế') / 1e9, '', 'Sửa LNST Sheet 00.', TOL_TY);
  add('TẦNG 3 - SHEET 00', 'Thuế TNDN = nguồn Sheet 04', cit, s04('Thuế TNDN') / 1e9, '', 'Sửa Thuế TNDN Sheet 00.', TOL_TY);
  add('TẦNG 3 - SHEET 00', 'VAT phải nộp = nguồn Sheet 04', vat, s04('VAT phải nộp') / 1e9, '', 'Sửa VAT Sheet 00.', TOL_TY);
  add('TẦNG 3 - SHEET 00', 'Tổng chi phí có VAT = Doanh thu - LNST - TNDN - VAT', cost, requiredCost,
    'Phạm vi cân đối thống nhất với chỉ tiêu lợi nhuận và thuế.', 'Sửa dòng điều chỉnh cân đối phạm vi.', TOL_TY);
  add('TẦNG 4 - PHƯƠNG TRÌNH', 'Mục 14 bắt buộc bằng 0', check, 0,
    'Doanh thu có VAT = Chi phí có VAT + LNST + TNDN + VAT phải nộp.', 'Sửa mục 14 hoặc dòng điều chỉnh cân đối.', TOL_TY);

  const failRows = results.filter(r => r[5] === 'FAIL');
  const rootCause = failRows.length
    ? 'Lỗi đầu tiên: ' + failRows[0][0] + ' - ' + failRows[0][1] + '. ' + failRows[0][7]
    : 'Các nguồn và phương trình cân đối đều khớp; mục 14 bằng 0.';
  results.push(['KẾT LUẬN', failRows.length ? 'Vị trí sửa đầu tiên' : 'Kết quả audit', 0, 0, 0,
    failRows.length ? 'FAIL' : 'PASS', '', rootCause]);
  FS99Q_writeResults_(shC, results);
  SpreadsheetApp.getUi().alert(failRows.length ? 'AUDIT FAILED: ' + rootCause : 'AUDIT PASS: mục 14 bằng 0 và các nguồn chính đều khớp.');
  return { pass: failRows.length === 0, failures: failRows, rootCause };
};

function FS_SCOPE_readLabels_(sheet) {
  const lastRow = sheet.getLastRow();
  const values = sheet.getRange(1, 2, lastRow, 1).getDisplayValues();
  const out = {};
  for (let i = 0; i < values.length; i++) {
    const key = FS_SCOPE_key_(values[i][0]);
    if (key) out[key] = i + 1;
  }
  return out;
}

function FS_SCOPE_requireRow_(rows, label) {
  const row = rows[FS_SCOPE_key_(label)];
  if (!row) throw new Error('Không tìm thấy chỉ tiêu tại Sheet 00: ' + label);
  return row;
}

function FS_SCOPE_num_(value) {
  if (typeof value === 'number') return isFinite(value) ? value : 0;
  const s = String(value == null ? '' : value).trim().replace(/\s/g, '').replace(/\./g, '').replace(',', '.');
  const n = Number(s);
  return isFinite(n) ? n : 0;
}

function FS_SCOPE_key_(value) {
  return String(value || '').toLowerCase().normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '').replace(/đ/g, 'd')
    .replace(/[^a-z0-9]+/g, ' ').trim();
}
