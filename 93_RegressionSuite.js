/*************************************************
 * 93_RegressionSuite.gs
 * Kiểm thử hồi quy không thay đổi dữ liệu cho mô hình FS.
 *************************************************/

function FS93_runRegressionSuite_(options) {
  const cfg = Object.assign({ showAlert: false, throwOnError: true }, options || {});
  const ss = SpreadsheetApp.getActive();
  const sh00 = ss.getSheetByName('00. Tổng hợp');
  const sh02 = ss.getSheetByName('02. Doanh thu');
  const sh03 = ss.getSheetByName('03. Chi phí & Vốn');
  const sh04 = ss.getSheetByName('04. Dòng tiền & Lợi nhuận');
  const issues = [];
  const warnings = [];
  const tol = 1000;

  [
    ['00. Tổng hợp', sh00],
    ['02. Doanh thu', sh02],
    ['03. Chi phí & Vốn', sh03],
    ['04. Dòng tiền & Lợi nhuận', sh04]
  ].forEach(x => { if (!x[1]) issues.push('Thiếu sheet "' + x[0] + '".'); });

  if (!issues.length) {
    try { FS97_assertLandCostConfig_(); }
    catch (err) { issues.push('Tiền đất: ' + err.message); }

    try { FS94_assertCITConsistency_(); }
    catch (err) { issues.push('Thuế TNDN: ' + err.message); }

    FS93_checkSheetStructures_(sh00, sh02, sh03, sh04, issues);
    FS93_checkCashFlowRows_(sh04, issues, warnings, tol);
    FS93_checkSheetLinks_(sh02, sh03, sh04, issues, tol);
    FS93_checkValuationFormulas_(sh00, issues, warnings);
  }

  const result = {
    ok: issues.length === 0,
    issues: issues,
    warnings: warnings,
    checkedAt: new Date()
  };

  const message = [
    'KIỂM THỬ HỒI QUY FS: ' + (result.ok ? 'ĐẠT' : 'KHÔNG ĐẠT'),
    'Lỗi: ' + issues.length + '; Cảnh báo: ' + warnings.length,
    issues.length ? '\nLỖI:\n- ' + issues.slice(0, 30).join('\n- ') : '',
    warnings.length ? '\nCẢNH BÁO:\n- ' + warnings.slice(0, 30).join('\n- ') : ''
  ].filter(Boolean).join('\n');

  if (cfg.showAlert) SpreadsheetApp.getUi().alert(message);
  if (!result.ok && cfg.throwOnError) throw new Error(message);
  return result;
}

function FS93_RunRegressionSuite() {
  return FS93_runRegressionSuite_({ showAlert: true, throwOnError: false });
}

function FS93_checkSheetStructures_(sh00, sh02, sh03, sh04, issues) {
  if (sh02.getLastColumn() < 32) issues.push('Sheet 02 thiếu cấu trúc 32 cột A:AF.');
  if (sh03.getLastColumn() < 35) issues.push('Sheet 03 thiếu cấu trúc 35 cột A:AI.');
  if (sh04.getLastColumn() < 39) issues.push('Sheet 04 thiếu cấu trúc 39 cột A:AM.');
  if (sh00.getLastRow() < 42) issues.push('Sheet 00 thiếu vùng chỉ tiêu đến dòng 42.');
}

function FS93_checkCashFlowRows_(sh04, issues, warnings, tol) {
  const n = Math.max(0, sh04.getLastRow() - 2);
  if (!n) {
    issues.push('Sheet 04 chưa có dữ liệu dòng tiền.');
    return;
  }

  const rows = sh04.getRange(3, 1, n, 39).getValues();
  let prevDebt = 0;
  let prevCash = 0;

  rows.forEach((r, i) => {
    const rowNo = i + 3;
    const month = FS93_num_(r[0]);
    if (!month) return;

    const revenueCash = FS93_num_(r[6]);
    const costAfterVat = FS93_num_(r[8]);
    const vatPayable = FS93_num_(r[10]);
    const cit = FS93_num_(r[11]);
    const beforeFinancing = FS93_num_(r[15]);
    const need = FS93_num_(r[16]);
    const equityDistribution = FS93_num_(r[17]);
    const equityNew = FS93_num_(r[18]);
    const equityReturn = FS93_num_(r[19]);
    const loanDraw = FS93_num_(r[21]);
    const interest = FS93_num_(r[22]);
    const principal = FS93_num_(r[23]);
    const debtEnd = FS93_num_(r[24]);
    const cashEnd = FS93_num_(r[25]);
    const cashAvailable = FS93_num_(r[26]);
    const fcff = FS93_num_(r[35]);
    const fcfe = FS93_num_(r[36]);
    const isFinal = i === rows.length - 1;

    const expectedBeforeFinancing = revenueCash - costAfterVat - vatPayable - cit;
    if (Math.abs(beforeFinancing - expectedBeforeFinancing) > tol) {
      issues.push('Sheet 04 dòng ' + rowNo + ': Dòng tiền trước tài trợ không khớp.');
    }
    if (Math.abs(fcff - beforeFinancing) > tol) {
      issues.push('Sheet 04 dòng ' + rowNo + ': FCFF không khớp dòng tiền trước tài trợ.');
    }

    const cashBeforeFunding = prevCash + beforeFinancing;
    const expectedNeed = Math.max(0, -cashBeforeFunding);
    if (Math.abs(need - expectedNeed) > tol) {
      issues.push('Sheet 04 dòng ' + rowNo + ': Nhu cầu vốn chưa dùng đúng tiền đầu kỳ.');
    }
    if (Math.abs(need - equityNew - loanDraw) > tol) {
      issues.push('Sheet 04 dòng ' + rowNo + ': Nhu cầu vốn không khớp CSH góp mới + giải ngân vay.');
    }
    if (Math.abs(equityReturn) > tol) {
      issues.push('Sheet 04 dòng ' + rowNo + ': Không được phát sinh CSH nộp lại từ tiền đã phân phối.');
    }

    const expectedPrincipal = Math.min(prevDebt + loanDraw + interest, Math.max(0, cashBeforeFunding));
    if (Math.abs(principal - expectedPrincipal) > tol) {
      issues.push('Sheet 04 dòng ' + rowNo + ': Trả gốc không khớp tiền dư và dư nợ sau vốn hóa lãi.');
    }

    const expectedDebt = Math.max(0, prevDebt + loanDraw + interest - principal);
    if (Math.abs(debtEnd - expectedDebt) > tol) {
      issues.push('Sheet 04 dòng ' + rowNo + ': Dư nợ cuối kỳ không khớp dư nợ đầu kỳ + giải ngân + lãi vay - trả gốc.');
    }

    const expectedAvailable = Math.max(0, cashBeforeFunding - principal);
    if (Math.abs(cashAvailable - expectedAvailable) > tol) {
      issues.push('Sheet 04 dòng ' + rowNo + ': Tiền khả dụng sau trả nợ không khớp.');
    }

    const expectedDistribution = isFinal ? cashAvailable : 0;
    if (Math.abs(equityDistribution - expectedDistribution) > tol) {
      issues.push('Sheet 04 dòng ' + rowNo + ': Phân phối CSH chỉ được phát sinh tại kỳ cuối.');
    }

    const expectedCashEnd = Math.max(0, cashAvailable - equityDistribution);
    if (Math.abs(cashEnd - expectedCashEnd) > tol) {
      issues.push('Sheet 04 dòng ' + rowNo + ': Tiền cuối kỳ không khớp.');
    }

    const expectedFcfe = cashAvailable - prevCash - equityNew;
    if (Math.abs(fcfe - expectedFcfe) > tol) {
      issues.push('Sheet 04 dòng ' + rowNo + ': FCFE không khớp biến động tiền khả dụng sau trả nợ trừ vốn CSH góp mới.');
    }

    if (need > tol && principal > tol) {
      issues.push('Sheet 04 dòng ' + rowNo + ': Vừa huy động vốn vừa trả gốc trong cùng kỳ.');
    }
    if (cit < -tol) issues.push('Sheet 04 dòng ' + rowNo + ': Thuế TNDN âm.');
    if (interest < -tol || loanDraw < -tol || principal < -tol || debtEnd < -tol) {
      issues.push('Sheet 04 dòng ' + rowNo + ': Chỉ tiêu vay có giá trị âm.');
    }
    if (cashEnd < -tol || cashAvailable < -tol) {
      issues.push('Sheet 04 dòng ' + rowNo + ': Chỉ tiêu tiền có giá trị âm.');
    }
    if (vatPayable < -tol && !isFinal) {
      issues.push('Sheet 04 dòng ' + rowNo + ': Hoàn VAT xuất hiện trước kỳ cuối.');
    }

    prevDebt = debtEnd;
    prevCash = cashEnd;
  });

  if (prevDebt > tol) warnings.push('Cuối mô hình vẫn còn dư nợ vay: ' + Math.round(prevDebt).toLocaleString('vi-VN') + ' đồng.');
  if (prevCash > tol) issues.push('Cuối mô hình vẫn còn tiền chưa phân phối cho CSH.');
}

function FS93_checkSheetLinks_(sh02, sh03, sh04, issues, tol) {
  const months = Math.max(sh03.getLastRow() - 1, sh04.getLastRow() - 2, 0);
  if (!months) return;

  const rows02 = sh02.getLastRow() > 1
    ? sh02.getRange(2, 1, sh02.getLastRow() - 1, 32).getValues()
    : [];
  const citByMonth = {};
  rows02.forEach(r => {
    const m = FS93_num_(r[0]);
    if (!m) return;
    citByMonth[m] = (citByMonth[m] || 0) + FS93_num_(r[31]);
  });

  const rows03 = sh03.getRange(2, 1, Math.min(months, Math.max(0, sh03.getLastRow() - 1)), 35).getValues();
  const rows04 = sh04.getRange(3, 1, Math.min(months, Math.max(0, sh04.getLastRow() - 2)), 39).getValues();
  const n = Math.min(rows03.length, rows04.length);

  for (let i = 0; i < n; i++) {
    const month = i + 1;
    const cit02 = citByMonth[month] || 0;
    const cit03 = FS93_num_(rows03[i][6]);
    const cit04 = FS93_num_(rows04[i][11]);
    if (Math.abs(cit02 - cit03) > tol || Math.abs(cit02 - cit04) > tol) {
      issues.push('Tháng ' + month + ': Thuế TNDN Sheet 02–03–04 không khớp.');
    }

    [23, 25, 26, 28].forEach((col03, j) => {
      const col04 = [21, 22, 23, 24][j];
      if (Math.abs(FS93_num_(rows03[i][col03]) - FS93_num_(rows04[i][col04])) > tol) {
        issues.push('Tháng ' + month + ': Dữ liệu tài trợ Sheet 03–04 không khớp tại nhóm cột ' + (j + 1) + '.');
      }
    });
  }
}

function FS93_checkValuationFormulas_(sh00, issues, warnings) {
  const npvProject = String(sh00.getRange('D33').getFormula() || '');
  const irrProject = String(sh00.getRange('D34').getFormula() || '');
  const npvEquity = String(sh00.getRange('D36').getFormula() || '');
  const irrEquity = String(sh00.getRange('D37').getFormula() || '');

  if (!/\$E\$21/.test(npvProject)) issues.push('NPV dự án chưa dùng WACC tại Sheet 00!E21.');
  if (!/AK3:AK/i.test(npvEquity)) issues.push('NPV vốn CSH chưa dùng dòng FCFE cột AK.');
  if (!/AJ3:AJ/i.test(npvProject)) issues.push('NPV dự án chưa dùng dòng FCFF cột AJ.');
  if (!/IRR\(/i.test(irrProject) || !/\^12/.test(irrProject)) warnings.push('IRR dự án chưa thể hiện quy đổi tháng sang năm hiệu dụng.');
  if (!/IRR\(/i.test(irrEquity) || !/\^12/.test(irrEquity)) warnings.push('IRR vốn CSH chưa thể hiện quy đổi tháng sang năm hiệu dụng.');
}

function FS93_num_(v) {
  const n = Number(v);
  return isFinite(n) ? n : 0;
}
