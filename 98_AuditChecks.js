/*************************************************
 * 98_AuditChecks.gs
 * Kiểm soát độc lập mô hình tài chính.
 * Chỉ đọc dữ liệu, không ghi đè sheet và không thay đổi giao diện.
 *************************************************/

function FS_AUDIT_RUN_ALL() {
  const ss = SpreadsheetApp.getActive();
  const issues = [];

  FS_AUDIT_checkTerminology_(ss, issues);
  FS_AUDIT_checkSheet03_(ss, issues);
  FS_AUDIT_checkSheet04_(ss, issues);
  FS_AUDIT_checkKpi_(ss, issues);

  const message = issues.length
    ? 'AUDIT phát hiện ' + issues.length + ' vấn đề:\n\n' + issues.slice(0, 30).join('\n') +
      (issues.length > 30 ? '\n\n... còn ' + (issues.length - 30) + ' vấn đề khác.' : '')
    : 'AUDIT: Không phát hiện sai lệch theo các kiểm soát hiện có.';

  SpreadsheetApp.getUi().alert(message);
  return issues;
}

function FS_AUDIT_checkTerminology_(ss, issues) {
  const tech = ss.getSheetByName('01. Kỹ thuật');
  if (!tech) {
    issues.push('[P0] Thiếu sheet 01. Kỹ thuật.');
    return;
  }

  const legacyPatterns = [
    'Chi phí XD/TB/khác',
    'Tiền SDĐ/thuê đất',
    'CP XD/TB trực tiếp trước VAT',
    'CPVH thuê'
  ].map(FS_AUDIT_norm_);

  const values = tech.getDataRange().getDisplayValues();
  values.forEach((row, r) => row.forEach((value, c) => {
    const key = FS_AUDIT_norm_(value);
    if (key && legacyPatterns.includes(key)) {
      issues.push('[P1] Thuật ngữ chưa chuẩn tại 01. Kỹ thuật!' + FS_AUDIT_a1_(r + 1, c + 1) + ': ' + value);
    }
  }));
}

function FS_AUDIT_checkSheet03_(ss, issues) {
  const sh = ss.getSheetByName('03. Chi phí & Vốn');
  if (!sh || sh.getLastRow() < 2) {
    issues.push('[P0] Sheet 03 chưa có dữ liệu để kiểm tra.');
    return;
  }

  const data = sh.getRange(2, 1, sh.getLastRow() - 1, Math.min(35, sh.getLastColumn())).getValues();
  const tol = 10;
  let prevDebt = 0;
  let prevVatCredit = 0;

  data.forEach((r, i) => {
    const row = i + 2;
    const vatOut = FS_AUDIT_num_(r[5]);
    const vatIn = FS_AUDIT_num_(r[14]);
    const vatCreditOpen = FS_AUDIT_num_(r[16]);
    const vatPayable = FS_AUDIT_num_(r[17]);
    const vatCreditClose = FS_AUDIT_num_(r[18]);
    const drawdown = FS_AUDIT_num_(r[23]);
    const principal = FS_AUDIT_num_(r[26]);
    const debtClose = FS_AUDIT_num_(r[28]);

    if (Math.abs(vatCreditOpen - prevVatCredit) > tol) {
      issues.push('[P0] Sheet 03 dòng ' + row + ': VAT khấu trừ đầu kỳ không khớp cuối kỳ trước.');
    }

    const isLastRow = i === data.length - 1;
    const expectedPayable = Math.max(0, vatOut - vatCreditOpen - vatIn);
    const expectedCredit = Math.max(0, vatCreditOpen + vatIn - vatOut);

    if (!isLastRow) {
      if (Math.abs(vatPayable - expectedPayable) > tol) {
        issues.push('[P0] Sheet 03 dòng ' + row + ': VAT phải nộp không khớp công thức khấu trừ.');
      }
      if (Math.abs(vatCreditClose - expectedCredit) > tol) {
        issues.push('[P0] Sheet 03 dòng ' + row + ': VAT còn khấu trừ cuối kỳ không khớp.');
      }
    } else if (expectedCredit > tol) {
      if (Math.abs(vatPayable + expectedCredit) > tol || Math.abs(vatCreditClose) > tol) {
        issues.push('[P0] Sheet 03 dòng cuối: số hoàn VAT không khớp VAT còn được khấu trừ.');
      }
    }

    const expectedDebt = Math.max(0, prevDebt + drawdown - principal);
    if (Math.abs(debtClose - expectedDebt) > tol) {
      issues.push('[P0] Sheet 03 dòng ' + row + ': Dư nợ cuối kỳ không khớp giải ngân và trả gốc.');
    }

    prevVatCredit = vatCreditClose;
    prevDebt = debtClose;
  });
}

function FS_AUDIT_checkSheet04_(ss, issues) {
  const sh = ss.getSheetByName('04. Dòng tiền & Lợi nhuận');
  if (!sh || sh.getLastRow() < 3) {
    issues.push('[P0] Sheet 04 chưa có dữ liệu để kiểm tra.');
    return;
  }

  const data = sh.getRange(3, 1, sh.getLastRow() - 2, Math.min(39, sh.getLastColumn())).getValues();
  const tol = 10;

  data.forEach((r, i) => {
    const row = i + 3;
    const fcff = FS_AUDIT_num_(r[35]);
    const fcfe = FS_AUDIT_num_(r[36]);
    const beforeFunding = FS_AUDIT_num_(r[15]);
    const distribution = FS_AUDIT_num_(r[17]);
    const equityNew = FS_AUDIT_num_(r[18]);
    const equityReturn = FS_AUDIT_num_(r[19]);

    if (Math.abs(fcff - beforeFunding) > tol) {
      issues.push('[P0] Sheet 04 dòng ' + row + ': FCFF không bằng dòng tiền trước tài trợ.');
    }
    if (Math.abs(fcfe - (distribution - equityNew - equityReturn)) > tol) {
      issues.push('[P0] Sheet 04 dòng ' + row + ': FCFE không khớp dòng tiền của chủ sở hữu.');
    }
  });
}

function FS_AUDIT_checkKpi_(ss, issues) {
  const sh = ss.getSheetByName('00. Tổng hợp');
  if (!sh) {
    issues.push('[P1] Thiếu sheet 00. Tổng hợp.');
    return;
  }

  ['D33', 'D34', 'D36', 'D37'].forEach(a1 => {
    const value = Number(sh.getRange(a1).getValue());
    if (!isFinite(value)) issues.push('[P0] 00. Tổng hợp!' + a1 + ': KPI không phải số hữu hạn.');
  });
}

function FS_AUDIT_num_(v) {
  const n = Number(v);
  return isFinite(n) ? n : 0;
}

function FS_AUDIT_norm_(v) {
  return String(v || '').toLowerCase().normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '').replace(/đ/g, 'd')
    .replace(/\s+/g, ' ').trim();
}

function FS_AUDIT_a1_(row, col) {
  let s = '';
  while (col > 0) {
    const rem = (col - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    col = Math.floor((col - 1) / 26);
  }
  return s + row;
}
