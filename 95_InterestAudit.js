/*************************************************
 * 95_InterestAudit.gs
 * Kiểm tra độc lập lãi vay, dư nợ và đồng bộ Sheet 03-04.
 * Chỉ đọc dữ liệu; không thay đổi mô hình.
 *************************************************/

function FS_AUDIT_LAI_VAY() {
  const ss = SpreadsheetApp.getActive();
  const tech = ss.getSheetByName('01. Kỹ thuật');
  const sh03 = ss.getSheetByName('03. Chi phí & Vốn');
  const sh04 = ss.getSheetByName('04. Dòng tiền & Lợi nhuận');

  if (!tech || !sh03 || !sh04) {
    throw new Error('Thiếu một trong các sheet 01. Kỹ thuật, 03. Chi phí & Vốn hoặc 04. Dòng tiền & Lợi nhuận.');
  }

  const laiSuatNam = FS95_rate_(FS04_getInfoValue_(tech, 'Lãi suất vay năm'));
  const laiSuatThang = Math.pow(1 + laiSuatNam, 1 / 12) - 1;
  const tol = 1000;
  const issues = [];
  const warnings = [];

  const rows04 = Math.max(0, sh04.getLastRow() - 2);
  const rows03 = Math.max(0, sh03.getLastRow() - 1);
  if (!rows04 || !rows03) throw new Error('Sheet 03 hoặc Sheet 04 chưa có dữ liệu.');

  const data04 = sh04.getRange(3, 1, rows04, 25).getValues();
  const data03 = sh03.getRange(2, 1, rows03, 29).getValues();
  const map03 = {};

  data03.forEach(r => {
    const t = Number(r[0]) || 0;
    if (!t) return;
    map03[t] = {
      giaiNgan: FS95_num_(r[23]), // X
      laiVay: FS95_num_(r[25]),  // Z
      traGoc: FS95_num_(r[26]),  // AA
      duNo: FS95_num_(r[28])     // AC
    };
  });

  let prevDebt = 0;
  let totalInterest = 0;

  data04.forEach((r, i) => {
    const rowNo = i + 3;
    const t = Number(r[0]) || 0;
    if (!t) return;

    const giaiNgan = FS95_num_(r[21]); // V
    const laiVay = FS95_num_(r[22]);  // W
    const traGoc = FS95_num_(r[23]);  // X
    const duNo = FS95_num_(r[24]);    // Y
    const expectedInterest = prevDebt * laiSuatThang;
    const expectedDebt = Math.max(0, prevDebt + giaiNgan - traGoc);

    if (Math.abs(laiVay - expectedInterest) > tol) {
      issues.push(`Sheet 04 dòng ${rowNo}: lãi vay chưa khớp dư nợ đầu kỳ × lãi suất tháng.`);
    }
    if (Math.abs(duNo - expectedDebt) > tol) {
      issues.push(`Sheet 04 dòng ${rowNo}: dư nợ cuối kỳ chưa khớp dư nợ đầu kỳ + giải ngân - trả gốc.`);
    }

    const s03 = map03[t];
    if (!s03) {
      issues.push(`Tháng ${t}: không tìm thấy dòng tương ứng tại Sheet 03.`);
    } else {
      if (Math.abs(s03.giaiNgan - giaiNgan) > tol) issues.push(`Tháng ${t}: giải ngân Sheet 03 chưa đồng bộ Sheet 04.`);
      if (Math.abs(s03.laiVay - laiVay) > tol) issues.push(`Tháng ${t}: lãi vay Sheet 03 chưa đồng bộ Sheet 04.`);
      if (Math.abs(s03.traGoc - traGoc) > tol) issues.push(`Tháng ${t}: trả gốc Sheet 03 chưa đồng bộ Sheet 04.`);
      if (Math.abs(s03.duNo - duNo) > tol) issues.push(`Tháng ${t}: dư nợ Sheet 03 chưa đồng bộ Sheet 04.`);
    }

    if (giaiNgan < -tol || laiVay < -tol || traGoc < -tol || duNo < -tol) {
      issues.push(`Sheet 04 dòng ${rowNo}: có giá trị tài trợ âm.`);
    }

    totalInterest += laiVay;
    prevDebt = duNo;
  });

  if (laiSuatNam <= 0 && totalInterest > tol) {
    issues.push('Lãi suất vay năm bằng 0 nhưng mô hình vẫn phát sinh chi phí lãi vay.');
  }
  if (laiSuatNam > 0 && prevDebt > tol) {
    warnings.push('Cuối mô hình vẫn còn dư nợ vay; cần xác nhận đây là giả định chủ động.');
  }

  const report = [
    `KIỂM TRA LÃI VAY: ${issues.length} lỗi, ${warnings.length} cảnh báo.`,
    `Lãi suất năm: ${(laiSuatNam * 100).toFixed(4)}%.`,
    `Tổng chi phí lãi vay: ${Math.round(totalInterest).toLocaleString('vi-VN')} đồng.`
  ];
  if (issues.length) report.push('\nLỖI:\n- ' + issues.join('\n- '));
  if (warnings.length) report.push('\nCẢNH BÁO:\n- ' + warnings.join('\n- '));
  if (!issues.length && !warnings.length) report.push('\nKhông phát hiện bất thường về lãi vay và dư nợ.');

  SpreadsheetApp.getUi().alert(report.join('\n'));
  return { issues, warnings, totalInterest, endingDebt: prevDebt };
}

function FS95_num_(value) {
  const n = Number(value);
  return isFinite(n) ? n : 0;
}

function FS95_rate_(value) {
  const n = FS95_num_(value);
  return n > 1 ? n / 100 : n;
}
