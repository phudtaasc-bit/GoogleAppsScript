/*************************************************
 * ZZZZZZZZ_SummaryRevenueCheckFix.js
 *
 * Nguyên tắc:
 * - Chỉ tiêu 1 trên Sheet 00 luôn lấy trực tiếp từ Sheet 02.
 * - Đẳng thức 1 = 2 + 3 + 12 + 13 chỉ là phép đối chiếu khác phạm vi,
 *   không dùng để ghi đè doanh thu hoặc ép kết quả bằng 0.
 * - Chi phí bán hàng, vận hành và bảo trì tại mục III đều trình bày sau VAT.
 * - VAT từng khoản chi được tính theo tỷ lệ cấu hình tại Sheet 01. Kỹ thuật;
 *   không phân bổ toàn bộ VAT đầu vào của dự án cho chi phí bán hàng.
 *************************************************/

const FSZZZZZZZZ_BASE_LAP_SHEET00_ = FS_lapSheet00;

FS_lapSheet00 = function() {
  FSZZZZZZZZ_BASE_LAP_SHEET00_();
  FSZZZZZZZZ_restoreSummaryRevenueAndCheck_();
};

/**
 * Chạy riêng khi chỉ cần sửa Sheet 00 hiện tại, không lập lại mô hình.
 */
function FSZZZZZZZZ_suaTongHopDoanhThuVaKiemTra() {
  FSZZZZZZZZ_restoreSummaryRevenueAndCheck_();
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert(
    'Đã cập nhật doanh thu, chi phí bán hàng/vận hành/bảo trì sau VAT và tiêu chí đối chiếu trên Sheet 00.'
  );
}

function FSZZZZZZZZ_restoreSummaryRevenueAndCheck_() {
  const ss = SpreadsheetApp.getActive();
  const sh00 = ss.getSheetByName('00. Tổng hợp');
  const tech = ss.getSheetByName('01. Kỹ thuật');
  const sh02 = ss.getSheetByName('02. Doanh thu');
  const sh03 = ss.getSheetByName('03. Chi phí & Vốn');
  if (!sh00 || !tech || !sh02 || !sh03) {
    throw new Error('Thiếu Sheet 00. Tổng hợp, 01. Kỹ thuật, 02. Doanh thu hoặc 03. Chi phí & Vốn.');
  }

  const headers02 = sh02
    .getRange(1, 1, 1, sh02.getLastColumn())
    .getDisplayValues()[0];
  const headers03 = sh03
    .getRange(1, 1, 1, sh03.getLastColumn())
    .getDisplayValues()[0];

  const cashCol = FSZZZZZZZZ_findHeader_(headers02, [
    'Dòng tiền huy động từ KH',
    'Dòng tiền huy động từ khách hàng'
  ]);
  if (cashCol < 1) {
    throw new Error('Không tìm thấy cột Dòng tiền huy động từ KH trên Sheet 02.');
  }

  const sellingCol = FSZZZZZZZZ_findHeader_(headers03, [
    'Chi phí bán hàng trước VAT'
  ]);
  const opCol = FSZZZZZZZZ_findHeader_(headers03, [
    'Chi phí vận hành thuê trước VAT',
    'Chi phí vận hành trước VAT'
  ]);
  const maintCol = FSZZZZZZZZ_findHeader_(headers03, [
    'Chi phí bảo trì trước VAT'
  ]);

  const sellingVatRate = FSZZZZZZZZ_getCommonCostVatRate_(tech, [
    'Chi phí bán hàng'
  ], 0);
  const opVatRate = FSZZZZZZZZ_getCommonCostVatRate_(tech, [
    'Chi phí vận hành',
    'Chi phí vận hành thuê'
  ], sellingVatRate);
  const maintVatRate = FSZZZZZZZZ_getCommonCostVatRate_(tech, [
    'Chi phí bảo trì',
    'Chi phí bảo hành, bảo trì'
  ], opVatRate);

  const cashLetter = FSZZZZZZZZ_colLetter_(cashCol);
  const s02 = `'02. Doanh thu'`;
  const s03 = `'03. Chi phí & Vốn'`;

  // Chỉ tiêu 1: luôn link trực tiếp từ Sheet chi tiết.
  sh00.getRange('D25').setFormula(
    `=SUM(${s02}!${cashLetter}2:${cashLetter})/1000000000`
  );

  // Chi phí bán hàng sau VAT: VAT đúng của chính chi phí bán hàng.
  // Không dùng vatAlloc('L') vì công thức cũ phân bổ cả VAT đầu vào dự án vào cột bán hàng.
  sh00.getRange('B32').setValue('Chi phí bán hàng (sau VAT)');
  if (sellingCol > 0) {
    const letter = FSZZZZZZZZ_colLetter_(sellingCol);
    sh00.getRange('D32').setFormula(
      `=SUM(${s03}!${letter}2:${letter})*(1+${sellingVatRate})/1000000000`
    );
  } else {
    sh00.getRange('D32').setValue(0);
  }

  // D44 chỉ phản ánh chênh lệch giữa các chỉ tiêu không đồng nhất phạm vi kinh tế.
  const checkRow = 44;
  sh00.getRange(checkRow, 1, 1, 5).clearContent();
  sh00.getRange(checkRow, 1).setValue('14');
  sh00.getRange(checkRow, 2).setValue('Chênh lệch đối chiếu phạm vi');
  sh00.getRange(checkRow, 3).setValue('tỷ đồng');
  sh00.getRange(checkRow, 4).setFormula('=D25-(D30+D33+D42+D43)');
  sh00.getRange(checkRow, 5).setValue(
    'Tham chiếu: Doanh thu có VAT - (Tổng chi phí có VAT + LNST + Thuế TNDN + VAT phải nộp). ' +
    'Các chỉ tiêu không cùng phạm vi kinh tế; kết quả không bắt buộc bằng 0 và không dùng để điều chỉnh số liệu.'
  );

  sh00.getRange(checkRow, 1, 1, 5)
    .setFontFamily('Times New Roman')
    .setFontSize(11)
    .setBorder(true, true, true, true, true, true)
    .setVerticalAlignment('middle');
  sh00.getRange(checkRow, 1, 1, 3).setFontWeight('bold');
  sh00.getRange(checkRow, 4)
    .setNumberFormat('#,##0.0;[Red]-#,##0.0')
    .setFontWeight('bold');
  sh00.getRange(checkRow, 5).setWrap(true);
  sh00.setRowHeight(checkRow, 48);

  // Bổ sung hai cấu phần chi phí sau bảng hiện hữu để không làm dịch chuyển D25:D44.
  const detailRows = [
    [45, '2.3', 'Chi phí vận hành (sau VAT)', opCol, opVatRate],
    [46, '2.4', 'Chi phí bảo trì (sau VAT)', maintCol, maintVatRate]
  ];

  detailRows.forEach(item => {
    const row = item[0];
    const tt = item[1];
    const label = item[2];
    const col = item[3];
    const vatRate = item[4];

    sh00.getRange(row, 1, 1, 5).clearContent();
    sh00.getRange(row, 1).setValue(tt);
    sh00.getRange(row, 2).setValue(label);
    sh00.getRange(row, 3).setValue('tỷ đồng');

    if (col > 0) {
      const letter = FSZZZZZZZZ_colLetter_(col);
      sh00.getRange(row, 4).setFormula(
        `=SUM(${s03}!${letter}2:${letter})*(1+${vatRate})/1000000000`
      );
    } else {
      sh00.getRange(row, 4).setValue(0);
    }

    sh00.getRange(row, 5).setValue(
      `Nguồn: Sheet 03; VAT theo cấu hình Sheet 01. Kỹ thuật (${vatRate * 100}%)`
    );
    sh00.getRange(row, 1, 1, 5)
      .setFontFamily('Times New Roman')
      .setFontSize(11)
      .setBorder(true, true, true, true, true, true)
      .setVerticalAlignment('middle');
    sh00.getRange(row, 4).setNumberFormat('#,##0.0;[Red]-#,##0.0');
  });

  // Tổng chi phí có VAT bao gồm vốn đầu tư, bán hàng, vận hành và bảo trì sau VAT.
  sh00.getRange('D30').setFormula('=D31+D32+D45+D46');
}

function FSZZZZZZZZ_getCommonCostVatRate_(tech, names, fallback) {
  const lastRow = tech.getLastRow();
  const lastCol = Math.max(3, Math.min(tech.getLastColumn(), 8));
  if (lastRow < 1) return FSZZZZZZZZ_rate_(fallback);

  const values = tech.getRange(1, 1, lastRow, lastCol).getValues();
  const targets = names.map(FSZZZZZZZZ_norm_);
  let inBlock = false;

  for (let r = 0; r < values.length; r++) {
    const first = FSZZZZZZZZ_norm_(values[r][0]);
    if (first === 'chi_phi_chung' || first === 'chi phi chung') {
      inBlock = true;
      continue;
    }
    if (!inBlock) continue;

    // Sang block kỹ thuật khác thì dừng.
    if (r > 0 && /^[A-Z0-9_]{4,}$/.test(String(values[r][0] || '').trim()) && first !== 'chi_phi_chung') {
      break;
    }

    if (targets.indexOf(first) >= 0) {
      // Cột C trong block CHI_PHI_CHUNG là tỷ lệ VAT.
      return FSZZZZZZZZ_rate_(values[r][2]);
    }
  }

  return FSZZZZZZZZ_rate_(fallback);
}

function FSZZZZZZZZ_rate_(value) {
  let n = Number(value);
  if (!isFinite(n)) n = 0;
  if (Math.abs(n) > 1) n = n / 100;
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
  return String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd')
    .replace(/Đ/g, 'D')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

function FSZZZZZZZZ_colLetter_(col) {
  let out = '';
  while (col > 0) {
    const rem = (col - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    col = Math.floor((col - 1) / 26);
  }
  return out;
}
