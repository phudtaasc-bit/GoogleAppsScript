/*************************************************
 * 96_ValuationAudit.js
 * Bộ kiểm tra chỉ đọc cho FCFF, FCFE, NPV và IRR.
 * Không sửa dữ liệu, không thay đổi giao diện mô hình.
 *************************************************/

function FS96_auditDinhGia() {
  const ss = SpreadsheetApp.getActive();
  const sh04 = ss.getSheetByName('04. Dòng tiền & Lợi nhuận');
  const tech = ss.getSheetByName('01. Kỹ thuật');
  if (!sh04) throw new Error('Không tìm thấy sheet "04. Dòng tiền & Lợi nhuận".');
  if (!tech) throw new Error('Không tìm thấy sheet "01. Kỹ thuật".');

  const issues = [];
  const warnings = [];
  const tol = 10;
  const lastRow = sh04.getLastRow();
  if (lastRow < 3) throw new Error('Sheet 04 chưa có dữ liệu.');

  const data = sh04.getRange(3, 1, lastRow - 2, 39).getValues();
  let coAmFCFF = false, coDuongFCFF = false;
  let coAmFCFE = false, coDuongFCFE = false;
  let prevCash = 0;

  data.forEach((r, i) => {
    const rowNo = i + 3;
    const p = FS96_num_(r[15]);        // P - Dòng tiền trước tài trợ / FCFF
    const sEquity = FS96_num_(r[18]);  // S - CSH góp mới
    const vDraw = FS96_num_(r[21]);    // V - Giải ngân vay
    const xPrincipal = FS96_num_(r[23]); // X - Trả gốc
    const zCashEnd = FS96_num_(r[25]); // Z - Tiền cuối kỳ sau phân phối
    const aaAvailable = FS96_num_(r[26]); // AA - Tiền khả dụng sau trả nợ
    const ajFCFF = FS96_num_(r[35]);   // AJ
    const akFCFE = FS96_num_(r[36]);   // AK

    if (Math.abs(ajFCFF - p) > tol) {
      issues.push(`Dòng ${rowNo}: FCFF không bằng Dòng tiền trước tài trợ.`);
    }

    // Lãi vay được vốn hóa nên không phải dòng tiền chi trong kỳ.
    const fcfeExpected = ajFCFF + vDraw - xPrincipal;
    if (Math.abs(akFCFE - fcfeExpected) > tol) {
      issues.push(`Dòng ${rowNo}: FCFE không bằng FCFF + Giải ngân vay - Trả gốc.`);
    }

    // Cầu nối độc lập để phát hiện ghi nhận trùng tiền giữ lại.
    const fcfeCashBridge = aaAvailable - prevCash - sEquity;
    if (Math.abs(akFCFE - fcfeCashBridge) > tol) {
      issues.push(`Dòng ${rowNo}: FCFE không khớp biến động tiền khả dụng trừ vốn CSH góp mới.`);
    }

    if (ajFCFF < -tol) coAmFCFF = true;
    if (ajFCFF > tol) coDuongFCFF = true;
    if (akFCFE < -tol) coAmFCFE = true;
    if (akFCFE > tol) coDuongFCFE = true;

    prevCash = zCashEnd;
  });

  if (!(coAmFCFF && coDuongFCFF)) {
    warnings.push('FCFF không có cả dòng âm và dòng dương; IRR dự án có thể không xác định hoặc không có ý nghĩa.');
  }
  if (!(coAmFCFE && coDuongFCFE)) {
    warnings.push('FCFE không có cả dòng âm và dòng dương; IRR vốn CSH có thể không xác định hoặc không có ý nghĩa.');
  }

  const rateProject = FS96_findInfo_(tech, [
    'Tỷ suất chiết khấu dự án',
    'WACC',
    'Tỷ suất chiết khấu'
  ]);
  const rateEquity = FS96_findInfo_(tech, [
    'Chi phí vốn chủ sở hữu',
    'Tỷ suất chiết khấu vốn CSH',
    'Tỷ suất chiết khấu FCFE',
    'Tỷ suất chiết khấu'
  ]);

  if (!rateProject.found) {
    warnings.push('Chưa tìm thấy suất chiết khấu dự án/WACC để tính NPV của FCFF.');
  }
  if (!rateEquity.found) {
    warnings.push('Chưa tìm thấy chi phí vốn chủ sở hữu để tính NPV của FCFE.');
  }
  if (rateProject.found && rateEquity.found && Math.abs(rateProject.value - rateEquity.value) < 1e-12) {
    warnings.push('Suất chiết khấu FCFF và FCFE đang bằng nhau. Chỉ phù hợp khi đây là chủ ý của mô hình.');
  }

  const result = [];
  if (issues.length) result.push('LỖI:\n- ' + issues.join('\n- '));
  if (warnings.length) result.push('CẢNH BÁO:\n- ' + warnings.join('\n- '));
  if (!result.length) result.push('Không phát hiện sai lệch FCFF/FCFE hoặc cảnh báo NPV/IRR.');

  SpreadsheetApp.getUi().alert('Audit FCFF / FCFE / NPV / IRR', result.join('\n\n'), SpreadsheetApp.getUi().ButtonSet.OK);
  return { issues, warnings };
}

function FS96_findInfo_(sheet, aliases) {
  const data = sheet.getDataRange().getDisplayValues();
  for (let r = 0; r < data.length; r++) {
    for (let c = 0; c < data[r].length; c++) {
      const current = FS96_key_(data[r][c]);
      if (!current) continue;
      for (const alias of aliases) {
        if (current === FS96_key_(alias)) {
          const value = c + 1 < data[r].length ? sheet.getRange(r + 1, c + 2).getValue() : '';
          return { found: true, value: FS96_rate_(value), label: data[r][c] };
        }
      }
    }
  }
  return { found: false, value: 0, label: '' };
}

function FS96_num_(v) {
  const n = Number(v);
  return isFinite(n) ? n : 0;
}

function FS96_rate_(v) {
  if (typeof v === 'number') return v > 1 ? v / 100 : v;
  const n = Number(String(v || '').replace('%', '').replace(',', '.').trim());
  if (!isFinite(n)) return 0;
  return n > 1 ? n / 100 : n;
}

function FS96_key_(v) {
  return String(v || '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd')
    .replace(/[^a-z0-9]/g, '');
}
