/*************************************************
 * 94_CITAudit.gs
 * Kiểm soát Thuế TNDN theo nguyên tắc mô hình:
 * - Tính riêng theo từng sản phẩm đầu ra.
 * - Lợi nhuận chịu thuế và Thuế TNDN không âm.
 * - Không bù trừ lỗ giữa sản phẩm hoặc giữa các tháng.
 * - Sheet 03 và Sheet 04 chỉ tổng hợp Thuế TNDN từ Sheet 02.
 *************************************************/

function FS_AUDIT_TNDN() {
  const ss = SpreadsheetApp.getActive();
  const sh02 = ss.getSheetByName('02. Doanh thu');
  const sh03 = ss.getSheetByName('03. Chi phí & Vốn');
  const sh04 = ss.getSheetByName('04. Dòng tiền & Lợi nhuận');

  if (!sh02) throw new Error('Không tìm thấy sheet "02. Doanh thu".');
  if (!sh03) throw new Error('Không tìm thấy sheet "03. Chi phí & Vốn".');
  if (!sh04) throw new Error('Không tìm thấy sheet "04. Dòng tiền & Lợi nhuận".');
  if (sh02.getLastColumn() < 32) throw new Error('Sheet 02 thiếu cấu trúc 32 cột A:AF.');
  if (sh03.getLastColumn() < 7) throw new Error('Sheet 03 thiếu cột Thuế TNDN tại G.');
  if (sh04.getLastColumn() < 12) throw new Error('Sheet 04 thiếu cột Thuế TNDN tại L.');

  const issues = [];
  const warnings = [];
  const tol = 10;
  const taxByMonth02 = {};

  const lastRow02 = sh02.getLastRow();
  if (lastRow02 >= 2) {
    const data02 = sh02.getRange(2, 1, lastRow02 - 1, 32).getValues();

    data02.forEach((r, i) => {
      const rowNo = i + 2;
      const month = FS94_num_(r[0]);       // A
      const product = String(r[4] || '').trim(); // E
      if (!month || !product) return;

      const revenue = FS94_num_(r[16]);    // Q
      const taxRate = FS94_rate_(r[20]);   // U
      const taxCost = FS94_num_(r[29]);    // AD
      const taxable = FS94_num_(r[30]);    // AE
      const cit = FS94_num_(r[31]);        // AF

      const expectedTaxable = Math.max(0, revenue - taxCost);
      const expectedCit = expectedTaxable * taxRate;

      if (taxable < -tol) {
        issues.push(`Sheet 02 dòng ${rowNo} / ${product}: Lợi nhuận chịu thuế âm.`);
      }
      if (cit < -tol) {
        issues.push(`Sheet 02 dòng ${rowNo} / ${product}: Thuế TNDN âm.`);
      }
      if (Math.abs(taxable - expectedTaxable) > tol) {
        issues.push(
          `Sheet 02 dòng ${rowNo} / ${product}: Lợi nhuận chịu thuế không bằng MAX(0; Doanh thu - Tổng giá vốn tính thuế).`
        );
      }
      if (Math.abs(cit - expectedCit) > tol) {
        issues.push(
          `Sheet 02 dòng ${rowNo} / ${product}: Thuế TNDN không bằng Lợi nhuận chịu thuế × Thuế suất sản phẩm.`
        );
      }
      if (taxRate < 0 || taxRate > 1) {
        issues.push(`Sheet 02 dòng ${rowNo} / ${product}: Thuế suất TNDN ngoài khoảng 0%–100%.`);
      }
      if (revenue <= tol && cit > tol) {
        warnings.push(`Sheet 02 dòng ${rowNo} / ${product}: Không có doanh thu nhưng vẫn phát sinh Thuế TNDN.`);
      }

      taxByMonth02[month] = (taxByMonth02[month] || 0) + cit;
    });
  }

  const taxByMonth03 = FS94_readMonthly_(sh03, 2, 1, 7);
  const taxByMonth04 = FS94_readMonthly_(sh04, 3, 1, 12);
  const months = new Set(
    Object.keys(taxByMonth02)
      .concat(Object.keys(taxByMonth03))
      .concat(Object.keys(taxByMonth04))
  );

  months.forEach(key => {
    const month = Number(key);
    const t02 = taxByMonth02[month] || 0;
    const t03 = taxByMonth03[month] || 0;
    const t04 = taxByMonth04[month] || 0;

    if (t03 < -tol) issues.push(`Tháng ${month}: Thuế TNDN tại Sheet 03 âm.`);
    if (t04 < -tol) issues.push(`Tháng ${month}: Thuế TNDN tại Sheet 04 âm.`);
    if (Math.abs(t03 - t02) > tol) {
      issues.push(`Tháng ${month}: Thuế TNDN Sheet 03 không khớp tổng theo sản phẩm tại Sheet 02.`);
    }
    if (Math.abs(t04 - t02) > tol) {
      issues.push(`Tháng ${month}: Thuế TNDN Sheet 04 không khớp tổng theo sản phẩm tại Sheet 02.`);
    }
  });

  const report = [`KIỂM TRA THUẾ TNDN: ${issues.length} lỗi, ${warnings.length} cảnh báo.`];
  if (issues.length) report.push('\nLỖI:\n- ' + issues.slice(0, 40).join('\n- '));
  if (warnings.length) report.push('\nCẢNH BÁO:\n- ' + warnings.slice(0, 20).join('\n- '));
  if (!issues.length && !warnings.length) {
    report.push('\nThuế TNDN đang tính riêng theo từng sản phẩm, không âm và tổng hợp nhất quán giữa Sheet 02–03–04.');
  }

  SpreadsheetApp.getUi().alert(report.join('\n'));
  return { issues, warnings };
}

function FS94_readMonthly_(sheet, startRow, monthCol, taxCol) {
  const out = {};
  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) return out;

  const width = Math.max(monthCol, taxCol);
  const values = sheet.getRange(startRow, 1, lastRow - startRow + 1, width).getValues();
  values.forEach(r => {
    const month = FS94_num_(r[monthCol - 1]);
    if (!month) return;
    out[month] = (out[month] || 0) + FS94_num_(r[taxCol - 1]);
  });
  return out;
}

function FS94_num_(value) {
  const n = Number(value);
  return isFinite(n) ? n : 0;
}

function FS94_rate_(value) {
  const n = FS94_num_(value);
  return n > 1 ? n / 100 : n;
}
