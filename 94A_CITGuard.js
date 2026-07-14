/*************************************************
 * 94A_CITGuard.gs
 * Chốt kiểm tra Thuế TNDN trước khi lập các sheet tổng hợp.
 * Không hiển thị UI; chỉ ném lỗi khi mô hình không nhất quán.
 *
 * Bản ổn định:
 * - Flush và chờ Sheets tính lại trước khi đọc.
 * - Đọc lại tối đa 3 lần để loại cảnh báo do giá trị cũ.
 * - Ép kiểu Number và so sánh theo sai số cho phép.
 *************************************************/

function FS94_assertCITConsistency_() {
  const ss = SpreadsheetApp.getActive();
  const sh02 = ss.getSheetByName('02. Doanh thu');
  const sh03 = ss.getSheetByName('03. Chi phí & Vốn');
  const sh04 = ss.getSheetByName('04. Dòng tiền & Lợi nhuận');

  if (!sh02 || !sh03 || !sh04) {
    throw new Error('Thiếu Sheet 02, 03 hoặc 04 để kiểm tra Thuế TNDN.');
  }

  const tol = 10;
  const maxAttempts = 3;
  let lastErrors = [];

  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    SpreadsheetApp.flush();
    if (attempt > 1) Utilities.sleep(1200);
    SpreadsheetApp.flush();

    const errors = FS94G_collectCITErrors_(sh02, sh03, sh04, tol);
    if (!errors.length) return true;

    lastErrors = errors;
    if (attempt < maxAttempts) Utilities.sleep(1200);
  }

  throw new Error(
    'Dừng lập sheet tổng hợp vì Thuế TNDN chưa nhất quán sau khi đã đọc lại ' +
    maxAttempts + ' lần:\n- ' +
    lastErrors.slice(0, 30).join('\n- ') +
    (lastErrors.length > 30 ? `\n- ... còn ${lastErrors.length - 30} lỗi khác.` : '')
  );
}

function FS94G_collectCITErrors_(sh02, sh03, sh04, tol) {
  const errors = [];
  const taxByMonth02 = {};

  if (sh02.getLastRow() >= 2) {
    const rows = sh02.getRange(2, 1, sh02.getLastRow() - 1, 32).getValues();
    rows.forEach((r, i) => {
      const month = FS94G_num_(r[0]);
      const product = String(r[4] || '').trim();
      if (!month || !product) return;

      const revenue = FS94G_num_(r[16]);
      const taxRate = FS94G_rate_(r[20]);
      const taxCost = FS94G_num_(r[29]);
      const taxable = FS94G_num_(r[30]);
      const cit = FS94G_num_(r[31]);
      const expectedTaxable = Math.max(0, revenue - taxCost);
      const expectedCit = expectedTaxable * taxRate;
      const rowNo = i + 2;

      if (taxable < -tol || cit < -tol) {
        errors.push(`Sheet 02 dòng ${rowNo} / ${product}: Lợi nhuận chịu thuế hoặc Thuế TNDN âm.`);
      }
      if (Math.abs(taxable - expectedTaxable) > tol) {
        errors.push(`Sheet 02 dòng ${rowNo} / ${product}: Lợi nhuận chịu thuế sai công thức theo sản phẩm.`);
      }
      if (Math.abs(cit - expectedCit) > tol) {
        errors.push(`Sheet 02 dòng ${rowNo} / ${product}: Thuế TNDN sai công thức theo sản phẩm.`);
      }

      taxByMonth02[month] = (taxByMonth02[month] || 0) + cit;
    });
  }

  const taxByMonth03 = FS94G_readMonthly_(sh03, 2, 7);
  const taxByMonth04 = FS94G_readMonthly_(sh04, 3, 12);
  const months = new Set(
    Object.keys(taxByMonth02)
      .concat(Object.keys(taxByMonth03))
      .concat(Object.keys(taxByMonth04))
  );

  months.forEach(key => {
    const month = Number(key);
    const t02 = FS94G_num_(taxByMonth02[month]);
    const t03 = FS94G_num_(taxByMonth03[month]);
    const t04 = FS94G_num_(taxByMonth04[month]);

    if (t03 < -tol || t04 < -tol) {
      errors.push(`Tháng ${month}: Thuế TNDN âm tại Sheet 03 hoặc Sheet 04.`);
    }
    if (Math.abs(t03 - t02) > tol) {
      errors.push(
        `Tháng ${month}: Thuế TNDN Sheet 03 không khớp Sheet 02 ` +
        `(S02=${Math.round(t02)}; S03=${Math.round(t03)}; lệch=${Math.round(t03 - t02)}).`
      );
    }
    if (Math.abs(t04 - t02) > tol) {
      errors.push(
        `Tháng ${month}: Thuế TNDN Sheet 04 không khớp Sheet 02 ` +
        `(S02=${Math.round(t02)}; S04=${Math.round(t04)}; lệch=${Math.round(t04 - t02)}).`
      );
    }
  });

  return errors;
}

function FS94G_readMonthly_(sheet, startRow, taxCol) {
  const out = {};
  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) return out;

  const values = sheet.getRange(startRow, 1, lastRow - startRow + 1, taxCol).getValues();
  values.forEach(r => {
    const month = FS94G_num_(r[0]);
    if (!month) return;
    out[month] = (out[month] || 0) + FS94G_num_(r[taxCol - 1]);
  });
  return out;
}

function FS94G_num_(value) {
  if (value === null || value === '' || typeof value === 'undefined') return 0;
  const n = Number(value);
  return Number.isFinite(n) ? n : 0;
}

function FS94G_rate_(value) {
  const n = FS94G_num_(value);
  return n > 1 ? n / 100 : n;
}
