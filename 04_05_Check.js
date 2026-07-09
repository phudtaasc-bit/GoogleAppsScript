/*************************************************
 * 04_05_Check.gs
 * Kiểm tra logic Sheet 04 - Dòng tiền
 *************************************************/

function FS04_check() {
  const ss = SpreadsheetApp.getActive();
  const sh = FS04_getSheetByPrefix_(ss, '04');
  if (!sh) throw new Error('Không tìm thấy sheet bắt đầu bằng "04".');

  const lastRow = sh.getLastRow();
  if (lastRow < 3) throw new Error('Sheet 04 chưa có dữ liệu.');

  const startRow = 3;
  const numRows = lastRow - 2;
  const data = sh.getRange(startRow, 1, numRows, 32).getValues();

  const issues = [];
  const tol = 10;

  data.forEach((r, i) => {
    const rowNo = startRow + i;
    const thang = r[0];

    const dtTruocVAT = n_(r[4]);     // E
    const vatDauRa = n_(r[5]);       // F
    const tienKH = n_(r[6]);         // G

    const chiTruocVAT = n_(r[7]);    // H
    const chiSauVAT = n_(r[8]);      // I
    const vatDauVao = n_(r[9]);      // J

    const vatPhaiNop = n_(r[10]);    // K
    const thueTNDN = n_(r[11]);      // L

    const giaVon = n_(r[12]);        // M
    const lnChiuThue = n_(r[13]);    // N
    const lnSauThue = n_(r[14]);     // O

    const dongTienTruocTT = n_(r[15]); // P
    const nhuCauVon = n_(r[16]);       // Q
    const dongTienSauTT = n_(r[17]);   // R

    const vonCSH = n_(r[18]);        // S
    const giaiNgan = n_(r[19]);      // T
    const tienCuoiKy = n_(r[24]);    // Y

    // 1. Tiền KH = Doanh thu + VAT đầu ra
    if (Math.abs(tienKH - dtTruocVAT - vatDauRa) > tol) {
      issues.push(`Dòng ${rowNo} / tháng ${thang}: Dòng tiền KH ≠ Doanh thu + VAT đầu ra.`);
    }

    // 2. Chi sau VAT = Chi trước VAT + VAT đầu vào
    if (Math.abs(chiSauVAT - chiTruocVAT - vatDauVao) > tol) {
      issues.push(`Dòng ${rowNo} / tháng ${thang}: Tổng chi sau VAT ≠ Chi trước VAT + VAT đầu vào.`);
    }

    // 3. Lợi nhuận sau thuế = Lợi nhuận chịu thuế - Thuế TNDN
    if (Math.abs(lnSauThue - (lnChiuThue - thueTNDN)) > tol) {
      issues.push(`Dòng ${rowNo} / tháng ${thang}: Lợi nhuận sau thuế chưa khớp.`);
    }

    // 4. Dòng tiền trước tài trợ = Tiền KH - Chi sau VAT - VAT phải nộp - Thuế TNDN
    const cfBefore = tienKH - chiSauVAT - vatPhaiNop - thueTNDN;
    if (Math.abs(dongTienTruocTT - cfBefore) > tol) {
      issues.push(`Dòng ${rowNo} / tháng ${thang}: Dòng tiền trước tài trợ chưa khớp.`);
    }

    // 5. Nhu cầu vốn không âm
    if (nhuCauVon < -tol) {
      issues.push(`Dòng ${rowNo} / tháng ${thang}: Nhu cầu vốn âm.`);
    }

    // 6. Tiền cuối kỳ không âm
    if (tienCuoiKy < -tol) {
      issues.push(`Dòng ${rowNo} / tháng ${thang}: Tiền cuối kỳ âm.`);
    }

    // 7. Giá vốn không âm
    if (giaVon < -tol) {
      issues.push(`Dòng ${rowNo} / tháng ${thang}: Giá vốn âm.`);
    }

    // 8. Vốn CSH và giải ngân không âm
    if (vonCSH < -tol || giaiNgan < -tol) {
      issues.push(`Dòng ${rowNo} / tháng ${thang}: Vốn góp hoặc giải ngân âm.`);
    }
  });

  if (issues.length === 0) {
    SpreadsheetApp.getUi().alert('✅ Check Sheet 04: Không phát hiện lỗi.');
  } else {
    SpreadsheetApp.getUi().alert(
      '⚠ Check Sheet 04 phát hiện lỗi:\n\n' +
      issues.slice(0, 20).join('\n') +
      (issues.length > 20 ? `\n\n... còn ${issues.length - 20} lỗi khác.` : '')
    );
  }
}

function n_(v) {
  const num = Number(v);
  return isNaN(num) ? 0 : num;
}