/*************************************************
 * 03_05_Check.gs
 * Kiểm tra logic Sheet 03
 *************************************************/

function FS03_checkV2() {
  const ss = SpreadsheetApp.getActive();
  const sh = FS_getSheetByPrefix_(ss, '03');
  if (!sh) throw new Error('Không tìm thấy sheet bắt đầu bằng "03".');

  const lastRow = sh.getLastRow();
  if (lastRow < 3) throw new Error('Sheet 03 chưa có dữ liệu.');

  const startRow = 3;
  const numRows = lastRow - 2;

  const data = sh.getRange(startRow, 1, numRows, 35).getValues();

  const issues = [];
  const tolerance = 10; // cho phép lệch làm tròn 10 đồng

  data.forEach((r, i) => {
    const rowNo = startRow + i;
    const thangSo = r[0];

    const nhuCauVon = FS03_num_(r[20]);      // U
    const vonCSH = FS03_num_(r[21]);         // V
    const luyKeVonCSH = FS03_num_(r[22]);    // W
    const giaiNgan = FS03_num_(r[23]);       // X
    const luyKeGiaiNgan = FS03_num_(r[24]);  // Y
    const laiVay = FS03_num_(r[25]);         // Z
    const traGoc = FS03_num_(r[26]);         // AA
    const luyKeTraGoc = FS03_num_(r[27]);    // AB
    const duNo = FS03_num_(r[28]);           // AC
    const dongTienSauTaiTro = FS03_num_(r[29]); // AD
    const tienCuoiKy = FS03_num_(r[30]);     // AE

    const vatDauVao = FS03_num_(r[14]);      // O
    const vatKTDauKy = FS03_num_(r[16]);     // Q
    const vatPhaiNop = FS03_num_(r[17]);     // R
    const vatKTCuoiKy = FS03_num_(r[18]);    // S

    // 1. Nhu cầu vốn = Vốn CSH + Giải ngân vay
    if (Math.abs(nhuCauVon - vonCSH - giaiNgan) > tolerance) {
      issues.push(`Dòng ${rowNo} / tháng ${thangSo}: Nhu cầu vốn ≠ Vốn CSH + Giải ngân vay.`);
    }

    // 2. Lũy kế giải ngân - lũy kế trả gốc = dư nợ cuối kỳ
    if (Math.abs(luyKeGiaiNgan - luyKeTraGoc - duNo) > tolerance) {
      issues.push(`Dòng ${rowNo} / tháng ${thangSo}: Dư nợ cuối kỳ không khớp lũy kế giải ngân - lũy kế trả gốc.`);
    }

    // 3. Không âm
    if (duNo < -tolerance) {
      issues.push(`Dòng ${rowNo} / tháng ${thangSo}: Dư nợ cuối kỳ âm.`);
    }

    if (tienCuoiKy < -tolerance) {
      issues.push(`Dòng ${rowNo} / tháng ${thangSo}: Tiền cuối kỳ âm.`);
    }

    if (vatPhaiNop < -tolerance) {
      issues.push(`Dòng ${rowNo} / tháng ${thangSo}: VAT phải nộp âm.`);
    }

    if (vatKTCuoiKy < -tolerance) {
      issues.push(`Dòng ${rowNo} / tháng ${thangSo}: VAT còn khấu trừ cuối kỳ âm.`);
    }

    // 4. VAT logic
    const vatPayCalc = Math.max(0, FS03_num_(r[5]) - vatKTDauKy - vatDauVao); // F - Q - O
    const vatCreditCalc = Math.max(0, vatKTDauKy + vatDauVao - FS03_num_(r[5]));

    if (Math.abs(vatPhaiNop - vatPayCalc) > tolerance) {
      issues.push(`Dòng ${rowNo} / tháng ${thangSo}: VAT phải nộp tính chưa đúng.`);
    }

    if (Math.abs(vatKTCuoiKy - vatCreditCalc) > tolerance) {
      issues.push(`Dòng ${rowNo} / tháng ${thangSo}: VAT còn khấu trừ cuối kỳ tính chưa đúng.`);
    }
  });

  if (issues.length === 0) {
    SpreadsheetApp.getUi().alert('✅ Check Sheet 03: Không phát hiện lỗi.');
  } else {
    const msg = issues.slice(0, 20).join('\n') +
      (issues.length > 20 ? `\n\n... còn ${issues.length - 20} lỗi khác.` : '');

    SpreadsheetApp.getUi().alert('⚠ Check Sheet 03 phát hiện lỗi:\n\n' + msg);
  }
}

function FS03_num_(v) {
  const n = Number(v);
  return isNaN(n) ? 0 : n;
}