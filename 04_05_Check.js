/*************************************************
 * 04_05_Check.gs
 * Kiểm tra logic Sheet 04 - Dòng tiền
 *************************************************/

function FS04_check() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('04. Dòng tiền & Lợi nhuận');
  if (!sh) throw new Error('Không tìm thấy sheet "04. Dòng tiền & Lợi nhuận".');

  const lastRow = sh.getLastRow();
  if (lastRow < 3) throw new Error('Sheet 04 chưa có dữ liệu.');
  if (sh.getLastColumn() < 39) throw new Error('Sheet 04 thiếu cấu trúc 39 cột A:AM.');

  const startRow = 3;
  const numRows = lastRow - 2;
  const data = sh.getRange(startRow, 1, numRows, 39).getValues();

  const issues = [];
  const tol = 10;
  let prevDebt = 0;
  let prevCash = 0;

  data.forEach((r, i) => {
    const rowNo = startRow + i;
    const thang = FS04_checkNum_(r[0]);
    if (!thang) return;

    const dtTruocVAT = FS04_checkNum_(r[4]);       // E
    const vatDauRa = FS04_checkNum_(r[5]);         // F
    const tienKH = FS04_checkNum_(r[6]);           // G
    const chiTruocVAT = FS04_checkNum_(r[7]);      // H
    const chiSauVAT = FS04_checkNum_(r[8]);        // I
    const vatDauVao = FS04_checkNum_(r[9]);        // J
    const vatPhaiNop = FS04_checkNum_(r[10]);      // K
    const thueTNDN = FS04_checkNum_(r[11]);        // L
    const giaVon = FS04_checkNum_(r[12]);          // M
    const lnChiuThue = FS04_checkNum_(r[13]);      // N
    const lnSauThue = FS04_checkNum_(r[14]);       // O
    const dongTienTruocTT = FS04_checkNum_(r[15]); // P
    const nhuCauVon = FS04_checkNum_(r[16]);       // Q
    const phanPhoiCSH = FS04_checkNum_(r[17]);     // R
    const cshGopMoi = FS04_checkNum_(r[18]);       // S
    const cshNopLai = FS04_checkNum_(r[19]);       // T
    const tongDongCSH = FS04_checkNum_(r[20]);     // U
    const giaiNganVay = FS04_checkNum_(r[21]);     // V
    const laiVay = FS04_checkNum_(r[22]);          // W
    const traGoc = FS04_checkNum_(r[23]);          // X
    const duNoCuoiKy = FS04_checkNum_(r[24]);      // Y
    const tienCuoiKy = FS04_checkNum_(r[25]);      // Z
    const tienKhaDung = FS04_checkNum_(r[26]);     // AA
    const fcff = FS04_checkNum_(r[35]);            // AJ
    const fcfe = FS04_checkNum_(r[36]);            // AK
    const laKyCuoi = i === data.length - 1;

    if (Math.abs(tienKH - dtTruocVAT - vatDauRa) > tol) {
      issues.push(`Dòng ${rowNo} / tháng ${thang}: Dòng tiền huy động từ KH ≠ Doanh thu trước VAT + VAT đầu ra.`);
    }

    if (Math.abs(chiSauVAT - chiTruocVAT - vatDauVao) > tol) {
      issues.push(`Dòng ${rowNo} / tháng ${thang}: Tổng chi sau VAT ≠ Tổng chi trước VAT + VAT đầu vào.`);
    }

    if (Math.abs(lnSauThue - (lnChiuThue - thueTNDN)) > tol) {
      issues.push(`Dòng ${rowNo} / tháng ${thang}: Lợi nhuận sau thuế chưa khớp.`);
    }

    const cfBefore = tienKH - chiSauVAT - vatPhaiNop - thueTNDN;
    if (Math.abs(dongTienTruocTT - cfBefore) > tol) {
      issues.push(`Dòng ${rowNo} / tháng ${thang}: Dòng tiền trước tài trợ chưa khớp.`);
    }

    if (Math.abs(fcff - dongTienTruocTT) > tol) {
      issues.push(`Dòng ${rowNo} / tháng ${thang}: FCFF không khớp dòng tiền trước tài trợ.`);
    }

    const cashBeforeFunding = prevCash + dongTienTruocTT;
    const needCalc = Math.max(0, -cashBeforeFunding);
    if (Math.abs(nhuCauVon - needCalc) > tol) {
      issues.push(`Dòng ${rowNo} / tháng ${thang}: Nhu cầu vốn chưa trừ đúng tiền đầu kỳ.`);
    }

    if (Math.abs(nhuCauVon - cshGopMoi - giaiNganVay) > tol) {
      issues.push(`Dòng ${rowNo} / tháng ${thang}: Nhu cầu vốn ≠ CSH góp mới + Giải ngân vay.`);
    }

    if (Math.abs(tongDongCSH - cshGopMoi - cshNopLai) > tol) {
      issues.push(`Dòng ${rowNo} / tháng ${thang}: Tổng dòng CSH vào dự án chưa khớp.`);
    }

    if (Math.abs(cshNopLai) > tol) {
      issues.push(`Dòng ${rowNo} / tháng ${thang}: Không được phát sinh CSH nộp lại từ tiền đã phân phối.`);
    }

    const principalCalc = Math.min(prevDebt + giaiNganVay + laiVay, Math.max(0, cashBeforeFunding));
    if (Math.abs(traGoc - principalCalc) > tol) {
      issues.push(`Dòng ${rowNo} / tháng ${thang}: Trả gốc chưa khớp tiền dư và dư nợ sau vốn hóa lãi.`);
    }

    const debtCalc = Math.max(0, prevDebt + giaiNganVay + laiVay - traGoc);
    if (Math.abs(duNoCuoiKy - debtCalc) > tol) {
      issues.push(`Dòng ${rowNo} / tháng ${thang}: Dư nợ cuối kỳ ≠ dư nợ đầu kỳ + giải ngân + lãi vay - trả gốc.`);
    }

    const availableCalc = Math.max(0, cashBeforeFunding - traGoc);
    if (Math.abs(tienKhaDung - availableCalc) > tol) {
      issues.push(`Dòng ${rowNo} / tháng ${thang}: Tiền khả dụng sau trả nợ chưa khớp.`);
    }

    const distributionCalc = laKyCuoi ? tienKhaDung : 0;
    if (Math.abs(phanPhoiCSH - distributionCalc) > tol) {
      issues.push(`Dòng ${rowNo} / tháng ${thang}: Chỉ kỳ cuối mới được phân phối toàn bộ tiền khả dụng.`);
    }

    if (Math.abs(tienCuoiKy - Math.max(0, tienKhaDung - phanPhoiCSH)) > tol) {
      issues.push(`Dòng ${rowNo} / tháng ${thang}: Tiền cuối kỳ chưa khớp tiền khả dụng trừ phân phối.`);
    }

    const fcfeCalc = fcff + giaiNganVay - traGoc;
    if (Math.abs(fcfe - fcfeCalc) > tol) {
      issues.push(`Dòng ${rowNo} / tháng ${thang}: FCFE ≠ FCFF + giải ngân vay - trả gốc.`);
    }

    // Kiểm tra tương đương đại số để phát hiện ghi nhận trùng tiền giữ lại.
    const fcfeCashBridge = tienKhaDung - prevCash - cshGopMoi;
    if (Math.abs(fcfe - fcfeCashBridge) > tol) {
      issues.push(`Dòng ${rowNo} / tháng ${thang}: FCFE không khớp cầu nối biến động tiền khả dụng.`);
    }

    if (nhuCauVon > tol && traGoc > tol) {
      issues.push(`Dòng ${rowNo} / tháng ${thang}: Vừa huy động vốn vừa trả gốc trong cùng kỳ.`);
    }

    // VAT âm chỉ hợp lệ ở kỳ cuối và được hiểu là khoản hoàn thuế.
    if (vatPhaiNop < -tol && !laKyCuoi) {
      issues.push(`Dòng ${rowNo} / tháng ${thang}: VAT phải nộp âm ngoài kỳ cuối; cần kiểm tra bù trừ VAT.`);
    }
    if (thueTNDN < -tol) issues.push(`Dòng ${rowNo} / tháng ${thang}: Thuế TNDN âm.`);
    if (giaVon < -tol) issues.push(`Dòng ${rowNo} / tháng ${thang}: Tổng giá vốn tính thuế âm.`);
    if (nhuCauVon < -tol || cshGopMoi < -tol || giaiNganVay < -tol) issues.push(`Dòng ${rowNo} / tháng ${thang}: Nhu cầu vốn, vốn CSH hoặc giải ngân vay âm.`);
    if (laiVay < -tol || traGoc < -tol || duNoCuoiKy < -tol) issues.push(`Dòng ${rowNo} / tháng ${thang}: Lãi vay, trả gốc hoặc dư nợ âm.`);
    if (tienCuoiKy < -tol || tienKhaDung < -tol) issues.push(`Dòng ${rowNo} / tháng ${thang}: Tiền cuối kỳ hoặc tiền khả dụng âm.`);

    prevDebt = duNoCuoiKy;
    prevCash = tienCuoiKy;
  });

  if (issues.length === 0) {
    SpreadsheetApp.getUi().alert('Check Sheet 04: Không phát hiện lỗi logic.');
  } else {
    SpreadsheetApp.getUi().alert(
      'Check Sheet 04 phát hiện lỗi:\n\n' +
      issues.slice(0, 30).join('\n') +
      (issues.length > 30 ? `\n\n... còn ${issues.length - 30} lỗi khác.` : '')
    );
  }
}

function FS04_checkNum_(v) {
  const num = Number(v);
  return isFinite(num) ? num : 0;
}
