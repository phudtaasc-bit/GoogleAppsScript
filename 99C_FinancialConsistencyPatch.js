/*************************************************
 * 99C_FinancialConsistencyPatch.js
 * Vá nhất quán dòng tiền khách hàng, FCFE và IRR.
 * Không thay đổi cấu trúc các sheet hiện có.
 *************************************************/

function FS99C_syncRevenueFromSheet02_() {
  const ss = SpreadsheetApp.getActive();
  const sh02 = ss.getSheetByName('02. Doanh thu');
  const sh03 = ss.getSheetByName('03. Chi phí & Vốn');

  if (!sh02 || !sh03) {
    throw new Error('Thiếu Sheet 02 hoặc Sheet 03 để đồng bộ doanh thu.');
  }

  const byMonth = {};
  const lastRow02 = sh02.getLastRow();

  if (lastRow02 >= 2) {
    const rows = sh02.getRange(2, 1, lastRow02 - 1, 32).getValues();
    rows.forEach(r => {
      const month = Number(r[0]) || 0;        // A - Tháng số
      if (!month) return;

      if (!byMonth[month]) {
        byMonth[month] = { customerCash: 0, vatOut: 0, cit: 0, sellingCost: 0 };
      }

      byMonth[month].customerCash += Number(r[19]) || 0; // T
      byMonth[month].vatOut += Number(r[18]) || 0;       // S
      byMonth[month].cit += Math.max(0, Number(r[31]) || 0); // AF
      byMonth[month].sellingCost += Number(r[22]) || 0;  // W
    });
  }

  const lastRow03 = sh03.getLastRow();
  if (lastRow03 < 2) return;

  const rowCount = lastRow03 - 1;
  const months = sh03.getRange(2, 1, rowCount, 1).getValues();
  const outE = [];
  const outF = [];
  const outG = [];
  const outL = [];

  months.forEach(r => {
    const month = Number(r[0]) || 0;
    const x = byMonth[month] || { customerCash: 0, vatOut: 0, cit: 0, sellingCost: 0 };
    outE.push([x.customerCash]);
    outF.push([x.vatOut]);
    outG.push([x.cit]);
    outL.push([x.sellingCost]);
  });

  sh03.getRange(2, 5, rowCount, 1).setValues(outE);  // E - Dòng tiền KH
  sh03.getRange(2, 6, rowCount, 1).setValues(outF);  // F - VAT đầu ra
  sh03.getRange(2, 7, rowCount, 1).setValues(outG);  // G - Thuế TNDN
  sh03.getRange(2, 12, rowCount, 1).setValues(outL); // L - Chi phí bán hàng
}

function FS99C_patchSheet04_() {
  const ss = SpreadsheetApp.getActive();
  const sh04 = ss.getSheetByName('04. Dòng tiền & Lợi nhuận');
  const tech = ss.getSheetByName('01. Kỹ thuật');

  if (!sh04 || !tech) {
    throw new Error('Thiếu Sheet 04 hoặc Sheet 01 để vá dòng tiền.');
  }

  const soThang = Number(FS04_getInfoValue_(tech, 'Số tháng mô hình')) || 0;
  if (!soThang) throw new Error('Thiếu Số tháng mô hình.');

  const endRow = soThang + 2;
  const s02 = `'02. Doanh thu'`;

  // Sheet 04 lấy dòng tiền khách hàng trực tiếp từ nguồn duy nhất là Sheet 02 cột T.
  sh04.getRange('G3').setFormula(
    `=MAP(A3:A${endRow};LAMBDA(t;IF(t="";"";SUMIF(${s02}!A:A;t;${s02}!T:T))))`
  );

  // FCFE chỉ gồm dòng tiền thực tế giữa dự án và chủ sở hữu:
  // âm khi CSH góp mới, dương khi phân phối/hoàn vốn.
  sh04.getRange(3, 37, soThang, 1).setFormulaR1C1('=RC[-19]-RC[-18]'); // AK = R - S

  SpreadsheetApp.flush();
}

function FS99C_patchSummaryIRR_() {
  const ss = SpreadsheetApp.getActive();
  const sh00 = ss.getSheetByName('00. Tổng hợp');
  const sh04 = ss.getSheetByName('04. Dòng tiền & Lợi nhuận');
  const tech = ss.getSheetByName('01. Kỹ thuật');

  if (!sh00 || !sh04 || !tech) return;

  const soThang = Number(FS04_getInfoValue_(tech, 'Số tháng mô hình')) || 0;
  if (!soThang) return;
  const endRow = soThang + 2;
  const s04 = `'04. Dòng tiền & Lợi nhuận'`;

  // Không che lỗi IRR bằng IFERROR(...;0). Nếu chuỗi không có đổi dấu, hiển thị #N/A.
  sh00.getRange('D35').setFormula(
    `=IF(OR(COUNTIF(${s04}!AJ3:AJ${endRow};">0")=0;COUNTIF(${s04}!AJ3:AJ${endRow};"<0")=0);NA();(1+IRR(${s04}!AJ3:AJ${endRow};0,01))^12-1)`
  );
  sh00.getRange('D38').setFormula(
    `=IF(OR(COUNTIF(${s04}!AK3:AK${endRow};">0")=0;COUNTIF(${s04}!AK3:AK${endRow};"<0")=0);NA();(1+IRR(${s04}!AK3:AK${endRow};0,01))^12-1)`
  );
  sh00.getRange('E35').setValue('IRR tháng quy đổi năm; không che lỗi bằng 0%');
  sh00.getRange('E38').setValue('FCFE = phân phối cho CSH - vốn CSH góp mới');
  SpreadsheetApp.flush();
}

function FS99C_lapSheet00() {
  FS_lapSheet00();
  FS99C_patchSummaryIRR_();
}

function FS99C_lapSheet04() {
  FS99C_syncRevenueFromSheet02_();
  FS_lapSheet04();
  FS99C_patchSheet04_();
}

function FS99C_chayMoHinh_Buoc1() {
  FS_chayMoHinh_Buoc1();
  FS99C_syncRevenueFromSheet02_();
  FS_lapSheet04();
  FS99C_patchSheet04_();
}

function FS99C_chayMoHinh_Buoc2() {
  FS99C_syncRevenueFromSheet02_();
  FS_chayMoHinh_Buoc2();
  FS99C_syncRevenueFromSheet02_();
  FS99C_patchSheet04_();
}

function FS99C_chayMoHinh_Buoc3() {
  FS99C_syncRevenueFromSheet02_();
  FS99C_patchSheet04_();
  FS_chayMoHinh_Buoc3_SafeCIT();
  FS99C_patchSummaryIRR_();
}

function FS99C_chayTuDongTheoTrangThai() {
  const state = FS_getRunState_();
  if (!state.status) return FS99C_chayMoHinh_Buoc1();
  if (state.status === 'INITIALIZED' || state.status === 'ITERATING') return FS99C_chayMoHinh_Buoc2();
  if (state.status === 'CONVERGED') return FS99C_chayMoHinh_Buoc3();
  return FS_chayToanBoMoHinh_Safe();
}
