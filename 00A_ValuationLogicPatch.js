/*************************************************
 * 00A_ValuationLogicPatch.gs
 * Vá tối thiểu chỉ tiêu NPV/IRR trên Sheet 00:
 * - FCFF dùng WACC của dự án.
 * - FCFE dùng chi phí vốn chủ sở hữu (Tỷ suất chiết khấu).
 * - Dòng tiền theo tháng nên quy đổi lãi suất năm sang tháng hiệu dụng.
 * - IRR tháng quy đổi về IRR năm hiệu dụng.
 *************************************************/

function FS_lapSheet00_Patched() {
  FS_lapSheet00();

  const ss = SpreadsheetApp.getActive();
  const tech = ss.getSheetByName('01. Kỹ thuật');
  const sh00 = ss.getSheetByName('00. Tổng hợp');
  const sh04 = ss.getSheetByName('04. Dòng tiền & Lợi nhuận');

  if (!tech || !sh00 || !sh04) {
    throw new Error('Thiếu Sheet 00, 01 hoặc 04 để cập nhật NPV/IRR.');
  }

  const soThang = Number(FS00_getInfoValue_(tech, 'Số tháng mô hình')) || 0;
  if (!soThang) throw new Error('Thiếu "Số tháng mô hình" để tính NPV/IRR.');

  const endRow04 = soThang + 2;
  const equityRateA1 = FS00_getInfoCellA1_(tech, 'Tỷ suất chiết khấu');
  if (!equityRateA1) {
    throw new Error('Thiếu "Tỷ suất chiết khấu" dùng cho FCFE tại 01. Kỹ thuật.');
  }

  const techName = `'01. Kỹ thuật'`;
  const s04 = `'04. Dòng tiền & Lợi nhuận'`;

  // E21 của Sheet 00 là WACC được tính từ cơ cấu nguồn vốn hiện hành.
  const monthlyWacc = `((1+$E$21)^(1/12)-1)`;
  const monthlyCostOfEquity = `((1+${techName}!${equityRateA1})^(1/12)-1)`;

  sh00.getRange('D33').setFormula(
    `=IFERROR(NPV(${monthlyWacc};${s04}!AJ3:AJ${endRow04})/1000000000;0)`
  );
  sh00.getRange('D34').setFormula(
    `=IFERROR((1+IRR(${s04}!AJ3:AJ${endRow04}))^12-1;0)`
  );

  sh00.getRange('D36').setFormula(
    `=IFERROR(NPV(${monthlyCostOfEquity};${s04}!AK3:AK${endRow04})/1000000000;0)`
  );
  sh00.getRange('D37').setFormula(
    `=IFERROR((1+IRR(${s04}!AK3:AK${endRow04}))^12-1;0)`
  );

  sh00.getRange('E33').setValue('FCFF chiết khấu theo WACC');
  sh00.getRange('E34').setValue('IRR tháng quy đổi năm hiệu dụng');
  sh00.getRange('E36').setValue('FCFE chiết khấu theo chi phí vốn CSH');
  sh00.getRange('E37').setValue('IRR tháng quy đổi năm hiệu dụng');

  SpreadsheetApp.flush();
}
