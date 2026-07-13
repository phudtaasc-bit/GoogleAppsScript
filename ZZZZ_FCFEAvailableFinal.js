/*************************************************
 * ZZZZ_FCFEAvailableFinal.js
 * Final override for FCFE after all prior patches.
 *
 * Agreed model rule:
 * - Residual cash after operating costs, taxes, interest and principal repayment
 *   is recognized as FCFE in the month it arises.
 * - The same residual cash remains in the project cash balance and is used first
 *   for future expenditure; only any remaining funding gap is raised according
 *   to the input equity/debt ratio.
 * - Therefore FCFE is FCFF plus net borrowing, not merely equity distribution
 *   minus new equity contribution.
 *************************************************/

const FSZZZZ_BASE_LAP_SHEET04_ = FS_lapSheet04;
const FSZZZZ_BASE_LAP_SHEET00_ = FS_lapSheet00;

FS_lapSheet04 = function() {
  FSZZZZ_BASE_LAP_SHEET04_();

  const ss = SpreadsheetApp.getActive();
  const tech = ss.getSheetByName('01. Kỹ thuật');
  const sh04 = ss.getSheetByName('04. Dòng tiền & Lợi nhuận');
  if (!tech || !sh04) return;

  const soThang = Number(FS04_getInfoValue_(tech, 'Số tháng mô hình')) || 0;
  if (!soThang) return;

  const startRow = 3;
  const endRow = soThang + 2;
  const rowCount = endRow - startRow + 1;

  // AK = FCFF (P) + Giải ngân vay (V) - Trả gốc (X).
  // R1C1 from AK: P = RC[-21], V = RC[-15], X = RC[-13].
  const fcfeFormulas = Array.from({ length: rowCount }, () => [
    '=RC[-21]+RC[-15]-RC[-13]'
  ]);
  sh04.getRange(startRow, 37, rowCount, 1).setFormulasR1C1(fcfeFormulas);

  // AM = FCFE lũy kế.
  const cumulativeFormulas = Array.from({ length: rowCount }, (_, i) => [
    `=SUM($AK$${startRow}:AK${startRow + i})`
  ]);
  sh04.getRange(startRow, 39, rowCount, 1).setFormulas(cumulativeFormulas);

  sh04.getRange(2, 37).setValue('FCFE khả dụng cho CSH');
  sh04.getRange(2, 39).setValue('FCFE lũy kế');
  SpreadsheetApp.flush();
};

FS_lapSheet00 = function() {
  FSZZZZ_BASE_LAP_SHEET00_();

  const ss = SpreadsheetApp.getActive();
  const tech = ss.getSheetByName('01. Kỹ thuật');
  const sh = ss.getSheetByName('00. Tổng hợp');
  if (!tech || !sh) return;

  const soThang = Number(FS00_getInfoValue_(tech, 'Số tháng mô hình')) || 40;
  const endRow04 = soThang + 2;
  const techName = `'01. Kỹ thuật'`;
  const s04 = `'04. Dòng tiền & Lợi nhuận'`;
  const equityRateA1 = FS00_getInfoCellA1_(tech, 'Tỷ suất chiết khấu');
  if (!equityRateA1) return;

  const monthlyWacc = `((1+$E$21)^(1/12)-1)`;
  const monthlyCostOfEquity = `((1+${techName}!${equityRateA1})^(1/12)-1)`;

  // Month 1 is time zero; do not discount it by one extra month.
  sh.getRange('D34').setFormula(
    `=IFERROR((${s04}!AJ3+NPV(${monthlyWacc};${s04}!AJ4:AJ${endRow04}))/1000000000;0)`
  );
  sh.getRange('D37').setFormula(
    `=IFERROR((${s04}!AK3+NPV(${monthlyCostOfEquity};${s04}!AK4:AK${endRow04}))/1000000000;0)`
  );

  sh.getRange('E37').setValue('FCFE khả dụng cho CSH chiết khấu theo chi phí vốn CSH');
  sh.getRange('E38').setValue('IRR của FCFE khả dụng cho CSH');
  sh.getRange('E39').setValue('Theo FCFE khả dụng lũy kế');
  SpreadsheetApp.flush();
};
