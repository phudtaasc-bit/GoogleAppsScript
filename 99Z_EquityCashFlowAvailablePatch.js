/*************************************************
 * 99Z_EquityCashFlowAvailablePatch.js
 *
 * Quy ước FCFE theo yêu cầu mô hình:
 * - Dòng tiền còn lại sau hoạt động, thuế và trả nợ là dòng tiền sẵn có cho CSH,
 *   nên được ghi nhận vào FCFE tại tháng phát sinh.
 * - Tiền này vẫn được giữ trong dự án và chuyển sang tiền đầu kỳ tháng sau.
 * - Khi tháng sau thiếu tiền, ưu tiên dùng tiền giữ lại; chỉ phần thiếu còn lại
 *   mới huy động CSH và vốn vay theo tỷ lệ đầu vào.
 * - Không cộng thêm khoản phân phối cuối kỳ vào FCFE để tránh ghi nhận hai lần.
 *************************************************/

const FS99Z_BASE_LAP_SHEET04_ = FS_lapSheet04;
const FS99Z_BASE_LAP_SHEET00_ = FS_lapSheet00;

FS_lapSheet04 = function() {
  FS99Z_BASE_LAP_SHEET04_();

  const ss = SpreadsheetApp.getActive();
  const tech = ss.getSheetByName('01. Kỹ thuật');
  const sh04 = ss.getSheetByName('04. Dòng tiền & Lợi nhuận');
  if (!tech || !sh04) return;

  const soThang = Number(FS04_getInfoValue_(tech, 'Số tháng mô hình')) || 0;
  if (!soThang) return;

  const endRow = soThang + 2;

  // FCFE khả dụng cho CSH trong tháng:
  // FCFF + giải ngân vay - trả gốc.
  // Tiền dư vẫn được giữ ở cột Z và ưu tiên sử dụng cho các tháng sau.
  for (let r = 3; r <= endRow; r++) {
    sh04.getRange(r, 37).setFormula(`=P${r}+V${r}-X${r}`); // AK
    sh04.getRange(r, 39).setFormula(`=SUM($AK$3:AK${r})`); // AM
  }

  sh04.getRange(2, 37).setValue('FCFE khả dụng cho CSH');
  SpreadsheetApp.flush();
};

FS_lapSheet00 = function() {
  FS99Z_BASE_LAP_SHEET00_();

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

  // Tháng số 1 là thời điểm 0, nên tách dòng đầu khỏi hàm NPV.
  sh.getRange('D34').setFormula(
    `=IFERROR((${s04}!AJ3+NPV(${monthlyWacc};${s04}!AJ4:AJ${endRow04}))/1000000000;0)`
  );
  sh.getRange('D37').setFormula(
    `=IFERROR((${s04}!AK3+NPV(${monthlyCostOfEquity};${s04}!AK4:AK${endRow04}))/1000000000;0)`
  );

  sh.getRange('E37').setValue('FCFE khả dụng cho CSH chiết khấu theo chi phí vốn CSH');
  sh.getRange('E38').setValue('IRR của dòng tiền khả dụng cho CSH');
  sh.getRange('E39').setValue('Theo FCFE khả dụng lũy kế');
  SpreadsheetApp.flush();
};
