/*************************************************
 * 99P_PerformancePipeline.js
 * Pipeline cập nhật mô hình theo từng bước độc lập để tránh timeout 6 phút.
 * Không thay đổi bản chất logic tài chính.
 *************************************************/

function FS99P_buoc1_CapNhatChiPhiNguon() {
  const ss = SpreadsheetApp.getActive();
  const tech = ss.getSheetByName('01. Kỹ thuật');
  const sh02 = ss.getSheetByName('02. Doanh thu');
  const sh03 = ss.getSheetByName('03. Chi phí & Vốn');

  if (!tech || !sh02 || !sh03) {
    throw new Error('Thiếu Sheet 01. Kỹ thuật, 02. Doanh thu hoặc 03. Chi phí & Vốn.');
  }

  const vatRates = FSZZZZZZZ_readCommonCostVatRates_(tech);
  const opVatRate = FSZZZZZZZ_pickRate_(vatRates, [
    'Chi phí vận hành',
    'Chi phí bán hàng',
    'Chi phí XD/TB/khác'
  ]);

  FSZZZZZZZ_rebuildSheet02TaxCost_(sh02);
  FSZZZZZZZ_rebuildSheet03Costs_(sh03, vatRates, opVatRate);
  SpreadsheetApp.flush();

  SpreadsheetApp.getActive().toast(
    'Bước 1 hoàn thành: đã cập nhật giá vốn, tổng chi trước VAT, VAT đầu vào và tổng chi sau VAT.',
    'FS Pipeline',
    8
  );
}

function FS99P_buoc2_LapDongTien() {
  const ss = SpreadsheetApp.getActive();

  FS_lapSheet04();
  FSZZZZZZZ_syncOperatingMaintenanceTo04_(ss);
  FSZZZZZZZ_forceFCFE_(ss);
  SpreadsheetApp.flush();

  SpreadsheetApp.getActive().toast(
    'Bước 2 hoàn thành: đã lập Sheet 04 và đồng bộ chi phí vận hành, bảo trì, FCFE.',
    'FS Pipeline',
    8
  );
}

function FS99P_buoc3_LapTongHopNam() {
  FS_lapSheet04A();
  SpreadsheetApp.flush();

  SpreadsheetApp.getActive().toast(
    'Bước 3 hoàn thành: đã lập Sheet 04A theo nguồn duy nhất từ Sheet 04.',
    'FS Pipeline',
    8
  );
}

function FS99P_buoc4_LapTongHopDuAn() {
  const ss = SpreadsheetApp.getActive();

  FS_lapSheet00();
  FSZZZZZZZZ_suaTongHopDoanhThuVaKiemTra();
  SpreadsheetApp.flush();

  SpreadsheetApp.getActive().toast(
    'Bước 4 hoàn thành: đã lập Sheet 00 và cập nhật dòng kiểm tra cân đối riêng.',
    'FS Pipeline',
    8
  );
}

function FS99P_buoc5_KiemTraNhanh() {
  const ss = SpreadsheetApp.getActive();
  const sh03 = ss.getSheetByName('03. Chi phí & Vốn');
  const sh04 = ss.getSheetByName('04. Dòng tiền & Lợi nhuận');
  const sh00 = ss.getSheetByName('00. Tổng hợp');

  if (!sh03 || !sh04 || !sh00) {
    throw new Error('Thiếu Sheet 03, 04 hoặc 00 để kiểm tra.');
  }

  const tolerance = 1;
  const errors = [];

  const last03 = sh03.getLastRow();
  if (last03 >= 2) {
    const data03 = sh03.getRange(2, 1, last03 - 1, Math.max(38, sh03.getLastColumn())).getValues();
    data03.forEach((r, i) => {
      const totalBefore = Number(r[13]) || 0; // N
      const vatIn = Number(r[14]) || 0;       // O
      const totalAfter = Number(r[15]) || 0;  // P
      if (Math.abs(totalAfter - totalBefore - vatIn) > tolerance) {
        errors.push(`Sheet 03 dòng ${i + 2}: P khác N + O`);
      }
    });
  }

  const last04 = sh04.getLastRow();
  if (last04 >= 3) {
    const data04 = sh04.getRange(3, 1, last04 - 2, Math.max(42, sh04.getLastColumn())).getValues();
    data04.forEach((r, i) => {
      const p = Number(r[15]) || 0;  // P
      const v = Number(r[21]) || 0;  // V
      const x = Number(r[23]) || 0;  // X
      const ak = Number(r[36]) || 0; // AK
      if (Math.abs(ak - (p + v - x)) > tolerance) {
        errors.push(`Sheet 04 dòng ${i + 3}: AK khác P + V - X`);
      }
    });
  }

  const balance = Number(sh00.getRange('D44').getValue()) || 0;
  if (Math.abs(balance) > 0.001) {
    errors.push(`Sheet 00 D44 còn chênh lệch ${balance}`);
  }

  const message = errors.length
    ? `Có ${errors.length} sai lệch.\n\n${errors.slice(0, 12).join('\n')}`
    : 'Kiểm tra nhanh đạt: P=N+O, AK=P+V-X và cân đối Sheet 00 không có sai lệch.';

  SpreadsheetApp.getUi().alert(message);
}
