function FS_lapSheet99() {
  const ss = SpreadsheetApp.getActive();
  const rows = [['Kiểm tra', 'Trạng thái', 'Sai lệch', 'Ghi chú']];
  const add = (name, a, b, note) => {
    const d = Math.abs(FS_num_(a) - FS_num_(b));
    rows.push([name, d <= 1 ? 'PASS' : 'FAIL', d, note || '']);
  };

  const r = ss.getSheetByName(FS_CFG.SHEETS.REVENUE);
  const c = ss.getSheetByName(FS_CFG.SHEETS.COST);
  const p = ss.getSheetByName(FS_CFG.SHEETS.PROFIT);
  const f = FS_getSheet_(ss, FS_CFG.SHEETS.CASH, FS_CFG.SHEETS.CASH_LEGACY);
  if (!r || !c || !p || !f) throw new Error('Thiếu sheet để kiểm tra.');

  const sum = (sh, col) => sh.getLastRow() < 2 ? 0 : FS_sum_(sh.getRange(2, col, sh.getLastRow() - 1, 1).getValues().flat());
  add('Doanh thu 02 ↔ 03A', sum(r, 16), sum(p, 8));
  add('Dòng tiền KH 02 ↔ 03', sum(r, 19), sum(c, 8));
  add('VAT đầu ra 02 ↔ 03', sum(r, 18), sum(c, 9));
  add('Tổng chi 03 ↔ 04', sum(c, 19), sum(f, 7));
  add('VAT đầu vào 03 ↔ 04', sum(c, 20), sum(f, 8));
  add('Thuế TNDN 03A ↔ 04', sum(p, 20), sum(f, 13));
  add('LNST 03A ↔ 04', sum(p, 21), sum(f, 14));
  add('Lãi vay 03A ↔ 04', sum(p, 16), sum(f, 17));

  const costTotal = sum(p, 15);
  const interestTotal = sum(p, 16);
  const taxTotal = sum(p, 20);
  const patTotal = sum(p, 21);
  const identity = sum(p, 8) - costTotal - interestTotal - taxTotal - patTotal;
  rows.push([
    'Cân đối lợi nhuận trước VAT', Math.abs(identity) <= 1 ? 'PASS' : 'FAIL', Math.abs(identity),
    'Doanh thu = chi phí tính thuế + lãi vay + thuế TNDN + LNST'
  ]);

  rows.push(['Tiền SDĐ tại Sheet 03', sum(c, 12) >= 0 ? 'PASS' : 'FAIL', '', 'Chi theo tiến độ chi phí đầu vào']);
  rows.push(['Tiền thuê đất tại Sheet 03', sum(c, 13) >= 0 ? 'PASS' : 'FAIL', '', 'Chi theo tiến độ chi phí đầu vào']);
  rows.push(['Chi phí vận hành', sum(c, 16) >= 0 ? 'PASS' : 'FAIL', '', 'Doanh thu thuê × tỷ lệ CPVH từng sản phẩm']);
  rows.push(['Chi phí bảo trì', sum(c, 17) >= 0 ? 'PASS' : 'FAIL', '', 'Doanh thu thuê × tỷ lệ bảo trì từng sản phẩm']);

  const state = JSON.parse(PropertiesService.getDocumentProperties().getProperty('FS_CONVERGENCE') || '{}');
  rows.push(['Hội tụ tài trợ', state.converged ? 'PASS' : 'FAIL', state.maxDiff || '', state.iterations ? 'Số vòng: ' + state.iterations : '']);

  const sh = FS_getOrCreateSheet_(ss, FS_CFG.SHEETS.CHECKS, FS_CFG.SHEETS.CHECKS_LEGACY);
  FS_resetSheet_(sh, rows.length + 3, 4);
  sh.getRange(1, 1).setValue('99. KIỂM TRA MÔ HÌNH');
  sh.getRange(3, 1, rows.length, 4).setValues(rows);
  sh.getRange(3, 1, 1, 4).setFontWeight('bold').setBackground('#d9e1f2');
  sh.getRange(4, 3, rows.length - 1, 1).setNumberFormat('#,##0');
  sh.autoResizeColumns(1, 4);
  return rows;
}
