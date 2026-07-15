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
  add('Thuế TNDN 03A ↔ 04', sum(p, 18), sum(f, 13));
  add('LNST 03A ↔ 04', sum(p, 19), sum(f, 14));
  add('Lãi vay 03A ↔ 04', sum(p, 14), sum(f, 17));

  const identity = sum(p, 8) - sum(p, 13) - sum(p, 14) - sum(p, 18) - sum(p, 19);
  rows.push([
    'Cân đối lợi nhuận trước VAT', Math.abs(identity) <= 1 ? 'PASS' : 'FAIL', Math.abs(identity),
    'Doanh thu = chi phí tính thuế + lãi vay + thuế TNDN + LNST'
  ]);

  const state = JSON.parse(PropertiesService.getDocumentProperties().getProperty('FS_CONVERGENCE') || '{}');
  rows.push(['Hội tụ tài trợ', state.converged ? 'PASS' : 'FAIL', state.maxDiff || '', state.iterations ? 'Số vòng: ' + state.iterations : '']);

  const sh = FS_getOrCreateSheet_(ss, FS_CFG.SHEETS.CHECKS, FS_CFG.SHEETS.CHECKS_LEGACY);
  FS_resetSheet_(sh, rows.length + 3, 4);
  sh.getRange(1, 1).setValue('99. KIỂM TRA MÔ HÌNH');
  sh.getRange(3, 1, rows.length, 4).setValues(rows);
  sh.getRange(3, 1, 1, 4).setFontWeight('bold').setBackground('#d9e1f2');
  sh.autoResizeColumns(1, 4);
  return rows;
}
