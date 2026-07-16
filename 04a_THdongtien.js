function FS_lapSheet04A() {
  const ss = SpreadsheetApp.getActive();
  const src = FS_getSheet_(ss, FS_CFG.SHEETS.CASH, FS_CFG.SHEETS.CASH_LEGACY);
  if (!src) throw new Error('Thiếu Sheet 04. Dòng tiền & Tài trợ.');

  const data = src.getLastRow() > 1 ? src.getRange(2, 1, src.getLastRow() - 1, 24).getValues() : [];
  const map = {};
  data.forEach(r => {
    const y = FS_num_(r[2]);
    if (!map[y]) map[y] = Array(20).fill(0);
    for (let c = 4; c <= 23; c++) map[y][c - 4] += FS_num_(r[c]);
  });

  const years = Object.keys(map).sort();
  const rows = years.map(y => [FS_num_(y)].concat(map[y]));
  const sh = FS_getOrCreateSheet_(ss, FS_CFG.SHEETS.SUMMARY_CASH, FS_CFG.SHEETS.SUMMARY_CASH_LEGACY);
  FS_resetSheet_(sh, rows.length + 1, 21);
  sh.getRange(1, 1, 1, 21).setValues([[
    'Năm', 'Dòng tiền thu khách hàng', 'VAT đầu ra', 'Tổng chi trước VAT', 'VAT đầu vào',
    'Tổng chi sau VAT', 'VAT khấu trừ đầu kỳ', 'VAT phải nộp', 'VAT khấu trừ cuối kỳ',
    'Thuế TNDN', 'LNST', 'Dòng tiền trước tài trợ', 'Nhu cầu vốn', 'Lãi vay',
    'Vốn góp CSH', 'Giải ngân vay', 'Trả gốc', 'Dư nợ cuối kỳ cộng dồn',
    'Tiền cuối kỳ cộng dồn', 'FCFF', 'FCFE'
  ]]);
  if (rows.length) sh.getRange(2, 1, rows.length, 21).setValues(rows);
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, 21).setFontWeight('bold').setBackground('#fff2cc').setWrap(true);
  if (rows.length) sh.getRange(2, 2, rows.length, 20).setNumberFormat('#,##0');
  sh.autoResizeColumns(1, 21);
}
