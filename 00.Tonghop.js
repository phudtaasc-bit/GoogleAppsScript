function FS_lapSheet00() {
  const ss = SpreadsheetApp.getActive();
  const r = ss.getSheetByName(FS_CFG.SHEETS.REVENUE);
  const c = ss.getSheetByName(FS_CFG.SHEETS.COST);
  const p = ss.getSheetByName(FS_CFG.SHEETS.PROFIT);
  const f = FS_getSheet_(ss, FS_CFG.SHEETS.CASH, FS_CFG.SHEETS.CASH_LEGACY);
  if (!r || !c || !p || !f) throw new Error('Thiếu sheet nguồn tổng hợp.');

  const sum = (sh, col) => sh.getLastRow() < 2 ? 0 : FS_sum_(sh.getRange(2, col, sh.getLastRow() - 1, 1).getValues().flat());
  const values = [
    ['CHỈ TIÊU', 'GIÁ TRỊ (tỷ đồng)'],
    ['Doanh thu trước VAT', sum(r, 16) / 1e9],
    ['VAT đầu ra', sum(r, 18) / 1e9],
    ['Tổng chi trước VAT', sum(c, 18) / 1e9],
    ['VAT đầu vào', sum(c, 19) / 1e9],
    ['VAT phải nộp', sum(f, 11) / 1e9],
    ['Lãi vay', sum(f, 17) / 1e9],
    ['Thuế TNDN', sum(p, 18) / 1e9],
    ['Lợi nhuận sau thuế', sum(p, 19) / 1e9],
    ['Vốn góp CSH', sum(f, 18) / 1e9],
    ['Giải ngân vay', sum(f, 19) / 1e9],
    ['Trả gốc', sum(f, 20) / 1e9],
    ['FCFF', sum(f, 23) / 1e9],
    ['FCFE', sum(f, 24) / 1e9]
  ];

  const sh = FS_getOrCreateSheet_(ss, FS_CFG.SHEETS.SUMMARY);
  FS_resetSheet_(sh, 30, 5);
  sh.getRange('A1:E1').merge().setValue('BẢNG TỔNG HỢP PHÂN TÍCH HIỆU QUẢ ĐẦU TƯ')
    .setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange(3, 1, values.length, 2).setValues(values);
  sh.getRange(3, 1, 1, 2).setFontWeight('bold').setBackground('#d9e1f2');
  sh.getRange(4, 2, values.length - 1, 1).setNumberFormat('#,##0.000');
  sh.autoResizeColumns(1, 5);
}
