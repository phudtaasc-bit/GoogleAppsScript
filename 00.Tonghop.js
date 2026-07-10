function FS_lapSheet00() {
  const ss = SpreadsheetApp.getActive();

  const tech = ss.getSheetByName('01. Kỹ thuật');
  const sh02 = ss.getSheetByName('02. Doanh thu');
  const sh03 = ss.getSheetByName('03. Chi phí & Vốn');
  const sh04 = ss.getSheetByName('04. Dòng tiền & Lợi nhuận');

  if (!tech) throw new Error('Không tìm thấy sheet "01. Kỹ thuật".');
  if (!sh02) throw new Error('Không tìm thấy sheet "02. Doanh thu".');
  if (!sh03) throw new Error('Không tìm thấy sheet "03. Chi phí & Vốn".');
  if (!sh04) throw new Error('Không tìm thấy sheet "04. Dòng tiền & Lợi nhuận".');

  const dashboard = ss.getSheetByName('00. Dashboard');
  if (dashboard) ss.deleteSheet(dashboard);

  let sh = ss.getSheetByName('00. Tổng hợp');
  if (!sh) sh = ss.insertSheet('00. Tổng hợp', 0);

  ss.setActiveSheet(sh);
  ss.moveActiveSheet(1);

  sh.showRows(1, sh.getMaxRows());
  sh.showColumns(1, sh.getMaxColumns());
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart();
  sh.setFrozenRows(0);
  sh.setFrozenColumns(0);
  sh.clearContents();
  sh.clearFormats();

  const soThang = Number(FS00_getInfoValue_(tech, 'Số tháng mô hình')) || 40;
  const endRow04 = soThang + 2;

  const techName = `'01. Kỹ thuật'`;
  const s02 = `'02. Doanh thu'`;
  const s03 = `'03. Chi phí & Vốn'`;
  const s04 = `'04. Dòng tiền & Lợi nhuận'`;

  const equityRateA1 = FS00_getInfoCellA1_(tech, 'Tỷ suất chiết khấu');
  const loanRateA1 = FS00_getInfoCellA1_(tech, 'Lãi suất vay năm') || 'B13';
  if (!equityRateA1) {
    throw new Error('Thiếu "Tỷ suất chiết khấu" dùng cho FCFE tại 01. Kỹ thuật.');
  }

  const monthlyWacc = `((1+$E$21)^(1/12)-1)`;
  const monthlyCostOfEquity = `((1+${techName}!${equityRateA1})^(1/12)-1)`;

  const vatBase =
    `(SUM(${s03}!H3:H)+SUM(${s03}!K3:K)+SUM(${s03}!L3:L)+SUM(${s03}!M3:M))`;

  const vatAlloc = (col) =>
    `IFERROR(SUM(${s03}!O3:O)*SUM(${s03}!${col}3:${col})/${vatBase};0)`;

  sh.getRange('A1:E1').merge().setValue('BẢNG TỔNG HỢP PHÂN TÍCH HIỆU QUẢ ĐẦU TƯ');
  sh.getRange('A2:E2').merge().setFormula(`="DỰ ÁN: "&${techName}!B2`);

  sh.getRange('A4:E4').merge().setValue('I. TỔNG VỐN ĐẦU TƯ DỰ ÁN');
  sh.getRange('A5:E5').setValues([['TT', 'Nội dung', 'Sau VAT (tỷ đồng)', 'Tỷ lệ', 'Ghi chú']]);

  sh.getRange('A6:E13').setValues([
    ['I', 'Chi phí XD/TB/khác', '', '', ''],
    ['II', 'Chi phí GPMB', '', '', ''],
    ['III', 'Tiền SDĐ/thuê đất', '', '', ''],
    ['IV', 'Chi phí HTKT', '', '', ''],
    ['V', 'Chi phí dự phòng', '', '', ''],
    ['VI', 'Chi phí lãi vay', '', '', ''],
    ['VII', 'TỔNG VỐN ĐẦU TƯ', '', '', ''],
    ['', 'TỔNG VỐN ĐẦU TƯ không gồm tiền SDĐ', '', '', '']
  ]);

  sh.getRange('C6').setFormula(`=(SUM(${s03}!H3:H)+${vatAlloc('H')})/1000000000`);
  sh.getRange('C7').setFormula(`=SUM(${s03}!I3:I)/1000000000`);
  sh.getRange('C8').setFormula(`=SUM(${s03}!J3:J)/1000000000`);
  sh.getRange('C9').setFormula(`=(SUM(${s03}!K3:K)+${vatAlloc('K')})/1000000000`);
  sh.getRange('C10').setFormula(`=(SUM(${s03}!M3:M)+${vatAlloc('M')})/1000000000`);
  sh.getRange('C11').setFormula(`=SUM(${s04}!W3:W${endRow04})/1000000000`);
  sh.getRange('C12').setFormula('=SUM(C6:C11)');
  sh.getRange('C13').setFormula('=C12-C8');

  sh.getRange('D6:D11').setFormulaR1C1('=IF(R12C3=0;0;RC[-1]/R12C3)');
  sh.getRange('D12').setFormula('=100%');

  sh.getRange('A15:E15').merge().setValue('II. CƠ CẤU NGUỒN VỐN');
  sh.getRange('A16:E16').setValues([['TT', 'Nội dung', 'Giá trị (tỷ đồng)', 'Tỷ trọng (%)', 'Chi phí vốn']]);

  sh.getRange('A17:E21').setValues([
    ['1', 'Vốn chủ sở hữu', '', '', ''],
    ['2', 'Vốn vay ngân hàng', '', '', ''],
    ['3', 'Vốn huy động khách hàng', '', '', ''],
    ['', 'Tổng nguồn vốn', '', '', ''],
    ['', 'WACC', '', '', '']
  ]);

  sh.getRange('C17').setFormula(`=SUM(${s04}!S3:S${endRow04})/1000000000`);
  sh.getRange('C18').setFormula(`=SUM(${s04}!V3:V${endRow04})/1000000000`);
  sh.getRange('C19').setFormula('=MAX(0;C12-C17-C18)');
  sh.getRange('C20').setFormula('=SUM(C17:C19)');

  sh.getRange('D17:D19').setFormulaR1C1('=IF(R20C3=0;0;RC[-1]/R20C3)');
  sh.getRange('D20').setFormula('=100%');

  sh.getRange('E17').setFormula(`=${techName}!${equityRateA1}`);
  sh.getRange('E18').setFormula(`=${techName}!${loanRateA1}`);
  sh.getRange('E19').setValue(0);
  sh.getRange('E21').setFormula('=IF(C20=0;0;SUMPRODUCT(C17:C19;E17:E19)/C20)');

  sh.getRange('A23:E23').merge().setValue('III. ĐÁNH GIÁ HIỆU QUẢ ĐẦU TƯ');
  sh.getRange('A24:E24').setValues([['TT', 'Nội dung', 'Đơn vị', 'Giá trị', 'Ghi chú']]);

  sh.getRange('A25:E42').setValues([
    ['1', 'Tổng doanh thu có VAT', 'tỷ đồng', '', ''],
    ['1.1', 'Phần NOXH', 'tỷ đồng', '', ''],
    ['1.2', 'Phần thấp tầng / Liền kề', 'tỷ đồng', '', ''],
    ['1.3', 'Phần TMDV / Cho thuê', 'tỷ đồng', '', ''],
    ['2', 'Tổng chi phí có VAT', 'tỷ đồng', '', ''],
    ['2.1', 'Tổng vốn đầu tư dự án', 'tỷ đồng', '', ''],
    ['2.2', 'Chi phí bán hàng', 'tỷ đồng', '', ''],
    ['3', 'Lợi nhuận sau thuế', 'tỷ đồng', '', ''],
    ['4', 'NPV dự án', 'tỷ đồng', '', 'FCFF chiết khấu theo WACC'],
    ['5', 'IRR dự án', '%', '', 'IRR tháng quy đổi năm hiệu dụng'],
    ['6', 'Thời gian hoàn vốn dự án', 'tháng', '', 'Theo FCFF lũy kế'],
    ['7', 'NPV vốn CSH', 'tỷ đồng', '', 'FCFE chiết khấu theo chi phí vốn CSH'],
    ['8', 'IRR vốn CSH', '%', '', 'IRR tháng quy đổi năm hiệu dụng'],
    ['9', 'Thời gian hoàn vốn - Vốn CSH', 'tháng', '', 'Theo FCFE lũy kế'],
    ['10', 'Đỉnh dư nợ vay', 'tỷ đồng', '', ''],
    ['11', 'Tổng lãi vay', 'tỷ đồng', '', ''],
    ['12', 'Tổng Thuế TNDN', 'tỷ đồng', '', ''],
    ['13', 'Tổng VAT phải nộp', 'tỷ đồng', '', '']
  ]);

  sh.getRange('D25').setFormula(`=SUM(${s04}!G3:G${endRow04})/1000000000`);
  sh.getRange('D26').setFormula(`=SUMIF(${s02}!E:E;"*NOXH*";${s02}!T:T)/1000000000`);
  sh.getRange('D27').setFormula(`=SUMIF(${s02}!E:E;"*Liền kề*";${s02}!T:T)/1000000000`);
  sh.getRange('D28').setFormula(`=(SUMIF(${s02}!E:E;"*TMDV*";${s02}!T:T)+SUMIF(${s02}!F:F;"*Cho thuê*";${s02}!T:T))/1000000000`);

  sh.getRange('D29').setFormula('=D30+D31');
  sh.getRange('D30').setFormula('=C12');
  sh.getRange('D31').setFormula(`=(SUM(${s03}!L3:L)+${vatAlloc('L')})/1000000000`);
  sh.getRange('D32').setFormula(`=SUM(${s04}!O3:O${endRow04})/1000000000`);

  sh.getRange('D33').setFormula(`=IFERROR(NPV(${monthlyWacc};${s04}!AJ3:AJ${endRow04})/1000000000;0)`);
  sh.getRange('D34').setFormula(`=IFERROR((1+IRR(${s04}!AJ3:AJ${endRow04}))^12-1;0)`);
  sh.getRange('D35').setValue('');

  sh.getRange('D36').setFormula(`=IFERROR(NPV(${monthlyCostOfEquity};${s04}!AK3:AK${endRow04})/1000000000;0)`);
  sh.getRange('D37').setFormula(`=IFERROR((1+IRR(${s04}!AK3:AK${endRow04}))^12-1;0)`);
  sh.getRange('D38').setValue('');

  sh.getRange('D39').setFormula(`=MAX(${s04}!Y3:Y${endRow04})/1000000000`);
  sh.getRange('D40').setFormula(`=SUM(${s04}!W3:W${endRow04})/1000000000`);
  sh.getRange('D41').setFormula(`=SUM(${s04}!L3:L${endRow04})/1000000000`);
  sh.getRange('D42').setFormula(`=SUM(${s04}!K3:K${endRow04})/1000000000`);

  SpreadsheetApp.flush();

  sh.getRange('D35').setValue(FS00_calcPaybackMonths_(sh04, 38, 3, endRow04));
  sh.getRange('D38').setValue(FS00_calcPaybackMonths_(sh04, 39, 3, endRow04));

  SpreadsheetApp.flush();
  FS00_formatTongHop_(sh);
}

function FS00_calcPaybackMonths_(sheet, col, startRow, endRow) {
  const values = sheet.getRange(startRow, col, endRow - startRow + 1, 1)
    .getValues()
    .flat()
    .map(v => Number(v))
    .filter(v => !isNaN(v));

  for (let i = 1; i < values.length; i++) {
    const prev = values[i - 1];
    const cur = values[i];

    if (prev < 0 && cur >= 0) {
      const fraction = Math.abs(prev) / (Math.abs(prev) + cur);
      return i + fraction;
    }
  }

  return 0;
}

function FS00_formatTongHop_(sh) {
  sh.setFrozenRows(2);
  sh.setFrozenColumns(0);

  sh.getRange('A1:E42')
    .setFontFamily('Times New Roman')
    .setFontSize(11)
    .setVerticalAlignment('middle')
    .setWrap(true)
    .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);

  sh.getRange('A1:E1').setFontSize(14).setFontWeight('bold').setHorizontalAlignment('left');
  sh.getRange('A2:E2').setFontSize(12).setFontWeight('bold').setHorizontalAlignment('left');

  ['A4:E4', 'A15:E15', 'A23:E23'].forEach(r => {
    sh.getRange(r).setBackground('#FFC000').setFontWeight('bold').setHorizontalAlignment('left');
  });

  ['A5:E5', 'A16:E16', 'A24:E24'].forEach(r => {
    sh.getRange(r).setBackground('#A6A6A6').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');
  });

  sh.getRange('A6:A42').setHorizontalAlignment('center');
  sh.getRange('B6:B42').setHorizontalAlignment('left');
  sh.getRange('C6:D42').setHorizontalAlignment('center');
  sh.getRange('E6:E42').setHorizontalAlignment('center');

  sh.getRange('C6:C13').setNumberFormat('#,##0.0');
  sh.getRange('D6:D13').setNumberFormat('0.0%');

  sh.getRange('C17:C20').setNumberFormat('#,##0.0');
  sh.getRange('D17:D20').setNumberFormat('0.0%');
  sh.getRange('E17:E21').setNumberFormat('0.0%');

  sh.getRange('D25:D33').setNumberFormat('#,##0.0');
  sh.getRange('D34').setNumberFormat('0.0%');
  sh.getRange('D35').setNumberFormat('0.00');
  sh.getRange('D36').setNumberFormat('#,##0.0');
  sh.getRange('D37').setNumberFormat('0.0%');
  sh.getRange('D38').setNumberFormat('0.00');
  sh.getRange('D39:D42').setNumberFormat('#,##0.0');

  ['A12:E13', 'A20:E21', 'A25:E25', 'A29:E29', 'A32:E42'].forEach(r => {
    sh.getRange(r).setFontWeight('bold');
  });

  sh.getRange('A33:E35').setBackground('#FCE4D6').setFontWeight('bold').setFontColor('#FF0000');
  sh.getRange('A36:E38').setBackground('#E2F0D9').setFontWeight('bold').setFontColor('#FF0000');

  sh.getRange('C12:D13').setFontColor('#FF0000').setFontWeight('bold');
  sh.getRange('E21').setFontColor('#FF0000').setFontWeight('bold');

  sh.setColumnWidth(1, 45);
  sh.setColumnWidth(2, 285);
  sh.setColumnWidth(3, 110);
  sh.setColumnWidth(4, 110);
  sh.setColumnWidth(5, 170);

  for (let r = 1; r <= 42; r++) sh.setRowHeight(r, 24);
  sh.setRowHeight(1, 30);
}

function FS00_getInfoValue_(sheet, label) {
  const row = FS00_findRow_(sheet, label);
  return row ? sheet.getRange(row, 2).getValue() : '';
}

function FS00_getInfoCellA1_(sheet, label) {
  const row = FS00_findRow_(sheet, label);
  return row ? sheet.getRange(row, 2).getA1Notation() : '';
}

function FS00_findRow_(sheet, text) {
  const data = sheet.getDataRange().getDisplayValues();
  const target = FS00_norm_(text);
  for (let r = 0; r < data.length; r++) {
    if (data[r].some(v => FS00_norm_(v) === target)) return r + 1;
  }
  return null;
}

function FS00_norm_(v) {
  return String(v || '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd')
    .replace(/\s+/g, ' ')
    .trim();
}
