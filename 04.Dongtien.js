/*************************************************
 * 04_DongTien.gs
 * Sheet 04 - Dòng tiền & Lợi nhuận
 *
 * Nguyên tắc tài trợ:
 * - Tiền cuối kỳ trước được dùng trước khi huy động vốn mới.
 * - Khi thiếu tiền, huy động CSH và vốn vay theo tỷ lệ đầu vào.
 * - Lãi vay tính trên dư nợ đầu kỳ và được vốn hóa vào dư nợ.
 * - Khi có tiền dư, ưu tiên trả gốc; phần còn lại giữ sang kỳ sau.
 * - Chỉ kỳ cuối mô hình mới phân phối tiền còn lại cho CSH.
 * - FCFE phản ánh dòng tiền khả dụng cho CSH, kể cả khi tiền được giữ lại.
 *************************************************/

function FS_lapSheet04() {
  const ss = SpreadsheetApp.getActive();

  const tech = ss.getSheetByName('01. Kỹ thuật');
  const sh02 = ss.getSheetByName('02. Doanh thu');
  const sh03 = ss.getSheetByName('03. Chi phí & Vốn');
  let sh04 = ss.getSheetByName('04. Dòng tiền & Lợi nhuận');

  if (!tech) throw new Error('Không tìm thấy sheet "01. Kỹ thuật".');
  if (!sh02) throw new Error('Không tìm thấy sheet "02. Doanh thu".');
  if (!sh03) throw new Error('Không tìm thấy sheet "03. Chi phí & Vốn".');
  if (!sh04) sh04 = ss.insertSheet('04. Dòng tiền & Lợi nhuận');

  const soThang = Number(FS04_getInfoValue_(tech, 'Số tháng mô hình'));
  const startDateA1 = FS04_getInfoCellA1_(tech, 'Ngày bắt đầu dự án');
  const loanRateA1 = FS04_getInfoCellA1_(tech, 'Lãi suất vay năm');
  const loanRatioA1 = FS04_getInfoCellA1_(tech, 'Tỷ lệ vốn vay');

  if (!soThang || !startDateA1) {
    throw new Error('Thiếu "Số tháng mô hình" hoặc "Ngày bắt đầu dự án".');
  }

  if (!loanRateA1 || !loanRatioA1) {
    throw new Error('Thiếu "Lãi suất vay năm" hoặc "Tỷ lệ vốn vay".');
  }

  const endRow = soThang + 2;
  const lastCol = 39;

  if (sh04.getMaxRows() < endRow) {
    sh04.insertRowsAfter(sh04.getMaxRows(), endRow - sh04.getMaxRows());
  }

  if (sh04.getMaxColumns() < lastCol) {
    sh04.insertColumnsAfter(sh04.getMaxColumns(), lastCol - sh04.getMaxColumns());
  }

  const techName = `'01. Kỹ thuật'`;
  const s02 = `'02. Doanh thu'`;
  const s03 = `'03. Chi phí & Vốn'`;

  sh04.showRows(1, sh04.getMaxRows());
  sh04.showColumns(1, sh04.getMaxColumns());
  sh04.getRange(1, 1, sh04.getMaxRows(), sh04.getMaxColumns()).breakApart();
  sh04.clearContents();
  sh04.clearFormats();

  sh04.getRange(1, 1, 1, lastCol).setValues([[
    'THỜI GIAN','','','',
    'DOANH THU & CHI PHÍ','','','','',
    'THUẾ','','',
    'LỢI NHUẬN','','',
    'DÒNG TIỀN DỰ ÁN','','',
    'NGUỒN VỐN CSH','','',
    'KHOẢN VAY','','','',
    'KẾT QUẢ TIỀN','','',
    'LŨY KẾ','','','','','','',
    'CHỈ TIÊU NPV/IRR','','',''
  ]]);

  sh04.getRange(2, 1, 1, lastCol).setValues([[
    'Tháng số','Tháng','Năm','Quý',
    'Doanh thu trước VAT','VAT đầu ra','Dòng tiền huy động từ KH','Tổng chi trước VAT','Tổng chi sau VAT',
    'VAT đầu vào','VAT phải nộp','Thuế TNDN',
    'Tổng giá vốn tính thuế','Lợi nhuận chịu thuế','Lợi nhuận sau thuế',
    'Dòng tiền trước tài trợ','Nhu cầu vốn sau tiền đầu kỳ','Dòng tiền phân phối cho CSH',
    'CSH góp mới','CSH nộp lại từ tiền đã phân phối','Tổng dòng CSH vào dự án',
    'Giải ngân vay','Lãi vay vốn hóa','Trả gốc','Dư nợ cuối kỳ',
    'Tiền cuối kỳ','Tiền khả dụng sau trả nợ','Ghi chú',
    'Lũy kế dòng tiền KH','Lũy kế chi sau VAT','Lũy kế CSH góp mới','Lũy kế CSH nộp lại','Lũy kế giải ngân vay','Lũy kế Thuế TNDN','Lũy kế lợi nhuận sau thuế',
    'FCFF_TIPV','FCFE_EPV','FCFF lũy kế','FCFE lũy kế'
  ]]);

  sh04.getRange('A3').setFormula(`=MAKEARRAY(${soThang};1;LAMBDA(r;c;r))`);
  sh04.getRange('B3').setFormula(`=ARRAYFORMULA(EDATE(${techName}!${startDateA1};A3:A${endRow}-1))`);
  sh04.getRange('C3').setFormula(`=ARRAYFORMULA(YEAR(B3:B${endRow}))`);
  sh04.getRange('D3').setFormula(`=ARRAYFORMULA("Q"&ROUNDUP(MONTH(B3:B${endRow})/3;0)&"/"&YEAR(B3:B${endRow}))`);

  sh04.getRange('E3').setFormula(`=MAP(A3:A${endRow};LAMBDA(t;IF(t="";"";SUMIF(${s02}!A:A;t;${s02}!Q:Q))))`);
  sh04.getRange('F3').setFormula(`=MAP(A3:A${endRow};LAMBDA(t;IF(t="";"";SUMIF(${s03}!A:A;t;${s03}!F:F))))`);
  sh04.getRange('G3').setFormula(`=MAP(A3:A${endRow};LAMBDA(t;IF(t="";"";SUMIF(${s03}!A:A;t;${s03}!E:E))))`);
  sh04.getRange('H3').setFormula(`=MAP(A3:A${endRow};LAMBDA(t;IF(t="";"";SUMIF(${s03}!A:A;t;${s03}!N:N))))`);
  sh04.getRange('I3').setFormula(`=MAP(A3:A${endRow};LAMBDA(t;IF(t="";"";SUMIF(${s03}!A:A;t;${s03}!P:P))))`);
  sh04.getRange('J3').setFormula(`=MAP(A3:A${endRow};LAMBDA(t;IF(t="";"";SUMIF(${s03}!A:A;t;${s03}!O:O))))`);
  sh04.getRange('K3').setFormula(`=MAP(A3:A${endRow};LAMBDA(t;IF(t="";"";SUMIF(${s03}!A:A;t;${s03}!R:R))))`);
  sh04.getRange('L3').setFormula(`=MAP(A3:A${endRow};LAMBDA(t;IF(t="";"";SUMIF(${s02}!A:A;t;${s02}!AF:AF))))`);
  sh04.getRange('M3').setFormula(`=MAP(A3:A${endRow};LAMBDA(t;IF(t="";"";SUMIF(${s02}!A:A;t;${s02}!AD:AD))))`);
  sh04.getRange('N3').setFormula(`=MAP(A3:A${endRow};LAMBDA(t;IF(t="";"";SUMIF(${s02}!A:A;t;${s02}!AE:AE))))`);
  sh04.getRange('O3').setFormula(`=ARRAYFORMULA(N3:N${endRow}-L3:L${endRow})`);
  sh04.getRange('P3').setFormula(`=ARRAYFORMULA(G3:G${endRow}-I3:I${endRow}-K3:K${endRow}-L3:L${endRow})`);

  for (let r = 3; r <= endRow; r++) {
    const prevDebt = r === 3 ? '0' : `Y${r - 1}`;
    const prevCash = r === 3 ? '0' : `Z${r - 1}`;
    const cashBeforeFunding = `(${prevCash}+P${r})`;
    const debtBeforeRepayment = `(${prevDebt}+V${r}+W${r})`;

    // Lãi vay vốn hóa: tính trên dư nợ đầu kỳ.
    sh04.getRange(r, 23).setFormula(`=${prevDebt}*((1+${techName}!${loanRateA1})^(1/12)-1)`);

    // Chỉ huy động khi tiền đầu kỳ cộng dòng tiền hoạt động không đủ chi trả.
    sh04.getRange(r, 17).setFormula(`=MAX(0;-${cashBeforeFunding})`);
    sh04.getRange(r, 19).setFormula(`=Q${r}*(1-${techName}!${loanRatioA1})`);
    sh04.getRange(r, 20).setValue(0);
    sh04.getRange(r, 21).setFormula(`=S${r}+T${r}`);
    sh04.getRange(r, 22).setFormula(`=Q${r}*${techName}!${loanRatioA1}`);

    // Tiền dư hoạt động được dùng trả gốc. Không dùng vốn góp hoặc giải ngân mới để trả gốc ngay.
    sh04.getRange(r, 24).setFormula(`=MIN(${debtBeforeRepayment};MAX(0;${cashBeforeFunding}))`);
    sh04.getRange(r, 25).setFormula(`=MAX(0;${debtBeforeRepayment}-X${r})`);

    // Tiền còn lại sau trả gốc được giữ lại; chỉ phân phối tại kỳ cuối.
    sh04.getRange(r, 27).setFormula(`=MAX(0;${cashBeforeFunding}-X${r})`);
    sh04.getRange(r, 18).setFormula(r === endRow ? `=AA${r}` : '=0');
    sh04.getRange(r, 26).setFormula(`=MAX(0;AA${r}-R${r})`);

    sh04.getRange(r, 28).setFormula(
      `=IF(R${r}>0;"Phân phối cuối mô hình";IF(X${r}>0;"Trả gốc vay";IF(Q${r}>0;"Huy động mới theo tỷ lệ vốn";IF(Z${r}>0;"Giữ tiền sang kỳ sau";""))))`
    );

    sh04.getRange(r, 29).setFormula(`=SUM($G$3:G${r})`);
    sh04.getRange(r, 30).setFormula(`=SUM($I$3:I${r})`);
    sh04.getRange(r, 31).setFormula(`=SUM($S$3:S${r})`);
    sh04.getRange(r, 32).setFormula(`=SUM($T$3:T${r})`);
    sh04.getRange(r, 33).setFormula(`=SUM($V$3:V${r})`);
    sh04.getRange(r, 34).setFormula(`=SUM($L$3:L${r})`);
    sh04.getRange(r, 35).setFormula(`=SUM($O$3:O${r})`);

    // FCFF không phụ thuộc cấu trúc tài trợ.
    sh04.getRange(r, 36).setFormula(`=P${r}`);

    // FCFE = biến động tiền thuộc CSH sau trả nợ - vốn CSH góp mới.
    // Tiền giữ lại tăng là FCFE dương; sử dụng tiền giữ lại kỳ sau tạo FCFE âm tương ứng.
    sh04.getRange(r, 37).setFormula(`=AA${r}-${prevCash}-S${r}`);
    sh04.getRange(r, 38).setFormula(`=SUM($AJ$3:AJ${r})`);
    sh04.getRange(r, 39).setFormula(`=SUM($AK$3:AK${r})`);
  }

  SpreadsheetApp.flush();
  FS04_format_(sh04, endRow);
}

function FS04_format_(sh, endRow) {
  const lastCol = 39;
  const dataRows = endRow - 2;

  sh.setFrozenRows(2);
  sh.setFrozenColumns(4);

  sh.getRange(1, 1, endRow, lastCol)
    .setFontFamily('Arial')
    .setFontSize(10)
    .setVerticalAlignment('middle')
    .setWrap(true)
    .setBorder(true, true, true, true, true, true, '#999999', SpreadsheetApp.BorderStyle.SOLID);

  const groups = [
    [1, 4, 'THỜI GIAN', '#D9E1F2'],
    [5, 5, 'DOANH THU & CHI PHÍ', '#D9EAF7'],
    [10, 3, 'THUẾ', '#E2F0D9'],
    [13, 3, 'LỢI NHUẬN', '#EADCF8'],
    [16, 3, 'DÒNG TIỀN DỰ ÁN', '#B7DEE8'],
    [19, 3, 'NGUỒN VỐN CSH', '#FCE4D6'],
    [22, 4, 'KHOẢN VAY', '#EADCF8'],
    [26, 3, 'KẾT QUẢ TIỀN', '#D9EAD3'],
    [29, 7, 'LŨY KẾ', '#F2F2F2'],
    [36, 4, 'CHỈ TIÊU NPV/IRR', '#FFC000']
  ];

  groups.forEach(([col, width, title, color]) => {
    sh.getRange(1, col, 1, width)
      .breakApart()
      .merge()
      .setValue(title)
      .setBackground(color)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
  });

  groups.forEach(([col, width, , color]) => {
    sh.getRange(2, col, 1, width)
      .setBackground(color)
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
  });

  if (dataRows > 0) {
    sh.getRange(3, 1, dataRows, 4).setHorizontalAlignment('center');
    sh.getRange(3, 2, dataRows, 1).setNumberFormat('dd/mm/yyyy');
    sh.getRange(3, 5, dataRows, lastCol - 4).setNumberFormat('#,##0').setHorizontalAlignment('right');
  }

  sh.autoResizeColumns(1, lastCol);
  for (let c = 5; c <= lastCol; c++) sh.setColumnWidth(c, 120);
}

function FS04_getInfoValue_(sheet, label) {
  const row = FS04_findRow_(sheet, label);
  return row ? sheet.getRange(row, 2).getValue() : '';
}

function FS04_getInfoCellA1_(sheet, label) {
  const row = FS04_findRow_(sheet, label);
  return row ? sheet.getRange(row, 2).getA1Notation() : '';
}

function FS04_findRow_(sheet, text) {
  const data = sheet.getDataRange().getDisplayValues();
  const target = FS04_norm_(text);
  for (let r = 0; r < data.length; r++) {
    if (data[r].some(v => FS04_norm_(v) === target)) return r + 1;
  }
  return null;
}

function FS04_norm_(v) {
  return String(v || '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd')
    .replace(/\s+/g, ' ')
    .trim();
}
