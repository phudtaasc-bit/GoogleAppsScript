function FS05_DEBUG_FCFE_BASE_LIGHT() {
  const ss = SpreadsheetApp.getActive();
  const sh04 = ss.getSheetByName('04. Dòng tiền & Lợi nhuận');
  const tech = ss.getSheetByName('01. Kỹ thuật');

  let out = ss.getSheetByName('05D. Debug FCFE');
  if (!out) out = ss.insertSheet('05D. Debug FCFE');
  out.clearContents();

  const n = 12; // chỉ debug 12 tháng đầu
  const data = sh04.getRange(3, 1, n, 39).getValues();

  const loanRate = FS05D_percent_(FS05D_getInfoValue_(tech, 'Lãi suất vay năm'));
  const loanRatio = FS05D_percent_(FS05D_getInfoValue_(tech, 'Tỷ lệ vốn vay'));
  const repayStart = Number(FS05D_getInfoValue_(tech, 'Tháng bắt đầu trả gốc')) || 0;
  const monthlyLoanRate = Math.pow(1 + loanRate, 1 / 12) - 1;

  const rows = [[
    'Tháng',
    'P Sheet',
    'W Sheet', 'W Calc', 'ΔW',
    'T Sheet', 'T Calc', 'ΔT',
    'Q Sheet', 'Q Calc', 'ΔQ',
    'S Sheet', 'S Calc', 'ΔS',
    'V Sheet', 'V Calc', 'ΔV',
    'X Sheet', 'X Calc', 'ΔX',
    'Y Sheet', 'Y Calc', 'ΔY',
    'R Sheet', 'R Calc', 'ΔR',
    'AA Sheet', 'AA Calc', 'ΔAA',
    'AK Sheet', 'AK Calc', 'ΔAK'
  ]];

  let prevDebt = 0;
  let prevAA = 0;

  data.forEach(r => {
    const month = Number(r[0]) || 0;

    const P = Number(r[35]) || 0;     // AJ
    const sheetQ = Number(r[16]) || 0;
    const sheetR = Number(r[17]) || 0;
    const sheetS = Number(r[18]) || 0;
    const sheetT = Number(r[19]) || 0;
    const sheetV = Number(r[21]) || 0;
    const sheetW = Number(r[22]) || 0;
    const sheetX = Number(r[23]) || 0;
    const sheetY = Number(r[24]) || 0;
    const sheetAA = Number(r[26]) || 0;
    const sheetAK = Number(r[36]) || 0;

    const W = prevDebt * monthlyLoanRate;
    const cashAfterInterest = P - W;
    const T = Math.min(prevAA, Math.max(0, -cashAfterInterest));
    const Q = Math.max(0, -cashAfterInterest - T);
    const S = Q * (1 - loanRatio);
    const V = Q * loanRatio;
    const cashAfterFunding = cashAfterInterest + S + T + V;
    const X = month < repayStart ? 0 : Math.min(prevDebt + V, Math.max(0, cashAfterFunding));
    const Y = Math.max(0, prevDebt + V - X);
    const R = Math.max(0, cashAfterFunding - X);
    const AA = Math.max(0, prevAA + R - T);
    const AK = R - S - T;

    rows.push([
      month,
      P,
      sheetW, W, sheetW - W,
      sheetT, T, sheetT - T,
      sheetQ, Q, sheetQ - Q,
      sheetS, S, sheetS - S,
      sheetV, V, sheetV - V,
      sheetX, X, sheetX - X,
      sheetY, Y, sheetY - Y,
      sheetR, R, sheetR - R,
      sheetAA, AA, sheetAA - AA,
      sheetAK, AK, sheetAK - AK
    ]);

    prevDebt = Y;
    prevAA = AA;
  });

  out.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  out.getRange(1, 1, 1, rows[0].length).setFontWeight('bold');

  SpreadsheetApp.getUi().alert('Đã tạo debug nhẹ 12 tháng đầu.');
}
function FS05D_getInfoValue_(sheet, label) {
  const row = FS05D_findRow_(sheet, label);
  return row ? sheet.getRange(row, 2).getValue() : '';
}

function FS05D_findRow_(sheet, text) {
  const data = sheet.getDataRange().getDisplayValues();
  const target = FS05D_norm_(text);

  for (let r = 0; r < data.length; r++) {
    if (data[r].some(v => FS05D_norm_(v) === target)) return r + 1;
  }

  return null;
}

function FS05D_norm_(v) {
  return String(v || '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd')
    .replace(/\s+/g, ' ')
    .trim();
}

function FS05D_percent_(v) {
  if (typeof v === 'number') return v > 1 ? v / 100 : v;

  const n = Number(
    String(v || '')
      .replace('%', '')
      .replace(',', '.')
      .trim()
  );

  if (!isFinite(n)) return 0;

  return n > 1 ? n / 100 : n;
}