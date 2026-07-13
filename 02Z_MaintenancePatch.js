/*************************************************
 * 02Z_MaintenancePatch.js
 * Bổ sung chi phí bảo trì cho sản phẩm cho thuê.
 *
 * Cột lõi Sheet 02 A:AF được giữ nguyên.
 * Cột bổ sung:
 * - AG: Tỷ lệ chi phí bảo trì
 * - AH: Chi phí vận hành thuê trước VAT
 * - AI: Chi phí bảo trì trước VAT
 *************************************************/

const FS02M_BASE_LAP_SHEET02_ = FS_lapSheet02;
const FS02M_BASE_READ_MONTHLY_ = FS03V21_docTongTheoThangTuSheet02_;
const FS02M_BASE_LAP_SHEET03_PATCHED_ = FS_lapSheet03_Patched;

FS_lapSheet02 = function() {
  FS02M_BASE_LAP_SHEET02_();
  FS02M_applyMaintenance_();
};

FS03V21_docTongTheoThangTuSheet02_ = function(sh02, soThang) {
  const out = FS02M_BASE_READ_MONTHLY_(sh02, soThang);
  const lastRow = sh02.getLastRow();
  if (lastRow < 2) return out;

  const width = Math.min(35, sh02.getLastColumn());
  const data = sh02.getRange(2, 1, lastRow - 1, width).getValues();
  data.forEach(r => {
    const monthNo = Number(r[0]) || 0;
    if (!monthNo || !out[monthNo]) return;

    const operating = Number(r[33]) || 0;   // AH
    const maintenance = Number(r[34]) || 0; // AI
    out[monthNo].chiVanHanh = (out[monthNo].chiVanHanh || 0) + operating;
    out[monthNo].chiBaoTri = (out[monthNo].chiBaoTri || 0) + maintenance;
  });

  return out;
};

FS_lapSheet03_Patched = function() {
  FS02M_BASE_LAP_SHEET03_PATCHED_();
  FS02M_applyOperationsToSheet03_();
};

function FS02M_applyMaintenance_() {
  const ss = SpreadsheetApp.getActive();
  const input = ss.getSheetByName('01. Đầu vào');
  const sh02 = ss.getSheetByName('02. Doanh thu');

  if (!input || !sh02 || sh02.getLastRow() < 2) return;

  const configs = FS02M_readConfigs_(input);
  const rowCount = sh02.getLastRow() - 1;
  const values = sh02.getRange(2, 1, rowCount, 32).getValues();

  const productBases = {};
  const rentStartMonth = {};

  values.forEach(r => {
    const product = FS02M_key_(r[4]);
    if (!product) return;

    const method = FS02M_key_(r[5]);
    const monthNo = Number(r[0]) || 0;
    const activeRate = Number(r[13]) || 0;

    if (method === FS02M_key_('Cho thuê') && activeRate > 0) {
      if (!rentStartMonth[product] || monthNo < rentStartMonth[product]) {
        rentStartMonth[product] = monthNo;
      }
    }

    const allocatedBase =
      (Number(r[21]) || 0) +
      (Number(r[24]) || 0) +
      (Number(r[25]) || 0) +
      (Number(r[26]) || 0) +
      (Number(r[27]) || 0);

    productBases[product] = (productBases[product] || 0) + allocatedBase;
  });

  const outX = [];
  const outAD = [];
  const outAE = [];
  const outAF = [];
  const outRate = [];
  const outOperating = [];
  const outMaintenance = [];

  values.forEach(r => {
    const product = FS02M_key_(r[4]);
    const method = FS02M_key_(r[5]);
    const monthNo = Number(r[0]) || 0;
    const activeRate = Number(r[13]) || 0;
    const currentOperatingCost = Number(r[23]) || 0;

    let maintenanceRate = 0;
    let maintenance = 0;
    const cfg = configs[product];
    const startMonth = rentStartMonth[product];

    if (
      cfg &&
      method === FS02M_key_('Cho thuê') &&
      activeRate > 0 &&
      startMonth &&
      monthNo >= startMonth
    ) {
      const operationYear = Math.floor((monthNo - startMonth) / 12) + 1;
      maintenanceRate = FS02M_rateForYear_(cfg.tiers, operationYear);
      maintenance = (productBases[product] || 0) * maintenanceRate / 12;
    }

    const operatingAndMaintenance = currentOperatingCost + maintenance;
    const oldTaxCost = Number(r[29]) || 0;
    const taxCost = oldTaxCost + maintenance;
    const revenue = Number(r[16]) || 0;
    const taxableProfit = Math.max(0, revenue - taxCost);
    const citRate = FS02M_rate_(r[20]);
    const cit = taxableProfit * citRate;

    outX.push([operatingAndMaintenance]);
    outAD.push([taxCost]);
    outAE.push([taxableProfit]);
    outAF.push([cit]);
    outRate.push([maintenanceRate]);
    outOperating.push([currentOperatingCost]);
    outMaintenance.push([maintenance]);
  });

  if (sh02.getMaxColumns() < 35) {
    sh02.insertColumnsAfter(sh02.getMaxColumns(), 35 - sh02.getMaxColumns());
  }

  sh02.getRange(2, 24, rowCount, 1).setValues(outX);
  sh02.getRange(2, 30, rowCount, 1).setValues(outAD);
  sh02.getRange(2, 31, rowCount, 1).setValues(outAE);
  sh02.getRange(2, 32, rowCount, 1).setValues(outAF);
  sh02.getRange(2, 33, rowCount, 1).setValues(outRate);
  sh02.getRange(2, 34, rowCount, 1).setValues(outOperating);
  sh02.getRange(2, 35, rowCount, 1).setValues(outMaintenance);

  sh02.getRange(1, 24).setValue('Chi phí vận hành & bảo trì thuê trước VAT');
  sh02.getRange(1, 33, 1, 3).setValues([[
    'Tỷ lệ chi phí bảo trì',
    'Chi phí vận hành thuê trước VAT',
    'Chi phí bảo trì trước VAT'
  ]]);

  sh02.getRange(1, 33, 1, 3)
    .setFontWeight('bold')
    .setBackground('#d9ead3')
    .setHorizontalAlignment('center')
    .setWrap(true);
  sh02.getRange(2, 33, rowCount, 1).setNumberFormat('0.00%');
  sh02.getRange(2, 34, rowCount, 2).setNumberFormat('#,##0');
  sh02.autoResizeColumns(33, 3);
}

function FS02M_applyOperationsToSheet03_() {
  const ss = SpreadsheetApp.getActive();
  const sh02 = ss.getSheetByName('02. Doanh thu');
  const sh03 = ss.getSheetByName('03. Chi phí & Vốn');
  if (!sh02 || !sh03 || sh03.getLastRow() < 2) return;

  const monthMap = {};
  if (sh02.getLastRow() >= 2 && sh02.getLastColumn() >= 35) {
    const data02 = sh02.getRange(2, 1, sh02.getLastRow() - 1, 35).getValues();
    data02.forEach(r => {
      const t = Number(r[0]) || 0;
      if (!t) return;
      if (!monthMap[t]) monthMap[t] = { operating: 0, maintenance: 0 };
      monthMap[t].operating += Number(r[33]) || 0;
      monthMap[t].maintenance += Number(r[34]) || 0;
    });
  }

  const rowCount = sh03.getLastRow() - 1;
  if (sh03.getMaxColumns() < 38) {
    sh03.insertColumnsAfter(sh03.getMaxColumns(), 38 - sh03.getMaxColumns());
  }

  const months = sh03.getRange(2, 1, rowCount, 1).getValues();
  const base = sh03.getRange(2, 14, rowCount, 7).getValues(); // N:T
  const outOperating = [];
  const outMaintenance = [];
  const outCombined = [];
  const outN = [];
  const outP = [];
  const outT = [];

  months.forEach((r, i) => {
    const t = Number(r[0]) || 0;
    const item = monthMap[t] || { operating: 0, maintenance: 0 };
    const combined = item.operating + item.maintenance;

    const totalBeforeVat = (Number(base[i][0]) || 0) + combined; // N
    const totalAfterVat = (Number(base[i][2]) || 0) + combined;  // P
    const cashBeforeFunding = (Number(base[i][6]) || 0) - combined; // T

    outOperating.push([item.operating]);
    outMaintenance.push([item.maintenance]);
    outCombined.push([combined]);
    outN.push([totalBeforeVat]);
    outP.push([totalAfterVat]);
    outT.push([cashBeforeFunding]);
  });

  sh03.getRange(2, 14, rowCount, 1).setValues(outN);
  sh03.getRange(2, 16, rowCount, 1).setValues(outP);
  sh03.getRange(2, 20, rowCount, 1).setValues(outT);
  sh03.getRange(2, 36, rowCount, 1).setValues(outOperating);
  sh03.getRange(2, 37, rowCount, 1).setValues(outMaintenance);
  sh03.getRange(2, 38, rowCount, 1).setValues(outCombined);

  sh03.getRange(1, 36, 1, 3).setValues([[
    'Chi phí vận hành thuê trước VAT',
    'Chi phí bảo trì trước VAT',
    'Tổng chi phí vận hành & bảo trì trước VAT'
  ]]);
  sh03.getRange(1, 36, 1, 3)
    .setFontWeight('bold')
    .setBackground('#d9ead3')
    .setHorizontalAlignment('center')
    .setWrap(true);
  sh03.getRange(2, 36, rowCount, 3).setNumberFormat('#,##0');
  sh03.autoResizeColumns(36, 3);
}

function FS02M_readConfigs_(sheet) {
  const values = sheet.getDataRange().getDisplayValues();
  let headerRow = -1;
  let productCol = -1;
  let maintenanceCol = -1;

  for (let r = 0; r < values.length; r++) {
    const normalized = values[r].map(FS02M_key_);
    const p = normalized.indexOf(FS02M_key_('Loại sản phẩm'));
    const m = normalized.indexOf(FS02M_key_('Chi phí bảo trì'));
    if (p >= 0 && m >= 0) {
      headerRow = r;
      productCol = p;
      maintenanceCol = m;
      break;
    }
  }

  if (headerRow < 0) return {};

  const out = {};
  let blankCount = 0;

  for (let r = headerRow + 1; r < values.length; r++) {
    const product = String(values[r][productCol] || '').trim();
    const text = String(values[r][maintenanceCol] || '').trim();

    if (!product) {
      blankCount++;
      if (blankCount >= 3) break;
      continue;
    }

    blankCount = 0;
    const tiers = FS02M_parseTiers_(text);
    if (tiers.length) out[FS02M_key_(product)] = { tiers };
  }

  return out;
}

function FS02M_parseTiers_(text) {
  const s = String(text || '').replace(/\r/g, '\n').replace(/;/g, '\n');
  const tiers = [];
  const rangeRegex = /năm\s*(\d+)\s*-\s*(\d+)\s*:\s*([0-9.,]+)\s*%/gi;
  let m;
  while ((m = rangeRegex.exec(s)) !== null) {
    tiers.push({ from: Number(m[1]), to: Number(m[2]), rate: FS02M_percentText_(m[3]) });
  }
  const onwardRegex = /năm\s*(\d+)\s*trở\s*đi\s*:\s*([0-9.,]+)\s*%/gi;
  while ((m = onwardRegex.exec(s)) !== null) {
    tiers.push({ from: Number(m[1]), to: Infinity, rate: FS02M_percentText_(m[2]) });
  }
  return tiers.sort((a, b) => a.from - b.from);
}

function FS02M_rateForYear_(tiers, yearNo) {
  for (const tier of tiers || []) {
    if (yearNo >= tier.from && yearNo <= tier.to) return tier.rate;
  }
  return 0;
}

function FS02M_percentText_(value) {
  const n = Number(String(value || '').replace(/\./g, '').replace(',', '.'));
  return isFinite(n) ? n / 100 : 0;
}

function FS02M_rate_(value) {
  const n = Number(value) || 0;
  return n > 1 ? n / 100 : n;
}

function FS02M_key_(value) {
  return String(value || '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd')
    .replace(/[^a-z0-9]/g, '');
}
