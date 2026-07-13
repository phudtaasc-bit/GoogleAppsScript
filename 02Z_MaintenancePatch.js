/*************************************************
 * 02Z_MaintenancePatch.js
 * Bổ sung chi phí bảo trì cho sản phẩm cho thuê mà không thay đổi
 * cấu trúc 32 cột của Sheet 02.
 *
 * Nguyên tắc:
 * - Đọc cấu hình tại cột "Chi phí bảo trì" của mục D. CHI TIẾT SẢN PHẨM.
 * - Tỷ lệ bảo trì áp dụng theo năm vận hành kể từ tháng bắt đầu phát sinh thuê.
 * - Cơ sở "giá vốn" là tổng nguyên giá phân bổ của sản phẩm:
 *   CPXD + GPMB + HTKT + tiền đất + dự phòng; không gồm lãi vay.
 * - Chi phí bảo trì theo năm được phân bổ đều 12 tháng.
 * - Gộp vào cột X "Chi phí vận hành thuê trước VAT" để giữ nguyên layout.
 * - Đồng thời cập nhật AD, AE, AF để Thuế TNDN phản ánh chi phí bảo trì.
 *************************************************/

const FS02M_BASE_LAP_SHEET02_ = FS_lapSheet02;

FS_lapSheet02 = function() {
  FS02M_BASE_LAP_SHEET02_();
  FS02M_applyMaintenance_();
};

function FS02M_applyMaintenance_() {
  const ss = SpreadsheetApp.getActive();
  const input = ss.getSheetByName('01. Đầu vào');
  const sh02 = ss.getSheetByName('02. Doanh thu');

  if (!input || !sh02 || sh02.getLastRow() < 2) return;

  const configs = FS02M_readConfigs_(input);
  if (!Object.keys(configs).length) return;

  const rowCount = sh02.getLastRow() - 1;
  const values = sh02.getRange(2, 1, rowCount, 32).getValues();

  const productBases = {};
  const rentStartMonth = {};

  values.forEach(r => {
    const product = FS02M_key_(r[4]);
    if (!product || !configs[product]) return;

    const method = FS02M_key_(r[5]);
    const monthNo = Number(r[0]) || 0;
    const activeRate = Number(r[13]) || 0;

    if (method === FS02M_key_('Cho thuê') && activeRate > 0) {
      if (!rentStartMonth[product] || monthNo < rentStartMonth[product]) {
        rentStartMonth[product] = monthNo;
      }
    }

    // V + Y + Z + AA + AB: CPXD, GPMB, HTKT, tiền đất, dự phòng.
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

  values.forEach(r => {
    const product = FS02M_key_(r[4]);
    const method = FS02M_key_(r[5]);
    const monthNo = Number(r[0]) || 0;
    const activeRate = Number(r[13]) || 0;
    const currentOperatingCost = Number(r[23]) || 0; // X

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
      const annualRate = FS02M_rateForYear_(cfg.tiers, operationYear);
      maintenance = (productBases[product] || 0) * annualRate / 12;
    }

    const operatingAndMaintenance = currentOperatingCost + maintenance;
    const oldTaxCost = Number(r[29]) || 0; // AD
    const taxCost = oldTaxCost + maintenance;
    const revenue = Number(r[16]) || 0;    // Q
    const taxableProfit = Math.max(0, revenue - taxCost);
    const citRate = FS02M_rate_(r[20]);    // U
    const cit = taxableProfit * citRate;

    outX.push([operatingAndMaintenance]);
    outAD.push([taxCost]);
    outAE.push([taxableProfit]);
    outAF.push([cit]);
  });

  sh02.getRange(2, 24, rowCount, 1).setValues(outX);  // X
  sh02.getRange(2, 30, rowCount, 1).setValues(outAD); // AD
  sh02.getRange(2, 31, rowCount, 1).setValues(outAE); // AE
  sh02.getRange(2, 32, rowCount, 1).setValues(outAF); // AF
  sh02.getRange(1, 24).setValue('Chi phí vận hành & bảo trì thuê trước VAT');
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
  const s = String(text || '')
    .replace(/\r/g, '\n')
    .replace(/;/g, '\n');
  const tiers = [];

  // Dạng: Năm 1-3: 0,2% giá vốn
  const rangeRegex = /năm\s*(\d+)\s*-\s*(\d+)\s*:\s*([0-9.,]+)\s*%/gi;
  let m;
  while ((m = rangeRegex.exec(s)) !== null) {
    tiers.push({
      from: Number(m[1]),
      to: Number(m[2]),
      rate: FS02M_percentText_(m[3])
    });
  }

  // Dạng: Năm 30 trở đi: 2,5% giá vốn
  const onwardRegex = /năm\s*(\d+)\s*trở\s*đi\s*:\s*([0-9.,]+)\s*%/gi;
  while ((m = onwardRegex.exec(s)) !== null) {
    tiers.push({
      from: Number(m[1]),
      to: Infinity,
      rate: FS02M_percentText_(m[2])
    });
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
