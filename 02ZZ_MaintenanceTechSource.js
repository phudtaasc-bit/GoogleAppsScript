/*************************************************
 * 02ZZ_MaintenanceTechSource.js
 * Nguồn cấu hình bảo trì dùng trong mô hình là Sheet 01. Kỹ thuật.
 * Sheet 01. Đầu vào chỉ là nơi nhập liệu gốc.
 *************************************************/

FS02M_readConfigs_ = function(sheetIgnored) {
  const ss = SpreadsheetApp.getActive();
  const tech = ss.getSheetByName('01. Kỹ thuật');
  if (!tech) return {};

  const data = tech.getDataRange().getValues();
  let headerRow = -1;
  let productCol = -1;
  let maintenanceCol = -1;
  let leaseYearsCol = -1;

  for (let r = 0; r < data.length; r++) {
    const normalized = data[r].map(FS02M_key_);
    const p = normalized.indexOf(FS02M_key_('Loại sản phẩm'));
    const m = normalized.indexOf(FS02M_key_('Chi phí bảo trì'));
    if (p >= 0 && m >= 0) {
      headerRow = r;
      productCol = p;
      maintenanceCol = m;
      leaseYearsCol = normalized.indexOf(FS02M_key_('Thời gian thuê (năm)'));
      break;
    }
  }

  if (headerRow < 0) return {};

  const out = {};
  let blankCount = 0;

  for (let r = headerRow + 1; r < data.length; r++) {
    const product = String(data[r][productCol] || '').trim();
    if (!product) {
      blankCount++;
      if (blankCount >= 3) break;
      continue;
    }

    if (/^[A-Z_]+$/.test(product)) break;
    blankCount = 0;

    const text = String(data[r][maintenanceCol] || '').trim();
    const tiers = FS02M_parseTiers_(text);
    if (!tiers.length) continue;

    out[FS02M_key_(product)] = {
      tiers,
      leaseYears: leaseYearsCol >= 0 ? Number(data[r][leaseYearsCol]) || 0 : 0
    };
  }

  return out;
};
