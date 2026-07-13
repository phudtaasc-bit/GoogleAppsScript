/*************************************************
 * 01Z_MaintenanceTechBridge.js
 * Giữ nguyên Sheet 01. Đầu vào; chuẩn hóa thông tin bảo trì sang Sheet 01. Kỹ thuật.
 *************************************************/

const FS01M_BASE_BUILD_TECH_ = FS_taoKyThuatTuDauVao;

FS_taoKyThuatTuDauVao = function() {
  FS01M_BASE_BUILD_TECH_();
  FS01M_writeMaintenanceToTech_();
};

function FS01M_writeMaintenanceToTech_() {
  const ss = SpreadsheetApp.getActive();
  const input = ss.getSheetByName('01. Đầu vào');
  const tech = ss.getSheetByName('01. Kỹ thuật');
  if (!input || !tech) return;

  const table = FS_getTable_(input, 'D. CHI TIẾT SẢN PHẨM', 'Loại sản phẩm');
  if (!table) return;

  const blockRow = FS_findRowContains_(tech, 'SAN_PHAM');
  if (!blockRow) return;

  const headerRow = blockRow + 1;
  const dataStartRow = blockRow + 2;

  const sourceByProduct = {};
  table.rows.forEach(row => {
    const product = FS_getByHeaderAny_(row, table.headers, ['Loại sản phẩm', 'Sản phẩm']);
    if (!product) return;

    const maintenanceText = FS_getByHeaderAny_(row, table.headers, [
      'Chi phí bảo trì',
      'Chi phi bao tri',
      'Bảo trì'
    ]);

    const leaseYears = FS_getByHeaderAny_(row, table.headers, [
      'Thời gian thuê (năm)',
      'Thời gian thuê',
      'Số năm thuê'
    ]);

    sourceByProduct[FS01M_key_(product)] = {
      maintenanceText: String(maintenanceText || '').trim(),
      leaseYears: Number(leaseYears) || 0
    };
  });

  const existingHeaders = tech.getRange(headerRow, 1, 1, Math.max(12, tech.getLastColumn())).getDisplayValues()[0];
  const productCol = existingHeaders.findIndex(h => FS01M_key_(h) === FS01M_key_('Loại sản phẩm')) + 1;
  if (!productCol) return;

  const headers = [
    'Chi phí bảo trì',
    'BT năm 1-3',
    'BT năm 4-10',
    'BT năm 11-20',
    'BT năm 21-30',
    'BT từ năm 30',
    'Thời gian thuê (năm)'
  ];

  const startCol = 13; // M:S, không làm thay đổi A:L đang được các module cũ sử dụng.
  tech.getRange(headerRow, startCol, 1, headers.length).setValues([headers]);

  const rows = [];
  for (let r = dataStartRow; r <= tech.getLastRow(); r++) {
    const product = String(tech.getRange(r, productCol).getDisplayValue() || '').trim();
    if (!product || /^[A-Z_]+$/.test(product)) break;

    const cfg = sourceByProduct[FS01M_key_(product)] || { maintenanceText: '', leaseYears: 0 };
    const tiers = typeof FS02M_parseTiers_ === 'function' ? FS02M_parseTiers_(cfg.maintenanceText) : [];

    rows.push([
      cfg.maintenanceText,
      FS01M_rateAt_(tiers, 1),
      FS01M_rateAt_(tiers, 4),
      FS01M_rateAt_(tiers, 11),
      FS01M_rateAt_(tiers, 21),
      FS01M_rateAt_(tiers, 30),
      cfg.leaseYears
    ]);
  }

  if (!rows.length) return;

  tech.getRange(dataStartRow, startCol, rows.length, headers.length).setValues(rows);
  tech.getRange(dataStartRow, startCol + 1, rows.length, 5).setNumberFormat('0.00%');
  tech.getRange(dataStartRow, startCol + 6, rows.length, 1).setNumberFormat('0');
  tech.getRange(headerRow, startCol, 1, headers.length).setFontWeight('bold');
  tech.autoResizeColumns(startCol, headers.length);
}

function FS01M_rateAt_(tiers, yearNo) {
  for (const tier of tiers || []) {
    if (yearNo >= tier.from && yearNo <= tier.to) return Number(tier.rate) || 0;
  }
  return 0;
}

function FS01M_key_(value) {
  return String(value || '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd')
    .replace(/[^a-z0-9]/g, '');
}
