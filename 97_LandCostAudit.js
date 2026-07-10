/*************************************************
 * 97_LandCostAudit.gs
 * Kiểm soát không phá hủy đối với Tiền SDĐ / Tiền thuê đất.
 * Chỉ đọc dữ liệu và báo lỗi; không ghi đè mô hình.
 *************************************************/

function FS_AUDIT_TIEN_DAT() {
  const ss = SpreadsheetApp.getActive();
  const tech = ss.getSheetByName('01. Kỹ thuật');
  if (!tech) throw new Error('Không tìm thấy sheet "01. Kỹ thuật".');

  const costs = FS97_readBlock_(tech, 'CHI_PHI_CHUNG');
  const schedules = FS97_readBlock_(tech, 'TIEN_DO_CHI_PHI');
  const products = FS97_readBlock_(tech, 'SAN_PHAM');

  const issues = [];
  const warnings = [];

  const landCosts = costs.rows
    .map((r, i) => ({
      row: costs.startRow + i,
      name: String(r[0] || '').trim(),
      beforeVat: FS97_num_(r[1]),
      vatRate: FS97_rate_(r[2])
    }))
    .filter(x => FS97_landCategory_(x.name));

  const landSchedules = schedules.rows
    .map((r, i) => ({
      row: schedules.startRow + i,
      name: String(r[0] || '').trim(),
      start: FS97_num_(r[1]),
      duration: FS97_num_(r[2]),
      rate: FS97_rate_(r[3]),
      type: String(r[4] || '').trim(),
      category: FS97_landCategory_(r[0])
    }))
    .filter(x => x.category);

  const productCategories = new Set(
    products.rows
      .map(r => FS97_productLandCategory_(String(r[0] || ''), String(r[1] || '')))
      .filter(Boolean)
  );

  const groupedCosts = FS97_groupBy_(landCosts, x => FS97_landCategory_(x.name));
  Object.keys(groupedCosts).forEach(category => {
    const rows = groupedCosts[category];
    if (rows.length > 1) {
      issues.push(
        `${FS97_label_(category)} có ${rows.length} dòng chi phí nguồn: ` +
        rows.map(x => `dòng ${x.row} (${x.name})`).join(', ') +
        '. Mỗi nhóm chỉ được có một dòng giá trị nguồn.'
      );
    }
  });

  const groupedSchedules = FS97_groupBy_(landSchedules, x => x.category);
  Object.keys(groupedSchedules).forEach(category => {
    const rows = groupedSchedules[category];
    const totalRate = rows.reduce((s, x) => s + x.rate, 0);

    if (totalRate > 1.000001) {
      issues.push(`${FS97_label_(category)} có tổng tỷ lệ tiến độ ${(totalRate * 100).toFixed(2)}%, vượt 100%.`);
    } else if (totalRate < 0.999999) {
      warnings.push(`${FS97_label_(category)} có tổng tỷ lệ tiến độ ${(totalRate * 100).toFixed(2)}%, chưa đủ 100%.`);
    }

    rows.forEach(x => {
      if (x.start < 1) issues.push(`Dòng ${x.row}: tháng bắt đầu của ${x.name} phải từ 1 trở lên.`);
      if (x.duration < 1) issues.push(`Dòng ${x.row}: thời gian của ${x.name} phải từ 1 tháng trở lên.`);
      if (x.rate <= 0) issues.push(`Dòng ${x.row}: tỷ lệ của ${x.name} phải lớn hơn 0.`);
      if (!FS97_validScheduleType_(x.type)) {
        warnings.push(`Dòng ${x.row}: loại tiến độ "${x.type}" chưa thuộc danh mục Một lần/Phân bổ.`);
      }
    });
  });

  productCategories.forEach(category => {
    if (!groupedCosts[category] || !groupedCosts[category].length) {
      warnings.push(`Có sản phẩm thuộc ${FS97_label_(category)} nhưng chưa có dòng chi phí nguồn tương ứng.`);
    }
    if (!groupedSchedules[category] || !groupedSchedules[category].length) {
      warnings.push(`Có sản phẩm thuộc ${FS97_label_(category)} nhưng chưa có tiến độ chi phí tương ứng.`);
    }
  });

  if (groupedSchedules.GENERAL && Object.keys(groupedSchedules).some(k => k !== 'GENERAL')) {
    issues.push(
      'Đang đồng thời tồn tại tiến độ tổng "Tiền SDĐ/thuê đất" và tiến độ riêng theo sản phẩm. ' +
      'Cấu hình này có nguy cơ áp dụng lặp toàn bộ pool tiền đất.'
    );
  }

  const report = [];
  report.push(`KIỂM TRA TIỀN ĐẤT: ${issues.length} lỗi, ${warnings.length} cảnh báo.`);
  if (issues.length) report.push('\nLỖI:\n- ' + issues.join('\n- '));
  if (warnings.length) report.push('\nCẢNH BÁO:\n- ' + warnings.join('\n- '));
  if (!issues.length && !warnings.length) report.push('\nKhông phát hiện bất thường cấu hình tiền đất.');

  SpreadsheetApp.getUi().alert(report.join('\n'));
  return { issues, warnings };
}

function FS97_landCategory_(value) {
  const k = FS97_key_(value);
  if (!k) return '';
  if (k === FS97_key_('Tiền SDĐ/thuê đất')) return 'GENERAL';
  if (FS97_has_(k, ['tien sdd lien ke', 'tien su dung dat lien ke', 'tien sd dat lien ke'])) return 'SDD_LIEN_KE';
  if (FS97_has_(k, ['tien thue dat chung cu', 'tien thue dat can ho'])) return 'THUE_DAT_CHUNG_CU';
  if (FS97_has_(k, ['tien thue dat tmdv', 'tien thue dat thuong mai dich vu', 'tien thue dat tm dv'])) return 'THUE_DAT_TMDV';
  if (FS97_has_(k, ['tien thue dat cho'])) return 'THUE_DAT_CHO';
  return '';
}

function FS97_productLandCategory_(productName, method) {
  const k = FS97_key_(productName);
  if (FS97_has_(k, ['lien ke'])) return 'SDD_LIEN_KE';
  if (FS97_has_(k, ['chung cu', 'can ho'])) return 'THUE_DAT_CHUNG_CU';
  if (FS97_has_(k, ['tmdv', 'thuong mai', 'thuong mai dich vu'])) return 'THUE_DAT_TMDV';
  if (FS97_has_(k, ['cho'])) return 'THUE_DAT_CHO';
  return FS97_key_(method) === FS97_key_('Cho thuê') ? '' : '';
}

function FS97_label_(category) {
  return ({
    GENERAL: 'Tiền SDĐ/thuê đất',
    SDD_LIEN_KE: 'Tiền SDĐ - Liền kề',
    THUE_DAT_CHUNG_CU: 'Tiền thuê đất - Chung cư',
    THUE_DAT_TMDV: 'Tiền thuê đất - TMDV',
    THUE_DAT_CHO: 'Tiền thuê đất - Chợ'
  })[category] || category;
}

function FS97_readBlock_(sheet, blockName) {
  const data = sheet.getDataRange().getDisplayValues();
  const target = FS97_key_(blockName);
  let blockRow = 0;

  for (let r = 0; r < data.length; r++) {
    if (data[r].some(v => FS97_key_(v) === target)) {
      blockRow = r + 1;
      break;
    }
  }
  if (!blockRow) throw new Error('Không tìm thấy block ' + blockName + '.');

  const headerRow = blockRow + 1;
  const lastCol = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn())
    .getDisplayValues()[0]
    .reduce((last, v, i) => String(v || '').trim() ? i + 1 : last, 1);

  const rows = [];
  let blank = 0;
  for (let r = blockRow + 2; r <= sheet.getLastRow(); r++) {
    const first = String(sheet.getRange(r, 1).getDisplayValue() || '').trim();
    if (FS97_isBlock_(first)) break;
    const row = sheet.getRange(r, 1, 1, lastCol).getValues()[0];
    if (!row.some(v => String(v || '').trim() !== '')) {
      blank++;
      if (blank >= 3) break;
      continue;
    }
    blank = 0;
    rows.push(row);
  }

  return { blockRow, headerRow, startRow: blockRow + 2, rows };
}

function FS97_groupBy_(items, keyFn) {
  return items.reduce((out, item) => {
    const key = keyFn(item);
    if (!key) return out;
    if (!out[key]) out[key] = [];
    out[key].push(item);
    return out;
  }, {});
}

function FS97_validScheduleType_(value) {
  const k = FS97_key_(value);
  return !k || k === FS97_key_('Một lần') || k === FS97_key_('Phân bổ');
}

function FS97_has_(key, terms) {
  return terms.some(x => key.indexOf(FS97_key_(x)) >= 0);
}

function FS97_isBlock_(value) {
  return ['THONG_TIN_CHUNG', 'CHI_PHI_CHUNG', 'SAN_PHAM', 'KE_HOACH_BAN_THU_TIEN', 'TIEN_DO_CHI_PHI']
    .some(x => FS97_key_(value) === FS97_key_(x));
}

function FS97_num_(value) {
  if (typeof value === 'number') return isFinite(value) ? value : 0;
  const n = Number(String(value || '').replace(/\s/g, '').replace(/,/g, '.').replace(/%/g, ''));
  return isFinite(n) ? n : 0;
}

function FS97_rate_(value) {
  const n = FS97_num_(value);
  return n > 1 ? n / 100 : n;
}

function FS97_key_(value) {
  return String(value || '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd')
    .replace(/[^a-z0-9]+/g, ' ')
    .trim();
}
