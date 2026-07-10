/*************************************************
 * 97A_LandCostGuard.gs
 * Chốt chặn không phá hủy trước khi lập lại Sheet 03/toàn mô hình.
 *************************************************/

function FS97_assertLandCostConfig_() {
  const ss = SpreadsheetApp.getActive();
  const tech = ss.getSheetByName('01. Kỹ thuật');
  if (!tech) throw new Error('Không tìm thấy sheet "01. Kỹ thuật".');

  const costs = FS97_readBlock_(tech, 'CHI_PHI_CHUNG');
  const schedules = FS97_readBlock_(tech, 'TIEN_DO_CHI_PHI');

  const landCosts = costs.rows
    .map((r, i) => ({
      row: costs.startRow + i,
      name: String(r[0] || '').trim(),
      value: FS97_num_(r[1]),
      category: FS97_landCategory_(r[0])
    }))
    .filter(x => x.category && Math.abs(x.value) > 0);

  const landSchedules = schedules.rows
    .map((r, i) => ({
      row: schedules.startRow + i,
      name: String(r[0] || '').trim(),
      rate: FS97_rate_(r[3]),
      category: FS97_landCategory_(r[0])
    }))
    .filter(x => x.category && x.rate > 0);

  const errors = [];
  const groupedCosts = FS97_groupBy_(landCosts, x => x.category);
  const groupedSchedules = FS97_groupBy_(landSchedules, x => x.category);

  Object.keys(groupedCosts).forEach(category => {
    if (groupedCosts[category].length > 1) {
      errors.push(
        `${FS97_label_(category)} có nhiều hơn một dòng chi phí nguồn: ` +
        groupedCosts[category].map(x => `dòng ${x.row}`).join(', ')
      );
    }
  });

  Object.keys(groupedSchedules).forEach(category => {
    const totalRate = groupedSchedules[category].reduce((s, x) => s + x.rate, 0);
    if (totalRate > 1.000001) {
      errors.push(`${FS97_label_(category)} có tổng tỷ lệ tiến độ ${(totalRate * 100).toFixed(2)}%, vượt 100%.`);
    }
  });

  const hasGeneralCost = Boolean(groupedCosts.GENERAL && groupedCosts.GENERAL.length);
  const hasSpecificCost = Object.keys(groupedCosts).some(k => k !== 'GENERAL' && groupedCosts[k].length);
  const hasGeneralSchedule = Boolean(groupedSchedules.GENERAL && groupedSchedules.GENERAL.length);
  const hasSpecificSchedule = Object.keys(groupedSchedules).some(k => k !== 'GENERAL' && groupedSchedules[k].length);

  if ((hasGeneralCost && hasSpecificCost) || (hasGeneralSchedule && hasSpecificSchedule)) {
    errors.push(
      'Không được đồng thời khai báo dòng tổng "Tiền SDĐ/thuê đất" và các dòng chi tiết theo Liền kề/Chung cư/TMDV/Chợ. ' +
      'Cấu hình này có nguy cơ tính lặp toàn bộ chi phí đất.'
    );
  }

  if (errors.length) {
    throw new Error(
      'Dừng chạy mô hình để tránh sai chi phí đất:\n- ' + errors.join('\n- ') +
      '\n\nHãy sửa cấu hình tại 01. Kỹ thuật rồi chạy lại.'
    );
  }

  return true;
}

function FS_lapSheet03_Safe() {
  FS97_assertLandCostConfig_();
  return FS_lapSheet03();
}

function FS_chayToanBoMoHinh_Safe() {
  FS97_assertLandCostConfig_();
  return FS_chayToanBoMoHinh_CoLaiVay();
}
