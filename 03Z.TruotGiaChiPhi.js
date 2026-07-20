const FS03_ESC_CFG = Object.freeze({
  TECH: '01A. Kỹ thuật',
  RATE_LABEL: 'Tỷ lệ trượt chi phí/năm',
  ELIGIBLE_ITEMS: Object.freeze({
    construction: true,
    infrastructure: true
  })
});

/**
 * Ghi đè hàm đọc chi phí theo tiến độ của 03.Chiphi.js.
 *
 * Chỉ áp dụng trượt giá cho:
 * - construction: Chi phí XD/TB/khác.
 * - infrastructure: Chi phí HTKT.
 *
 * Không áp dụng trực tiếp cho:
 * - GPMB.
 * - Tiền SDĐ.
 * - Tiền thuê đất.
 * - Chi phí bán hàng.
 * - Chi phí vận hành.
 * - Chi phí bảo trì.
 * - Chi phí dự phòng.
 *
 * Công thức hệ số tại tháng số t:
 * (1 + tỷ lệ trượt chi phí/năm)^((t - 1) / 12)
 */
function FS03_scheduled_(index, itemKey, monthNo) {
  const baseValue = FS03_num_(index[itemKey + '|' + monthNo]);

  if (!baseValue || !FS03_ESC_CFG.ELIGIBLE_ITEMS[itemKey]) {
    return baseValue;
  }

  const annualEscalationRate = FS03_ESC_readAnnualRate_();

  if (!(annualEscalationRate > 0)) {
    return baseValue;
  }

  const factor = Math.pow(
    1 + annualEscalationRate,
    Math.max(0, FS03_num_(monthNo) - 1) / 12
  );

  return baseValue * factor;
}

/**
 * Đọc và cache tỷ lệ trượt chi phí/năm từ 01A. Kỹ thuật.
 * Cache chỉ tồn tại trong một lần thực thi Apps Script.
 */
function FS03_ESC_readAnnualRate_() {
  if (FS03_ESC_readAnnualRate_.cachedValue != null) {
    return FS03_ESC_readAnnualRate_.cachedValue;
  }

  const ss = SpreadsheetApp.getActive();
  const tech = ss.getSheetByName(FS03_ESC_CFG.TECH);

  if (!tech) {
    throw new Error('Không tìm thấy sheet "01A. Kỹ thuật" để đọc tỷ lệ trượt chi phí.');
  }

  const rawValue = FS03_readInfoValue_(
    tech,
    FS03_ESC_CFG.RATE_LABEL
  );

  const rate = FS03_rate_(rawValue);

  if (rate < 0) {
    throw new Error('Tỷ lệ trượt chi phí/năm không được âm.');
  }

  FS03_ESC_readAnnualRate_.cachedValue = rate;
  return rate;
}
