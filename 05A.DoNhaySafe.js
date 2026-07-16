const FS05_SAFE_CFG = Object.freeze({
  SOURCE_SHEETS: [
    '01A. Kỹ thuật',
    '02. Doanh thu',
    '03. Chi phí & Vốn',
    '03A. Lợi nhuận & Thuế',
    '04. Dòng tiền & Tài trợ',
    '00. Tổng hợp'
  ]
});

/**
 * Điểm vào duy nhất từ menu.
 * Sheet 05 chỉ được đọc dữ liệu nguồn và ghi kết quả vào chính Sheet 05.
 */
function FS05_lapBangDoNhay_AnToan() {
  const ss = SpreadsheetApp.getActive();
  const before = FS05_safeSnapshot_(ss);

  try {
    FS05_lapBangDoNhay();
  } finally {
    SpreadsheetApp.flush();
  }

  const after = FS05_safeSnapshot_(ss);
  const changed = FS05_safeCompare_(before, after);

  if (changed.length) {
    throw new Error(
      'Độ nhạy đã làm thay đổi sheet nguồn: ' + changed.join(', ') +
      '. Không sử dụng kết quả Sheet 05 cho đến khi chạy lại toàn bộ mô hình.'
    );
  }

  SpreadsheetApp.getUi().alert(
    'Đã hoàn thành bảng độ nhạy. Các sheet nguồn 01A, 02, 03, 03A, 04 và 00 không thay đổi.'
  );
}

function FS05_safeSnapshot_(ss) {
  const out = {};

  FS05_SAFE_CFG.SOURCE_SHEETS.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) throw new Error('Thiếu sheet nguồn "' + name + '".');

    const range = sheet.getDataRange();
    const payload = JSON.stringify({
      rows: range.getNumRows(),
      columns: range.getNumColumns(),
      values: range.getValues().map(row => row.map(FS05_safeSerializable_)),
      formulas: range.getFormulas()
    });

    const digest = Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256,
      payload,
      Utilities.Charset.UTF_8
    );

    out[name] = digest.map(byte => {
      const value = byte < 0 ? byte + 256 : byte;
      return ('0' + value.toString(16)).slice(-2);
    }).join('');
  });

  return out;
}

function FS05_safeCompare_(before, after) {
  return FS05_SAFE_CFG.SOURCE_SHEETS.filter(name => before[name] !== after[name]);
}

function FS05_safeSerializable_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return { type: 'date', value: value.getTime() };
  }
  if (typeof value === 'number') {
    return { type: 'number', value: isFinite(value) ? value : null };
  }
  if (typeof value === 'boolean') {
    return { type: 'boolean', value: value };
  }
  return { type: 'text', value: String(value == null ? '' : value) };
}
