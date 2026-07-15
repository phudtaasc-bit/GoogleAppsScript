/*************************************************
 * ZZZZZZZZZZZZ_AuditCrossSheetDetail.js
 * Bổ sung audit đối chiếu chi tiết giữa Sheet 02, 03 và 04.
 * Không thay đổi logic tính toán; chỉ đọc và ghi thêm kết quả vào 99. Checks.
 *************************************************/

const FS99X_BASE_AUDIT_ = FS99Q_chayAuditReconciliation;

FS99Q_chayAuditReconciliation = function() {
  FS99X_BASE_AUDIT_();
  FS99X_appendCrossSheetChecks_();
};

function FS99X_appendCrossSheetChecks_() {
  const ss = SpreadsheetApp.getActive();
  const sh99 = ss.getSheetByName('99. Checks');
  if (!sh99) throw new Error('Không tìm thấy Sheet 99. Checks.');

  const definitions = [
    {
      name: 'Doanh thu / Dòng tiền khách hàng',
      s02: ['Dòng tiền huy động từ KH', 'Tổng dòng tiền huy động từ KH', 'Tổng doanh thu có VAT'],
      s03: ['Dòng tiền huy động từ KH', 'Dòng tiền khách hàng'],
      s04: ['Dòng tiền huy động từ KH', 'Dòng tiền khách hàng']
    },
    {
      name: 'Chi phí XD/TB',
      s02: ['CP XD/TB trực tiếp trước VAT', 'Chi phí XD/TB/khác trước VAT'],
      s03: ['Chi XD/TB/khác trước VAT', 'Chi phí XD/TB/khác trước VAT', 'CP XD/TB trực tiếp trước VAT'],
      s04: ['Chi XD/TB/khác trước VAT', 'Chi phí XD/TB/khác trước VAT']
    },
    {
      name: 'Chi phí GPMB',
      s02: ['Chi phí GPMB phân bổ trước VAT', 'Chi phí GPMB trước VAT'],
      s03: ['Chi GPMB trước VAT', 'Chi phí GPMB trước VAT'],
      s04: ['Chi GPMB trước VAT', 'Chi phí GPMB trước VAT']
    },
    {
      name: 'Chi phí HTKT',
      s02: ['Chi phí HTKT phân bổ trước VAT', 'Chi phí HTKT trước VAT'],
      s03: ['Chi HTKT trước VAT', 'Chi phí HTKT trước VAT'],
      s04: ['Chi HTKT trước VAT', 'Chi phí HTKT trước VAT']
    },
    {
      name: 'Tiền SDĐ/thuê đất',
      s02: ['Tiền SDĐ/thuê đất phân bổ trước VAT', 'Tiền SDĐ/thuê đất trước VAT'],
      s03: ['Tiền SDĐ/thuê đất trước VAT', 'Tiền SDĐ/thuê đất'],
      s04: ['Tiền SDĐ/thuê đất trước VAT', 'Tiền SDĐ/thuê đất']
    },
    {
      name: 'Chi phí bán hàng',
      s02: ['Chi phí bán hàng trước VAT'],
      s03: ['Chi phí bán hàng trước VAT', 'Chi bán hàng trước VAT'],
      s04: ['Chi phí bán hàng trước VAT', 'Chi bán hàng trước VAT']
    },
    {
      name: 'Chi phí vận hành',
      s02: ['Chi phí vận hành thuê trước VAT', 'Chi phí vận hành trước VAT'],
      s03: ['Chi phí vận hành thuê trước VAT', 'Chi phí vận hành trước VAT'],
      s04: ['Chi phí vận hành thuê trước VAT', 'Chi phí vận hành trước VAT']
    },
    {
      name: 'Chi phí bảo trì',
      s02: ['Chi phí bảo trì trước VAT'],
      s03: ['Chi phí bảo trì trước VAT'],
      s04: ['Chi phí bảo trì trước VAT']
    },
    {
      name: 'Chi phí dự phòng',
      s02: ['Chi phí dự phòng phân bổ trước VAT', 'Chi phí dự phòng trước VAT'],
      s03: ['Chi phí dự phòng trước VAT', 'Chi dự phòng trước VAT'],
      s04: ['Chi phí dự phòng trước VAT', 'Chi dự phòng trước VAT']
    },
    {
      name: 'Lãi vay',
      s02: ['Chi phí lãi vay phân bổ', 'Lãi vay phân bổ'],
      s03: ['Chi phí lãi vay', 'Lãi vay', 'Lãi vay vốn hóa'],
      s04: ['Lãi vay vốn hóa', 'Lãi vay']
    },
    {
      name: 'VAT phải nộp',
      s02: ['VAT phải nộp'],
      s03: ['VAT phải nộp'],
      s04: ['VAT phải nộp']
    },
    {
      name: 'Thuế TNDN',
      s02: ['Thuế TNDN tạm tính', 'Thuế TNDN'],
      s03: ['Thuế TNDN', 'Thuế TNDN tạm tính'],
      s04: ['Thuế TNDN', 'Thuế TNDN tạm tính']
    }
  ];

  const totals02 = FS99X_readTotals_('02. Doanh thu', definitions.map(x => x.s02));
  const totals03 = FS99X_readTotals_('03. Chi phí & Vốn', definitions.map(x => x.s03));
  const totals04 = FS99X_readTotals_('04. Dòng tiền & Lợi nhuận', definitions.map(x => x.s04));

  let startRow = sh99.getLastRow() + 2;
  sh99.getRange(startRow, 1, 1, 7).merge();
  sh99.getRange(startRow, 1)
    .setValue('TẦNG 6 - ĐỐI CHIẾU CHI TIẾT SHEET 02 / 03 / 04')
    .setFontWeight('bold')
    .setBackground('#CFE2F3');
  startRow++;

  const rows = [];
  definitions.forEach((d, i) => {
    const v02 = totals02[i];
    const v03 = totals03[i];
    const v04 = totals04[i];
    rows.push(FS99X_makeRow_(d.name, 'Sheet 02', v02, 'Sheet 03', v03));
    rows.push(FS99X_makeRow_(d.name, 'Sheet 03', v03, 'Sheet 04', v04));
  });

  sh99.getRange(startRow, 1, rows.length, 7).setValues(rows);
  sh99.getRange(startRow, 3, rows.length, 3).setNumberFormat('#,##0.000');

  for (let r = 0; r < rows.length; r++) {
    const status = rows[r][5];
    const bg = status === 'PASS' ? '#D9EAD3' : (status === 'MISSING' ? '#FFF2CC' : '#F4CCCC');
    sh99.getRange(startRow + r, 1, 1, 7).setBackground(bg);
  }
  sh99.autoResizeColumns(1, 7);
}

function FS99X_makeRow_(indicator, leftName, left, rightName, right) {
  if (!left.found || !right.found) {
    const missing = [];
    if (!left.found) missing.push(leftName);
    if (!right.found) missing.push(rightName);
    return [
      'TẦNG 6 - ĐỐI CHIẾU 02/03/04',
      indicator + ': ' + leftName + ' = ' + rightName,
      left.found ? left.value / 1e9 : '',
      right.found ? right.value / 1e9 : '',
      '',
      'MISSING',
      'Không có chỉ tiêu tại: ' + missing.join(', ') + '. Giá trị được để trống.'
    ];
  }

  const diff = left.value - right.value;
  const pass = Math.abs(diff) <= 1;
  return [
    'TẦNG 6 - ĐỐI CHIẾU 02/03/04',
    indicator + ': ' + leftName + ' = ' + rightName,
    left.value / 1e9,
    right.value / 1e9,
    diff / 1e9,
    pass ? 'PASS' : 'FAIL',
    'Nguồn: ' + left.header + ' ↔ ' + right.header + '. Đơn vị tỷ đồng.'
  ];
}

function FS99X_readTotals_(sheetName, aliasGroups) {
  const sh = SpreadsheetApp.getActive().getSheetByName(sheetName);
  if (!sh) throw new Error('Không tìm thấy sheet ' + sheetName + '.');

  const headerRow = FS99X_detectHeaderRow_(sh, aliasGroups);
  if (headerRow < 1) {
    return aliasGroups.map(() => ({ found: false, value: 0, header: '' }));
  }

  const headers = sh.getRange(headerRow, 1, 1, sh.getLastColumn()).getDisplayValues()[0];
  const dataStart = headerRow + 1;
  const n = Math.max(0, sh.getLastRow() - headerRow);
  const values = n > 0 ? sh.getRange(dataStart, 1, n, sh.getLastColumn()).getValues() : [];

  return aliasGroups.map(aliases => {
    const colIndex = FS99X_findHeader_(headers, aliases);
    if (colIndex < 0) return { found: false, value: 0, header: '' };
    const total = values.reduce((s, row) => s + (Number(row[colIndex]) || 0), 0);
    return { found: true, value: total, header: headers[colIndex] };
  });
}

function FS99X_detectHeaderRow_(sh, aliasGroups) {
  const maxRows = Math.min(8, sh.getLastRow());
  let bestRow = -1;
  let bestHits = 0;

  for (let r = 1; r <= maxRows; r++) {
    const headers = sh.getRange(r, 1, 1, sh.getLastColumn()).getDisplayValues()[0];
    let hits = 0;
    aliasGroups.forEach(a => {
      if (FS99X_findHeader_(headers, a) >= 0) hits++;
    });
    if (hits > bestHits) {
      bestHits = hits;
      bestRow = r;
    }
  }

  return bestHits > 0 ? bestRow : -1;
}

function FS99X_findHeader_(headers, aliases) {
  const normalizedHeaders = headers.map(FS99X_key_);
  const normalizedAliases = aliases.map(FS99X_key_).filter(k => k.length >= 3);

  for (const key of normalizedAliases) {
    const idx = normalizedHeaders.findIndex(h => h && h === key);
    if (idx >= 0) return idx;
  }

  for (const key of normalizedAliases) {
    const idx = normalizedHeaders.findIndex(h => {
      if (!h || h.length < 3) return false;
      return h.indexOf(key) >= 0 || key.indexOf(h) >= 0;
    });
    if (idx >= 0) return idx;
  }

  return -1;
}

function FS99X_key_(v) {
  return String(v || '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd')
    .replace(/[^a-z0-9]+/g, ' ')
    .trim();
}
