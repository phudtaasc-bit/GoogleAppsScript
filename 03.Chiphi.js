function FS_lapSheet03() {
  const ss = SpreadsheetApp.getActive();
  const tech = FS_getSheet_(ss, FS_CFG.SHEETS.TECH, FS_CFG.SHEETS.TECH_LEGACY);
  const rev = ss.getSheetByName(FS_CFG.SHEETS.REVENUE);
  if (!tech || !rev) throw new Error('Cần lập 01A. Kỹ thuật và 02. Doanh thu trước.');

  const info = FS_readInfo_(tech);
  const months = FS_num_(info['Số tháng mô hình']);
  const costs = FS_costMap_(tech);
  const sched = FS_readBlock_(tech, 'TIEN_DO_CHI_PHI').map(r => ({
    item: String(r[0] || ''), start: Math.max(1, FS_num_(r[1])),
    duration: Math.max(1, FS_num_(r[2])), rate: FS_rate_(r[3]), type: String(r[4] || '')
  }));
  const revRows = rev.getLastRow() > 1 ? rev.getRange(2, 1, rev.getLastRow() - 1, 20).getValues() : [];

  const items = {
    xd: FS_costItem_(costs, ['Chi phí XD/TB/khác', 'Chi phí xây dựng & thiết bị']),
    gpmb: FS_costItem_(costs, ['Chi phí GPMB']),
    land: FS_costItem_(costs, ['Tiền SDĐ/thuê đất']),
    htkt: FS_costItem_(costs, ['Chi phí HTKT']),
    sell: FS_costItem_(costs, ['Chi phí bán hàng']),
    reserve: FS_costItem_(costs, ['Chi phí dự phòng']),
    op: FS_costItem_(costs, ['Chi phí vận hành']),
    maint: FS_costItem_(costs, ['Chi phí bảo trì'])
  };

  function scheduled(itemKey, monthNo) {
    const it = items[itemKey];
    const candidates = sched.filter(x => FS_key_(x.item) === FS_key_(it.name));
    if (!candidates.length) return 0;
    return candidates.reduce((sum, x) => {
      if (monthNo < x.start || monthNo >= x.start + x.duration) return sum;
      const oneTime = FS_norm_(x.type).includes('mot lan');
      const factor = oneTime ? (monthNo === x.start ? 1 : 0) : 1 / x.duration;
      return sum + it.before * x.rate * factor;
    }, 0);
  }

  const productMeta = {};
  revRows.forEach(r => {
    const code = String(r[4] || '');
    if (!productMeta[code]) productMeta[code] = {
      code, name: r[5], type: r[6], area: FS_num_(r[7]), saleTotal: 0, rentTotal: 0
    };
    productMeta[code].saleTotal += FS_num_(r[13]);
    productMeta[code].rentTotal += FS_num_(r[14]);
  });
  const products = Object.values(productMeta);
  const totalArea = FS_sum_(products.map(p => p.area)) || 1;

  const rows = [];
  for (let t = 1; t <= months; t++) {
    const monthRows = revRows.filter(r => FS_num_(r[0]) === t);
    const xdMonth = scheduled('xd', t);
    const gpmbMonth = scheduled('gpmb', t);
    const landMonth = scheduled('land', t);
    const htktMonth = scheduled('htkt', t);
    const reserveMonth = (xdMonth + htktMonth) * items.reserve.rate;

    monthRows.forEach(r => {
      const code = String(r[4] || '');
      const meta = productMeta[code];
      const areaShare = FS_ratio_(meta.area, totalArea);
      const saleRevenue = FS_num_(r[13]);
      const rentRevenue = FS_num_(r[14]);

      const xd = xdMonth * areaShare;
      const gpmb = gpmbMonth * areaShare;
      const land = landMonth * areaShare;
      const htkt = htktMonth * areaShare;
      const sell = r[6] === 'Bán' ? saleRevenue * items.sell.rate : 0;
      const op = r[6] === 'Cho thuê' ? rentRevenue * items.op.rate : 0;
      const maint = r[6] === 'Cho thuê' ? rentRevenue * items.maint.rate : 0;
      const reserve = reserveMonth * areaShare;

      const vatIn =
        xd * items.xd.vat + gpmb * items.gpmb.vat + land * items.land.vat +
        htkt * items.htkt.vat + sell * items.sell.vat + op * items.op.vat +
        maint * items.maint.vat + reserve * items.reserve.vat;
      const totalBefore = xd + gpmb + land + htkt + sell + op + maint + reserve;

      rows.push([
        r[0], r[1], r[2], r[3], code, r[5], r[6], r[18], r[17],
        xd, gpmb, land, htkt, sell, op, maint, reserve,
        totalBefore, vatIn, totalBefore + vatIn
      ]);
    });
  }

  const sh = FS_getOrCreateSheet_(ss, FS_CFG.SHEETS.COST);
  FS_resetSheet_(sh, rows.length + 1, 20);
  sh.getRange(1, 1, 1, 20).setValues([[
    'Tháng số', 'Tháng', 'Năm', 'Quý', 'Mã SP', 'Tên sản phẩm', 'Loại hình',
    'Dòng tiền thu khách hàng', 'VAT đầu ra', 'XD/TB trước VAT', 'GPMB trước VAT',
    'Tiền SDĐ/thuê đất trước VAT', 'HTKT trước VAT', 'Chi phí bán hàng trước VAT',
    'Chi phí vận hành trước VAT', 'Chi phí bảo trì trước VAT', 'Chi phí dự phòng trước VAT',
    'Tổng chi trước VAT', 'VAT đầu vào', 'Tổng chi sau VAT'
  ]]);
  if (rows.length) sh.getRange(2, 1, rows.length, 20).setValues(rows);
  FS03_format_(sh, rows.length + 1);
}

function FS03_format_(sh, endRow) {
  sh.setFrozenRows(1);
  sh.setFrozenColumns(7);
  sh.getRange(1, 1, 1, 20).setFontWeight('bold').setBackground('#fce4d6').setWrap(true);
  if (endRow > 1) {
    sh.getRange(2, 2, endRow - 1, 1).setNumberFormat('MM/yyyy');
    sh.getRange(2, 8, endRow - 1, 13).setNumberFormat('#,##0.00');
  }
  sh.autoResizeColumns(1, 20);
}
