function FS_lapSheet03() {
  const ss = SpreadsheetApp.getActive();
  const tech = FS_getSheet_(ss, FS_CFG.SHEETS.TECH, FS_CFG.SHEETS.TECH_LEGACY);
  const rev = ss.getSheetByName(FS_CFG.SHEETS.REVENUE);
  if (!tech || !rev) throw new Error('Cần lập 01A. Kỹ thuật và 02. Doanh thu trước.');

  const info = FS_readInfo_(tech);
  const months = FS_num_(info['Số tháng mô hình']);
  const costs = FS_costMap_(tech);
  const sched = FS_readBlock_(tech, 'TIEN_DO_CHI_PHI').map(r => ({
    item: String(r[0] || ''),
    start: Math.max(1, FS_num_(r[1])),
    duration: Math.max(1, FS_num_(r[2])),
    rate: FS_rate_(r[3]),
    type: String(r[4] || '')
  }));
  const productRows = FS_readBlock_(tech, 'SAN_PHAM');
  const revRows = rev.getLastRow() > 1 ? rev.getRange(2, 1, rev.getLastRow() - 1, 20).getValues() : [];

  const items = {
    xd: FS_costItem_(costs, ['Chi phí XD/TB/khác', 'Chi phí xây dựng & thiết bị']),
    gpmb: FS_costItem_(costs, ['Chi phí GPMB']),
    landUse: FS_costItem_(costs, ['Tiền SDĐ', 'Tiền sử dụng đất']),
    landRent: FS_costItem_(costs, ['Tiền thuê đất']),
    htkt: FS_costItem_(costs, ['Chi phí HTKT']),
    sell: FS_costItem_(costs, ['Chi phí bán hàng']),
    reserve: FS_costItem_(costs, ['Chi phí dự phòng']),
    op: FS_costItem_(costs, ['Chi phí vận hành']),
    maint: FS_costItem_(costs, ['Chi phí bảo trì'])
  };

  const legacyLand = FS_costItem_(costs, ['Tiền SDĐ/thuê đất']);
  if (!items.landUse.before && !items.landRent.before && legacyLand.before) {
    throw new Error('Sheet 01A còn gộp "Tiền SDĐ/thuê đất". Hãy tách thành 2 dòng "Tiền SDĐ" và "Tiền thuê đất" trước khi lập Sheet 03.');
  }

  function scheduled(itemKey, monthNo) {
    const it = items[itemKey];
    if (!it.before) return 0;
    const candidates = sched.filter(x => FS_key_(x.item) === FS_key_(it.name));
    if (!candidates.length) return 0;
    return candidates.reduce((sum, x) => {
      if (monthNo < x.start || monthNo >= x.start + x.duration) return sum;
      const oneTime = FS_norm_(x.type).includes('mot lan');
      const factor = oneTime ? (monthNo === x.start ? 1 : 0) : 1 / x.duration;
      return sum + it.before * x.rate * factor;
    }, 0);
  }

  const productInput = {};
  productRows.forEach(r => {
    const name = String(r[0] || '').trim();
    if (!name) return;
    const code = FS_maSanPhamGoc_(name);
    const opText = r[9];
    const maintText = r[10];
    productInput[code] = {
      opRate: FS_namedRateFromText_(opText, ['CPVH', 'chi phí vận hành']) || items.op.rate,
      maintRate: FS_rateFromText_(maintText) || items.maint.rate,
      leaseYears: Math.max(0, FS_num_(r[11]))
    };
  });

  const productMeta = {};
  revRows.forEach(r => {
    const code = String(r[4] || '');
    if (!productMeta[code]) productMeta[code] = {
      code,
      name: r[5],
      type: r[6],
      area: FS_num_(r[7]),
      saleTotal: 0,
      rentTotal: 0,
      opRate: productInput[code]?.opRate || items.op.rate,
      maintRate: productInput[code]?.maintRate || items.maint.rate,
      leaseYears: productInput[code]?.leaseYears || 0
    };
    productMeta[code].saleTotal += FS_num_(r[13]);
    productMeta[code].rentTotal += FS_num_(r[14]);
  });
  const products = Object.values(productMeta);
  const totalArea = FS_sum_(products.map(p => p.area)) || 1;
  const saleArea = FS_sum_(products.filter(p => p.type === 'Bán').map(p => p.area)) || 1;
  const rentArea = FS_sum_(products.filter(p => p.type === 'Cho thuê').map(p => p.area)) || 1;

  const rows = [];
  for (let t = 1; t <= months; t++) {
    const monthRows = revRows.filter(r => FS_num_(r[0]) === t);
    const xdMonth = scheduled('xd', t);
    const gpmbMonth = scheduled('gpmb', t);
    const landUseMonth = scheduled('landUse', t);
    const landRentMonth = scheduled('landRent', t);
    const htktMonth = scheduled('htkt', t);
    const reserveMonth = (xdMonth + htktMonth) * items.reserve.rate;

    monthRows.forEach(r => {
      const code = String(r[4] || '');
      const meta = productMeta[code];
      const areaShare = FS_ratio_(meta.area, totalArea);
      const saleShare = meta.type === 'Bán' ? FS_ratio_(meta.area, saleArea) : 0;
      const rentShare = meta.type === 'Cho thuê' ? FS_ratio_(meta.area, rentArea) : 0;
      const saleRevenue = FS_num_(r[13]);
      const rentRevenue = FS_num_(r[14]);

      const xd = xdMonth * areaShare;
      const gpmb = gpmbMonth * areaShare;
      const landUse = landUseMonth * saleShare;
      const landRent = landRentMonth * rentShare;
      const htkt = htktMonth * areaShare;
      const sell = meta.type === 'Bán' ? saleRevenue * items.sell.rate : 0;
      const op = meta.type === 'Cho thuê' ? rentRevenue * meta.opRate : 0;
      const maint = meta.type === 'Cho thuê' ? rentRevenue * meta.maintRate : 0;
      const reserve = reserveMonth * areaShare;

      const vatIn =
        xd * items.xd.vat + gpmb * items.gpmb.vat + landUse * items.landUse.vat +
        landRent * items.landRent.vat + htkt * items.htkt.vat + sell * items.sell.vat +
        op * items.op.vat + maint * items.maint.vat + reserve * items.reserve.vat;
      const totalBefore = xd + gpmb + landUse + landRent + htkt + sell + op + maint + reserve;

      rows.push([
        r[0], r[1], r[2], r[3], code, r[5], r[6], r[18], r[17],
        xd, gpmb, landUse, landRent, htkt, sell, op, maint, reserve,
        totalBefore, vatIn, totalBefore + vatIn
      ]);
    });
  }

  const sh = FS_getOrCreateSheet_(ss, FS_CFG.SHEETS.COST);
  FS_resetSheet_(sh, rows.length + 1, 21);
  sh.getRange(1, 1, 1, 21).setValues([[
    'Tháng số', 'Tháng', 'Năm', 'Quý', 'Mã SP', 'Tên sản phẩm', 'Loại hình',
    'Dòng tiền thu khách hàng', 'VAT đầu ra', 'XD/TB trước VAT', 'GPMB trước VAT',
    'Tiền SDĐ trước VAT', 'Tiền thuê đất trước VAT', 'HTKT trước VAT',
    'Chi phí bán hàng trước VAT', 'Chi phí vận hành trước VAT',
    'Chi phí bảo trì trước VAT', 'Chi phí dự phòng trước VAT',
    'Tổng chi trước VAT', 'VAT đầu vào', 'Tổng chi sau VAT'
  ]]);
  if (rows.length) sh.getRange(2, 1, rows.length, 21).setValues(rows);
  FS03_format_(sh, rows.length + 1);
}

function FS03_format_(sh, endRow) {
  sh.setFrozenRows(1);
  sh.setFrozenColumns(7);
  sh.getRange(1, 1, 1, 21).setFontWeight('bold').setBackground('#fce4d6').setWrap(true);
  if (endRow > 1) {
    sh.getRange(2, 2, endRow - 1, 1).setNumberFormat('MM/yyyy');
    sh.getRange(2, 8, endRow - 1, 14).setNumberFormat('#,##0');
  }
  sh.autoResizeColumns(1, 21);
}
