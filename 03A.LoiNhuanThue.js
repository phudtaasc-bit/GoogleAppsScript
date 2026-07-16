const FS03A_CFG = Object.freeze({
  TECH: '01A. Kỹ thuật',
  REVENUE: '02. Doanh thu',
  COST: '03. Chi phí & Vốn',
  PROFIT: '03A. Lợi nhuận & Thuế'
});

function FS_lapSheet03A() {
  return FS_hoiTuTaiTro();
}

function FS03A_build_(interestByMonth, writeSheet) {
  const ss = SpreadsheetApp.getActive();
  const tech = ss.getSheetByName(FS03A_CFG.TECH);
  const revenueSheet = ss.getSheetByName(FS03A_CFG.REVENUE);
  const costSheet = ss.getSheetByName(FS03A_CFG.COST);
  if (!tech || !revenueSheet || !costSheet) {
    throw new Error('Cần lập "01A. Kỹ thuật", "02. Doanh thu" và "03. Chi phí & Vốn" trước.');
  }

  const products = FS03A_readProducts_(tech);
  const productByCode = {};
  products.forEach(p => { productByCode[p.code] = p; });

  const revenueRows = FS03A_readRevenue_(revenueSheet, productByCode);
  const costRows = FS03A_readCosts_(costSheet, productByCode);
  const costByKey = FS03A_indexByKey_(costRows, 'Sheet 03');
  const capitalizedInterestByCode = FS03A_allocateInterest_(costRows, interestByMonth || []);
  const pools = FS03A_buildPools_(products, revenueRows, costRows, capitalizedInterestByCode);
  const rows = [];

  revenueRows.forEach(revenue => {
    const product = productByCode[revenue.code];
    const key = revenue.code + '|' + revenue.monthNo;
    const cost = costByKey[key] || FS03A_emptyCost_(revenue);
    const pool = pools[revenue.code];

    let baseCostRecognized = 0;
    let capitalizedInterestRecognized = 0;
    let landUseRecognized = 0;
    let landRentRecognized = 0;

    if (product.group === 'Bán') {
      const rate = pool.totalSaleRevenue > 0 ? revenue.saleRevenue / pool.totalSaleRevenue : 0;
      baseCostRecognized = pool.baseCostPool * rate;
      capitalizedInterestRecognized = pool.interestPool * rate;
      landUseRecognized = pool.landUsePool * rate;
      landRentRecognized = pool.landRentPool * rate;
    } else {
      const leaseMonths = Math.max(1, product.leaseYears * 12);
      const active = pool.firstRentMonth > 0 &&
        revenue.monthNo >= pool.firstRentMonth &&
        revenue.monthNo < pool.firstRentMonth + leaseMonths;
      if (active) {
        baseCostRecognized = pool.baseCostPool / leaseMonths;
        capitalizedInterestRecognized = pool.interestPool / leaseMonths;
        landRentRecognized = pool.landRentPool / leaseMonths;
      }
    }

    const selling = cost.selling;
    const operating = cost.operating;
    const maintenance = cost.maintenance;
    const totalCost = baseCostRecognized + capitalizedInterestRecognized +
      landUseRecognized + landRentRecognized + selling + operating + maintenance;
    const pbt = revenue.totalRevenue - totalCost;
    const taxable = Math.max(0, pbt);
    const cit = taxable * revenue.citRate;
    const pat = pbt - cit;

    rows.push([
      revenue.monthNo, revenue.date, revenue.year, revenue.quarter,
      revenue.code, product.name, product.group,
      revenue.saleRevenue, revenue.rentRevenue, revenue.totalRevenue,
      baseCostRecognized, capitalizedInterestRecognized,
      landUseRecognized, landRentRecognized,
      selling, operating, maintenance,
      totalCost, pbt, taxable, revenue.citRate, cit, pat
    ]);
  });

  if (writeSheet !== false) FS03A_write_(ss, rows);
  return rows;
}

function FS03A_allocateInterest_(costRows, interestByMonth) {
  const monthTotals = {};
  costRows.forEach(r => {
    const eligible = r.construction + r.clearance + r.infrastructure + r.contingency + r.landUse + r.landRent;
    monthTotals[r.monthNo] = (monthTotals[r.monthNo] || 0) + eligible;
  });

  const result = {};
  costRows.forEach(r => {
    const eligible = r.construction + r.clearance + r.infrastructure + r.contingency + r.landUse + r.landRent;
    const total = monthTotals[r.monthNo] || 0;
    const interest = FS03A_num_(interestByMonth[r.monthNo - 1]);
    const allocated = total > 0 ? interest * eligible / total : 0;
    result[r.code] = (result[r.code] || 0) + allocated;
  });
  return result;
}

function FS03A_buildPools_(products, revenueRows, costRows, interestByCode) {
  const pools = {};
  products.forEach(p => {
    pools[p.code] = {
      totalSaleRevenue: 0, firstRentMonth: 0,
      baseCostPool: 0, interestPool: FS03A_num_(interestByCode[p.code]),
      landUsePool: 0, landRentPool: 0
    };
  });
  revenueRows.forEach(r => {
    const p = pools[r.code];
    p.totalSaleRevenue += r.saleRevenue;
    if (!p.firstRentMonth && r.rentRevenue > 0) p.firstRentMonth = r.monthNo;
  });
  costRows.forEach(r => {
    const p = pools[r.code];
    p.baseCostPool += r.construction + r.clearance + r.infrastructure + r.contingency;
    p.landUsePool += r.landUse;
    p.landRentPool += r.landRent;
  });
  return pools;
}

function FS03A_readProducts_(sheet) {
  const rows = FS03A_readBlock_(sheet, 'SAN_PHAM');
  const products = rows.map(r => ({
    code: String(r[0] || '').trim().toUpperCase(),
    name: String(r[1] || '').trim(),
    group: FS03A_group_(r[2]),
    leaseYears: Math.max(0, FS03A_num_(r[12]))
  })).filter(p => p.code || p.name);
  if (!products.length) throw new Error('Block SAN_PHAM không có dữ liệu sản phẩm.');
  if (products.some(p => !p.code)) throw new Error('Block SAN_PHAM còn thiếu Mã SP.');
  const codes = products.map(p => p.code);
  const dup = [...new Set(codes.filter((c, i) => codes.indexOf(c) !== i))];
  if (dup.length) throw new Error('Mã SP bị trùng trong SAN_PHAM: ' + dup.join(', '));
  const invalid = products.filter(p => !['Bán', 'Cho thuê'].includes(p.group));
  if (invalid.length) throw new Error('Nhóm sản phẩm không hợp lệ: ' + invalid.map(p => p.code).join(', '));
  const missing = products.filter(p => p.group === 'Cho thuê' && p.leaseYears <= 0);
  if (missing.length) throw new Error('Sản phẩm cho thuê chưa có Thời gian thuê: ' + missing.map(p => p.code).join(', '));
  return products;
}

function FS03A_readRevenue_(sheet, productByCode) {
  const table = FS03A_readTable_(sheet);
  const required = ['thangso','thang','nam','quy','masp','doanhthubantruocvat','doanhthuthuetruocvat','tongdoanhthutruocvat','thuesuattndn'];
  FS03A_require_(table.index, required, '02. Doanh thu');
  return table.values.map(row => {
    const code = String(row[table.index.masp] || '').trim().toUpperCase();
    if (!productByCode[code]) throw new Error('Sheet 02 có Mã SP không tồn tại: ' + code);
    return {
      monthNo: FS03A_num_(row[table.index.thangso]), date: row[table.index.thang],
      year: row[table.index.nam], quarter: row[table.index.quy], code,
      saleRevenue: FS03A_num_(row[table.index.doanhthubantruocvat]),
      rentRevenue: FS03A_num_(row[table.index.doanhthuthuetruocvat]),
      totalRevenue: FS03A_num_(row[table.index.tongdoanhthutruocvat]),
      citRate: FS03A_rate_(row[table.index.thuesuattndn])
    };
  });
}

function FS03A_readCosts_(sheet, productByCode) {
  const table = FS03A_readTable_(sheet);
  const required = ['thangso','thang','nam','quy','masp','xdtbtruocvat','gpmbtruocvat','htkttruocvat','tiensddtruocvat','tienthuedattruocvat','chiphiduphongtruocvat','chiphibanhangtruocvat','chiphivanhanhtruocvat','chiphibaotritruocvat'];
  FS03A_require_(table.index, required, '03. Chi phí & Vốn');
  return table.values.map(row => {
    const code = String(row[table.index.masp] || '').trim().toUpperCase();
    if (!productByCode[code]) throw new Error('Sheet 03 có Mã SP không tồn tại: ' + code);
    return {
      monthNo: FS03A_num_(row[table.index.thangso]), date: row[table.index.thang],
      year: row[table.index.nam], quarter: row[table.index.quy], code,
      construction: FS03A_num_(row[table.index.xdtbtruocvat]),
      clearance: FS03A_num_(row[table.index.gpmbtruocvat]),
      infrastructure: FS03A_num_(row[table.index.htkttruocvat]),
      landUse: FS03A_num_(row[table.index.tiensddtruocvat]),
      landRent: FS03A_num_(row[table.index.tienthuedattruocvat]),
      contingency: FS03A_num_(row[table.index.chiphiduphongtruocvat]),
      selling: FS03A_num_(row[table.index.chiphibanhangtruocvat]),
      operating: FS03A_num_(row[table.index.chiphivanhanhtruocvat]),
      maintenance: FS03A_num_(row[table.index.chiphibaotritruocvat])
    };
  });
}

function FS03A_indexByKey_(rows, source) {
  const out = {};
  rows.forEach(r => {
    const key = r.code + '|' + r.monthNo;
    if (out[key]) throw new Error(source + ' bị trùng khóa Mã SP + Tháng số: ' + key);
    out[key] = r;
  });
  return out;
}

function FS03A_emptyCost_(r) {
  return { monthNo:r.monthNo,date:r.date,year:r.year,quarter:r.quarter,code:r.code,
    construction:0,clearance:0,infrastructure:0,landUse:0,landRent:0,contingency:0,selling:0,operating:0,maintenance:0 };
}

function FS03A_write_(ss, rows) {
  let sh = ss.getSheetByName(FS03A_CFG.PROFIT);
  if (!sh) sh = ss.insertSheet(FS03A_CFG.PROFIT);
  sh.clear(); sh.clearFormats();
  const headers = [[
    'Tháng số','Tháng','Năm','Quý','Mã SP','Tên sản phẩm','Nhóm',
    'Doanh thu bán trước VAT','Doanh thu thuê trước VAT','Tổng doanh thu trước VAT',
    'Giá vốn XD/GPMB/HTKT/Dự phòng','Lãi vay vốn hóa phân bổ','Tiền SDĐ phân bổ','Tiền thuê đất phân bổ',
    'Chi phí bán hàng','Chi phí vận hành','Chi phí bảo trì','Tổng chi phí hạch toán',
    'Lợi nhuận trước thuế','Thu nhập chịu thuế','Thuế suất TNDN','Thuế TNDN','LNST'
  ]];
  sh.getRange(1,1,1,headers[0].length).setValues(headers);
  if (rows.length) sh.getRange(2,1,rows.length,headers[0].length).setValues(rows);
  sh.setFrozenRows(1); sh.setFrozenColumns(7);
  sh.getRange(1,1,1,headers[0].length).setFontWeight('bold').setBackground('#e2f0d9').setWrap(true);
  if (rows.length) {
    sh.getRange(2,2,rows.length,1).setNumberFormat('MM/yyyy');
    sh.getRange(2,8,rows.length,13).setNumberFormat('#,##0');
    sh.getRange(2,21,rows.length,1).setNumberFormat('0.00%');
    sh.getRange(2,22,rows.length,2).setNumberFormat('#,##0');
  }
  sh.autoResizeColumns(1,headers[0].length);
}

function FS03A_readTable_(sheet) {
  const cols = sheet.getLastColumn();
  const headers = sheet.getRange(1,1,1,cols).getDisplayValues()[0];
  const values = sheet.getLastRow() > 1 ? sheet.getRange(2,1,sheet.getLastRow()-1,cols).getValues() : [];
  const index = {}; headers.forEach((h,i) => { index[FS03A_key_(h)] = i; });
  return { values, index };
}
function FS03A_require_(index, required, name) {
  const missing = required.filter(k => index[k] == null);
  if (missing.length) throw new Error('Sheet "' + name + '" thiếu cột: ' + missing.join(', '));
}
function FS03A_readBlock_(sheet, marker) {
  const values = sheet.getDataRange().getValues();
  const target = FS03A_key_(marker); let start = -1;
  for (let i=0;i<values.length;i++) if (FS03A_key_(values[i][0])===target) { start=i; break; }
  if (start<0) throw new Error('Không tìm thấy block ' + marker + '.');
  const rows=[]; let blanks=0;
  for (let i=start+2;i<values.length;i++) {
    const row=values[i], first=String(row[0]||'').trim();
    const has=row.some(v=>String(v==null?'':v).trim()!=='');
    if (/^[A-Z0-9_]+$/.test(first)&&first.includes('_')) break;
    if (!has) { if (++blanks>=2) break; continue; }
    blanks=0; rows.push(row);
  }
  return rows;
}
function FS03A_group_(v) { const k=FS03A_key_(v); return k==='ban'?'Bán':k==='chothue'?'Cho thuê':String(v||'').trim(); }
function FS03A_num_(v) { if(typeof v==='number') return isFinite(v)?v:0; const t=String(v==null?'':v).trim().replace(/\s/g,''); if(!t)return 0; const n=Number(t.includes(',')&&t.includes('.')?t.replace(/\./g,'').replace(',','.'):t.replace(/,/g,'')); return isFinite(n)?n:0; }
function FS03A_rate_(v) { if(typeof v==='number') return v>1?v/100:v; const t=String(v==null?'':v).trim(); if(!t)return 0; const n=FS03A_num_(t.replace('%','')); return t.includes('%')||n>1?n/100:n; }
function FS03A_norm_(v) { return String(v==null?'':v).toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/đ/g,'d').replace(/\s+/g,' ').trim(); }
function FS03A_key_(v) { return FS03A_norm_(v).replace(/²/g,'2').replace(/\^2/g,'2').replace(/m\s*2/g,'m2').replace(/[^a-z0-9]/g,''); }
