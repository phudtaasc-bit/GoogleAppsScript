/*************************************************
 * 03_ChiPhiVonVay.gs - FS V2.1
 * Dán đè toàn bộ file 03_ChiPhiVonVay.gs hiện tại.
 * Giữ nguyên layout Sheet 03: A:AI = 35 cột như mô hình hiện tại.
 *************************************************/

function FS_lapSheet03() {
  const ss = SpreadsheetApp.getActive();
  const tech = ss.getSheetByName('01. Kỹ thuật');
  if (!tech) throw new Error('Không tìm thấy sheet "01. Kỹ thuật".');

  const sh02 = FS03V21_getSheetByPrefix_(ss, '02');
  if (!sh02) throw new Error('Không tìm thấy sheet bắt đầu bằng "02".');

  const sh03 = FS03V21_getSheetByPrefix_(ss, '03');
  if (!sh03) throw new Error('Không tìm thấy sheet bắt đầu bằng "03".');

  const thongTin = FS03V21_docThongTin_(tech);
  const sanPham = FS03V21_docSanPham_(tech);
  const chiPhiChung = FS03V21_docChiPhiChung_(tech);
  const tienDoChiPhi = FS03V21_docTienDoChiPhi_(tech);

  const soThang = Number(thongTin['Số tháng mô hình']) || 0;
  const ngayBatDau = thongTin['Ngày bắt đầu dự án'];
  const tyLeVay = FS03V21_tyLe_(thongTin['Tỷ lệ vốn vay']);
  const laiSuatNam = FS03V21_tyLe_(thongTin['Lãi suất vay năm']);
  const thangBatDauTraGoc = Number(thongTin['Tháng bắt đầu trả gốc']) || 999999;
  const tyLeTruotGiaNam = FS03V21_tyLe_(thongTin['Tỷ lệ trượt chi phí/năm']);

  if (!soThang || !ngayBatDau) throw new Error('Thiếu "Số tháng mô hình" hoặc "Ngày bắt đầu dự án".');

  const endRow = soThang + 1;
  FS03V21_chuanBiSheet_(sh03, endRow, 35);

  sh03.getRange(1, 1, 1, 35).setValues([[
    'Tháng số','Tháng','Năm','Quý',
    'Dòng tiền huy động từ KH','VAT đầu ra','Thuế TNDN',
    'Chi XD/TB/khác trước VAT','Chi GPMB trước VAT','Tiền SDĐ/thuê đất trước VAT','Chi HTKT trước VAT',
    'Chi phí bán hàng trước VAT','Chi phí dự phòng trước VAT',
    'Tổng chi trước VAT','VAT đầu vào','Tổng chi sau VAT',
    'VAT còn được khấu trừ đầu kỳ','VAT phải nộp','VAT còn được khấu trừ cuối kỳ',
    'Dòng tiền trước tài trợ',
    'Nhu cầu vốn','Vốn góp CSH','Lũy kế vốn góp CSH',
    'Giải ngân vay','Lũy kế giải ngân vay',
    'Lãi vay','Trả gốc','Lũy kế trả gốc','Dư nợ cuối kỳ',
    'Dòng tiền sau tài trợ','Tiền cuối kỳ',
    'Ghi chú 1','Ghi chú 2','Ghi chú 3','Ghi chú 4'
  ]]);

  const tongTheoSheet02 = FS03V21_docTongTheoThangTuSheet02_(sh02, soThang);
  const nhomNguyenGia = FS03V21_taoNhomNguyenGia_(sanPham, chiPhiChung);
  const tongChiPhiTheoTienDo = FS03V21_tinhChiPhiTheoTienDo_(soThang, tienDoChiPhi, chiPhiChung, nhomNguyenGia, tyLeTruotGiaNam);

  const out = [];
  let vatKhauTruDauKy = 0;
  let tienCuoiKy = 0;
  let duNoCuoiKy = 0;
  let luyKeVonGop = 0;
  let luyKeGiaiNgan = 0;
  let luyKeTraGoc = 0;

  for (let t = 1; t <= soThang; t++) {
    const ngay = FS03V21_congThang_(ngayBatDau, t - 1);
    const nam = ngay.getFullYear();
    const quy = 'Q' + Math.ceil((ngay.getMonth() + 1) / 3) + '/' + nam;

    const tu02 = tongTheoSheet02[t] || { dongTienKH: 0, vatDauRa: 0, thueTNDN: 0, chiBanHang: 0 };
    const chi = tongChiPhiTheoTienDo[t] || FS03V21_mauChiPhiThang_();

    const tongChiTruocVAT = chi.chiXD + chi.chiGPMB + chi.chiDat + chi.chiHTKT + tu02.chiBanHang + chi.chiDuPhong;
    const vatDauVao = chi.vatXD + chi.vatGPMB + chi.vatDat + chi.vatHTKT + chi.vatDuPhong;
    const tongChiSauVAT = tongChiTruocVAT + vatDauVao;

    let vatPhaiNop = Math.max(0, tu02.vatDauRa - vatKhauTruDauKy - vatDauVao);
    let vatKhauTruCuoiKy = Math.max(0, vatKhauTruDauKy + vatDauVao - tu02.vatDauRa);
    if (t === soThang && vatKhauTruCuoiKy > 0) {
      vatPhaiNop = -vatKhauTruCuoiKy;
      vatKhauTruCuoiKy = 0;
    }

    const dongTienTruocTaiTro = tu02.dongTienKH - tongChiSauVAT - vatPhaiNop - tu02.thueTNDN;
    const laiVay = duNoCuoiKy * (Math.pow(1 + laiSuatNam, 1 / 12) - 1);
    const nhuCauVon = Math.max(0, -(tienCuoiKy + dongTienTruocTaiTro - laiVay));
    const vonGopCSH = nhuCauVon * (1 - tyLeVay);
    const giaiNganVay = nhuCauVon * tyLeVay;
    const tienSauTaiTroTruocGoc = tienCuoiKy + dongTienTruocTaiTro + vonGopCSH + giaiNganVay - laiVay;
    const traGoc = t < thangBatDauTraGoc ? 0 : Math.min(duNoCuoiKy + giaiNganVay, Math.max(0, tienSauTaiTroTruocGoc));
    const duNoMoi = Math.max(0, duNoCuoiKy + giaiNganVay - traGoc);
    const dongTienSauTaiTro = dongTienTruocTaiTro + vonGopCSH + giaiNganVay - laiVay - traGoc;
    const tienCuoiKyMoi = tienCuoiKy + dongTienSauTaiTro;

    luyKeVonGop += vonGopCSH;
    luyKeGiaiNgan += giaiNganVay;
    luyKeTraGoc += traGoc;

    out.push([
      t, ngay, nam, quy,
      tu02.dongTienKH, tu02.vatDauRa, tu02.thueTNDN,
      chi.chiXD, chi.chiGPMB, chi.chiDat, chi.chiHTKT,
      tu02.chiBanHang, chi.chiDuPhong,
      tongChiTruocVAT, vatDauVao, tongChiSauVAT,
      vatKhauTruDauKy, vatPhaiNop, vatKhauTruCuoiKy,
      dongTienTruocTaiTro,
      nhuCauVon, vonGopCSH, luyKeVonGop,
      giaiNganVay, luyKeGiaiNgan,
      laiVay, traGoc, luyKeTraGoc, duNoMoi,
      dongTienSauTaiTro, tienCuoiKyMoi,
      '', '', '', ''
    ]);

    vatKhauTruDauKy = vatKhauTruCuoiKy;
    tienCuoiKy = tienCuoiKyMoi;
    duNoCuoiKy = duNoMoi;
  }

  if (out.length) sh03.getRange(2, 1, out.length, 35).setValues(out);
  FS03_formatSheet_(sh03, endRow);
}

function FS03V21_mauChiPhiThang_() {
  return { chiXD: 0, chiGPMB: 0, chiDat: 0, chiHTKT: 0, chiDuPhong: 0, vatXD: 0, vatGPMB: 0, vatDat: 0, vatHTKT: 0, vatDuPhong: 0 };
}

/*************************************************
 * NHÓM NGUYÊN GIÁ
 *************************************************/
function FS03V21_taoNhomNguyenGia_(sanPham, chiPhiChung) {
  const tongCPXDCoSo = FS03V21_tong_(sanPham.map(sp => sp.dtkd * sp.cpxdM2));
  const tongDienTichDat = FS03V21_tong_(sanPham.map(sp => sp.dienTichDat));
  const tongGPMB = FS03V21_layChiPhi_(chiPhiChung, ['Chi phí GPMB']);
  const tongHTKT = FS03V21_layChiPhi_(chiPhiChung, ['Chi phí HTKT']);
  const tyLeDuPhong = FS03V21_layTyLe_(chiPhiChung, ['Chi phí dự phòng']);
  const vatGPMB = FS03V21_layVAT_(chiPhiChung, ['Chi phí GPMB']);
  const vatHTKT = FS03V21_layVAT_(chiPhiChung, ['Chi phí HTKT']);
  const vatDuPhong = FS03V21_layVAT_(chiPhiChung, ['Chi phí dự phòng']);

  const ds = {};
  sanPham.forEach(sp => {
    const ma = FS03V21_key_(sp.ten);
    const coSoCPXD = sp.dtkd * sp.cpxdM2;
    const tyTrongCPXD = tongCPXDCoSo ? coSoCPXD / tongCPXDCoSo : 0;
    const tyTrongDat = tongDienTichDat ? sp.dienTichDat / tongDienTichDat : 0;
    const chiXD = coSoCPXD;
    const chiHTKT = tongHTKT * tyTrongCPXD;
    const chiGPMB = tongGPMB * tyTrongDat;
    const chiDat = FS03V21_layChiPhiDatTheoSanPham_(sp, chiPhiChung);
    const chiDuPhong = (chiXD + chiHTKT) * tyLeDuPhong;
    const vatDat = FS03V21_layVATDatTheoSanPham_(sp, chiPhiChung);

    ds[ma] = {
      ten: sp.ten,
      hinhThuc: sp.hinhThuc,
      nhomTaiSan: FS03V21_laChoThue_(sp.hinhThuc) ? 'TAI_SAN_DAU_TU' : 'TAI_SAN_DE_BAN',
      chiXD, chiGPMB, chiHTKT, chiDat, chiDuPhong,
      vatXD: 0,
      vatGPMB: chiGPMB * vatGPMB,
      vatHTKT: chiHTKT * vatHTKT,
      vatDat: chiDat * vatDat,
      vatDuPhong: chiDuPhong * vatDuPhong,
      tongTruocLaiVay: chiXD + chiGPMB + chiHTKT + chiDat + chiDuPhong,
      thoiGianPhanBo: Math.max(1, sp.thoiGianThue || 600)
    };
  });
  return ds;
}

function FS03V21_layChiPhiDatTheoSanPham_(sp, chiPhiChung) {
  const k = FS03V21_key_(sp.ten);
  if (FS03V21_chua_(k, ['lienke'])) return FS03V21_layChiPhi_(chiPhiChung, ['Tiền SDĐ liền kề','Tiền SDĐ đất liền kề','Tiền SD đất liền kề']);
  if (FS03V21_chua_(k, ['chungcu','canho'])) return FS03V21_layChiPhi_(chiPhiChung, ['Tiền thuê đất chung cư','Tiền thuê đất Chung cư']);
  if (FS03V21_chua_(k, ['tmdv','thuongmai','thuongmaidichvu'])) return FS03V21_layChiPhi_(chiPhiChung, ['Tiền thuê đất TMDV','Tiền thuê đất thương mại dịch vụ']);
  if (FS03V21_chua_(k, ['cho'])) return FS03V21_layChiPhi_(chiPhiChung, ['Tiền thuê đất Chợ','Tiền thuê đất chợ']);
  return 0;
}

function FS03V21_layVATDatTheoSanPham_(sp, chiPhiChung) {
  const k = FS03V21_key_(sp.ten);
  if (FS03V21_chua_(k, ['lienke'])) return FS03V21_layVAT_(chiPhiChung, ['Tiền SDĐ liền kề','Tiền SDĐ đất liền kề','Tiền SD đất liền kề']);
  if (FS03V21_chua_(k, ['chungcu','canho'])) return FS03V21_layVAT_(chiPhiChung, ['Tiền thuê đất chung cư','Tiền thuê đất Chung cư']);
  if (FS03V21_chua_(k, ['tmdv','thuongmai','thuongmaidichvu'])) return FS03V21_layVAT_(chiPhiChung, ['Tiền thuê đất TMDV','Tiền thuê đất thương mại dịch vụ']);
  if (FS03V21_chua_(k, ['cho'])) return FS03V21_layVAT_(chiPhiChung, ['Tiền thuê đất Chợ','Tiền thuê đất chợ']);
  return 0;
}

/*************************************************
 * CHI PHÍ THEO TIẾN ĐỘ
 *************************************************/
function FS03V21_tinhChiPhiTheoTienDo_(soThang, tienDo, chiPhiChung, nhomNguyenGia, tyLeTruotGiaNam) {
  const out = {};
  for (let t = 1; t <= soThang; t++) out[t] = FS03V21_mauChiPhiThang_();
  const tongPool = FS03V21_congNhomNguyenGia_(nhomNguyenGia);

  const danhMuc = [
    { ten: 'Chi phí XD/TB/khác', truong: 'chiXD', truongVAT: 'vatXD', tong: tongPool.chiXD, vat: tongPool.vatXD, truotGia: true },
    { ten: 'Chi phí GPMB', truong: 'chiGPMB', truongVAT: 'vatGPMB', tong: tongPool.chiGPMB, vat: tongPool.vatGPMB, truotGia: false },
    { ten: 'Tiền SDĐ/thuê đất', truong: 'chiDat', truongVAT: 'vatDat', tong: tongPool.chiDat, vat: tongPool.vatDat, truotGia: false },
    { ten: 'Chi phí HTKT', truong: 'chiHTKT', truongVAT: 'vatHTKT', tong: tongPool.chiHTKT, vat: tongPool.vatHTKT, truotGia: true },
    { ten: 'Chi phí dự phòng', truong: 'chiDuPhong', truongVAT: 'vatDuPhong', tong: tongPool.chiDuPhong, vat: tongPool.vatDuPhong, truotGia: false }
  ];

  danhMuc.forEach(dm => {
    const cacDong = tienDo.filter(td => FS03V21_cungTenKhoanMuc_(td.khoanMuc, dm.ten));
    if (!cacDong.length || !dm.tong) return;
    cacDong.forEach(td => {
      const batDau = Math.max(1, Number(td.thangBatDau) || 1);
      const thoiGian = Math.max(1, Number(td.thoiGian) || 1);
      const tyLe = FS03V21_tyLe_(td.tyLe) || 0;
      const loai = FS03V21_key_(td.loai);
      const ketThuc = Math.min(soThang, batDau + thoiGian - 1);
      const soKy = Math.max(1, ketThuc - batDau + 1);
      for (let t = batDau; t <= ketThuc; t++) {
        const heSoTruotGia = dm.truotGia ? Math.pow(1 + tyLeTruotGiaNam, (t - 1) / 12) : 1;
        const giaTri = loai === FS03V21_key_('Một lần') ? (t === batDau ? dm.tong * tyLe * heSoTruotGia : 0) : dm.tong * tyLe / soKy * heSoTruotGia;
        const vat = dm.tong ? giaTri * (dm.vat / dm.tong) : 0;
        out[t][dm.truong] += giaTri;
        out[t][dm.truongVAT] += vat;
      }
    });
  });
  return out;
}

function FS03V21_congNhomNguyenGia_(nhomNguyenGia) {
  const out = FS03V21_mauChiPhiThang_();
  Object.keys(nhomNguyenGia).forEach(k => {
    const p = nhomNguyenGia[k];
    Object.keys(out).forEach(key => out[key] += Number(p[key]) || 0);
  });
  return out;
}

/*************************************************
 * ĐỌC SHEET 02
 *************************************************/
function FS03V21_docTongTheoThangTuSheet02_(sh02, soThang) {
  const out = {};
  for (let t = 1; t <= soThang; t++) out[t] = { dongTienKH: 0, vatDauRa: 0, thueTNDN: 0, chiBanHang: 0 };
  const lastRow = sh02.getLastRow();
  if (lastRow < 2) return out;
  const data = sh02.getRange(2, 1, lastRow - 1, Math.min(32, sh02.getLastColumn())).getValues();
  data.forEach(r => {
    const t = Number(r[0]) || 0;
    if (!t || !out[t]) return;
    out[t].dongTienKH += FS03V21_so_(r[19]);
    out[t].vatDauRa += FS03V21_so_(r[18]);
    out[t].thueTNDN += FS03V21_so_(r[31]);
    out[t].chiBanHang += FS03V21_so_(r[22]);
  });
  return out;
}

/*************************************************
 * ĐỌC 01. KỸ THUẬT
 *************************************************/
function FS03V21_docThongTin_(sheet) {
  const blockRow = FS03V21_timDongChinhXac_(sheet, 'THONG_TIN_CHUNG');
  if (!blockRow) throw new Error('Không tìm thấy block THONG_TIN_CHUNG.');
  const out = {};
  for (let r = blockRow + 1; r <= sheet.getLastRow(); r++) {
    const label = String(sheet.getRange(r, 1).getDisplayValue() || '').trim();
    if (!label || FS03V21_laBlock_(label)) break;
    out[label] = sheet.getRange(r, 2).getValue();
  }
  return out;
}

function FS03V21_docSanPham_(sheet) {
  const block = FS03V21_docBlock_(sheet, 'SAN_PHAM');
  return block.rows.filter(r => String(r[0] || '').trim()).map(r => ({
    ten: String(r[0] || '').trim(), hinhThuc: String(r[1] || '').trim(), dtkd: FS03V21_so_(r[2]),
    giaBan: FS03V21_so_(r[3]), giaThue: FS03V21_so_(r[4]), cpxdM2: FS03V21_so_(r[5]),
    vatDauRa: FS03V21_tyLe_(r[6]), thueTNDN: FS03V21_tyLe_(r[7]), lapDay: FS03V21_tyLe_(r[8]),
    cpvh: FS03V21_tyLe_(r[9]), ghiChu: String(r[10] || ''), dienTichDat: FS03V21_so_(r[11]),
    thoiGianThue: FS03V21_docThoiGianThueTuGhiChu_(String(r[10] || ''))
  }));
}

function FS03V21_docChiPhiChung_(sheet) {
  const block = FS03V21_docBlock_(sheet, 'CHI_PHI_CHUNG');
  const out = {};
  block.rows.forEach(r => {
    const ten = String(r[0] || '').trim();
    if (!ten) return;
    out[FS03V21_key_(ten)] = { ten, truocVAT: FS03V21_so_(r[1]), vat: FS03V21_tyLe_(r[2]), sauVAT: FS03V21_so_(r[3]), ghiChu: String(r[4] || ''), tyLe: FS03V21_tyLe_(r[5]) };
  });
  return out;
}

function FS03V21_docTienDoChiPhi_(sheet) {
  const block = FS03V21_docBlock_(sheet, 'TIEN_DO_CHI_PHI');
  return block.rows.filter(r => String(r[0] || '').trim()).map(r => ({
    khoanMuc: String(r[0] || '').trim(), thangBatDau: FS03V21_so_(r[1]), thoiGian: FS03V21_so_(r[2]), tyLe: FS03V21_tyLe_(r[3]), loai: String(r[4] || '').trim()
  }));
}

function FS03V21_docBlock_(sheet, blockName) {
  const blockRow = FS03V21_timDongChinhXac_(sheet, blockName);
  if (!blockRow) throw new Error('Không tìm thấy block ' + blockName);
  const headerRow = blockRow + 1;
  const lastCol = FS03V21_cotCuoiDong_(sheet, headerRow);
  const rows = [];
  let blank = 0;
  for (let r = blockRow + 2; r <= sheet.getLastRow(); r++) {
    const first = String(sheet.getRange(r, 1).getDisplayValue() || '').trim();
    if (FS03V21_laBlock_(first)) break;
    const row = sheet.getRange(r, 1, 1, lastCol).getValues()[0];
    const hasData = row.some(v => String(v || '').trim() !== '');
    if (!hasData) {
      blank++;
      if (blank >= 3) break;
      continue;
    }
    blank = 0;
    rows.push(row);
  }
  return { blockRow, headerRow, rows };
}

/*************************************************
 * TIỆN ÍCH CHI PHÍ
 *************************************************/
function FS03V21_layChiPhi_(map, dsTen) {
  for (const ten of dsTen) {
    const item = map[FS03V21_key_(ten)];
    if (item) return Number(item.truocVAT) || 0;
  }
  return 0;
}
function FS03V21_layVAT_(map, dsTen) {
  for (const ten of dsTen) {
    const item = map[FS03V21_key_(ten)];
    if (item) return FS03V21_tyLe_(item.vat);
  }
  return 0;
}
function FS03V21_layTyLe_(map, dsTen) {
  for (const ten of dsTen) {
    const item = map[FS03V21_key_(ten)];
    if (item) return FS03V21_tyLe_(item.tyLe);
  }
  return 0;
}
function FS03V21_cungTenKhoanMuc_(a, b) {
  const ka = FS03V21_key_(a), kb = FS03V21_key_(b);
  if (ka === kb) return true;
  if (kb === FS03V21_key_('Tiền SDĐ/thuê đất')) {
    return ['Tiền SDĐ/thuê đất','Tiền SDĐ liền kề','Tiền SDĐ đất liền kề','Tiền SD đất liền kề','Tiền thuê đất chung cư','Tiền thuê đất TMDV','Tiền thuê đất Chợ'].some(x => ka === FS03V21_key_(x));
  }
  return false;
}

/*************************************************
 * HELPER GIỮ TÊN CŨ ĐỂ TƯƠNG THÍCH
 *************************************************/
function FS03_capNhatNguonVonTuSheet04() {
  const ss = SpreadsheetApp.getActive();
  const sh03 = ss.getSheetByName('03. Chi phí & Vốn');
  const sh04 = ss.getSheetByName('04. Dòng tiền & Lợi nhuận');
  if (!sh03 || !sh04) return;
  const lastRow03 = sh03.getLastRow(), lastRow04 = sh04.getLastRow();
  if (lastRow03 < 2 || lastRow04 < 3) return;
  const data04 = sh04.getRange(3, 1, lastRow04 - 2, Math.min(25, sh04.getLastColumn())).getValues();
  const map04 = {};
  data04.forEach(r => {
    const thang = Number(r[0]) || 0;
    if (!thang) return;
    map04[thang] = { giaiNganVay: FS03V21_so_(r[21]), laiVay: FS03V21_so_(r[22]), traGoc: FS03V21_so_(r[23]), duNo: FS03V21_so_(r[24]) };
  });
  const thang03 = sh03.getRange(2, 1, lastRow03 - 1, 1).getValues();
  const outX = [], outZ = [], outAA = [], outAC = [];
  thang03.forEach(r => {
    const t = Number(r[0]) || 0;
    const v = map04[t] || { giaiNganVay: 0, laiVay: 0, traGoc: 0, duNo: 0 };
    outX.push([v.giaiNganVay]); outZ.push([v.laiVay]); outAA.push([v.traGoc]); outAC.push([v.duNo]);
  });
  sh03.getRange(2, 24, outX.length, 1).setValues(outX);
  sh03.getRange(2, 26, outZ.length, 1).setValues(outZ);
  sh03.getRange(2, 27, outAA.length, 1).setValues(outAA);
  sh03.getRange(2, 29, outAC.length, 1).setValues(outAC);
}

function FS03_formatSheet_(sh, endRow) {
  const lastCol = 35;
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, lastCol).setFontWeight('bold').setBackground('#d9ead3').setHorizontalAlignment('center').setWrap(true);
  sh.getRange(1, 1, endRow, lastCol).setFontFamily('Arial').setFontSize(10).setVerticalAlignment('middle');
  if (endRow > 1) {
    const rows = endRow - 1;
    sh.getRange(2, 1, rows, 1).setNumberFormat('0');
    sh.getRange(2, 2, rows, 1).setNumberFormat('dd/mm/yyyy');
    sh.getRange(2, 3, rows, 1).setNumberFormat('0');
    sh.getRange(2, 5, rows, 27).setNumberFormat('#,##0');
  }
  sh.autoResizeColumns(1, lastCol);
}
function FS03_getSheetByPrefix_(ss, prefix) { return FS03V21_getSheetByPrefix_(ss, prefix); }
function FS03_getBlockInfo_(sheet, blockName) {
  const b = FS03V21_docBlock_(sheet, blockName);
  return { blockRow: b.blockRow, startRow: b.blockRow + 2, endRow: b.blockRow + 1 + b.rows.length, rowCount: b.rows.length };
}
function FS03_getInfoValue_(sheet, label) { const info = FS03V21_docThongTin_(sheet); return info[label] || ''; }
function FS03_getInfoCellA1_(sheet, label) { const row = FS03_findRow_(sheet, label); return row ? sheet.getRange(row, 2).getA1Notation() : ''; }
function FS03_findRow_(sheet, text) { return FS03V21_timDongChua_(sheet, text); }
function FS03_norm_(v) { return FS03V21_norm_(v); }

/*************************************************
 * TIỆN ÍCH CHUNG
 *************************************************/
function FS03V21_chuanBiSheet_(sh, endRow, lastCol) {
  sh.showRows(1, sh.getMaxRows());
  sh.showColumns(1, sh.getMaxColumns());
  if (sh.getMaxRows() < endRow) sh.insertRowsAfter(sh.getMaxRows(), endRow - sh.getMaxRows());
  if (sh.getMaxColumns() < lastCol) sh.insertColumnsAfter(sh.getMaxColumns(), lastCol - sh.getMaxColumns());
  sh.clearContents();
  sh.clearFormats();
}
function FS03V21_getSheetByPrefix_(ss, prefix) { return ss.getSheets().find(s => String(s.getName()).trim().startsWith(prefix)); }
function FS03V21_timDongChinhXac_(sheet, text) {
  const data = sheet.getDataRange().getDisplayValues(), target = FS03V21_key_(text);
  for (let r = 0; r < data.length; r++) if (data[r].some(v => FS03V21_key_(v) === target)) return r + 1;
  return null;
}
function FS03V21_timDongChua_(sheet, text) {
  const data = sheet.getDataRange().getDisplayValues(), target = FS03V21_norm_(text);
  for (let r = 0; r < data.length; r++) if (data[r].some(v => FS03V21_norm_(v).includes(target))) return r + 1;
  return null;
}
function FS03V21_cotCuoiDong_(sheet, row) {
  const values = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  let last = 1;
  values.forEach((v, i) => { if (String(v || '').trim() !== '') last = i + 1; });
  return last;
}
function FS03V21_laBlock_(v) { return ['THONG_TIN_CHUNG','CHI_PHI_CHUNG','SAN_PHAM','KE_HOACH_BAN_THU_TIEN','TIEN_DO_CHI_PHI'].includes(String(v || '').trim()); }
function FS03V21_laChoThue_(hinhThuc) { return FS03V21_key_(hinhThuc) === FS03V21_key_('Cho thuê'); }
function FS03V21_docThoiGianThueTuGhiChu_(ghiChu) {
  const s = String(ghiChu || '').toLowerCase();
  let m = s.match(/([0-9,.]+)\s*năm/);
  if (m) return Math.round(Number(m[1].replace(',', '.')) * 12);
  m = s.match(/([0-9,.]+)\s*tháng/);
  if (m) return Math.round(Number(m[1].replace(',', '.')));
  return 600;
}
function FS03V21_congThang_(dateValue, months) { const d = new Date(dateValue); return new Date(d.getFullYear(), d.getMonth() + months, d.getDate()); }
function FS03V21_tong_(arr) { return arr.reduce((s, v) => s + (Number(v) || 0), 0); }
function FS03V21_so_(v) {
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  const raw = String(v || '').trim();
  if (!raw) return 0;
  const s = raw.replace(/\s/g, '').replace(/%/g, '').replace(/,/g, '.');
  const n = Number(s);
  return isFinite(n) ? n : 0;
}
function FS03V21_tyLe_(v) {
  if (v === '' || v === null || typeof v === 'undefined') return 0;
  if (typeof v === 'number') return v > 1 ? v / 100 : v;
  const n = Number(String(v || '').trim().replace('%', '').replace(',', '.'));
  if (!isFinite(n)) return 0;
  return n > 1 ? n / 100 : n;
}
function FS03V21_chua_(key, terms) { const k = FS03V21_key_(key); return terms.some(t => k.indexOf(FS03V21_key_(t)) >= 0); }
function FS03V21_key_(v) {
  return String(v || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/đ/g, 'd').replace(/²/g, '2').replace(/\^2/g, '2').replace(/m\s*2/g, 'm2').replace(/m\s*²/g, 'm2').replace(/[^a-z0-9]/g, '');
}
function FS03V21_norm_(v) {
  return String(v || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/đ/g, 'd').replace(/\s+/g, ' ').trim();
}
