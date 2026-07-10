/*************************************************
 * 03A_ChiPhiLogicPatch.gs
 * Vá tối thiểu trên logic Sheet 03 hiện tại:
 * - Tách pool Tiền SDĐ/Tiền thuê đất theo nhóm sản phẩm.
 * - Không nhân lặp toàn bộ pool tiền đất cho từng dòng tiến độ riêng.
 * - Phân bổ nguồn tiền đất theo trọng số diện tích đất trong cùng nhóm.
 * - Giữ tương thích cấu hình tổng Tiền SDĐ/thuê đất khi chưa khai báo chi tiết.
 * - Bổ sung VAT đầu vào Chi phí xây dựng & thiết bị và Chi phí bán hàng.
 * - Giữ nguyên logic hoàn VAT âm tại kỳ cuối.
 *************************************************/

function FS_lapSheet03_Patched() {
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

  if (!soThang || !ngayBatDau) {
    throw new Error('Thiếu "Số tháng mô hình" hoặc "Ngày bắt đầu dự án".');
  }

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
  const nhomNguyenGia = FS03P_taoNhomNguyenGia_(sanPham, chiPhiChung);
  const tongChiPhiTheoTienDo = FS03P_tinhChiPhiTheoTienDo_(
    soThang,
    tienDoChiPhi,
    nhomNguyenGia,
    tyLeTruotGiaNam
  );
  const vatBanHangRate = FS03V21_layVAT_(chiPhiChung, ['Chi phí bán hàng']);

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

    const tu02 = tongTheoSheet02[t] || {
      dongTienKH: 0,
      vatDauRa: 0,
      thueTNDN: 0,
      chiBanHang: 0
    };
    const chi = tongChiPhiTheoTienDo[t] || FS03V21_mauChiPhiThang_();

    const tongChiTruocVAT =
      chi.chiXD + chi.chiGPMB + chi.chiDat + chi.chiHTKT +
      tu02.chiBanHang + chi.chiDuPhong;

    const vatBanHang = tu02.chiBanHang * vatBanHangRate;
    const vatDauVao =
      chi.vatXD + chi.vatGPMB + chi.vatDat + chi.vatHTKT +
      vatBanHang + chi.vatDuPhong;
    const tongChiSauVAT = tongChiTruocVAT + vatDauVao;

    let vatPhaiNop = Math.max(0, tu02.vatDauRa - vatKhauTruDauKy - vatDauVao);
    let vatKhauTruCuoiKy = Math.max(0, vatKhauTruDauKy + vatDauVao - tu02.vatDauRa);

    // Theo quy ước của mô hình: VAT còn được khấu trừ tại kỳ cuối là khoản hoàn thuế.
    if (t === soThang && vatKhauTruCuoiKy > 0) {
      vatPhaiNop = -vatKhauTruCuoiKy;
      vatKhauTruCuoiKy = 0;
    }

    const dongTienTruocTaiTro =
      tu02.dongTienKH - tongChiSauVAT - vatPhaiNop - tu02.thueTNDN;
    const laiVay = duNoCuoiKy * (Math.pow(1 + laiSuatNam, 1 / 12) - 1);
    const nhuCauVon = Math.max(0, -(tienCuoiKy + dongTienTruocTaiTro - laiVay));
    const vonGopCSH = nhuCauVon * (1 - tyLeVay);
    const giaiNganVay = nhuCauVon * tyLeVay;
    const tienSauTaiTroTruocGoc =
      tienCuoiKy + dongTienTruocTaiTro + vonGopCSH + giaiNganVay - laiVay;
    const traGoc = t < thangBatDauTraGoc
      ? 0
      : Math.min(duNoCuoiKy + giaiNganVay, Math.max(0, tienSauTaiTroTruocGoc));
    const duNoMoi = Math.max(0, duNoCuoiKy + giaiNganVay - traGoc);
    const dongTienSauTaiTro =
      dongTienTruocTaiTro + vonGopCSH + giaiNganVay - laiVay - traGoc;
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

function FS03P_taoNhomNguyenGia_(sanPham, chiPhiChung) {
  const tongCPXDCoSo = FS03V21_tong_(sanPham.map(sp => sp.dtkd * sp.cpxdM2));
  const tongDienTichDat = FS03V21_tong_(sanPham.map(sp => sp.dienTichDat));

  const tongGPMB = FS03V21_layChiPhi_(chiPhiChung, ['Chi phí GPMB']);
  const tongHTKT = FS03V21_layChiPhi_(chiPhiChung, ['Chi phí HTKT']);
  const tyLeDuPhong = FS03V21_layTyLe_(chiPhiChung, ['Chi phí dự phòng']);

  const vatXD = FS03V21_layVAT_(chiPhiChung, [
    'Chi phí xây dựng & thiết bị',
    'Chi phí xây dựng và thiết bị',
    'Chi phí XD/TB/khác'
  ]);
  const vatGPMB = FS03V21_layVAT_(chiPhiChung, ['Chi phí GPMB']);
  const vatHTKT = FS03V21_layVAT_(chiPhiChung, ['Chi phí HTKT']);
  const vatDuPhong = FS03V21_layVAT_(chiPhiChung, ['Chi phí dự phòng']);

  const landSources = FS03P_docNguonTienDat_(chiPhiChung);
  const hasSpecificLand = Object.keys(landSources.specific)
    .some(k => landSources.specific[k].exists);

  const categoryLandArea = {};
  sanPham.forEach(sp => {
    const cat = FS03P_landCategory_(sp.ten);
    if (!cat) return;
    categoryLandArea[cat] = (categoryLandArea[cat] || 0) + Math.max(0, sp.dienTichDat || 0);
  });

  const ds = {};
  sanPham.forEach(sp => {
    const ma = FS03V21_key_(sp.ten);
    const coSoCPXD = sp.dtkd * sp.cpxdM2;
    const tyTrongCPXD = tongCPXDCoSo ? coSoCPXD / tongCPXDCoSo : 0;
    const tyTrongDatTong = tongDienTichDat ? sp.dienTichDat / tongDienTichDat : 0;

    const chiXD = coSoCPXD;
    const chiHTKT = tongHTKT * tyTrongCPXD;
    const chiGPMB = tongGPMB * tyTrongDatTong;

    const category = FS03P_landCategory_(sp.ten);
    let chiDat = 0;
    let vatDatRate = 0;

    if (hasSpecificLand && category && landSources.specific[category] && landSources.specific[category].exists) {
      const source = landSources.specific[category];
      const categoryArea = categoryLandArea[category] || 0;
      const weight = categoryArea
        ? Math.max(0, sp.dienTichDat || 0) / categoryArea
        : FS03P_equalWeightInCategory_(sanPham, sp, category);
      chiDat = source.value * weight;
      vatDatRate = source.vatRate;
    } else if (!hasSpecificLand && landSources.general.exists) {
      chiDat = landSources.general.value * tyTrongDatTong;
      vatDatRate = landSources.general.vatRate;
    }

    const chiDuPhong = (chiXD + chiHTKT) * tyLeDuPhong;

    ds[ma] = {
      ten: sp.ten,
      hinhThuc: sp.hinhThuc,
      landCategory: category,
      nhomTaiSan: FS03V21_laChoThue_(sp.hinhThuc) ? 'TAI_SAN_DAU_TU' : 'TAI_SAN_DE_BAN',
      chiXD,
      chiGPMB,
      chiHTKT,
      chiDat,
      chiDuPhong,
      vatXD: chiXD * vatXD,
      vatGPMB: chiGPMB * vatGPMB,
      vatHTKT: chiHTKT * vatHTKT,
      vatDat: chiDat * vatDatRate,
      vatDuPhong: chiDuPhong * vatDuPhong,
      tongTruocLaiVay: chiXD + chiGPMB + chiHTKT + chiDat + chiDuPhong,
      thoiGianPhanBo: Math.max(1, sp.thoiGianThue || 600)
    };
  });

  return ds;
}

function FS03P_tinhChiPhiTheoTienDo_(soThang, tienDo, nhomNguyenGia, tyLeTruotGiaNam) {
  const out = {};
  for (let t = 1; t <= soThang; t++) out[t] = FS03V21_mauChiPhiThang_();

  const tongPool = FS03V21_congNhomNguyenGia_(nhomNguyenGia);
  const danhMuc = [
    { ten: 'Chi phí XD/TB/khác', truong: 'chiXD', truongVAT: 'vatXD', tong: tongPool.chiXD, vat: tongPool.vatXD, truotGia: true },
    { ten: 'Chi phí GPMB', truong: 'chiGPMB', truongVAT: 'vatGPMB', tong: tongPool.chiGPMB, vat: tongPool.vatGPMB, truotGia: false },
    { ten: 'Chi phí HTKT', truong: 'chiHTKT', truongVAT: 'vatHTKT', tong: tongPool.chiHTKT, vat: tongPool.vatHTKT, truotGia: true },
    { ten: 'Chi phí dự phòng', truong: 'chiDuPhong', truongVAT: 'vatDuPhong', tong: tongPool.chiDuPhong, vat: tongPool.vatDuPhong, truotGia: false }
  ];

  danhMuc.forEach(dm => {
    const rows = tienDo.filter(td => FS03V21_cungTenKhoanMuc_(td.khoanMuc, dm.ten));
    FS03P_apDungTienDo_(out, soThang, rows, dm, tyLeTruotGiaNam);
  });

  const landPools = FS03P_congPoolTienDat_(nhomNguyenGia);
  const landRows = tienDo
    .map(td => ({ td, category: FS03P_landScheduleCategory_(td.khoanMuc) }))
    .filter(x => x.category);

  const hasSpecificSchedule = landRows.some(x => x.category !== 'GENERAL');
  if (hasSpecificSchedule) {
    Object.keys(landPools).forEach(category => {
      if (category === 'GENERAL') return;
      const pool = landPools[category];
      const rows = landRows.filter(x => x.category === category).map(x => x.td);
      FS03P_apDungTienDo_(out, soThang, rows, {
        ten: category,
        truong: 'chiDat',
        truongVAT: 'vatDat',
        tong: pool.chiDat,
        vat: pool.vatDat,
        truotGia: false
      }, tyLeTruotGiaNam);
    });
  } else {
    const rows = landRows.filter(x => x.category === 'GENERAL').map(x => x.td);
    FS03P_apDungTienDo_(out, soThang, rows, {
      ten: 'Tiền SDĐ/thuê đất',
      truong: 'chiDat',
      truongVAT: 'vatDat',
      tong: tongPool.chiDat,
      vat: tongPool.vatDat,
      truotGia: false
    }, tyLeTruotGiaNam);
  }

  return out;
}

function FS03P_apDungTienDo_(out, soThang, rows, dm, tyLeTruotGiaNam) {
  if (!rows.length || !dm.tong) return;

  rows.forEach(td => {
    const batDau = Math.max(1, Number(td.thangBatDau) || 1);
    const thoiGian = Math.max(1, Number(td.thoiGian) || 1);
    const tyLe = FS03V21_tyLe_(td.tyLe) || 0;
    const loai = FS03V21_key_(td.loai);
    const ketThuc = Math.min(soThang, batDau + thoiGian - 1);
    const soKy = Math.max(1, ketThuc - batDau + 1);

    for (let t = batDau; t <= ketThuc; t++) {
      const heSoTruotGia = dm.truotGia
        ? Math.pow(1 + tyLeTruotGiaNam, (t - 1) / 12)
        : 1;
      const giaTri = loai === FS03V21_key_('Một lần')
        ? (t === batDau ? dm.tong * tyLe * heSoTruotGia : 0)
        : dm.tong * tyLe / soKy * heSoTruotGia;
      const vat = dm.tong ? giaTri * (dm.vat / dm.tong) : 0;

      out[t][dm.truong] += giaTri;
      out[t][dm.truongVAT] += vat;
    }
  });
}

function FS03P_docNguonTienDat_(chiPhiChung) {
  return {
    general: FS03P_findCost_(chiPhiChung, ['Tiền SDĐ/thuê đất']),
    specific: {
      SDD_LIEN_KE: FS03P_findCost_(chiPhiChung, [
        'Tiền SDĐ liền kề',
        'Tiền SDĐ đất liền kề',
        'Tiền SD đất liền kề',
        'Tiền sử dụng đất liền kề'
      ]),
      THUE_DAT_CHUNG_CU: FS03P_findCost_(chiPhiChung, [
        'Tiền thuê đất chung cư',
        'Tiền thuê đất căn hộ'
      ]),
      THUE_DAT_TMDV: FS03P_findCost_(chiPhiChung, [
        'Tiền thuê đất TMDV',
        'Tiền thuê đất thương mại dịch vụ'
      ]),
      THUE_DAT_CHO: FS03P_findCost_(chiPhiChung, ['Tiền thuê đất Chợ'])
    }
  };
}

function FS03P_findCost_(map, aliases) {
  for (const name of aliases) {
    const item = map[FS03V21_key_(name)];
    if (item) {
      return {
        exists: true,
        value: Number(item.truocVAT) || 0,
        vatRate: FS03V21_tyLe_(item.vat)
      };
    }
  }
  return { exists: false, value: 0, vatRate: 0 };
}

function FS03P_landCategory_(productName) {
  const k = FS03V21_key_(productName);
  if (FS03V21_chua_(k, ['lienke'])) return 'SDD_LIEN_KE';
  if (FS03V21_chua_(k, ['chungcu', 'canho'])) return 'THUE_DAT_CHUNG_CU';
  if (FS03V21_chua_(k, ['tmdv', 'thuongmai', 'thuongmaidichvu'])) return 'THUE_DAT_TMDV';
  if (FS03V21_chua_(k, ['cho'])) return 'THUE_DAT_CHO';
  return '';
}

function FS03P_landScheduleCategory_(name) {
  const k = FS03V21_key_(name);
  if (k === FS03V21_key_('Tiền SDĐ/thuê đất')) return 'GENERAL';
  if (FS03V21_chua_(k, ['tiensddlienke', 'tiensudungdatlienke', 'tiensddatlienke'])) return 'SDD_LIEN_KE';
  if (FS03V21_chua_(k, ['tienthuedatchungcu', 'tienthuedatcanho'])) return 'THUE_DAT_CHUNG_CU';
  if (FS03V21_chua_(k, ['tienthuedattmdv', 'tienthuedatthuongmaidichvu'])) return 'THUE_DAT_TMDV';
  if (FS03V21_chua_(k, ['tienthuedatcho'])) return 'THUE_DAT_CHO';
  return '';
}

function FS03P_congPoolTienDat_(nhomNguyenGia) {
  const out = {
    SDD_LIEN_KE: { chiDat: 0, vatDat: 0 },
    THUE_DAT_CHUNG_CU: { chiDat: 0, vatDat: 0 },
    THUE_DAT_TMDV: { chiDat: 0, vatDat: 0 },
    THUE_DAT_CHO: { chiDat: 0, vatDat: 0 }
  };

  Object.keys(nhomNguyenGia).forEach(key => {
    const p = nhomNguyenGia[key];
    const category = p.landCategory;
    if (!category || !out[category]) return;
    out[category].chiDat += Number(p.chiDat) || 0;
    out[category].vatDat += Number(p.vatDat) || 0;
  });

  return out;
}

function FS03P_equalWeightInCategory_(sanPham, current, category) {
  const same = sanPham.filter(sp => FS03P_landCategory_(sp.ten) === category);
  if (!same.length) return 0;
  return 1 / same.length;
}
