/*************************************************
 * 05B - Độ nhạy vốn CSH 11x11 - Batch Runner
 * Trục ngang: tăng/giảm lãi suất vay ±2 điểm %
 * Trục dọc: tăng/giảm vốn đầu tư
 *************************************************/

function FS05B_DoNhay_CSH() {
  FS05B_CSH_11x11_START();
}

function FS05B_CSH_11x11_START() {
  const ss = SpreadsheetApp.getActive();
  const tech = ss.getSheetByName('01. Kỹ thuật');
  const sh = ss.getSheetByName('05. Độ nhạy');
  const tongHop = ss.getSheetByName('00. Tổng hợp');

  if (!tech) throw new Error('Không tìm thấy sheet 01. Kỹ thuật.');
  if (!sh) throw new Error('Không tìm thấy sheet 05. Độ nhạy.');
  if (!tongHop) throw new Error('Không tìm thấy sheet 00. Tổng hợp.');

  FS05B_CSH_11x11_CLEAR_TRIGGERS_();

  const costFactors = [-0.10, -0.08, -0.05, -0.04, -0.02, 0, 0.02, 0.04, 0.05, 0.08, 0.10];

  // ±2 điểm %, chia 11 mức
  const loanDeltas = [-0.10, -0.08, -0.05, -0.04, -0.02, 0, 0.02, 0.04, 0.05, 0.08, 0.10];

  const loanCell = FS05B_FIND_VALUE_CELL_(tech, 'Lãi suất vay năm');
  const costCells = FS05B_FIND_COST_CELLS_(tech);

  const saved = {
    loan: FS05B_SAVE_CELL_(loanCell),
    costs: costCells.map(FS05B_SAVE_CELL_)
  };

  const baseNPV = Number(tongHop.getRange('D36').getValue()) || 0;
  const baseIRR = Number(tongHop.getRange('D37').getValue()) || 0;

  FS05B_WRITE_SHELL_11_(sh, 36, 'NPV VỐN CSH', baseNPV, costFactors, loanDeltas, '#,##0.0');
  FS05B_WRITE_SHELL_11_(sh, 52, 'IRR VỐN CSH', baseIRR, costFactors, loanDeltas, '0.0%');

  const state = {
    r: 0,
    c: 0,
    costFactors,
    loanDeltas,
    baseNPV,
    baseIRR,
    saved,
    startedAt: new Date().toISOString()
  };

  PropertiesService.getDocumentProperties().setProperty('FS05B_CSH_STATE', JSON.stringify(state));

  sh.getRange('O36').setValue('Đang chạy độ nhạy vốn CSH 11x11...');
  sh.getRange('O37').setValue('Tiến độ: 0/121');

  FS05B_CSH_11x11_RUN();
}

function FS05B_CSH_11x11_RUN() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) return;

  try {
    const props = PropertiesService.getDocumentProperties();
    const raw = props.getProperty('FS05B_CSH_STATE');
    if (!raw) return;

    const state = JSON.parse(raw);

    const ss = SpreadsheetApp.getActive();
    const tech = ss.getSheetByName('01. Kỹ thuật');
    const sh = ss.getSheetByName('05. Độ nhạy');
    const tongHop = ss.getSheetByName('00. Tổng hợp');

    const loanCell = tech.getRange(state.saved.loan.a1);
    const costCells = state.saved.costs.map(x => tech.getRange(x.a1));

    const startTime = Date.now();
    const maxMillis = 260000; // chạy tối đa khoảng 4 phút 20 giây/lượt

    let done = false;

    while (!done && Date.now() - startTime < maxMillis) {
      const r = state.r;
      const c = state.c;

      if (r >= state.costFactors.length) {
        done = true;
        break;
      }

      const costDelta = state.costFactors[r];
      const loanDelta = state.loanDeltas[c];

      // Ô trung tâm dùng Base, không cần rebuild
      if (r === 5 && c === 5) {
        sh.getRange(36 + 3 + r, 3 + c).setValue(state.baseNPV);
        sh.getRange(52 + 3 + r, 3 + c).setValue(state.baseIRR);
      } else {
        loanCell.setValue(Math.max(0, Number(state.saved.loan.value) * (1 + loanDelta)));

        state.saved.costs.forEach((item, i) => {
          costCells[i].setValue(Number(item.value) * (1 + costDelta));
        });

        SpreadsheetApp.flush();

        FS05B_REBUILD_MODEL_FAST_NO_ALERT_();

        const npv = Number(tongHop.getRange('D36').getValue()) || 0;
        const irr = Number(tongHop.getRange('D37').getValue()) || 0;

        sh.getRange(36 + 3 + r, 3 + c).setValue(npv);
        sh.getRange(52 + 3 + r, 3 + c).setValue(irr);

        FS05B_COLOR_CELL_(sh.getRange(36 + 3 + r, 3 + c), npv, state.baseNPV);
        FS05B_COLOR_CELL_(sh.getRange(52 + 3 + r, 3 + c), irr, state.baseIRR);
      }

      state.c++;

      if (state.c >= state.loanDeltas.length) {
        state.c = 0;
        state.r++;
      }

      const finished = Math.min(state.r * 11 + state.c, 121);
      sh.getRange('O37').setValue('Tiến độ: ' + finished + '/121');

      props.setProperty('FS05B_CSH_STATE', JSON.stringify(state));
    }

    if (done || state.r >= state.costFactors.length) {
      FS05B_RESTORE_CELL_(state.saved.loan);
      state.saved.costs.forEach(FS05B_RESTORE_CELL_);

      SpreadsheetApp.flush();
      FS05B_REBUILD_MODEL_FAST_NO_ALERT_();

      sh.getRange('O36').setValue('Đã chạy xong độ nhạy vốn CSH 11x11.');
      sh.getRange('O37').setValue('Tiến độ: 121/121');

      FS05B_CSH_11x11_CLEAR_TRIGGERS_();
      props.deleteProperty('FS05B_CSH_STATE');
      return;
    }

    FS05B_CSH_11x11_CREATE_TRIGGER_();

  } finally {
    lock.releaseLock();
  }
}

function FS05B_CSH_11x11_STOP() {
  const props = PropertiesService.getDocumentProperties();
  const raw = props.getProperty('FS05B_CSH_STATE');

  if (raw) {
    const state = JSON.parse(raw);
    FS05B_RESTORE_CELL_(state.saved.loan);
    state.saved.costs.forEach(FS05B_RESTORE_CELL_);
    props.deleteProperty('FS05B_CSH_STATE');
  }

  FS05B_CSH_11x11_CLEAR_TRIGGERS_();
  SpreadsheetApp.getUi().alert('Đã dừng batch độ nhạy vốn CSH.');
}

function FS05B_REBUILD_MODEL_FAST_NO_ALERT_() {
  FS_lapSheet03();
  SpreadsheetApp.flush();

  FS_lapSheet02();
  SpreadsheetApp.flush();

  FS_lapSheet04();
  SpreadsheetApp.flush();

  FS03_capNhatNguonVonTuSheet04();
  SpreadsheetApp.flush();

  FS_lapSheet02();
  SpreadsheetApp.flush();

  FS_lapSheet04();
  SpreadsheetApp.flush();

  FS03_capNhatNguonVonTuSheet04();
  SpreadsheetApp.flush();

  FS_lapSheet04A();
  FS_lapSheet00();
  SpreadsheetApp.flush();
}

function FS05B_WRITE_SHELL_11_(sh, startRow, title, base, costFactors, loanDeltas, fmt) {
  const rows = 11;
  const cols = 11;

  sh.getRange(startRow, 1, 14, 13).breakApart().clearContent().clearFormat();

  sh.getRange(startRow, 1, 1, 13)
    .setBackground('#FFC000')
    .setFontWeight('bold');

  sh.getRange(startRow, 2)
    .setValue(title)
    .setFontWeight('bold')
    .setFontSize(12);

  sh.getRange(startRow + 1, 3, 1, cols)
    .merge()
    .setValue('Tăng/giảm lãi suất vay')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  sh.getRange(startRow + 2, 2)
    .setValue(base)
    .setNumberFormat(fmt)
    .setBackground('#008000')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');

  sh.getRange(startRow + 2, 3, 1, cols)
    .setValues([loanDeltas])
    .setNumberFormat('0.0%')
    .setBackground('#A6A6A6')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');

  sh.getRange(startRow + 3, 1, rows, 1)
    .merge()
    .setValue('Tăng/giảm\nvốn đầu tư')
    .setFontWeight('bold')
    .setWrap(true)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  sh.getRange(startRow + 3, 2, rows, 1)
    .setValues(costFactors.map(x => [x]))
    .setNumberFormat('0%')
    .setFontWeight('bold');

  sh.getRange(startRow + 3, 3, rows, cols)
    .setNumberFormat(fmt)
    .setBackground('#FFFFFF');

  sh.getRange(startRow + 8, 8)
    .setValue(base)
    .setNumberFormat(fmt)
    .setBackground('#FFFF00')
    .setFontWeight('bold')
    .setBorder(true, true, true, true, true, true, '#FF0000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  sh.getRange(startRow + 2, 2, rows + 1, cols + 1)
    .setBorder(true, true, true, true, true, true);

  sh.getRange(startRow + 3, 1, rows, cols + 2)
    .setBorder(true, true, true, true, true, true);

  sh.getRange(startRow, 1, 14, 13)
    .setFontFamily('Times New Roman')
    .setFontSize(10)
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('center');
}

function FS05B_COLOR_CELL_(range, value, base) {
  range.setBackground(Number(value) < Number(base) ? '#E6B8B7' : '#FFFFFF');
}

function FS05B_CSH_11x11_CREATE_TRIGGER_() {
  FS05B_CSH_11x11_CLEAR_TRIGGERS_();

  ScriptApp.newTrigger('FS05B_CSH_11x11_RUN')
    .timeBased()
    .after(60 * 1000)
    .create();
}

function FS05B_CSH_11x11_CLEAR_TRIGGERS_() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'FS05B_CSH_11x11_RUN') {
      ScriptApp.deleteTrigger(t);
    }
  });
}

function FS05B_FIND_VALUE_CELL_(sheet, label) {
  const data = sheet.getDataRange().getDisplayValues();
  const target = FS05B_NORM_(label);

  for (let r = 0; r < data.length; r++) {
    if (data[r].some(v => FS05B_NORM_(v) === target)) {
      return sheet.getRange(r + 1, 2);
    }
  }

  throw new Error('Không tìm thấy: ' + label);
}

function FS05B_FIND_COST_CELLS_(tech) {
  const targetNames = [
    'Chi phí XD/TB/khác',
    'Chi phí GPMB',
    'Tiền SDĐ/thuê đất',
    'Chi phí HTKT'
  ].map(FS05B_NORM_);

  const data = tech.getDataRange().getDisplayValues();
  const cells = [];
  let inBlock = false;

  for (let r = 0; r < data.length; r++) {
    const first = FS05B_NORM_(data[r][0]);

    if (first === 'chi_phi_chung') {
      inBlock = true;
      continue;
    }

    if (inBlock && ['san_pham', 'ke_hoach_ban_thu_tien', 'tien_do_chi_phi'].includes(first)) {
      break;
    }

    if (inBlock && targetNames.includes(first)) {
      cells.push(tech.getRange(r + 1, 2));
    }
  }

  if (cells.length === 0) {
    throw new Error('Không tìm thấy các dòng vốn đầu tư trong CHI_PHI_CHUNG.');
  }

  return cells;
}

function FS05B_SAVE_CELL_(range) {
  return {
    a1: range.getA1Notation(),
    value: range.getValue(),
    formula: range.getFormula()
  };
}

function FS05B_RESTORE_CELL_(item) {
  const sh = SpreadsheetApp.getActive().getSheetByName('01. Kỹ thuật');
  const range = sh.getRange(item.a1);

  if (item.formula) {
    range.setFormula(item.formula);
  } else {
    range.setValue(item.value);
  }
}

function FS05B_NORM_(v) {
  return String(v || '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd')
    .replace(/\s+/g, ' ')
    .trim();
}