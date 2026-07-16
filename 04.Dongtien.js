const FS04_CFG = Object.freeze({
  TECH: '01A. Kỹ thuật',
  COST: '03. Chi phí & Vốn',
  CASH: '04. Dòng tiền & Tài trợ',
  CASH_LEGACY: '04. Dòng tiền'
});

function FS_lapSheet04() {
  return FS_hoiTuTaiTro();
}

function FS04_build_(taxByMonth, writeSheet) {
  const ss = SpreadsheetApp.getActive();
  const tech = ss.getSheetByName(FS04_CFG.TECH);
  const costSheet = ss.getSheetByName(FS04_CFG.COST);
  if (!tech || !costSheet) throw new Error('Cần lập "01A. Kỹ thuật" và "03. Chi phí & Vốn" trước.');

  const months = Math.max(0, FS04_num_(FS04_readInfoValue_(tech, 'Số tháng mô hình')));
  const loanRatio = FS04_rate_(FS04_readInfoValue_(tech, 'Tỷ lệ vốn vay'));
  const annualInterestRate = FS04_rate_(FS04_readInfoValue_(tech, 'Lãi suất vay năm'));
  if (!months) throw new Error('Số tháng mô hình phải lớn hơn 0.');
  if (loanRatio < 0 || loanRatio > 1) throw new Error('Tỷ lệ vốn vay phải nằm trong khoảng 0% đến 100%.');
  if (annualInterestRate < 0) throw new Error('Lãi suất vay năm không được âm.');

  const monthlyRate = Math.pow(1 + annualInterestRate, 1 / 12) - 1;
  const costByMonth = FS04_readCostByMonth_(costSheet, months);
  const rows = [];
  const interestByMonth = Array(months).fill(0);

  let openingCash = 0;
  let openingDebt = 0;
  let openingVatCredit = 0;

  for (let monthNo = 1; monthNo <= months; monthNo++) {
    const current = costByMonth[monthNo] || FS04_emptyCostMonth_(monthNo);
    const cit = FS04_num_((taxByMonth || [])[monthNo - 1]);
    const vatPayable = Math.max(0, current.vatOut - openingVatCredit - current.vatIn);
    const closingVatCredit = Math.max(0, openingVatCredit + current.vatIn - current.vatOut);

    const fcff = current.customerCash - current.costAfterVat - vatPayable - cit;
    const interest = openingDebt * monthlyRate;
    interestByMonth[monthNo - 1] = interest;
    const cashBeforeFinancing = openingCash + fcff - interest;

    let fundingNeed = 0;
    let equityContribution = 0;
    let loanDrawdown = 0;
    let principalRepayment = 0;
    let closingDebt = openingDebt;
    let closingCash = 0;

    if (cashBeforeFinancing < 0) {
      fundingNeed = -cashBeforeFinancing;
      loanDrawdown = fundingNeed * loanRatio;
      equityContribution = fundingNeed - loanDrawdown;
      closingDebt = openingDebt + loanDrawdown;
      closingCash = 0;
    } else {
      principalRepayment = Math.min(openingDebt, cashBeforeFinancing);
      closingDebt = Math.max(0, openingDebt - principalRepayment);
      closingCash = Math.max(0, cashBeforeFinancing - principalRepayment);
    }

    // FCFE là dòng tiền của riêng kỳ, không phải số dư tiền lũy kế.
    // Tiền tồn cuối kỳ vẫn chuyển nguyên sang kỳ sau để ưu tiên chi phí trước khi huy động mới.
    const fcfe = closingCash - openingCash - equityContribution;

    rows.push([
      monthNo, current.date, current.year, current.quarter,
      current.customerCash, current.vatOut,
      current.costBeforeVat, current.vatIn, current.costAfterVat,
      openingVatCredit, vatPayable, closingVatCredit,
      cit, fcff, openingCash, cashBeforeFinancing, fundingNeed,
      interest, equityContribution, loanDrawdown, principalRepayment,
      openingDebt, closingDebt, closingCash, fcfe
    ]);

    openingCash = closingCash;
    openingDebt = closingDebt;
    openingVatCredit = closingVatCredit;
  }

  if (writeSheet !== false) FS04_write_(ss, rows);
  return { rows, interestByMonth, monthlyRate, loanRatio };
}

function FS04_readCostByMonth_(sheet, months) {
  const table = FS04_readTable_(sheet);
  const required = ['thangso','thang','nam','quy','dongtienkhachhang','vatdaura','tongchitruocvat','vatdauvao','tongchisauvat'];
  FS04_require_(table.index, required, '03. Chi phí & Vốn');
  const out = {};
  for (let i=1;i<=months;i++) out[i]=FS04_emptyCostMonth_(i);
  table.values.forEach(row => {
    const m=FS04_num_(row[table.index.thangso]);
    if(m<1||m>months)return;
    const x=out[m];
    if(!x.date)x.date=row[table.index.thang];
    if(!x.year)x.year=row[table.index.nam];
    if(!x.quarter)x.quarter=row[table.index.quy];
    x.customerCash+=FS04_num_(row[table.index.dongtienkhachhang]);
    x.vatOut+=FS04_num_(row[table.index.vatdaura]);
    x.costBeforeVat+=FS04_num_(row[table.index.tongchitruocvat]);
    x.vatIn+=FS04_num_(row[table.index.vatdauvao]);
    x.costAfterVat+=FS04_num_(row[table.index.tongchisauvat]);
  });
  return out;
}

function FS04_emptyCostMonth_(m) {
  return {monthNo:m,date:'',year:'',quarter:'',customerCash:0,vatOut:0,costBeforeVat:0,vatIn:0,costAfterVat:0};
}

function FS04_write_(ss, rows) {
  let sh=ss.getSheetByName(FS04_CFG.CASH);
  const legacy=ss.getSheetByName(FS04_CFG.CASH_LEGACY);
  if(!sh&&legacy){legacy.setName(FS04_CFG.CASH);sh=legacy;}
  if(!sh)sh=ss.insertSheet(FS04_CFG.CASH);
  sh.clear();sh.clearFormats();
  const headers=[[
    'Tháng số','Tháng','Năm','Quý','Dòng tiền khách hàng','VAT đầu ra',
    'Tổng chi trước VAT','VAT đầu vào','Tổng chi sau VAT',
    'VAT khấu trừ đầu kỳ','VAT phải nộp','VAT khấu trừ cuối kỳ',
    'Thuế TNDN','FCFF','Tiền tồn đầu kỳ','Tiền trước tài trợ','Nhu cầu vốn',
    'Lãi vay','Vốn góp CSH','Giải ngân vay','Trả gốc',
    'Dư nợ đầu kỳ','Dư nợ cuối kỳ','Tiền tồn cuối kỳ','FCFE'
  ]];
  sh.getRange(1,1,1,headers[0].length).setValues(headers);
  if(rows.length)sh.getRange(2,1,rows.length,headers[0].length).setValues(rows);
  sh.setFrozenRows(1);sh.setFrozenColumns(4);
  sh.getRange(1,1,1,headers[0].length).setFontWeight('bold').setBackground('#ddebf7').setWrap(true);
  if(rows.length){
    sh.getRange(2,2,rows.length,1).setNumberFormat('MM/yyyy');
    sh.getRange(2,5,rows.length,headers[0].length-4).setNumberFormat('#,##0');
  }
  sh.autoResizeColumns(1,headers[0].length);
}

function FS04_readTable_(sheet){
  const cols=sheet.getLastColumn();
  const headers=sheet.getRange(1,1,1,cols).getDisplayValues()[0];
  const values=sheet.getLastRow()>1?sheet.getRange(2,1,sheet.getLastRow()-1,cols).getValues():[];
  const index={};headers.forEach((h,i)=>{index[FS04_key_(h)]=i;});
  return {values,index};
}
function FS04_require_(index,required,name){const missing=required.filter(k=>index[k]==null);if(missing.length)throw new Error('Sheet "'+name+'" thiếu cột: '+missing.join(', '));}
function FS04_readInfoValue_(sheet,label){const target=FS04_key_(label),v=sheet.getDataRange().getValues();for(let i=0;i<v.length;i++)if(FS04_key_(v[i][0])===target)return v[i][1];return '';}
function FS04_num_(v){if(typeof v==='number')return isFinite(v)?v:0;const t=String(v==null?'':v).trim().replace(/\s/g,'');if(!t)return 0;const n=Number(t.includes(',')&&t.includes('.')?t.replace(/\./g,'').replace(',','.'):t.replace(/,/g,''));return isFinite(n)?n:0;}
function FS04_rate_(v){if(typeof v==='number')return v>1?v/100:v;const t=String(v==null?'':v).trim();if(!t)return 0;const n=FS04_num_(t.replace('%',''));return t.includes('%')||n>1?n/100:n;}
function FS04_norm_(v){return String(v==null?'':v).toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/đ/g,'d').replace(/\s+/g,' ').trim();}
function FS04_key_(v){return FS04_norm_(v).replace(/²/g,'2').replace(/\^2/g,'2').replace(/m\s*2/g,'m2').replace(/[^a-z0-9]/g,'');}