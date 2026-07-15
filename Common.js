const FS_CFG = Object.freeze({
  SHEETS: { TECH:'01. Kỹ thuật', REVENUE:'02. Doanh thu', COST:'03. Chi phí & Vốn', PROFIT:'03A. Lợi nhuận & Thuế', CASH:'04. Dòng tiền & Lợi nhuận', SUMMARY_CASH:'04A. TH dòng tiền', SUMMARY:'00. Tổng hợp', CHECKS:'99. Checks' },
  CONVERGENCE: { tolerance: 1, maxIterations: 100 }
});
function FS_getOrCreateSheet_(ss,name){return ss.getSheetByName(name)||ss.insertSheet(name);}
function FS_resetSheet_(sh,rows,cols){if(sh.getMaxRows()<rows)sh.insertRowsAfter(sh.getMaxRows(),rows-sh.getMaxRows());if(sh.getMaxColumns()<cols)sh.insertColumnsAfter(sh.getMaxColumns(),cols-sh.getMaxColumns());sh.showRows(1,sh.getMaxRows());sh.showColumns(1,sh.getMaxColumns());sh.getRange(1,1,sh.getMaxRows(),sh.getMaxColumns()).breakApart();sh.clearContents();sh.clearFormats();}
function FS_norm_(v){return String(v??'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/đ/g,'d').replace(/Đ/g,'D').toLowerCase().replace(/[^a-z0-9]+/g,' ').trim();}
function FS_key_(v){return FS_norm_(v).replace(/\s+/g,'');}
function FS_num_(v){if(typeof v==='number')return isFinite(v)?v:0;const n=Number(String(v??'').replace(/\s/g,'').replace(/,/g,''));return isFinite(n)?n:0;}
function FS_rate_(v){if(typeof v==='number')return Math.abs(v)>1?v/100:v;const s=String(v??'').trim();const n=Number(s.replace('%','').replace(',','.'));return isFinite(n)?(s.includes('%')||Math.abs(n)>1?n/100:n):0;}
function FS_addMonths_(d,m){const x=new Date(d);return new Date(x.getFullYear(),x.getMonth()+m,x.getDate());}
function FS_findExactRow_(sh,text){const target=FS_norm_(text);const vals=sh.getRange(1,1,sh.getLastRow(),1).getDisplayValues();for(let i=0;i<vals.length;i++)if(FS_norm_(vals[i][0])===target)return i+1;return 0;}
function FS_readInfo_(tech){const r=FS_findExactRow_(tech,'THONG_TIN_CHUNG');if(!r)throw new Error('Không tìm thấy THONG_TIN_CHUNG.');const o={};for(let i=r+1;i<=tech.getLastRow();i++){const k=String(tech.getRange(i,1).getDisplayValue()||'').trim();if(!k||/^[A-Z_]+$/.test(k))break;o[k]=tech.getRange(i,2).getValue();}return o;}
function FS_readBlock_(tech,name){const r=FS_findExactRow_(tech,name);if(!r)throw new Error('Không tìm thấy block '+name);const headerRow=r+1;const lastCol=Math.max(1,tech.getRange(headerRow,1,1,tech.getLastColumn()).getDisplayValues()[0].reduce((n,v,i)=>String(v).trim()?i+1:n,0));const rows=[];for(let i=r+2,blank=0;i<=tech.getLastRow();i++){const first=String(tech.getRange(i,1).getDisplayValue()||'').trim();if(/^[A-Z_]+$/.test(first))break;const row=tech.getRange(i,1,1,lastCol).getValues()[0];if(!row.some(v=>String(v??'').trim()!=='')){if(++blank>=3)break;continue;}blank=0;rows.push(row);}return rows;}
function FS_costMap_(tech){const m={};FS_readBlock_(tech,'CHI_PHI_CHUNG').forEach(r=>{const n=String(r[0]||'').trim();if(n)m[FS_key_(n)]={name:n,before:FS_num_(r[1]),vat:FS_rate_(r[2]),after:FS_num_(r[3]),note:String(r[4]||''),rate:FS_rate_(r[5])};});return m;}
function FS_costItem_(m,names){for(const n of names){const x=m[FS_key_(n)];if(x)return x;}return {before:0,vat:0,rate:0,name:names[0]};}
function FS_productType_(name,method){const k=FS_key_(name),mk=FS_key_(method);if(mk.includes('chothue'))return 'RENT';if(k.includes('tmdv')&&mk.includes('ban'))return 'SALE';return mk.includes('ban')?'SALE':'RENT';}
