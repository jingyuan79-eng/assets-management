// 用假的 SpreadsheetApp 在本地跑一遍 Apps Script，验证每个 action 都能返回结果
import fs from 'fs';
const src = fs.readFileSync('AppsScript.gs','utf8');

function mkSheet(name, rows){
  return {
    _rows: rows,
    getName:()=>name,
    getDataRange(){ const r=this._rows; return { getValues:()=>r.map(x=>x.slice()) }; },
    clear(){ this._rows.length=0; },
    getLastRow(){ return this._rows.length; },
    _cell(r,c){ while(this._rows.length<r) this._rows.push([]);
                const row=this._rows[r-1]; while(row.length<c) row.push(""); return row; },
    // 支持单格与区域两种用法：getRange(r,c) / getRange(r,c,nRows,nCols)
    getRange(r,c,nr,nc){
      const self=this; nr=nr||1; nc=nc||1;
      return {
        getValue:()=> (self._rows[r-1]||[])[c-1] ?? "",
        setValue(v){ self._cell(r,c)[c-1]=v; },
        setFormula(v){ self._cell(r,c)[c-1]=v; },
        getValues(){ const out=[];
          for(let i=0;i<nr;i++){ const row=[];
            for(let j=0;j<nc;j++) row.push((self._rows[r-1+i]||[])[c-1+j] ?? "");
            out.push(row); }
          return out; },
        setValues(vals){ for(let i=0;i<nr;i++){ const row=self._cell(r+i, c+nc-1);
            for(let j=0;j<nc;j++) row[c-1+j]=vals[i][j]; } },
        clearContent(){ for(let i=0;i<nr;i++){ const row=self._cell(r+i, c+nc-1);
            for(let j=0;j<nc;j++) row[c-1+j]=""; } }
      };
    },
    appendRow(a){ this._rows.push(a.slice()); },
    deleteRow(r){ this._rows.splice(r-1,1); }
  };
}
const stock = mkSheet('Stock',[
  ['Symbol','Category','Share','Cost','Price'],
  ['ASMIY','Semiconductor',25,888,969.26],
  ['NVDA','Semiconductor',32,148.71,214.72],
  ['Cash','Gold.coin.cash','','',1326],
]);
const ledger = mkSheet('Ledger',[
  ['Date','category','amount'],
  [new Date('2026-08-12T00:00:00-07:00'),'Dining',27.31],
  [new Date('2026-08-22T00:00:00-07:00'),'Grocery',109.09],
]);
const monthlyTab=mkSheet('Ledger_monthly',[]);
const savTab=mkSheet('Savings',[['ID','Name','Type','Balance','Rate','Last Post','Next update','Maturity','Status'],
  ['','OS-9250','OS',30000,3.4,'2026-08-01','2026-09-01','','active']]);
const tabs={Stock:stock,Ledger:ledger,Ledger_monthly:monthlyTab,Savings:savTab};
globalThis.__tabs=tabs;

globalThis.SpreadsheetApp={ flush(){}, getActiveSpreadsheet:()=>({ getSheetByName:n=>tabs[n]||null,
  insertSheet:n=>{ tabs[n]=mkSheet(n,[]); return tabs[n]; } }) };
globalThis.ContentService={ MimeType:{JSON:'json'},
  createTextOutput:t=>({ _t:t, setMimeType(){ return this; }, getContent(){ return this._t; } }) };
globalThis.Utilities={ formatDate:(d,tz,f)=>{
  const p=new Date(d.getTime()-7*3600e3).toISOString();
  return f==='yyyy-MM'? p.slice(0,7) : p.slice(0,10); } };
globalThis.LockService={getScriptLock:()=>({waitLock(){},releaseLock(){}})};
globalThis.SpreadsheetApp_flush=1;
globalThis.Logger={ log:(...a)=>console.log('  ',...a) };

const runner = new Function('return (function(){' + src + '\nreturn {doGet,doPost,selfTest,readAll};})()');
const {doGet,doPost,selfTest} = runner();

const J = r => JSON.parse(r.getContent());
const ok = (label, cond, extra='') => console.log((cond?'✅':'❌'), label, extra);

// 1. 读取
let r = J(doGet({parameter:{}}));
ok('doGet 读取', r.status==='success', `stock=${r.stock.length} ledger=${r.ledger.length} 类别=${JSON.stringify(r.expense)}`);

// 2. Shortcut 记账（POST）
r = J(doPost({postData:{contents:JSON.stringify({action:'ledger',category:'Dining',amount:12.5})}}));
ok('Shortcut 记账', r.status==='success', `→ ${r.category} $${r.amount} 行${r.row}`);

// 3. category 为空时的兜底
r = J(doPost({postData:{contents:JSON.stringify({action:'ledger',amount:20})}}));
ok('空分类兜底', r.category==='UNCATEGORIZED', `→ ${r.category}`);

// 4. 固定支出去重
r = J(doGet({parameter:{action:'autoLedger',key:'fix:网费:2026-08',category:'Bill & utilities',amount:70}}));
ok('固定支出首次写入', r.status==='success');
r = J(doGet({parameter:{action:'autoLedger',key:'fix:网费:2026-08',category:'Bill & utilities',amount:70}}));
ok('固定支出重复被拦截', r.status==='skipped', `→ ${r.status}`);

// 5. 改 / 删
const before = ledger._rows.length;
r = J(doGet({parameter:{action:'updateLedger',row:2,amount:99,expectAmount:27.31}}));
ok('改一笔支出', r.status==='success' && ledger._rows[1][2]===99);
r = J(doGet({parameter:{action:'updateLedger',row:2,amount:1,expectAmount:27.31}}));
ok('金额对不上时拒绝', r.status==='error', `→ ${r.message}`);
r = J(doGet({parameter:{action:'deleteLedger',row:2,expectAmount:99}}));
ok('删一笔支出', r.status==='success' && ledger._rows.length===before-1);

// 6. 股票
r = J(doGet({parameter:{action:'updateStock',symbol:'NVDA',shares:30,cost:150}}));
ok('改持仓', r.status==='success' && stock._rows[2][2]===30);
r = J(doGet({parameter:{action:'addStock',symbol:'SOXX',category:'Semiconductor',shares:6,cost:531}}));
ok('新增持仓+公式', r.status==='success' && /GOOGLEFINANCE/.test(stock._rows.at(-1)[4]));
r = J(doGet({parameter:{action:'updateCash',amount:5000}}));
ok('改券商现金', r.status==='success' && stock._rows[3][4]===5000);
r = J(doGet({parameter:{action:'deleteStock',symbol:'SOXX'}}));
ok('删持仓', r.status==='success');

// 7. Ledger 表头自动补列
ok('自动补 note/key 表头', ledger._rows[0][3]==='note' && ledger._rows[0][4]==='key');

r = J(doGet({parameter:{}}));
ok('月度汇总返回', !!r.monthly && Object.keys(r.monthly).length>0, JSON.stringify(r.monthly));
ok('不再生成重复的 Monthly 表', !globalThis.__tabs || !globalThis.__tabs['Monthly']);

r = J(doGet({parameter:{}}));
ok('返回每月总计', !!r.monthly && Object.keys(r.monthly).length>0, JSON.stringify(r.monthly));
ok('写入 Ledger_monthly', monthlyTab._rows.length>1, JSON.stringify(monthlyTab._rows[0]||[]).slice(0,60));
const rowsBefore=JSON.stringify(monthlyTab._rows);
J(doGet({parameter:{}}));
ok('内容不变时不重复写', JSON.stringify(monthlyTab._rows)===rowsBefore);

console.log('\n-- selfTest 输出 --'); selfTest();

const rr = J(doGet({parameter:{}}));
console.log('DEBUG status=', rr.status, 'msg=', rr.message, 'keys=', Object.keys(rr).join(','));

// ===== 并发写同一个 key 只应落一行 =====
console.log('\n-- 固定支出并发防重 --');
const rowsBefore2=ledger._rows.length;
const k='fix:test:2026-09-01';
[1,2,3].forEach(()=>J(doGet({parameter:{action:'autoLedger',key:k,
  category:'Bill & utilities',amount:70,note:'网费（固定）'}})));
const rows=ledger._rows.filter(r=>String(r[4]||'')===k);
ok('三次请求只写一行', rows.length===1, `写入 ${rows.length} 行`);
ok("总行数只 +1", ledger._rows.length===rowsBefore2+1);
