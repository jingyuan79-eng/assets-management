// 用户的三笔真实国债，跨时间验证付息 / 复利 / 到期
import fs from 'fs';
const gs=fs.readFileSync('AppsScript.gs','utf8');
const ok=(l,c,e='')=>console.log((c?'✅':'❌'),l,e);
function mkSheet(name,rows){return{_rows:rows,getName:()=>name,
 getDataRange(){const r=this._rows;return{getValues:()=>r.map(x=>x.slice())};},
 clear(){this._rows.length=0;},getLastRow(){return this._rows.length;},
 _cell(r,c){while(this._rows.length<r)this._rows.push([]);const row=this._rows[r-1];
   while(row.length<c)row.push("");return row;},
 getRange(r,c,nr,nc){const self=this;nr=nr||1;nc=nc||1;return{
   getValue:()=>(self._rows[r-1]||[])[c-1]??"",setValue(v){self._cell(r,c)[c-1]=v;},
   setFormula(v){self._cell(r,c)[c-1]=v;},
   getValues(){const o=[];for(let i=0;i<nr;i++){const row=[];
     for(let j=0;j<nc;j++)row.push((self._rows[r-1+i]||[])[c-1+j]??"");o.push(row);}return o;},
   setValues(v){for(let i=0;i<nr;i++){const row=self._cell(r+i,c+nc-1);
     for(let j=0;j<nc;j++)row[c-1+j]=v[i][j];}}};},
 appendRow(a){this._rows.push(a.slice());},deleteRow(r){this._rows.splice(r-1,1);}};}

const bond=mkSheet('Bond',[
 ['ID','Name','Type','Start','Term','Rate','Principal','Balance','LastPost','Status'],
 ['','T-Note 3Y','T-Note','2024-08-15',3,3.75,10000,10000,'2026-08-15','active'],
 ['','I Bond','I-Bond','2026-06-01',30,4.26,8000,8000,'2026-06-01','active'],
 ['','T-Note 3Y (2026)','T-Note','2026-08-15',3,4.25,10000,10000,'2026-08-15','active'],
]);
const ledger=mkSheet('Ledger',[['Date','category','amount','note','key']]);
const tabs={Bond:bond,Ledger:ledger,Anchor:mkSheet('Anchor',[['date','amount','note'],['2026-08-01',10000,'初始值']]),
  Stock:mkSheet('Stock',[['Symbol','Category','Share','Cost','Price'],['Cash','Gold.coin.cash','','',1326]]),
  Savings:mkSheet('Savings',[['ID','Name','Type','Balance','Rate','Last Post','Next update','Maturity','Status']])};
globalThis.SpreadsheetApp={flush(){},getActiveSpreadsheet:()=>({getSheetByName:n=>tabs[n]||null,
  insertSheet(n){tabs[n]=mkSheet(n,[]);return tabs[n];}})};
globalThis.ContentService={MimeType:{JSON:'json'},createTextOutput:t=>({_t:t,setMimeType(){return this;},getContent(){return this._t;}})};
globalThis.Utilities={getUuid:()=>Math.random().toString(36).slice(2)+'0000000',
  formatDate:(d,tz,f)=>{const p=new Date(d.getTime()).toISOString();return f==='yyyy-MM'?p.slice(0,7):p.slice(0,10);}};
globalThis.Logger={log:()=>{}};
globalThis.LockService={getScriptLock:()=>({waitLock(){},releaseLock(){}})};
globalThis.ScriptApp={getProjectTriggers:()=>[],newTrigger:()=>({timeBased:()=>({atHour:()=>({everyDays:()=>({create:()=>{}})})})})};
let NOW=new Date('2026-08-25T00:00:00Z'); const RD=Date;
globalThis.Date=class extends RD{constructor(...a){if(a.length===0)super(NOW.getTime());else super(...a);}static now(){return NOW.getTime();}};
const api=new Function('return (function(){'+gs+'\nreturn {doGet,runBonds,readBondRows,bondSheet,cleanBackfill};})()')();
const ss=SpreadsheetApp.getActiveSpreadsheet();
const J=r=>JSON.parse(r.getContent());
const coupons=()=>ledger._rows.filter(r=>String(r[4]||'').startsWith('bond:'));

console.log('— 今天 2026-08-25：都还没到下一次付息 —');
let out=api.runBonds(ss);
ok('三笔都在', out.length===3, out.map(b=>b.name+' $'+b.balance).join(' | '));
ok('自动补 ID', out.every(b=>!!b.id));
ok('还没产生利息', coupons().length===0);
ok('到期日算对', out[0].maturity==='2027-08-15'&&out[2].maturity==='2029-08-15',
   out.map(b=>b.name+'→'+b.maturity).join(', '));
ok('I-Bond 余额未动', out[1].balance===8000);

console.log('\n— 跳到 2027-02-16：两笔 T-Note 应各付一次息 —');
NOW=new RD('2027-02-16T00:00:00Z');
api.runBonds(ss);
const c1=coupons();
ok('两笔付息', c1.length===2, c1.map(r=>r[3]+' $'+r[2]).join(' | '));
ok('3.75% 那笔 = 187.50', c1.some(r=>Math.abs(r[2]-187.5)<0.01));
ok('4.25% 那笔 = 212.50', c1.some(r=>Math.abs(r[2]-212.5)<0.01));
ok('记成收入进流动现金', c1.every(r=>r[1]==='Income · 国债利息'));
ok('付息日是 2027-02-15', c1.every(r=>r[0]==='2027-02-15'));
// 重复运行不重复付
api.runBonds(ss);
ok('重复运行不重复付息', coupons().length===2, `仍是 ${coupons().length} 笔`);

console.log('\n— I-Bond 复利：2026-12-01 第一次 —');
let ib=api.readBondRows(bond).find(b=>b.type==='I-Bond');
const expect1=8000*(1+0.0426/2);
ok('半年复利一次', Math.abs(ib.bal-expect1)<0.01, `期望 ${expect1.toFixed(2)}，实际 ${ib.bal.toFixed(2)}`);
ok('LastPost 推到 2026-12-01', ledgerDate(ib)==='2026-12-01', ledgerDate(ib));
function ledgerDate(b){ return b.post.toISOString().slice(0,10); }

console.log('\n— I-Bond 改利率：只影响之后的复利 —');
J(api.doGet({parameter:{action:'updateBond', id:ib.id, rate:3.98}}));
const balBefore=api.readBondRows(bond).find(b=>b.type==='I-Bond').bal;
ok('改利率不动已有余额', Math.abs(balBefore-expect1)<0.01, `$${balBefore.toFixed(2)}`);
NOW=new RD('2027-06-02T00:00:00Z');
api.runBonds(ss);
ib=api.readBondRows(bond).find(b=>b.type==='I-Bond');
const expect2=expect1*(1+0.0398/2);
ok('下一期用新利率', Math.abs(ib.bal-expect2)<0.01,
   `${expect1.toFixed(2)} ×(1+3.98%/2) = ${expect2.toFixed(2)}，实际 ${ib.bal.toFixed(2)}`);

console.log('\n— 手动校正 I-Bond 余额（利率改晚了的情况）—');
J(api.doGet({parameter:{action:'updateBond', id:ib.id, balance:8500}}));
ok('以手动值为准', api.readBondRows(bond).find(b=>b.type==='I-Bond').bal===8500);

console.log('\n— 2027-08-16：第一笔 T-Note 到期 —');
NOW=new RD('2027-08-16T00:00:00Z');
api.runBonds(ss);
const rows=api.readBondRows(bond);
const b1=rows.find(b=>b.name==='T-Note 3Y');
ok('已关闭', b1.closed, bond._rows[1][9]);
const mat=ledger._rows.filter(r=>String(r[4]||'').startsWith('bondmat:'));
ok('本金记为 Redeem', mat.length===1 && mat[0][1]==='Redeem · 国债' && mat[0][2]===10000,
   mat.map(r=>r[1]+' $'+r[2]).join());
const augC=coupons().filter(r=>r[0]==='2027-08-15');
ok('到期当天的最后一次付息也发了', augC.length===2, augC.map(r=>r[3]+' $'+r[2]).join(' | '));
const finalOut=J(api.doGet({parameter:{}}));
ok('到期的不再出现在活跃列表', finalOut.bond.filter(b=>b.status!=='closed').length===2,
   finalOut.bond.map(b=>b.name+':'+b.status).join(' | '));
const m8=finalOut.monthly['2027-08'];
ok('本金回流单列，不进收入也不进净收入',
   m8.redeem===10000 && m8.net===m8.income-m8.expense && m8.income===400,
   JSON.stringify(m8));
api.runBonds(ss);
ok('到期不重复入账', ledger._rows.filter(r=>String(r[4]||'').startsWith('bondmat:')).length===1);

console.log('\n— 2024 年买的国债不应补录历史利息 —');
{
  // 重置：只留 Anchor(2026-08-01)，Ledger 清空，Bond 空表
  NOW=new RD('2026-08-25T00:00:00Z');
  ledger._rows.length=1;
  bond._rows.length=1;
  const r=J(api.doGet({parameter:{action:'addBond', name:'T-Note 3Y', type:'T-Note',
    start:'2024-08-15', term:3, rate:3.75, principal:10000}}));
  ok('新增成功', r.status==='success', r.message||'');
  ok('进度直接推到最近一次付息日', bond._rows[1][8]==='2026-08-15', 'LastPost='+bond._rows[1][8]);
  api.runBonds(ss);
  ok('没有补录任何历史利息', coupons().length===0, `产生了 ${coupons().length} 笔`);

  // 再点一次保存（模拟连点）
  const dup=J(api.doGet({parameter:{action:'addBond', name:'T-Note 3Y', type:'T-Note',
    start:'2024-08-15', term:3, rate:3.75, principal:10000}}));
  ok('重复保存被拦下', dup.status==='error', dup.message||'');
  ok('Bond 表只有一行', bond._rows.length===2, `共 ${bond._rows.length-1} 笔`);

  // 到 2027-02-16 才该有第一笔
  NOW=new RD('2027-02-16T00:00:00Z');
  api.runBonds(ss);
  ok('2027-02-15 才付第一次息', coupons().length===1 && coupons()[0][0]==='2027-02-15',
     coupons().map(r=>r[0]+' $'+r[2]).join());
}

console.log('\n— 手填 LastPost 很早时也不补录记账前的账 —');
{
  NOW=new RD('2026-08-25T00:00:00Z');
  ledger._rows.length=1; bond._rows.length=1;
  bond.appendRow(['','老 T-Note','T-Note','2024-08-15',3,3.75,10000,10000,'2024-08-15','active']);
  api.runBonds(ss);
  // 2025-02/2025-08/2026-02 在起始日之前 → 跳过；2026-08-15 在之后 → 照常记
  ok('只记起始日之后的那一笔', coupons().length===1 && coupons()[0][0]==='2026-08-15',
     coupons().map(r=>r[0]).join(', ')||'0 笔');
  ok('起始日之前的三次被跳过', !coupons().some(r=>String(r[0])<'2026-08-01'));
  ok('进度仍然推到了 2026-08-15', bond._rows[1][8]==='2026-08-15', bond._rows[1][8]);
}

console.log('\n— cleanBackfill 能清掉已经误写的历史行 —');
{
  ledger._rows.length=1;
  ledger.appendRow(['2025-02-15','Income · 国债利息',187.5,'老息','bond:x:2025-02-15']);
  ledger.appendRow(['2026-02-15','Income · 国债利息',187.5,'老息','bond:x:2026-02-15']);
  ledger.appendRow(['2026-08-20','Dining',50,'','']);
  ledger.appendRow(['2026-08-22','Income · 国债利息',187.5,'新息','bond:x:2026-08-22']);
  api.cleanBackfill();
  const left=ledger._rows.slice(1);
  ok('起始日之前的自动行被删', !left.some(r=>String(r[0])<'2026-08-01'), left.map(r=>r[0]).join(', '));
  ok('手动记的账不动', left.some(r=>r[1]==='Dining'));
  ok('起始日之后的自动行保留', left.some(r=>String(r[4]||'').startsWith('bond:')));
}
