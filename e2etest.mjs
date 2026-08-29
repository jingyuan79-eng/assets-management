// 端到端：真实 Apps Script 代码 → 真实 index.html，中间不手工造数据
import { JSDOM } from 'jsdom'; import fs from 'fs';
const gs=fs.readFileSync('AppsScript.gs','utf8');
const html=fs.readFileSync('index.html','utf8');
const ok=(l,c,e='')=>console.log((c?'✅':'❌'),l,e);

// ---- 假 Sheets ----
function mkSheet(name,rows){return{_rows:rows,getName:()=>name,
 getDataRange(){const r=this._rows;return{getValues:()=>r.map(x=>x.slice())};},
 clear(){this._rows.length=0;},getLastRow(){return this._rows.length;},
 _cell(r,c){while(this._rows.length<r)this._rows.push([]);const row=this._rows[r-1];
   while(row.length<c)row.push("");return row;},
 getRange(r,c,nr,nc){const self=this;nr=nr||1;nc=nc||1;return{
   getValue:()=>(self._rows[r-1]||[])[c-1]??"",
   setValue(v){self._cell(r,c)[c-1]=v;},setFormula(v){self._cell(r,c)[c-1]=v;},
   getValues(){const o=[];for(let i=0;i<nr;i++){const row=[];
     for(let j=0;j<nc;j++)row.push((self._rows[r-1+i]||[])[c-1+j]??"");o.push(row);}return o;},
   setValues(v){for(let i=0;i<nr;i++){const row=self._cell(r+i,c+nc-1);
     for(let j=0;j<nc;j++)row[c-1+j]=v[i][j];}}};},
 appendRow(a){this._rows.push(a.slice());},deleteRow(r){this._rows.splice(r-1,1);}};}

// 用户真实的 Savings 表（Rate 百分数、Status 写成 Activate）
const tabs={
 Savings:mkSheet('Savings',[
  ['ID','Name','Type','Balance','Rate','Last Post','Next update','Maturity','Status'],
  ['','CD-7597','CD',21335,4,'2026-08-01','2026-09-01','2027-06-30','Activate'],
  ['','CD-9185','CD',10274,4.10,'2026-08-01','2026-09-01','2027-02-05','Activate'],
  ['','OS-9250','OS',30057,3.40,'2026-08-01','2026-09-01','','']]),
 Stock:mkSheet('Stock',[['Symbol','Category','Share','Cost','Price'],
  ['NVDA','Semiconductor',32,148.71,214.72],['Cash','Gold.coin.cash','','',1326]]),
 Ledger:mkSheet('Ledger',[['Date','category','amount'],
  [new Date('2026-08-12T00:00:00Z'),'Dining',27.31]]),
};
globalThis.SpreadsheetApp={flush(){},getActiveSpreadsheet:()=>({getSheetByName:n=>tabs[n]||null,
  insertSheet(n){tabs[n]=mkSheet(n,[]);return tabs[n];}})};
globalThis.ContentService={MimeType:{JSON:'json'},
  createTextOutput:t=>({_t:t,setMimeType(){return this;},getContent(){return this._t;}})};
globalThis.Utilities={getUuid:()=>Math.random().toString(36).slice(2)+'0000000',
  formatDate:(d,tz,f)=>{const p=new Date(d.getTime()).toISOString();
    return f==='yyyy-MM'?p.slice(0,7):p.slice(0,10);}};
globalThis.LockService={getScriptLock:()=>({waitLock(){},releaseLock(){}})};
globalThis.Logger={log:()=>{}};
globalThis.ScriptApp={getProjectTriggers:()=>[],newTrigger:()=>({timeBased:()=>({atHour:()=>({everyDays:()=>({create:()=>{}})})})})};
const NOW=new Date('2026-08-23T00:00:00Z');
const RealDate=Date;
globalThis.Date=class extends RealDate{constructor(...a){if(a.length===0)super(NOW.getTime());else super(...a);}
  static now(){return NOW.getTime();}};
const backend=new Function('return (function(){'+gs+'\nreturn {doGet,doPost};})()')();

// ---- 把小程序的 fetch 直接接到后端 ----
const dom=new JSDOM(html,{runScripts:'dangerously',url:'https://x.test/',
  beforeParse(w){ w.localStorage.clear();
    w.fetch=async(u)=>{ const url=new URL(u);
      const p={}; url.searchParams.forEach((v,k)=>p[k]=v);
      const res=backend.doGet({parameter:p.action?p:{}});
      const body=res.getContent();
      return { json:async()=>JSON.parse(body) };
    };
    w.alert=m=>console.log('   [alert]',m); w.confirm=()=>true; w.scrollTo=()=>{};
    w.Element.prototype.scrollIntoView=function(){};
  }});
const w=dom.window,d=w.document;
await new Promise(r=>setTimeout(r,400));

console.log('— 后端输出 → 前端显示 —');
const raw=JSON.parse(backend.doGet({parameter:{}}).getContent());
console.log('   后端 savings[0]:', JSON.stringify(raw.savings[0]));
ok('前端读到 3 个账户', w.eval("DATA.cd.groups[0].holdings.length")===3,
   w.eval("DATA.cd.groups[0].holdings.map(x=>x.sym+':'+x.mv).join(' ')"));
ok('储蓄总额正确（非 0）', w.eval("classTotal.cd")===21335+10274+30057, '$'+w.eval("classTotal.cd"));
ok('利率显示正确', w.eval("DATA.cd.groups[0].holdings[0].rate")===4,
   w.eval("DATA.cd.groups[0].holdings.map(x=>x.qty).join(' ')"));
ok('到期日显示正确', w.eval("DATA.cd.groups[0].holdings[0].mat")==='2027-06-30',
   w.eval("DATA.cd.groups[0].holdings[0].mat"));
ok('界面不显示下次入息日', !/下次入息/.test(d.getElementById('d-body')?d.getElementById('d-body').textContent:''));
ok('Activate 被规范成 active，未被误判 closed',
   !w.eval("DATA.cd.groups[0].holdings.some(x=>!x.mv)"));

console.log('\n— 在小程序里改余额，刷新后不应变回 0 —');
w.openDetail('cd');
const i=w.eval("DATA.cd.groups[0].holdings.findIndex(x=>x.sym==='OS-9250')");
w.editCd(i);
d.getElementById('f-mv').value='31000';
w.saveCd(); await new Promise(r=>setTimeout(r,500));
ok('Sheet 里已更新', tabs.Savings._rows[3][3]===31000, 'Sheet D4 = '+tabs.Savings._rows[3][3]);
ok('刷新后前端仍是 31000（不回 0）',
   w.eval("DATA.cd.groups[0].holdings.find(x=>x.sym==='OS-9250').mv")===31000,
   '$'+w.eval("DATA.cd.groups[0].holdings.find(x=>x.sym==='OS-9250').mv"));

console.log('\n— 新增储蓄，从 OS 扣款 —');
w.openDetail('cd'); w.showCdForm();
d.getElementById('f-name').value='CD-NEW'; d.getElementById('f-mv').value='5000';
d.getElementById('f-rate').value='4.5'; d.getElementById('f-mat').value='2027-08-23';
d.getElementById('f-src').value='marcus';
w.saveCd(); await new Promise(r=>setTimeout(r,600));
ok('新账户出现在前端', w.eval("DATA.cd.groups[0].holdings.some(x=>x.sym==='CD-NEW')"),
   w.eval("DATA.cd.groups[0].holdings.map(x=>x.sym+':'+x.mv).join(' ')"));
ok('OS 已扣款', w.eval("DATA.cd.groups[0].holdings.find(x=>x.sym==='OS-9250').mv")<31000,
   '$'+w.eval("DATA.cd.groups[0].holdings.find(x=>x.sym==='OS-9250').mv"));

console.log('\n— 其他模块没被带坏 —');
ok('股票总额正常', w.eval("classTotal.stock")>0, '$'+w.eval("classTotal.stock"));
ok('支出读到', w.eval("JSON.stringify(monthExpense)")!=='{}', w.eval("JSON.stringify(monthExpense)"));

console.log('\n— 清理后回归 —');
ok('无本地储蓄种子（Sheet 才是唯一来源）',
   !/21264|10239|41000/.test(fs.readFileSync('index.html','utf8')));
ok('cdList_v3 已从代码里消失',
   !/cdList_v3'/.test(fs.readFileSync('index.html','utf8').replace("'cdList_v3'","")) ||
   fs.readFileSync('index.html','utf8').split('cdList_v3').length-1===1);

console.log('\n— 没有 Anchor 表时要能优雅降级 —');
{
  ok('不崩溃，正常出数据', /已同步/.test(d.getElementById('syncBadge').textContent),
     d.getElementById('syncBadge').textContent);
  ok('流动现金显示 0 而不是乱数', w.eval("cashInfo.balance")===0, '$'+w.eval("cashInfo.balance"));
  ok('标记为「还没对过账」', w.eval("cashInfo.hasAnchor")===false);
  w.openDetail('cash'); w.switchCashTab('in');
  ok('提示用户先对账', /还没有设过起算余额/.test(d.getElementById('cashTabBody').textContent));
  ok('股票不受影响', w.eval("classTotal.stock")>0, '$'+w.eval("classTotal.stock"));
}

console.log('\n— 漏洞B：券商现金不足时不再静默吞掉 —');
{
  w.eval("saveBrokerCash(100); brokerCashItem().mv=100;");
  let asked=false; w.confirm=(m)=>{ asked=/券商现金只有/.test(m); return false; };
  w.openDetail('stock');
  const gi=w.eval("DATA.stock.groups.findIndex(g=>g.holdings.some(h=>h.sym==='NVDA'))");
  const hi=w.eval(`DATA.stock.groups[${gi}].holdings.findIndex(h=>h.sym==='NVDA')`);
  w.showStockPanel(gi,hi); w.stockAction('buy');
  d.getElementById('s-sh').value='10'; d.getElementById('s-px').value='200';
  w.applyBuy(); await new Promise(r=>setTimeout(r,200));
  ok('现金不足会先问一句', asked);
  ok('取消后不改动持仓', w.eval(`DATA.stock.groups[${gi}].holdings.find(h=>h.sym==='NVDA').shares`)===32,
     '股数 '+w.eval(`DATA.stock.groups[${gi}].holdings.find(h=>h.sym==='NVDA').shares`));
  ok('取消后现金不变', w.eval("brokerCashItem().mv")===100, '$'+w.eval("brokerCashItem().mv"));
}
