// 端到端：真实 Apps Script + 真实 index.html，验证锚点现金模型
import { JSDOM } from 'jsdom'; import fs from 'fs';
const gs=fs.readFileSync('AppsScript.gs','utf8');
const html=fs.readFileSync('index.html','utf8');
const ok=(l,c,e='')=>console.log((c?'✅':'❌'),l,e);
const wait=(ms=400)=>new Promise(r=>setTimeout(r,ms));

// 从 Sheet 直接推算权威余额：最新锚点 + 锚点当天及之后的流水
function expectedCash(){
  const anchors=tabs.Anchor._rows.slice(1).filter(r=>r[0])
    .sort((a,b)=>String(a[0])<String(b[0])?-1:1);
  const last=anchors[anchors.length-1];
  const ad=String(last[0]).slice(0,10);
  const sum=tabs.Ledger._rows.slice(1).reduce((s,r)=>{
    const ds=String(r[0]||'').slice(0,10);
    if(!ds||ds<ad) return s;
    const c=String(r[1]||'');
    const k=(c.startsWith('Income')||c.startsWith('Redeem'))?1:-1;
    return s+k*(parseFloat(r[2])||0);},0);
  return Math.round((last[1]+sum)*100)/100;
}


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

const tabs={
 Ledger:mkSheet('Ledger',[['Date','category','amount','note','key'],
   ['2026-08-01','Income · Payroll',2500,'工资','pay:2026-08-01'],
   ['2026-08-03','Bill & utilities',2625.84,'House loan（固定）','fix:b:2026-08-03'],
   ['2026-08-07','Dining',51.65,'',''],
   ['2026-08-15','Income · Payroll',2500,'工资','pay:2026-08-15'],
   ['2026-08-20','Transfer · 券商',2000,'',''],
 ]),
 Anchor:mkSheet('Anchor',[['date','amount','note'],['2026-08-01',10000,'初始值']]),
 Stock:mkSheet('Stock',[['Symbol','Category','Share','Cost','Price'],
   ['NVDA','Semiconductor',32,148.71,214.72],['Cash','Gold.coin.cash','','',1326]]),
 Savings:mkSheet('Savings',[['ID','Name','Type','Balance','Rate','Last Post','Next update','Maturity','Status'],
   ['sv_c','OS-9250','OS',30057,3.4,'2026-08-01','2026-09-01','','active']]),
};
globalThis.SpreadsheetApp={flush(){},getActiveSpreadsheet:()=>({getSheetByName:n=>tabs[n]||null,
  insertSheet(n){tabs[n]=mkSheet(n,[]);return tabs[n];}})};
globalThis.ContentService={MimeType:{JSON:'json'},
  createTextOutput:t=>({_t:t,setMimeType(){return this;},getContent(){return this._t;}})};
globalThis.Utilities={getUuid:()=>Math.random().toString(36).slice(2)+'0000000',
  formatDate:(d,tz,f)=>{const p=new Date(d.getTime()).toISOString();
    return f==='yyyy-MM'?p.slice(0,7):p.slice(0,10);}};
globalThis.Logger={log:()=>{}};
globalThis.LockService={getScriptLock:()=>({waitLock(){},releaseLock(){}})};
globalThis.ScriptApp={getProjectTriggers:()=>[],newTrigger:()=>({timeBased:()=>({atHour:()=>({everyDays:()=>({create:()=>{}})})})})};
const NOW=new Date('2026-08-25T00:00:00Z'); const RealDate=Date;
globalThis.Date=class extends RealDate{constructor(...a){if(a.length===0)super(NOW.getTime());else super(...a);}
  static now(){return NOW.getTime();}};
const backend=new Function('return (function(){'+gs+'\nreturn {doGet,doPost};})()')();

const dom=new JSDOM(html,{runScripts:'dangerously',url:'https://x.test/',
  beforeParse(w){ w.localStorage.clear();
    // jsdom 的 window 有自己的 Date，外面冻结 globalThis.Date 管不到它。
    // 不冻的话前端读的是真实系统时间，一跨月夹具里写死的日期就全部失效。
    const _RD=w.Date;
    w.Date=class extends _RD{constructor(...a){if(a.length===0)super(NOW.getTime());else super(...a);}
      static now(){return NOW.getTime();}};
    w.fetch=async(u)=>{const url=new URL(u);const p={};url.searchParams.forEach((v,k)=>p[k]=v);
      return {json:async()=>JSON.parse(backend.doGet({parameter:p.action?p:{}}).getContent())};};
    w.alert=()=>{}; w.confirm=()=>true; w.scrollTo=()=>{};
    w.Element.prototype.scrollIntoView=function(){};
  }});
const w=dom.window,d=w.document;
await wait();

console.log('— 锚点推导流动现金 —');
// 10000 + 2500 - 2625.84 - 51.65 + 2500 - 2000 = 10322.51
ok('余额=锚点+流水', Math.abs(w.eval("cashInfo.balance")-10322.51)<0.01, '$'+w.eval("cashInfo.balance"));
ok('现金卡片同步', w.eval("chaseItem().mv")===10323, '$'+w.eval("chaseItem().mv"));
ok('不再有 expenseDeducted', !w.localStorage.getItem('expenseDeducted_v2'));
ok('不再本地存 Chase 余额', !w.localStorage.getItem('cashList_v1'));

console.log('\n— 月度统计分收入/支出/投资转出 —');
const m=JSON.parse(w.eval("JSON.stringify(monthlyTotals)"))['2026-08'];
ok('收入 5000', m.income===5000, JSON.stringify(m));
ok('支出 2677.49', Math.abs(m.expense-2677.49)<0.01);
ok('投资转出 2000', m.transfer===2000);
ok('净收入 = 收入 − 支出 = 2322.51', Math.abs(m.net-(m.income-m.expense))<0.01, '净收入 $'+m.net);
ok('月末流动现金', Math.abs(m.endCash-10322.51)<0.01, '$'+m.endCash);
const mt=tabs.Ledger_monthly._rows;
ok('Ledger_monthly 表头含新列', mt[0].join(',').includes('净收入')&&mt[0].join(',').includes('本金回流')
   &&mt[0].join(',').includes('月末流动现金'), mt[0].slice(0,7).join(' | '));

console.log('\n— 支出/收入编辑即时反映到余额 —');
w.openDetail('cash'); w.switchCashTab('ex'); w.openExpenseDrill('Dining');
const b0=w.eval("cashInfo.balance");
w.showAddLedger();
d.getElementById('lg-amt').value='20'; d.getElementById('lg-note').value='咖啡';
w.saveLedgerAdd(); await wait(500);
ok('补记支出后余额减少', Math.abs(w.eval("cashInfo.balance")-(b0-20))<0.01, '$'+w.eval("cashInfo.balance"));
w.switchCashTab('in'); w.showBonusForm();
d.getElementById('bonus-amt').value='1000'; w.saveBonus(); await wait(500);
ok('Bonus 记为收入行', tabs.Ledger._rows.some(r=>r[1]==='Income · Bonus'&&r[2]===1000));
ok('余额增加 1000', Math.abs(w.eval("cashInfo.balance")-(b0-20+1000))<0.01, '$'+w.eval("cashInfo.balance"));

console.log('\n— 对账 —');
w.switchCashTab('in'); w.showReconcileForm();
ok('对账按钮存在', /对账/.test(d.getElementById('cashTabBody').innerHTML));
ok('重算历史按钮存在', /重算历史/.test(d.getElementById('cashTabBody').innerHTML));
const computed=w.eval("cashInfo.balance");
d.getElementById('rc-amt').value=String(computed-62);
d.getElementById('rc-mode').value='adjust';
w.submitReconcile(); await wait(700);
ok('对账后余额与 Sheet 推算一致', Math.abs(w.eval("cashInfo.balance")-expectedCash())<0.02,
   `小程序 $${w.eval("cashInfo.balance")}，Sheet 推算 $${expectedCash()}`);
ok('Anchor 表新增一行', tabs.Anchor._rows.length===3, JSON.stringify(tabs.Anchor._rows.at(-1)));
// 对账当天再补一笔，应正确扣减
const b1=w.eval("cashInfo.balance");
w.switchCashTab('ex'); w.openExpenseDrill('Dining'); w.showAddLedger();
d.getElementById('lg-amt').value='15'; w.saveLedgerAdd(); await wait(500);
ok('对账当天补记仍正确扣减', Math.abs(w.eval("cashInfo.balance")-(b1-15))<0.01, '$'+w.eval("cashInfo.balance"));

console.log('\n— 对账时记为支出 —');
w.switchCashTab('in'); w.showReconcileForm();
const c2=w.eval("cashInfo.balance");
d.getElementById('rc-amt').value=String(c2-30);
d.getElementById('rc-mode').value='expense';
w.updateReconcileDiff();
d.getElementById('rc-cat').value='Other or unexpected';
w.submitReconcile(); await wait(700);
ok('补了一笔支出行', tabs.Ledger._rows.some(r=>r[3]==='对账差额'&&r[1]==='Other or unexpected'));
ok('余额与 Sheet 推算一致', Math.abs(w.eval("cashInfo.balance")-expectedCash())<0.02,
   `小程序 $${w.eval("cashInfo.balance")}，Sheet 推算 $${expectedCash()}`);

console.log('\n— 券商现金按增量写 —');
const cashRow=()=>tabs.Stock._rows.find(r=>r[0]==='Cash')[4];
const bk0=cashRow();
w.openDetail('stock'); w.showDividendForm();
d.getElementById('dv-amt').value='42.5'; w.applyDividend(); await wait(500);
ok('分红按增量入账', Math.abs(cashRow()-(bk0+42.5))<0.01, `$${bk0} → $${cashRow()}`);

console.log('\n— 定期转账：到日子才执行，带 key 防重 —');
const before=w.eval("cashInfo.balance");
const past=new Date(NOW); past.setDate(past.getDate()-30);
w.saveRecurring([{id:'r1',src:'chase',tgt:'broker',amt:500,freq:'biweekly',
                  start:past.toISOString().slice(0,10),lastRun:null}]);
w.runRecurring(); await wait(900);
const xfers=tabs.Ledger._rows.filter(r=>String(r[4]||'').startsWith('xfer:r1:'));
ok('30天内每两周转 3 次', xfers.length===3, xfers.map(r=>r[0]).join(', '));
ok('流动现金相应减少', Math.abs(w.eval("cashInfo.balance")-(before-1500))<0.01, '$'+w.eval("cashInfo.balance"));
// 模拟本地状态丢失后重跑
w.saveRecurring([{id:'r1',src:'chase',tgt:'broker',amt:500,freq:'biweekly',
                  start:past.toISOString().slice(0,10),lastRun:null}]);
w.runRecurring(); await wait(900);
const xfers2=tabs.Ledger._rows.filter(r=>String(r[4]||'').startsWith('xfer:r1:'));
ok('状态丢失重跑也不重复转账', xfers2.length===3, `仍是 ${xfers2.length} 笔`);
// 未来的日子不提前执行
const fut=new Date(NOW); fut.setDate(fut.getDate()+10);
w.saveRecurring([{id:'r2',src:'chase',tgt:'broker',amt:800,freq:'monthly',
                  start:fut.toISOString().slice(0,10),lastRun:null}]);
w.runRecurring(); await wait(500);
ok('未来日期不提前执行', !tabs.Ledger._rows.some(r=>String(r[4]||'').startsWith('xfer:r2:')));

console.log('\n— 支出合计不能把收入算进去 —');
{
  // 当前 Ledger 里有 Payroll(收入) / Dining(支出) / Transfer·券商(转账) 三种
  w.openDetail('cash'); w.switchCashTab('ex');
  const cells={};
  [...d.querySelectorAll('.expcell:not(.exptotal)')].forEach(c=>{
    cells[c.querySelector('.nm').textContent]=c.querySelector('.amt').textContent;});
  const total=d.querySelector('.exptotal .amt').textContent;
  const me=JSON.parse(w.eval("JSON.stringify(monthExpense)"));
  const sumCats=Object.values(me).reduce((s,v)=>s+v,0);
  ok('monthExpense 只含支出分类',
     !Object.keys(me).some(k=>/^Income|^Transfer/.test(k)), Object.keys(me).join(' / '));
  ok('本月合计 = 各分类之和', total==='$'+Math.round(sumCats).toLocaleString('en-US'),
     `合计 ${total}，各类之和 $${Math.round(sumCats)}`);
  const mt=JSON.parse(w.eval("JSON.stringify(monthlyTotals)"))['2026-08'];
  ok('与后端算的支出一致', Math.abs(sumCats-mt.expense)<0.01,
     `前端 $${sumCats.toFixed(2)}，后端 $${mt.expense}`);
  ok('没有「不匹配」误报', !/不匹配/.test(d.getElementById('cashTabBody').textContent));
}

console.log('\n— 分类下钻不串入收入 —');
{
  w.openExpenseDrill('Bill & utilities');
  ok('只出现该分类的支出', !/Income|工资/.test(d.getElementById('cashTabBody').textContent));
  w.openExpenseDrill('__ALL__');
  const t=d.getElementById('cashTabBody').textContent;
  ok('全部明细里能看到收入并标 +', /\+\$/.test(t) && /−\$/.test(t));
}

console.log('\n— 收入 tab 小结 —');
{
  w.switchCashTab('in');
  const t=d.getElementById('cashTabBody').textContent;
  ok('显示本月收入与净收入（不显示支出）', /本月收入/.test(t)&&/净收入/.test(t)&&!/本月支出/.test(t),
     t.replace(/\s+/g,' ').slice(0,60));
}

console.log('\n— 界面调整 —');
{
  w.openDetail('cash'); w.switchCashTab('in');
  const t=d.getElementById('cashTabBody').textContent;
  ok('收入页没有「本月支出」', !/本月支出/.test(t));
  ok('有本月收入 / 净收入 / 对账三个按钮', /本月收入/.test(t)&&/净收入/.test(t)&&/对账/.test(t));
  ok('Payroll/Bonus 在上方', t.indexOf('Payroll')<t.indexOf('本月收入'), 'Payroll 位置靠前');
  ok('重算历史已移出主界面', !/重算历史/.test(t));
  w.showReconcileForm();
  const f=d.getElementById('incomeFormWrap').textContent;
  ok('对账标题是「手动修改余额」', /手动修改余额/.test(f));
  ok('保留上次对账日期', /上次对账|还没对过账/.test(f));
  ok('去掉「按流水算出来是」', !/按流水算出来/.test(f));
  ok('去掉「信用卡时间差等」', !/信用卡时间差/.test(f));
  ok('表单里三个按钮', /取消/.test(f)&&/确认对账/.test(f)&&/重算历史/.test(f));
}

console.log('\n— 股票现金文案 —');
{
  w.openDetail('stock');
  ok('说明改为点击添加分红', /券商账户现金 · 点击添加股票分红/.test(d.getElementById('d-body').textContent));
  w.showDividendForm();
  const f=d.getElementById('cashslot').textContent;
  ok('分红表单去掉多余标题', !/添加股票分红/.test(f)&&!/当前券商现金/.test(f), f.replace(/\s+/g,' ').slice(0,40));
}

console.log('\n— 全部明细的合计只算支出 —');
{
  w.openDetail('cash'); w.switchCashTab('ex');
  const grid=d.querySelector('.exptotal .amt').textContent;
  w.openExpenseDrill('__ALL__');
  const head=d.getElementById('cashTabBody').querySelector('.num');
  ok('全部明细合计 = 方块本月合计', head.textContent===grid, `明细 ${head.textContent} vs 方块 ${grid}`);
  ok('列表里仍能看到收入行（可订正）', /\+\$/.test(d.getElementById('cashTabBody').textContent));
}

console.log('\n— 工资写进 Ledger，靠 key 防重复 —');
{
  const payRows=()=>tabs.Ledger._rows.filter(r=>String(r[4]||'').startsWith('pay:'));
  const base=payRows().length;   // 初始数据里已有的工资行
  w.openDetail('cash'); w.switchCashTab('in'); w.showPayrollForm();
  const start=new Date(NOW); start.setDate(start.getDate()-30);
  d.getElementById('pay-amt').value='2500';
  d.getElementById('pay-freq').value='biweekly';
  d.getElementById('pay-start').value=start.toISOString().slice(0,10);
  const b0=w.eval("cashInfo.balance");
  w.savePayroll(); await wait(900);
  const added=payRows().length-base;
  ok('只补记账起始日之后的（2 次）', added===2,
     `新增 ${added} 笔：`+payRows().slice(base).map(r=>r[0]).join(', '));
  ok('起始日之前的不补发', !payRows().slice(base).some(r=>String(r[0])<'2026-08-01'));
  ok('分类是 Income · Payroll', payRows().every(r=>r[1]==='Income · Payroll'));
  ok('流动现金增加 5000', Math.abs(w.eval("cashInfo.balance")-(b0+5000))<0.01,
     `$${b0} → $${w.eval("cashInfo.balance")}`);
  // 再改一次金额，不应重复补发历史
  const b1=w.eval("cashInfo.balance");
  w.showPayrollForm();
  d.getElementById('pay-amt').value='2600';
  d.getElementById('pay-freq').value='biweekly';
  d.getElementById('pay-start').value=start.toISOString().slice(0,10);
  w.savePayroll(); await wait(900);
  ok('改金额不会重复补发', payRows().length===base+2, `仍是 ${payRows().length} 笔`);
  ok('余额不受影响', Math.abs(w.eval("cashInfo.balance")-b1)<0.01, '$'+w.eval("cashInfo.balance"));
  // 模拟换设备：清掉本地配置再设一次
  w.localStorage.removeItem('incomeConfig_v1');
  w.showPayrollForm();
  d.getElementById('pay-amt').value='2500';
  d.getElementById('pay-freq').value='biweekly';
  d.getElementById('pay-start').value=start.toISOString().slice(0,10);
  w.savePayroll(); await wait(900);
  ok('换设备重设也不重复发', payRows().length===base+2, `仍是 ${payRows().length} 笔`);
}

console.log('\n— 净收入 = 收入 − 支出（不含转账与回流）—');
{
  // 造一笔国债到期回流 + 一笔利息
  tabs.Ledger.appendRow(['2026-08-24','Redeem · 国债',10000,'到期本金','bondmat:t1']);
  tabs.Ledger.appendRow(['2026-08-24','Income · 国债利息',187.5,'T-Note 半年息','bond:t1:2026-08-24']);
  await w.syncStockFromSheet(); await wait(600);
  // 权威值：锚点 + 全部流水（含回流）
  const anchors=tabs.Anchor._rows.slice(1).filter(r=>r[0]);
  const latest=anchors.sort((a,b)=>String(a[0])<String(b[0])?-1:1).slice(-1)[0];
  const anchorAmt=latest[1], anchorDate=String(latest[0]).slice(0,10);
  const signed=tabs.Ledger._rows.slice(1).reduce((s,r)=>{
    const ds=String(r[0]||'').slice(0,10);
    if(!ds || ds<anchorDate) return s;
    const c=String(r[1]||''); const k=(c.startsWith('Income')||c.startsWith('Redeem'))?1:-1;
    return s + k*(parseFloat(r[2])||0);},0);
  const mt=JSON.parse(w.eval("JSON.stringify(monthlyTotals)"))['2026-08'];
  ok('回流单独统计', mt.redeem===10000, JSON.stringify(mt));
  ok('净收入只算收入减支出', Math.abs(mt.net-(mt.income-mt.expense))<0.01,
     `${mt.income} − ${mt.expense} = ${mt.net}`);
  ok('回流不进净收入', mt.net<10000, '净收入 $'+mt.net);
  ok('回流计入现金余额', Math.abs(w.eval("cashInfo.balance")-(anchorAmt+signed))<0.02,
     `期望 $${(anchorAmt+signed).toFixed(2)}，实际 $${w.eval("cashInfo.balance")}`);
  const head=tabs.Ledger_monthly._rows[0].join(',');
  ok('Sheet 列名改为净收入/本金回流', head.includes('净收入')&&head.includes('本金回流')&&!head.includes('净流入'),
     tabs.Ledger_monthly._rows[0].slice(0,7).join(' | '));
}

console.log('\n— 收入页能看到入息 —');
{
  w.openDetail('cash'); w.switchCashTab('in');
  const t=d.getElementById('cashTabBody').textContent;
  ok('显示「净收入」而非「净流入」', /净收入/.test(t)&&!/净流入/.test(t));
  ok('有本月收入明细区', /本月收入明细/.test(t));
  ok('国债半年息出现在里面', /国债利息|半年息/.test(t), t.match(/半年息[^　]{0,20}/)||'');
  ok('回流不出现在收入明细', !/到期本金/.test(t));
  const html=d.getElementById('cashTabBody').innerHTML;
  ok('净收入卡片用新配色 #D8E3E7', /#D8E3E7/.test(html));
  ok('对账按钮用新配色 #FBECDE', /#FBECDE/.test(html));
  // 点开一笔收入应能编辑
  const irow=w.eval("ledgerRows.filter(r=>catKind(r.category)==='income').slice(-1)[0].row");
  w.editLedger(irow);
  ok('收入行可点开编辑', !!d.getElementById('lg-amt'));
  const sel=d.getElementById('lg-cat');
  ok('分类下拉给的是收入类别', [...sel.options].every(o=>/^Income/.test(o.value)),
     [...sel.options].map(o=>o.value).join(' / '));
}
