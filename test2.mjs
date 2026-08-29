import { JSDOM } from 'jsdom';
import fs from 'fs';
const html = fs.readFileSync('index.html','utf8');

const stockRows=[
 {symbol:'ASMIY',category:'Semiconductor',shares:25,cost:888,price:969.26},
 {symbol:'NVDA',category:'Semiconductor',shares:32,cost:148.71,price:214.72},
 {symbol:'VOO',category:'IndexETF',shares:12.52,cost:679.47,price:703.71},
 {symbol:'Cash',category:'Gold.coin.cash',shares:'',cost:'',price:1326},
];
let ledger=[
 {row:2,date:'2026-08-12',category:'Dining',amount:27.31,note:'',key:''},
 {row:3,date:'2026-08-12',category:'Grocery & Household',amount:92.84,note:'',key:''},
 {row:4,date:'2026-08-21',category:'Health & Beauty',amount:69,note:'',key:''},
 {row:5,date:'2026-08-22',category:'Grocery',amount:109.09,note:'',key:''},  // 故意留一个不匹配
];
let nextRow=6;
const calls=[];
function payload(){ return {status:'success',stock:stockRows,ledger:JSON.parse(JSON.stringify(ledger)),
  expense:{},serverDate:'2026-08-23'}; }

global.fetch = async (url) => {
  const u=new URL(url);
  const a=u.searchParams.get('action');
  if(!a) return { json: async()=>payload() };
  const p=Object.fromEntries(u.searchParams);
  calls.push(p);
  if(a==='autoLedger'){
    if(ledger.some(r=>r.key===p.key)) return {json:async()=>({status:'skipped',key:p.key})};
    ledger.push({row:nextRow, date:p.date, category:p.category, amount:+p.amount, note:p.note, key:p.key});
    return {json:async()=>({status:'success',row:nextRow++})};
  }
  if(a==='addLedger'){ ledger.push({row:nextRow,date:p.date,category:p.category,amount:+p.amount,note:p.note,key:''});
    return {json:async()=>({status:'success',row:nextRow++})}; }
  if(a==='updateLedger'){ const r=ledger.find(x=>x.row==p.row);
    if(!r) return {json:async()=>({status:'error',message:'no row'})};
    if(Math.abs(r.amount-p.expectAmount)>0.005) return {json:async()=>({status:'error',message:'数据已变化'})};
    r.date=p.date; r.category=p.category; r.amount=+p.amount; r.note=p.note;
    return {json:async()=>({status:'success'})}; }
  if(a==='deleteLedger'){ const i=ledger.findIndex(x=>x.row==p.row);
    if(i<0) return {json:async()=>({status:'error'})};
    ledger.splice(i,1); return {json:async()=>({status:'success'})}; }
  return {json:async()=>({status:'success'})};
};

const dom=new JSDOM(html,{runScripts:'dangerously',url:'https://x.test/',
  beforeParse(w){ w.fetch=global.fetch; w.alert=m=>console.log('   [alert]',m);
    w.confirm=()=>true; w.scrollTo=()=>{}; w.Element.prototype.scrollIntoView=function(){}; }});
const w=dom.window,d=dom.window.document;
const wait=(ms=250)=>new Promise(r=>setTimeout(r,ms));
const ok=(l,c,e='')=>console.log((c?'✅':'❌'),l,e);
await wait(400);

ok('同步状态', /已同步/.test(d.getElementById('syncBadge').textContent), d.getElementById('syncBadge').textContent);

// 打开现金页 → 支出
w.openDetail('cash'); w.switchCashTab('ex');
const cells=[...d.querySelectorAll('.expcell:not(.exptotal) .nm')].map(e=>e.textContent);
const colOf=n=>{const c=[...d.querySelectorAll('.expcell:not(.exptotal)')].find(x=>x.querySelector('.nm').textContent===n);return c.getAttribute('style');};
ok('Other 用原 Gifts 的色', /#DCDDE3/.test(colOf('Other or unexpected')));
ok('Gifts 用原 Home 的色', /#E6DDF6/.test(colOf('Gifts & Families')));
ok('Home 用原合计的色', /#fbf1d9/.test(colOf('Home')));
const tcell=d.querySelector('.expcell.exptotal');
ok('合计在 Other 右边（最后一格）', d.querySelector('.expgrid').lastElementChild===tcell);
ok('合计不再跨格', !/grid-column:span 2/.test(html));
ok('合计无「点方块看明细」', !/点方块看明细/.test(tcell.textContent), tcell.textContent);
ok('11 个方块顺序正确', cells.length===11 && cells[0]==='Bill & utilities' && cells[1]==='Auto & Gas' && cells[10]==='Other or unexpected', cells.length+'个');
ok('不匹配警告', /Grocery/.test(d.getElementById('cashTabBody').textContent) && /不匹配/.test(d.getElementById('cashTabBody').textContent));
const tot=d.querySelector('.exptotal .amt').textContent;
ok('本月合计', tot==='$298', tot);

// 下钻
w.openExpenseDrill('Dining');
ok('下钻到 Dining', /本月这一类|27.31/.test(d.getElementById('cashTabBody').textContent));
// 改一笔
w.editLedger(2);
ok('编辑表单出现', !!d.getElementById('lg-amt'));
d.getElementById('lg-amt').value='50';
w.saveLedgerEdit(2,27.31); await wait(400);
ok('改金额成功', ledger.find(r=>r.row===2).amount===50, '→ $'+ledger.find(r=>r.row===2).amount);

// 补记
w.openExpenseDrill('Pet'); w.showAddLedger();
d.getElementById('lg-amt').value='33.5'; d.getElementById('lg-note').value='猫粮';
w.saveLedgerAdd(); await wait(400);
ok('补记一笔', ledger.some(r=>r.amount===33.5 && r.category==='Pet'));

// 删除
const cnt=ledger.length;
w.openExpenseDrill('Health & Beauty'); w.editLedger(4); w.deleteLedgerRow(4,69); await wait(400);
ok('删除一笔', ledger.length===cnt-1);

// 固定支出
w.saveFixed([{id:'f1',name:'房贷月供',amount:2025,category:'Bill & utilities',freq:'monthly',start:'2026-08-01',enabled:true},
             {id:'f2',name:'网费',amount:70,category:'Bill & utilities',freq:'monthly',start:'2026-08-05',enabled:true}]);
await w.runFixedExpenses(); await wait(300);
const fixWrote=ledger.filter(r=>(r.key||'').startsWith('fix:'));
ok('固定支出自动写入', fixWrote.length===2, fixWrote.map(r=>r.note+' $'+r.amount).join(', '));
const before=ledger.length;
await w.runFixedExpenses(); await wait(300);
ok('再次运行不重复写', ledger.length===before, ledger.length+' 条');

w.rebuildMonthExpense();
ok('房贷计入 Bill & utilities', Math.round(w.eval("monthExpense['Bill \u0026 utilities']"))===2095, '$'+Math.round(w.eval("monthExpense['Bill \u0026 utilities']")||0));

// 固定支出：内嵌在分类明细里
w.openDetail('cash'); w.switchCashTab('ex'); w.openExpenseDrill('Bill & utilities');
const bt=d.getElementById('cashTabBody').textContent;
ok('Bill 明细页有固定支出区', /固定支出/.test(bt) && /房贷月供/.test(bt) && /网费/.test(bt));
ok('折合每月显示', /折合每月 \$2,095/.test(bt), (bt.match(/折合每月 \$[\d,]+/)||[''])[0]);
w.openExpenseDrill('Auto & Gas');
ok('Auto & Gas 也有固定支出区', /添加固定支出/.test(d.getElementById('cashTabBody').textContent));
w.openExpenseDrill('Pet');
ok('Pet 也有固定支出区', /添加固定支出/.test(d.getElementById('cashTabBody').textContent));
w.openExpenseDrill('Dining');
ok('Dining 也有固定支出区（已扩展到 9 类）', /添加固定支出/.test(d.getElementById('cashTabBody').textContent));
w.openExpenseDrill('Gifts & Families');
ok('Gifts & Families 没有固定支出区', !/添加固定支出/.test(d.getElementById('cashTabBody').textContent));
// 在 Pet 里新增一项固定支出
w.openExpenseDrill('Pet'); w.showFixedForm('Pet');
d.getElementById('fx-name').value='宠物保险'; d.getElementById('fx-amt').value='45';
d.getElementById('fx-start').value='2026-08-10';
w.saveFixed_('Pet'); await wait(400);
ok('Pet 固定支出已保存并写账', ledger.some(r=>/宠物保险/.test(r.note||'') && r.category==='Pet'));

// 写失败提示
w.fetch = async (url)=>{ const u=new URL(url);
  if(!u.searchParams.get('action')) return {json:async()=>payload()};
  return {json:async()=>({status:'error',message:'数据已变化，请刷新后重试'})}; };
w.openDetail('cash'); w.switchCashTab('ex'); w.openExpenseDrill('__ALL__');
w.editLedger(3); d.getElementById('lg-amt').value='1'; w.saveLedgerEdit(3,999); await wait(400);
ok('写失败会报出来', /保存失败/.test(d.getElementById('syncBadge').textContent), d.getElementById('syncBadge').textContent);

// --- 股票/现金写入仍走同一条通道 ---
const calls2=[];
w.fetch = async (url)=>{ const u=new URL(url); const a=u.searchParams.get('action');
  if(!a) return {json:async()=>payload()};
  calls2.push(Object.fromEntries(u.searchParams));
  return {json:async()=>({status:'success'})}; };
w.openDetail('stock'); w.showStockPanel(0,0); w.deleteStock(); await wait(400);
ok('删股票写 Sheet', calls2.some(c=>c.action==='deleteStock'&&c.symbol==='ASMIY'), JSON.stringify(calls2.map(c=>c.action)));
calls2.length=0;
w.doTransfer('chase','broker',1000); await wait(400);
ok('转账按增量写券商现金', calls2.some(c=>c.action==='adjustCash'&&Number(c.delta)===1000),
   JSON.stringify(calls2.map(c=>c.action+'='+(c.delta||c.amount||''))));

// ===== 乐观更新：界面是否立刻反映，不等网络 =====
let hold; const slow = new Promise(r=>hold=r);
w.fetch = async (url)=>{ const u=new URL(url); const a=u.searchParams.get('action');
  if(!a) return {json:async()=>payload()};
  await slow;                                  // 模拟 Google 慢 5 秒
  return {json:async()=>({status:'success',row:99})}; };

w.openDetail('cash'); w.switchCashTab('ex'); w.openExpenseDrill('Dining');
const beforeTxt = d.getElementById('cashTabBody').textContent;
w.showAddLedger();
d.getElementById('lg-amt').value='7.77'; d.getElementById('lg-note').value='咖啡';
w.saveLedgerAdd();
await wait(30);                                // 只等 30ms，远小于网络时间
ok('补记后界面立刻更新', /7.77/.test(d.getElementById('cashTabBody').textContent));
ok('页脚显示写入中', /写入中/.test(d.getElementById('syncBadge').textContent), d.getElementById('syncBadge').textContent);
hold(); await wait(200);
ok('写入完成后转为已同步', /已同步/.test(d.getElementById('syncBadge').textContent));

// ===== 写失败要自动撤回 =====
w.fetch = async (url)=>{ const u=new URL(url);
  if(!u.searchParams.get('action')) return {json:async()=>payload()};
  return {json:async()=>({status:'error',message:'模拟失败'})}; };
const txt0 = d.getElementById('cashTabBody').textContent;
w.showAddLedger();
d.getElementById('lg-amt').value='999'; d.getElementById('lg-note').value='应当撤回';
w.saveLedgerAdd(); await wait(200);
ok('失败后自动撤回', !/应当撤回/.test(d.getElementById('cashTabBody').textContent));

// ===== 删除后本地行号跟着上移 =====
w.fetch = async (url)=>{ const u=new URL(url); const a=u.searchParams.get('action');
  if(!a) return {json:async()=>payload()};
  return {json:async()=>({status:'success'})}; };
const rowsBefore = w.eval('ledgerRows.map(r=>r.row).join(",")');
w.openExpenseDrill('__ALL__');
const target = w.eval('ledgerRows.filter(r=>r.row>0).sort((a,b)=>a.row-b.row)[0].row');
const higher = w.eval(`ledgerRows.filter(r=>r.row>${target}).length`);
w.editLedger(target);
w.deleteLedgerRow(target, w.eval(`ledgerRows.find(r=>r.row===${target}).amount`));
await wait(200);
const shifted = w.eval(`ledgerRows.filter(r=>r.row>=${target}).length`);
ok('删除后行号同步上移', shifted===higher, `删前>${target}的有${higher}条，删后≥${target}的有${shifted}条`);
