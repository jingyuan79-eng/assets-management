import { JSDOM } from 'jsdom'; import fs from 'fs';
const gs=fs.readFileSync('AppsScript.gs','utf8');
const html=fs.readFileSync('index.html','utf8');
const ok=(l,c,e='')=>console.log((c?'✅':'❌'),l,e);
const wait=(ms=350)=>new Promise(r=>setTimeout(r,ms));
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
const tabs={
 Bond:mkSheet('Bond',[['ID','Name','Type','Start','Term','Rate','Principal','Balance','LastPost','Status'],
  ['bd_1','T-Note 3Y','T-Note','2024-08-15',3,3.75,10000,10000,'2026-08-15','active'],
  ['bd_2','I Bond','I-Bond','2026-06-01',30,4.26,8000,8000,'2026-06-01','active'],
  ['bd_3','T-Note 3Y (2026)','T-Note','2026-08-15',3,4.25,10000,10000,'2026-08-15','active']]),
 Ledger:mkSheet('Ledger',[['Date','category','amount','note','key']]),
 Anchor:mkSheet('Anchor',[['date','amount','note'],['2026-08-01',10000,'初始值']]),
 Stock:mkSheet('Stock',[['Symbol','Category','Share','Cost','Price'],
  ['NVDA','Semiconductor',32,148.71,214.72],['VOO','IndexETF',12,679,703],
  ['Cash','Gold.coin.cash','','',1326]]),
 Savings:mkSheet('Savings',[['ID','Name','Type','Balance','Rate','Last Post','Next update','Maturity','Status'],
  ['sv_c','OS-9250','OS',30057,3.4,'2026-08-01','2026-09-01','','active']]),
};
globalThis.SpreadsheetApp={flush(){},getActiveSpreadsheet:()=>({getSheetByName:n=>tabs[n]||null,
  insertSheet(n){tabs[n]=mkSheet(n,[]);return tabs[n];}})};
globalThis.ContentService={MimeType:{JSON:'json'},createTextOutput:t=>({_t:t,setMimeType(){return this;},getContent(){return this._t;}})};
globalThis.Utilities={getUuid:()=>Math.random().toString(36).slice(2)+'0000000',
  formatDate:(d,tz,f)=>{const p=new Date(d.getTime()).toISOString();return f==='yyyy-MM'?p.slice(0,7):p.slice(0,10);}};
globalThis.Logger={log:()=>{}}; globalThis.LockService={getScriptLock:()=>({waitLock(){},releaseLock(){}})};
globalThis.ScriptApp={getProjectTriggers:()=>[],newTrigger:()=>({timeBased:()=>({atHour:()=>({everyDays:()=>({create:()=>{}})})})})};
const NOW=new Date('2026-08-25T00:00:00Z'); const RD=Date;
globalThis.Date=class extends RD{constructor(...a){if(a.length===0)super(NOW.getTime());else super(...a);}static now(){return NOW.getTime();}};
const backend=new Function('return (function(){'+gs+'\nreturn {doGet};})()')();

let reads=0, writes=0, delayMs=0;
const dom=new JSDOM(html,{runScripts:'dangerously',url:'https://x.test/',
  beforeParse(w){ w.localStorage.clear();
    w.Date=globalThis.Date;   // 与后端共用冻结时钟，测试才可重复
    w.fetch=async(u)=>{const url=new URL(u);const p={};url.searchParams.forEach((v,k)=>p[k]=v);
      if(p.action) writes++; else reads++;
      if(delayMs) await new Promise(r=>setTimeout(r,delayMs));
      return {json:async()=>JSON.parse(backend.doGet({parameter:p.action?p:{}}).getContent())};};
    w.alert=m=>console.log('   [alert]',m); w.confirm=()=>true; w.scrollTo=()=>{};
    w.Element.prototype.scrollIntoView=function(){};
  }});
const w=dom.window,d=w.document;
await wait();

console.log('— 国债行的显示 —');
w.openDetail('bond');
{
  // 只取第一张卡（持仓表），第二张是 Next payment 表
  const rows=[...d.querySelectorAll('#d-body .stkgroup')][0].querySelectorAll('.row');
  const txt=r=>({sym:r.querySelector('.sym').textContent, qty:r.querySelector('.qty').textContent,
                 desc:r.querySelector('.desc')?r.querySelector('.desc').textContent:'', mv:r.querySelector('.mv').textContent,
                 pc:r.querySelector('.pc')?r.querySelector('.pc').textContent:''});
  [...rows].forEach(r=>{const t=txt(r);console.log(`   ${t.sym} | ${t.qty} | ${t.mv} | ${t.pc}`);});
  ok('右下角是到期日不是占比', /^\d{4}\//.test(txt(rows[0]).pc), txt(rows[0]).pc);
  ok('行内不再有 Next 文案', [...rows].every(r=>txt(r).desc.trim()===''));
  const names=[...rows].map(r=>txt(r).sym);
  ok('I-Bond 只显示 I Bond', !names.some(n=>/I.?Bond 30Y/.test(n)), names.join(' / '));
}

console.log('\n— 新增国债：几次网络往返 —');
reads=0; writes=0; delayMs=200;
w.showBondForm();
d.getElementById('b-start').value='2026-08-20';
d.getElementById('b-term').value='2';
d.getElementById('b-mv').value='5000';
d.getElementById('b-rate').value='4.0';
const btn=d.querySelector('#bondAddWrap .save');
w.saveBond({target:btn});
await wait(60);
ok('立刻出现占位行', /保存中…/.test(d.getElementById('d-body').textContent));
await wait(900);
ok('只发 1 次请求（原来 2 次）', writes===1 && reads===0, `写 ${writes} 次，读 ${reads} 次`);
ok('新账户已入列', w.eval("DATA.bond.groups[0].holdings.some(x=>x.btype==='T-Bill'||x.sym==='T-Note 2Y')"),
   w.eval("DATA.bond.groups[0].holdings.map(x=>x.sym).join(' / ')"));
ok('占位行已消失', !/保存中…/.test(d.getElementById('d-body').textContent));

console.log('\n— 新增储蓄 —');
reads=0; writes=0;
w.openDetail('cd'); w.showCdForm();
d.getElementById('f-name').value='CD-TEST';
d.getElementById('f-mv').value='10000';
d.getElementById('f-rate').value='4.1';
d.getElementById('f-mat').value='2027-08-25';
const btn2=d.querySelector('#cdAddWrap .save');
w.saveCd({target:btn2});
await wait(60);
ok('立刻出现占位行', /保存中…/.test(d.getElementById('d-body').textContent));
await wait(900);
ok('只发 1 次请求', writes===1 && reads===0, `写 ${writes} 次，读 ${reads} 次`);
ok('新账户已入列', w.eval("DATA.cd.groups[0].holdings.some(x=>x.sym==='CD-TEST')"),
   w.eval("DATA.cd.groups[0].holdings.map(x=>x.sym).join(' / ')"));

console.log('\n— 删除 —');
reads=0; writes=0;
const before=w.eval("DATA.bond.groups[0].holdings.length");
w.openDetail('bond'); w.deleteBond(0);
await wait(60);
ok('立刻从列表消失', w.eval("DATA.bond.groups[0].holdings.length")===before-1);
await wait(900);
ok('删除也只 1 次请求', writes===1 && reads===0, `写 ${writes} 次，读 ${reads} 次`);
ok('Sheet 里也删了', tabs.Bond._rows.length===before,  `Bond 表剩 ${tabs.Bond._rows.length-1} 行`);

console.log('\n— 排版调整 —');
{
  w.openDetail('bond');
  const r0=[...d.querySelectorAll('#d-body .row')][0];
    const names=[...d.querySelectorAll('#d-body .sym')].map(e=>e.textContent);
  ok('I-Bond 只显示 I Bond', names.includes('I Bond'), names.join(' / '));
  const pcs=[...d.querySelectorAll('#d-body .stkrows .pc')].map(e=>e.textContent);
  ok('行内只剩到期日，无占比', pcs.every(p=>/^\d{4}\//.test(p)), pcs.join(' / '));
  const gp=d.querySelector('#d-body .stkpc').textContent;
  const expect=(w.eval("classTotal.bond")/w.eval("grand")*100).toFixed(1)+'%';
  ok('组标题显示占投资总额比例', gp===expect, `${gp}（应为 ${expect}）`);
}
{
  w.openDetail('cd');
  const pcs=[...d.querySelectorAll('#d-body .stkrows .pc')];
  ok('储蓄行不显示占比', pcs.length===0, `还剩 ${pcs.length} 个`);
  const gp=d.querySelector('#d-body .stkpc').textContent;
  ok('储蓄组标题也是占总额比例', gp===(w.eval("classTotal.cd")/w.eval("grand")*100).toFixed(1)+'%', gp);
}
{
  w.openDetail('stock');
  const pcs=[...d.querySelectorAll('#d-body .stkrows .pc')];
  ok('股票仍保留个股占比', pcs.length>0 && pcs.every(p=>/%$/.test(p.textContent)),
     pcs.map(p=>p.textContent).join(' / ')||'无');
}
ok('主页不再有半导体集中度', !d.getElementById('semiFill') && !/半导体集中度/.test(d.getElementById('scr-overview').textContent));

console.log('\n— 转账页 —');
{
  w.saveRecurring([{id:'r9',src:'chase',tgt:'broker',amt:600,freq:'biweekly',
                    start:'2026-08-07',lastRun:'2026-08-21'}]);
  w.openTransfer();
  const t=d.getElementById('transfer-body').textContent;
  ok('显示 Next transfer', /Next transfer/.test(t));
  ok('日期算对（8/21+14=9/4）', /2026\/09\/04/.test(t), (t.match(/2026\/\d\d\/\d\d/)||['无'])[0]);
  ok('保留频率与金额', /每两周 \$600/.test(t));
  const html=d.getElementById('transfer-body').innerHTML;
  ok('余额卡片改名 Checking-Chase', /Checking-Chase/.test(html)&&!/>流动现金</.test(html));
}

console.log('\n— 记账起始日：按「最早锚点所在月的 1 号」，不是当天 —');
{
  // 模拟你的情况：只有一条锚点，而且是用「对账」在 8/20 建的
  tabs.Anchor._rows.length=1;
  tabs.Anchor.appendRow(['2026-08-20',12000,'对账']);
  tabs.Ledger._rows.length=1;
  const r=JSON.parse(backend.doGet({parameter:{}}).getContent());
  ok('起始日回退到当月 1 号', r.cash.bookStart==='2026-08-01', '起始日 '+r.cash.bookStart);

  // 设一份 8/1 起、每两周的工资
  await w.syncStockFromSheet(); await wait(500);
  w.openDetail('cash'); w.switchCashTab('in'); w.showPayrollForm();
  d.getElementById('pay-amt').value='2500';
  d.getElementById('pay-freq').value='biweekly';
  d.getElementById('pay-start').value='2026-08-01';
  w.savePayroll(); await wait(1200);
  const pay=tabs.Ledger._rows.filter(x=>String(x[4]||'').startsWith('pay:'));
  // 冻结在 2026-08-25，8/1 起每两周 → 8/1、8/15 两笔
  const want=2;
  ok('8/1 起的工资被正确补记', pay.length===want,
     `应 ${want} 笔，实际 ${pay.length} 笔：`+pay.map(x=>x[0]).join(', '));
  ok('全部落在 8 月', pay.every(x=>String(x[0]).startsWith('2026-08')));
  w.switchCashTab('in');
  const t=d.getElementById('cashTabBody').textContent;
  ok('本月收入不再是 0', new RegExp('本月收入\\$'+(want*2500).toLocaleString('en-US')).test(t.replace(/\s/g,'')),
     t.replace(/\s+/g,'').slice(0,50));
  ok('收入明细能看到', /本月收入明细/.test(t)&&/Payroll/.test(t));
}

console.log('\n— 国债页 Next payment 表 —');
{
  tabs.Bond._rows.length=1;
  tabs.Bond.appendRow(['bd_1','T-Note 3Y','T-Note','2024-08-15',3,3.75,10000,10000,'2026-08-15','active']);
  tabs.Bond.appendRow(['bd_2','I Bond','I-Bond','2026-06-01',30,4.26,8000,8000,'2026-06-01','active']);
  tabs.Bond.appendRow(['bd_3','T-Note 3Y','T-Note','2026-08-15',3,4.25,10000,10000,'2026-08-15','active']);
  await w.syncStockFromSheet(); await wait(600);
  w.openDetail('bond');
  const descs=[...d.querySelectorAll('#d-body .stkrows .desc')].map(e=>e.textContent.trim());
  ok('行内 Next 文案已删除', descs.every(t=>t===''), JSON.stringify(descs));
  const groups=[...d.querySelectorAll('#d-body .stkgroup')];
  ok('多出 Next payment 表', groups.length===2 && /Next payment/.test(groups[1].textContent));
  const nrows=[...groups[1].querySelectorAll('.row')];
  ok('只列 T-Note，不含 I Bond', nrows.length===2 && !groups[1].textContent.includes('I Bond'),
     nrows.map(r=>r.textContent.replace(/\s+/g,' ')).join(' | '));
  const t=groups[1].textContent;
  ok('日期与金额正确', /2027\/02\/15/.test(t) && /\$187\.50/.test(t) && /\$212\.50/.test(t), t.replace(/\s+/g,' '));
  ok('表头显示合计 $400.00', /\$400\.00/.test(groups[1].querySelector('.stkpc').textContent),
     groups[1].querySelector('.stkpc').textContent);
}

console.log('\n— 转账页合并为一个表单 —');
{
  w.openTransfer();
  const body=d.getElementById('transfer-body');
  ok('余额卡标题改成 Checking-Chase', /Checking-Chase/.test(body.textContent)&&!/流动现金\$/.test(body.textContent));
  ok('只有一套 From/To/金额', !!d.getElementById('x-src')&&!!d.getElementById('x-tgt')&&!!d.getElementById('x-amt')
     && !d.getElementById('o-src') && !d.getElementById('r-src'));
  ok('顶部有两个切换按钮', /一次性转账[\s\S]*定期转账/.test(body.textContent));
  ok('默认一次性：不显示频率与首次日', !d.getElementById('x-freq') && !d.getElementById('x-start'));
  ok('只有一个确认按钮', [...body.querySelectorAll('.btns button')].length===1
     && /确认/.test(body.querySelector('.btns button').textContent));

  // 切到定期
  w.setXferMode('recurring');
  ok('切换后出现频率与首次转账日', !!d.getElementById('x-freq') && !!d.getElementById('x-start'));
  ok('字段沿用同一套 id', !!d.getElementById('x-src')&&!!d.getElementById('x-amt'));

  // 定期提交走 submitRecurring
  const n0=JSON.parse(w.eval("JSON.stringify(loadRecurring())")).length;
  d.getElementById('x-src').value='chase'; d.getElementById('x-tgt').value='broker';
  d.getElementById('x-amt').value='600';
  d.getElementById('x-freq').value='biweekly';
  d.getElementById('x-start').value='2026-09-01';
  w.submitTransfer();
  await wait(500);
  const list=JSON.parse(w.eval("JSON.stringify(loadRecurring())"));
  ok('保存为定期转账', list.length===n0+1 && list.at(-1).freq==='biweekly' && list.at(-1).amt===600,
     JSON.stringify(list.at(-1)));

  // 切回一次性并提交
  w.setXferMode('once');
  const bal0=w.eval("cashInfo.balance");
  d.getElementById('x-src').value='chase'; d.getElementById('x-tgt').value='broker';
  d.getElementById('x-amt').value='100';
  w.submitTransfer(); await wait(500);
  ok('一次性转账立刻扣款', Math.abs(w.eval("cashInfo.balance")-(bal0-100))<0.01,
     `$${bal0} → $${w.eval("cashInfo.balance")}`);
  ok('没有多存一条定期', JSON.parse(w.eval("JSON.stringify(loadRecurring())")).length===n0+1);
  ok('已设置列表显示 Next transfer', /Next transfer/.test(d.getElementById('transfer-body').textContent));
}

console.log('\n— 表单收紧 —');
{
  const css=fs.readFileSync('index.html','utf8');
  ok('输入框内边距 8px', /\.editform input,\.editform select\{[^}]*padding:8px 10px/.test(css));
  ok("字段间距 8px", /\.editform \.fld\{margin-bottom:8px;\}/.test(css));
  ok('表单内边距 12px', /\.editform\{[^}]*padding:12px/.test(css));
  ok('输入字号仍是 16px（避免 iOS 聚焦缩放）', /\.editform input,\.editform select\{[^}]*font-size:16px/.test(css));
  ok('不再有 15px 大号内联下拉样式', !/padding:15px 14px/.test(css));
}

console.log('\n— 转账页余额卡左右并排 —');
{
  w.openTransfer();
  const body=d.getElementById('transfer-body');
  const first=body.firstElementChild;
  ok('第一块是 flex 横向容器', /display:flex/.test(first.getAttribute('style')||''), first.getAttribute('style'));
  const cards=[...first.children];
  ok('两张卡并排', cards.length===2, `共 ${cards.length} 张`);
  ok('左 Checking-Chase / 右 Marcus OS',
     /Checking-Chase/.test(cards[0].textContent) && /Marcus OS/.test(cards[1].textContent),
     cards.map(c=>c.textContent.replace(/\s+/g,' ')).join('  |  '));
  ok('两张卡各占一半', cards.every(c=>/flex:1/.test(c.getAttribute('style')||'')));
}
