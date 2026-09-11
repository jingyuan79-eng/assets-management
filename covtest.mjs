// 补三处覆盖盲区：70 天流水窗口、整行回写不误伤其他列、首屏本地缓存。
// 这三处是 2026-08 改动最密集的地方，原来一条断言都没有。
import { JSDOM } from 'jsdom'; import fs from 'fs';
const gs=fs.readFileSync('AppsScript.gs','utf8');
const html=fs.readFileSync('index.html','utf8');
const ok=(l,c,e='')=>console.log((c?'✅':'❌'),l,e);
// 只取渲染出来的文字：内联 <script> 的源码不算页面内容，
// 否则代码里出现 isNaN、或注释里写了 NaN，都会让断言误报。
const visibleText=(doc)=>{
  const b=doc.body.cloneNode(true);
  [].forEach.call(b.querySelectorAll('script,style'),n=>n.remove());
  return b.textContent||'';
};
const wait=(ms=350)=>new Promise(r=>setTimeout(r,ms));

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
     for(let j=0;j<nc;j++)row[c-1+j]=v[i][j];}},
   clearContent(){for(let i=0;i<nr;i++){const row=self._cell(r+i,c+nc-1);
     for(let j=0;j<nc;j++)row[c-1+j]="";}}};},
 appendRow(a){this._rows.push(a.slice());},deleteRow(r){this._rows.splice(r-1,1);}};}

const pad=n=>String(n).padStart(2,'0');
const fmt=d=>`${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}`;
const daysAgo=n=>{const d=new Date();d.setHours(0,0,0,0);d.setDate(d.getDate()-n);return d;};
const ymOfD=d=>`${d.getFullYear()}-${pad(d.getMonth()+1)}`;

function backend(tabs){
  globalThis.SpreadsheetApp={flush(){},getActiveSpreadsheet:()=>({getSheetByName:n=>tabs[n]||null,
    insertSheet(n){tabs[n]=mkSheet(n,[]);return tabs[n];}})};
  globalThis.ContentService={MimeType:{JSON:'json'},
    createTextOutput:t=>({_t:t,setMimeType(){return this;},getContent(){return this._t;}})};
  globalThis.Utilities={formatDate:(d,tz,f)=>{
    const p=new Date(d.getTime()-7*3600e3).toISOString();
    return f==='yyyy-MM'?p.slice(0,7):p.slice(0,10);},
    getUuid:()=>'abcdefgh-1234'};
  globalThis.LockService={getScriptLock:()=>({waitLock(){},releaseLock(){}})};
  globalThis.Logger={log(){}};
  globalThis.ScriptApp={};
  return new Function('return (function(){'+gs+'\nreturn {doGet,doPost};})()')();
}
const J=r=>JSON.parse(r.getContent());

// ══════════════ 一、70 天流水窗口 ══════════════
// 前端 dueDates / bookFloor 用 60 天窗口判断「这笔要不要补记」，去重靠在
// ledgerRows 里查 key。后端返回的明细窗口必须盖过 60 天，否则每次打开都会
// 把早就写过的固定支出、工资重发一遍。
console.log('— 70 天流水窗口 —');
{
  const D5=daysAgo(5), D40=daysAgo(40), D65=daysAgo(65), D80=daysAgo(80), D200=daysAgo(200);
  const ledger=mkSheet('Ledger',[['Date','category','amount','note','key'],
    [fmt(D5),'Dining',10,'','fix:a:'+fmt(D5)],
    [fmt(D40),'Dining',20,'','fix:a:'+fmt(D40)],
    [fmt(D65),'Dining',30,'','fix:a:'+fmt(D65)],   // 60~70 天之间：修复前查不到的缝隙
    [fmt(D80),'Dining',40,'','fix:a:'+fmt(D80)],
    [fmt(D200),'Dining',50,'','fix:a:'+fmt(D200)]]);
  const tabs={Ledger:ledger,
    Stock:mkSheet('Stock',[['Symbol','Category','Share','Cost','Price']]),
    Anchor:mkSheet('Anchor',[['date','amount','note'],[fmt(daysAgo(300)),1000,'初始']]),
    Savings:mkSheet('Savings',[['ID','Name','Type','Balance','Rate','Last Post','Next update','Maturity','Status']]),
    Bond:mkSheet('Bond',[['ID','Name','Type','Start','Term','Rate','Principal','Balance','LastPost','Status']]),
    Ledger_monthly:mkSheet('Ledger_monthly',[])};
  const {doGet}=backend(tabs);
  const r=J(doGet({parameter:{}}));
  const has=d=>r.ledger.some(x=>x.date===fmt(d));
  const now=new Date(), prevM=new Date(now.getFullYear(),now.getMonth()-1,1);
  const months=[ymOfD(now),ymOfD(prevM)];
  const inWindow=d=>fmt(d)>=fmt(daysAgo(70)) || months.includes(ymOfD(d));

  ok('5 天前的流水在明细里', has(D5));
  ok('40 天前的流水在明细里', has(D40));
  ok('65 天前的流水在明细里（修复前落在缝隙里查不到）', has(D65), fmt(D65));
  ok('200 天前的流水不在明细里', !has(D200));
  ok('80 天前的按规则处理', has(D80)===inWindow(D80),
     `期望${inWindow(D80)?'在':'不在'} 实际${has(D80)?'在':'不在'}`);
  ok('明细窗口盖过前端 60 天的补记判断窗口',
     r.ledger.every(x=>x.date>=fmt(daysAgo(70))||months.includes(x.date.slice(0,7))));
  ok('ledgerMonths 仍只声明当月+上月', JSON.stringify(r.ledgerMonths)===JSON.stringify(months),
     JSON.stringify(r.ledgerMonths));
  ok('ledgerMonths 声明的月份必定被完整覆盖',
     months.every(m=>{const s=new Date(m+'-01T00:00:00');return fmt(s)>=fmt(daysAgo(70));}));
  ok('key 一并返回（前端靠它去重）', r.ledger.every(x=>'key' in x));
}

// ══════════════ 二、整行回写不误伤其他列 ══════════════
// writeSavRow / writeBondRow 为省往返改成整行 setValues，不由它们计算的列
// 从 raw 原样带回。一旦哪次改动漏了 raw，用户填的利率、名称就会被冲掉。
console.log('\n— 整行回写不误伤其他列 —');
{
  const sav=mkSheet('Savings',[
   ['ID','Name','Type','Balance','Rate','Last Post','Next update','Maturity','Status'],
   ['sv_1','CD-7597','CD',21335,4.1,fmt(daysAgo(3)),'','2027-06-30','active']]);
  const tabs={Savings:sav,
    Stock:mkSheet('Stock',[['Symbol','Category','Share','Cost','Price']]),
    Ledger:mkSheet('Ledger',[['Date','category','amount','note','key']]),
    Anchor:mkSheet('Anchor',[['date','amount','note']]),
    Bond:mkSheet('Bond',[['ID','Name','Type','Start','Term','Rate','Principal','Balance','LastPost','Status']]),
    Ledger_monthly:mkSheet('Ledger_monthly',[])};
  const {doGet}=backend(tabs);

  // 只改余额：其余列必须原封不动
  J(doGet({parameter:{action:'updateSaving',id:'sv_1',balance:22000}}));
  const row=sav._rows[1];
  ok('余额已更新', row[3]===22000, String(row[3]));
  ok('名称未被冲掉', row[1]==='CD-7597', String(row[1]));
  ok('类型未被冲掉', row[2]==='CD', String(row[2]));
  ok('利率保持用户写法 4.1（没被写成 0.041）', row[4]===4.1, String(row[4]));
  ok('到期日未被冲掉', row[7]==='2027-06-30', String(row[7]));

  // 改利率：新值必须落盘，不能被随后的整行回写还原
  J(doGet({parameter:{action:'updateSaving',id:'sv_1',rate:3.5}}));
  ok('改利率后新值落盘', Math.abs(sav._rows[1][4]-0.035)<1e-9, String(sav._rows[1][4]));
  ok('改利率不影响名称', sav._rows[1][1]==='CD-7597');
  ok('改利率不影响余额', sav._rows[1][3]===22000, String(sav._rows[1][3]));

  // 国债：runBonds 在读操作里推进付息，会触发 writeBondRow
  const start=new Date(); start.setHours(0,0,0,0); start.setMonth(start.getMonth()-7);
  const bond=mkSheet('Bond',[
   ['ID','Name','Type','Start','Term','Rate','Principal','Balance','LastPost','Status'],
   ['bd_1','T-Note 2y','T-Note',fmt(start),2,4.5,10000,10000,fmt(start),'active']]);
  tabs.Bond=bond;
  J(doGet({parameter:{}}));
  const b=bond._rows[1];
  ok('国债 LastPost 已推进', b[8]!==fmt(start), String(b[8]));
  ok('国债名称未被冲掉', b[1]==='T-Note 2y', String(b[1]));
  ok('国债起始日未被冲掉', b[3]===fmt(start), String(b[3]));
  ok('国债期限未被冲掉', b[4]===2, String(b[4]));
  ok('国债利率保持用户写法 4.5', b[5]===4.5, String(b[5]));
  ok('国债本金未被冲掉', b[6]===10000, String(b[6]));
}

// ══════════════ 三、首屏本地缓存 ══════════════
// 打开时先用上次同步结果填满界面，不空等网络；但必须标明数据时点，
// 绝不让过期余额冒充实时余额。
console.log('\n— 首屏本地缓存 —');
{
  const FRESH={status:'success',
    stock:[{symbol:'Cash',category:'Gold.coin.cash',shares:'',cost:'',price:1326}],
    expense:{},ledger:[],monthly:{},ledgerMonths:['2026-08','2026-07'],
    cash:{balance:12345,anchorDate:'2026-08-01',anchorAmount:10000,hasAnchor:true,bookStart:'2025-01-01'},
    savings:[{id:'sv_1',name:'OS-9250',type:'OS',balance:30000,rate:0.034,
              lastPost:'2026-08-01',nextUpdate:'2026-09-01',maturity:'',status:'active'}],
    bond:[{id:'bd_1',name:'T-Note',type:'T-Note',start:'2026-02-15',term:2,rate:0.045,
           principal:10000,balance:10000,lastPost:'2026-08-15',maturity:'2028-02-15',status:'active'}],
    notices:[],hsa:[],retire:[],serverDate:'2026-08-29'};
  const mkDom=(seed)=>new JSDOM(html,{runScripts:'dangerously',url:'https://x.test/',
    beforeParse(w){ w.localStorage.clear();
      if(seed) w.localStorage.setItem('sheetCache_v1', seed);
      w.fetch=()=>new Promise(res=>setTimeout(()=>res({json:async()=>FRESH}),300));
      w.alert=()=>{}; w.confirm=()=>true; w.scrollTo=()=>{};
      w.Element.prototype.scrollIntoView=function(){}; }});

  // ---- 首次安装：没有缓存，不能显示写死的假数字 ----
  const d1=mkDom(null), w1=d1.window;
  ok('首次安装：国债为 0，不闪硬编码种子', w1.eval("classTotal.bond")===0, '$'+w1.eval("classTotal.bond"));
  ok('首次安装：储蓄为 0', w1.eval("classTotal.cd")===0, '$'+w1.eval("classTotal.cd"));
  ok('首次安装：现金为 0', w1.eval("classTotal.cash")===0, '$'+w1.eval("classTotal.cash"));
  ok('首次安装：徽标显示同步中，而非数据时点',
     /同步中/.test(w1.document.getElementById('syncBadge').textContent),
     w1.document.getElementById('syncBadge').textContent);
  await wait(500);
  ok('同步后现金到位', w1.eval("cashInfo.balance")===12345, '$'+w1.eval("cashInfo.balance"));
  ok('同步后写入缓存', !!w1.localStorage.getItem('sheetCache_v1'));
  ok('同步后 cacheTs 归零（屏幕上不再是缓存）', w1.eval("cacheTs")===0);
  ok('ASOF 采用后端 serverDate', w1.eval("ASOF")==='2026-08-29', w1.eval("ASOF"));

  const cached=JSON.parse(w1.localStorage.getItem('sheetCache_v1'));
  ok('缓存含 cash/savings/bond 三块', !!cached.cash&&!!cached.savings&&!!cached.bond);
  ok('缓存带时间戳', typeof cached.ts==='number'&&cached.ts>0);

  // ---- 第二次打开：拿上一轮的缓存垫底 ----
  const seed=JSON.stringify(Object.assign({},cached,{ts:Date.now()-26*3600*1000}));
  const d2=mkDom(seed), w2=d2.window;
  ok('第二次打开：储蓄立即有数（不空等网络）', w2.eval("classTotal.cd")===30000, '$'+w2.eval("classTotal.cd"));
  ok('第二次打开：现金立即有数', w2.eval("cashInfo.balance")===12345, '$'+w2.eval("cashInfo.balance"));
  ok('第二次打开：国债立即有数', w2.eval("classTotal.bond")>0, '$'+w2.eval("classTotal.bond"));
  ok('cacheTs 标记「当前是缓存」', w2.eval("cacheTs")>0);
  const badge=w2.document.getElementById('syncBadge').textContent;
  ok('徽标标明数据时点，不冒充实时', /数据为|上次同步/.test(badge)&&/更新中/.test(badge), badge);
  await wait(500);
  ok('新数据到达后 cacheTs 归零', w2.eval("cacheTs")===0);
  ok('新数据到达后徽标转为已同步',
     /已同步/.test(w2.document.getElementById('syncBadge').textContent),
     w2.document.getElementById('syncBadge').textContent);

  // ---- localStorage 不可用（隐私模式）----
  // ⚠️ 已知限制：目前只有 saveSheetCache 包了 try/catch，saveBrokerCash /
  // persistStock 等写入点没有。所以 localStorage 一抛异常，rebuildStockGroups
  // 会中断，整个同步落进 catch，界面起得来但一条数据都拿不到。
  // 下面如实断言当前行为；等哪天把所有写入点都包上，改成断言「仍能拿到数据」。
  const errs=[];
  const d3=new JSDOM(html,{runScripts:'dangerously',url:'https://x.test/',
    beforeParse(w){ w.localStorage.clear();
      Object.defineProperty(w,'localStorage',{get(){throw new Error('隐私模式');}});
      w.fetch=()=>new Promise(res=>setTimeout(()=>res({json:async()=>FRESH}),300));
      w.alert=()=>{}; w.confirm=()=>true; w.scrollTo=()=>{};
      w.console.log=(...a)=>errs.push(a.join(' '));
      w.Element.prototype.scrollIntoView=function(){}; }});
  ok('隐私模式下界面仍能渲染，不白屏', !!d3.window.document.getElementById('syncBadge'));
  await wait(500);
  ok('隐私模式下同步失败会明确提示，不静默', /同步失败/.test(
     d3.window.document.getElementById('syncBadge').textContent),
     d3.window.document.getElementById('syncBadge').textContent);
  ok('⚠️ 已知限制：隐私模式下拿不到数据（写入点未全部包 try/catch）',
     errs.some(m=>/sheet sync failed/.test(m)),
     '—— 修好后请把这条改成断言能拿到数据');
}

// ══════════════ 四、HSA 是第 5 种资金性质 ══════════════
// HSA 的钱走独立账户，不从 Chase 出。记进 Ledger 但对流动现金贡献为 0，
// 也不能混进「本月支出」——没刷 HSA 卡的同名分类必须留在原支出板块。
console.log('\n— HSA 独立成账 —');
{
  const today=new Date(); const ym=`${today.getFullYear()}-${pad(today.getMonth()+1)}`;
  const d=n=>`${ym}-${pad(n)}`;
  const ledger=mkSheet('Ledger',[['Date','category','amount','note','key'],
    [d(3),'Dining',200,'',''],
    [d(5),'Health & Beauty',80,'没刷HSA卡',''],          // 留在支出板块
    [d(6),'HSA · Income · 供款',250,'本人',''],
    [d(6),'HSA · Income · 雇主补助',100,'雇主',''],
    [d(12),'HSA · Health & Beauty',60,'牙医',''],        // 同名分类，但刷了 HSA 卡
    [d(18),'HSA · 处方药',45.5,'','']]);
  const tabs={Ledger:ledger,
    Stock:mkSheet('Stock',[['Symbol','Category','Share','Cost','Price']]),
    Anchor:mkSheet('Anchor',[['date','amount','note'],[d(1),10000,'初始']]),
    Savings:mkSheet('Savings',[['ID','Name','Type','Balance','Rate','Last Post','Next update','Maturity','Status']]),
    Bond:mkSheet('Bond',[['ID','Name','Type','Start','Term','Rate','Principal','Balance','LastPost','Status']]),
    Ledger_monthly:mkSheet('Ledger_monthly',[])};
  const {doGet}=backend(tabs);
  const r=J(doGet({parameter:{}}));
  const m=r.monthly[ym]||{};

  ok('HSA 流水不减流动现金（锚点 10000 − 仅非 HSA 支出 280）', r.cash.balance===9720, '$'+r.cash.balance);
  ok('本月支出不含 HSA', m.expense===280, '$'+m.expense);
  ok('没刷卡的 Health & Beauty 仍在支出分类里', (r.expense['Health & Beauty']||0)===80,
     '$'+(r.expense['Health & Beauty']||0));
  ok('刷了卡的 HSA · Health & Beauty 不在支出分类里',
     !Object.keys(r.expense).some(k=>k.indexOf('HSA')===0), Object.keys(r.expense).join(','));
  ok('HSA 收入单独汇总', m.hsaIn===350, '$'+m.hsaIn);
  ok('HSA 支出单独汇总', m.hsaOut===105.5, '$'+m.hsaOut);
  ok('HSA 不进净收入', m.net===-280, '$'+m.net);
  ok('HSA 不进投资转出/本金回流', m.transfer===0 && m.redeem===0);
  ok('HSA 明细照常返回给前端', r.ledger.filter(x=>x.category.indexOf('HSA')===0).length===4);
  ok('Ledger_monthly 新增 HSA 两列',
     tabs.Ledger_monthly._rows[0].includes('HSA 收入') && tabs.Ledger_monthly._rows[0].includes('HSA 支出'));

  // 前端同一套规则
  const dom=new JSDOM(html,{runScripts:'dangerously',url:'https://x.test/',
    beforeParse(w){ w.localStorage.clear();
      w.fetch=()=>Promise.resolve({json:async()=>r});
      w.alert=()=>{}; w.confirm=()=>true; w.scrollTo=()=>{};
      w.Element.prototype.scrollIntoView=function(){}; }});
  const w=dom.window; await wait();
  ok('前端 catKind 与后端一致', w.eval("catKind('HSA · 处方药')")==='hsa');
  ok('前端 catSign 对 HSA 返回 0', w.eval("catSign('HSA · 处方药')")===0);
  ok('前端支出方块不含 HSA', w.eval("JSON.stringify(monthExpense)").indexOf('HSA')<0,
     w.eval("JSON.stringify(monthExpense)"));
  ok('HSA 收入合计', w.eval("hsaTotal('in')")===350, '$'+w.eval("hsaTotal('in')"));
  ok('HSA 支出合计', w.eval("hsaTotal('out')")===105.5, '$'+w.eval("hsaTotal('out')"));
  w.openDetail('hsa');
  ok('HSA 页默认停在支出', w.eval("hsaTab")==='ex');
  ok('HSA 页列出开销明细', /牙医|处方药/.test(w.document.getElementById('hsaTabBody').textContent));
  w.switchHsaTab('in');
  ok('切到收入页列出供款', /雇主|本人/.test(w.document.getElementById('hsaTabBody').textContent));
}

// ══════════════ 五、HSA 供款与 Shortcut 分类归一 ══════════════
console.log('\n— HSA 供款与分类归一 —');
{
  const ledger=mkSheet('Ledger',[['Date','category','amount','note','key']]);
  const tabs={Ledger:ledger,
    Stock:mkSheet('Stock',[['Symbol','Category','Share','Cost','Price']]),
    Anchor:mkSheet('Anchor',[['date','amount','note'],[fmt(daysAgo(30)),10000,'初始']]),
    Savings:mkSheet('Savings',[['ID','Name','Type','Balance','Rate','Last Post','Next update','Maturity','Status']]),
    Bond:mkSheet('Bond',[['ID','Name','Type','Start','Term','Rate','Principal','Balance','LastPost','Status']]),
    Ledger_monthly:mkSheet('Ledger_monthly',[])};
  const {doGet}=backend(tabs);

  // Shortcut 只发「HSA」，后端要补成带前缀的形式
  let r=J(doGet({parameter:{action:'addLedger',date:fmt(daysAgo(1)),category:'HSA',amount:30}}));
  ok('Shortcut 的「HSA」被补成 HSA · 医疗', r.category==='HSA · 医疗', r.category);
  ok('补前缀后归到 hsa 而非支出', ledger._rows[1][1]==='HSA · 医疗', String(ledger._rows[1][1]));

  // 普通分类不受影响
  r=J(doGet({parameter:{action:'addLedger',date:fmt(daysAgo(1)),category:'Dining',amount:20}}));
  ok('普通分类不被改写', r.category==='Dining', r.category);

  // autoLedger（固定支出）走同一套归一
  J(doGet({parameter:{action:'autoLedger',key:'fix:hsa1:x',date:fmt(daysAgo(1)),
                      category:'HSA',amount:120,note:'每月理疗'}}));
  ok('固定支出的 HSA 也补前缀',
     ledger._rows.some(x=>String(x[4])==='fix:hsa1:x' && x[1]==='HSA · 医疗'));

  const out=J(doGet({parameter:{}}));
  ok('HSA 开销不减流动现金（锚点 10000 − 仅 Dining 20）', out.cash.balance===9980, '$'+out.cash.balance);

  // 前端：供款配置驱动的自动入账
  const dom=new JSDOM(html,{runScripts:'dangerously',url:'https://x.test/',
    beforeParse(w){ w.localStorage.clear();
      w.__calls=[];
      w.fetch=(u)=>{ const url=new URL(u), a=url.searchParams.get('action');
        if(a){ w.__calls.push(Object.fromEntries(url.searchParams));
               return Promise.resolve({json:async()=>({status:'success',row:99})}); }
        return Promise.resolve({json:async()=>({status:'success',stock:[],expense:{},ledger:[],
          monthly:{},ledgerMonths:[],cash:{balance:0,hasAnchor:false},savings:[],bond:[],
          notices:[],hsa:[],retire:[],serverDate:'2026-08-29'})}); };
      w.alert=()=>{}; w.confirm=()=>true; w.scrollTo=()=>{};
      w.Element.prototype.scrollIntoView=function(){}; }});
  const w=dom.window; await wait();
  const d0=new Date(); d0.setHours(0,0,0,0); d0.setDate(d0.getDate()-14);
  const startStr=`${d0.getFullYear()}-${pad(d0.getMonth()+1)}-${pad(d0.getDate())}`;
  w.eval(`saveHsaCfg({employee:250,employer:100,freq:'biweekly',start:'${startStr}'})`);
  await w.eval("runHsaContrib()");
  const calls=w.__calls.filter(c=>c.action==='autoLedger');
  ok('供款按发薪日入账（本人一笔、雇主一笔）',
     calls.some(c=>c.key.indexOf('hsaee:')===0 && c.category==='HSA · Income · 供款' && +c.amount===250) &&
     calls.some(c=>c.key.indexOf('hsaer:')===0 && c.category==='HSA · Income · 雇主补助' && +c.amount===100),
     calls.map(c=>c.key).join(','));
  ok('两周起始日 → 至少两个发薪日', new Set(calls.map(c=>c.key.split(':')[1])).size>=2);
  const before=w.__calls.length;
  await w.eval("runHsaContrib()");
  ok('已入账的不再重发（key 去重）', w.__calls.length===before, `新增 ${w.__calls.length-before} 次`);
  ok('未配置时不发请求', w.eval("(function(){saveHsaCfg({});return hsaPayDatesDue(loadHsaCfg()).length;})()")===0);
}

// ══════════════ 六、HSA 合计直接显示在 tab 上 ══════════════
// tab 的方向标识是 in/ex，流水方向是 in/out，取合计时必须换算 ——
// 弄混的话支出 tab 会永远显示 $0（曾经就是这么错的）。
console.log('\n— HSA tab 上的合计 —');
{
  const today=new Date(); const ym=`${today.getFullYear()}-${pad(today.getMonth()+1)}`;
  const d=n=>`${ym}-${pad(n)}`;
  const LED=[
    {row:2,date:d(7),category:'HSA · Income · 供款',amount:250,note:'本人',key:'hsaee:'+d(7)},
    {row:3,date:d(7),category:'HSA · Income · 雇主补助',amount:100,note:'雇主',key:'hsaer:'+d(7)},
    {row:4,date:d(12),category:'HSA · 医疗',amount:60,note:'牙医',key:''},
    {row:5,date:d(18),category:'HSA · 处方药',amount:45.5,note:'',key:''}];
  const dom=new JSDOM(html,{runScripts:'dangerously',url:'https://x.test/',
    beforeParse(w){ w.localStorage.clear();
      w.fetch=()=>Promise.resolve({json:async()=>({status:'success',stock:[],expense:{},
        ledger:LED,monthly:{},ledgerMonths:[ym],cash:{balance:0,hasAnchor:false},
        savings:[],bond:[],notices:[],hsa:[],retire:[],serverDate:'2026-08-29'})});
      w.alert=()=>{}; w.confirm=()=>true; w.scrollTo=()=>{};
      w.Element.prototype.scrollIntoView=function(){}; }});
  const w=dom.window; await wait();
  w.openDetail('hsa');
  const tabs=()=>[].map.call(w.document.querySelectorAll('#hsaWrap .cashtab'),t=>t.textContent);
  ok('收入 tab 带出本月合计', tabs()[0]==='本月收入$350', tabs()[0]);   // 250 本人 + 100 雇主
  ok('支出 tab 带出本月合计（不是 $0）', tabs()[1]==='本月医疗开销$106', tabs()[1]);
  ok('页面里不再有单独的汇总框',
     !/本月存入\$/.test(w.document.getElementById('hsaTabBody').textContent));
  w.switchHsaTab('in');
  ok('切 tab 后合计仍在', tabs()[0]==='本月收入$350' && tabs()[1]==='本月医疗开销$106', tabs().join(' / '));
}

// ══════════════ 七、HSA 余额：锚点 + 流水 + 发薪日再分配 ══════════════
// Cash 随 HSA 收支增减；收入入账那天先给 Investment 结息，再把 Cash
// 超过 Floor 的部分扫进 Investment。全量重放，补记旧账也能自愈。
console.log('\n— HSA 余额引擎 —');
{
  const HHEAD=['Name','Anchor Date','Anchor Amount','Rate','Floor','Sweep','Balance(自动算)','Updated'];
  function runH(ledgerRows, hsaRows){
    const L=mkSheet('Ledger',[['Date','category','amount','note','key'],...ledgerRows]);
    const H=mkSheet('HSA',[HHEAD,...hsaRows]);
    const tabs={Ledger:L,HSA:H,
      Stock:mkSheet('Stock',[['Symbol','Category','Share','Cost','Price']]),
      Anchor:mkSheet('Anchor',[['date','amount','note']]),
      Savings:mkSheet('Savings',[['ID','Name','Type','Balance','Rate','Last Post','Next update','Maturity','Status']]),
      Bond:mkSheet('Bond',[['ID','Name','Type','Start','Term','Rate','Principal','Balance','LastPost','Status']]),
      Ledger_monthly:mkSheet('Ledger_monthly',[])};
    const {doGet}=backend(tabs);
    return {r:J(doGet({parameter:{}})).hsa, H};
  }
  const DAILY=Math.pow(1.10,1/365)-1;
  const near=(n,a,b,t)=>ok(n, Math.abs(a-b)<(t||0.05), `${a}`);

  // 发薪日：入账 → 结息 → 扫到 Floor
  let {r,H}=runH([['2026-08-15','HSA · Income · 供款',250,'',''],
                  ['2026-08-15','HSA · Income · 雇主补助',100,'','']],
                 [['Cash','2026-08-01',1900,'',2000,'biweekly','',''],
                  ['Investment','2026-08-01',10000,10,'','','','']]);
  near('Cash 被扫到 Floor', r.cash, 2000);
  near('Investment = 结息 14 天 + 扫入 250', r.investment, 10000*Math.pow(1+DAILY,14)+250);
  ok('记录了扫账日', r.lastSweep==='2026-08-15', r.lastSweep);
  ok('余额回写进 Sheet 的 Balance 列', Number(H._rows[1][6])===r.cash);
  ok('锚点列不被引擎改写', String(H._rows[1][1])==='2026-08-01' && Number(H._rows[1][2])===1900);

  // 花超了：没到 Floor 就不转
  r=runH([['2026-08-10','HSA · 医疗',800,'',''],
          ['2026-08-15','HSA · Income · 供款',250,'',''],
          ['2026-08-15','HSA · Income · 雇主补助',100,'','']],
         [['Cash','2026-08-01',1900,'',2000,'biweekly','',''],
          ['Investment','2026-08-01',10000,10,'','','','']]).r;
  near('Cash = 1900 − 800 + 350', r.cash, 1450);
  near('Investment 只结息不扫入', r.investment, 10000*Math.pow(1+DAILY,14));
  ok('没到 Floor 就没有扫账日', r.lastSweep==='', r.lastSweep);

  // 只有开销的日子不触发再分配
  r=runH([['2026-08-20','HSA · 处方药',45.5,'','']],
         [['Cash','2026-08-01',2500,'',2000,'biweekly','',''],
          ['Investment','2026-08-01',10000,10,'','','','']]).r;
  near('Cash 直接减，没被扫', r.cash, 2454.5);
  near('Investment 一分没动', r.investment, 10000);

  // 两个发薪日：第二次从上次结息日续滚
  r=runH([['2026-08-01','HSA · Income · 供款',250,'',''],
          ['2026-08-15','HSA · Income · 供款',250,'','']],
         [['Cash','2026-08-01',2000,'',2000,'biweekly','',''],
          ['Investment','2026-08-01',10000,10,'','','','']]).r;
  near('两次扫入且中间正确结息', r.investment, 10250*Math.pow(1+DAILY,14)+250);

  // 补记更早的开销 → 自愈
  const base=[['2026-08-15','HSA · Income · 供款',250,'','']];
  const seed=[['Cash','2026-08-01',1900,'',2000,'biweekly','',''],
              ['Investment','2026-08-01',10000,10,'','','','']];
  near('补记前 Cash 被扫到 Floor', runH(base,seed).r.cash, 2000);
  const after=runH([['2026-08-05','HSA · 医疗',300,'',''],...base],seed).r;
  near('补记 8/05 的 300 后自愈重算', after.cash, 1850);
  ok('自愈后不再触发扫账', after.lastSweep==='');

  // 前端显示
  const info={cash:2000, investment:10286.62, rate:0.10, floor:2000,
              anchorDate:'2026-08-01', invPost:'2026-08-15', updated:'2026-08-29', ready:true};
  const dom=new JSDOM(html,{runScripts:'dangerously',url:'https://x.test/',
    beforeParse(w){ w.localStorage.clear();
      w.fetch=()=>Promise.resolve({json:async()=>({status:'success',stock:[],expense:{},ledger:[],
        monthly:{},ledgerMonths:[],cash:{balance:0,hasAnchor:false},savings:[],bond:[],
        notices:[],hsa:info,retire:[],serverDate:'2026-08-29'})});
      w.alert=()=>{}; w.confirm=()=>true; w.scrollTo=()=>{};
      w.Element.prototype.scrollIntoView=function(){}; }});
  const w=dom.window; await wait();
  ok('前端不再有写死的 HSA 种子', w.eval("DATA.hsa.groups[0].holdings.length")===2);
  ok('Cash 用服务器算的值', w.eval("DATA.hsa.groups[0].holdings.find(x=>x.sym==='Cash').mv")===2000);
  ok('Investment 用服务器算的值',
     w.eval("DATA.hsa.groups[0].holdings.find(x=>x.sym==='Investment').mv")===10287);
  ok('HSA 总额 = 两者之和', w.eval("classTotal.hsa")===12287, '$'+w.eval("classTotal.hsa"));
  w.openDetail('hsa');
  ok('Cash 行标出自动补仓频率',
     /自动补仓 · 每两周/.test(w.document.getElementById('d-body').textContent));
  ok('Investment 行灰字是「基金定投」',
     /基金定投/.test(w.document.getElementById('d-body').textContent));
}

// ══════════════ 八、HSA 表的旧表头迁移 ══════════════
// 表可能是更早手工建的（表头 Account / Amount）。只判断 B1 是否为空不够 ——
// 旧表 B1 写着「Amount」，非空，会被误判成「表头没问题」而带错语义继续算。
console.log('\n— HSA 旧表头迁移 —');
{
  const HHEAD=['Name','Anchor Date','Anchor Amount','Rate','Floor','Sweep','Balance(自动算)','Updated'];
  function runOn(hsaRows){
    const H=mkSheet('HSA',hsaRows);
    const tabs={HSA:H, Ledger:mkSheet('Ledger',[['Date','category','amount','note','key']]),
      Stock:mkSheet('Stock',[['Symbol','Category','Share','Cost','Price']]),
      Anchor:mkSheet('Anchor',[['date','amount','note']]),
      Savings:mkSheet('Savings',[['ID','Name','Type','Balance','Rate','Last Post','Next update','Maturity','Status']]),
      Bond:mkSheet('Bond',[['ID','Name','Type','Start','Term','Rate','Principal','Balance','LastPost','Status']]),
      Ledger_monthly:mkSheet('Ledger_monthly',[])};
    const {doGet}=backend(tabs);
    return {r:J(doGet({parameter:{}})).hsa, H};
  }
  // 旧表：表头 Account / Amount，行序也反着（Investment 在前）
  let {r,H}=runOn([['Account','Amount','','','','','',''],
                   ['Investment','','','','','','',''],
                   ['Cash','','','','','','','']]);
  ok('旧表头被改写成新表头', H._rows[0][1]==='Anchor Date', String(H._rows[0][1]));
  ok('Investment 的 Rate 补上默认 10', Number(H._rows[1][3])===10, String(H._rows[1][3]));
  ok('Cash 的 Floor 补上默认 2000', Number(H._rows[2][4])===2000, String(H._rows[2][4]));
  ok('行序颠倒也能按名字认出来', r && r.ready===true);
  ok('Floor 不会因为留空而变成 0（否则 Cash 会被全额扫走）', r.floor===2000, String(r.floor));
  ok('Rate 不会因为留空而变成 0', Math.abs(r.rate-0.10)<1e-9, String(r.rate));

  // 已经是新表头的，不重复改写、也不覆盖你填的值
  const {H:H2}=runOn([HHEAD,
                      ['Cash','2026-08-01',1500,'',1000,'monthly','',''],
                      ['Investment','2026-08-01',5000,7,'','','','']]);
  ok('已填的 Floor 不被默认值覆盖', Number(H2._rows[1][4])===1000, String(H2._rows[1][4]));
  ok('已填的 Rate 不被默认值覆盖', Number(H2._rows[2][3])===7, String(H2._rows[2][3]));
  ok('已填的锚点不被动', String(H2._rows[1][1])==='2026-08-01' && Number(H2._rows[1][2])===1500);
}

// ══════════════ 九、HSA 页的固定支出要能存进去并立刻显示 ══════════════
// 固定支出区块在现金页和 HSA 页都有。这组函数原本只调 renderCashTabBody()，
// 在 HSA 页上是空操作 —— 数据存了、请求发了，但界面不动，看起来像「点保存没反应」。
console.log('\n— HSA 页的固定支出 —');
{
  const today=new Date(); const ym=`${today.getFullYear()}-${pad(today.getMonth()+1)}`;
  const calls=[];
  const dom=new JSDOM(html,{runScripts:'dangerously',url:'https://x.test/',
    beforeParse(w){ w.localStorage.clear();
      w.fetch=(u)=>{ const url=new URL(u), a=url.searchParams.get('action');
        if(a){ calls.push(Object.fromEntries(url.searchParams));
               return Promise.resolve({json:async()=>({status:'success',row:50})}); }
        return Promise.resolve({json:async()=>({status:'success',stock:[],expense:{},ledger:[],
          monthly:{},ledgerMonths:[ym],cash:{balance:0,hasAnchor:false},savings:[],bond:[],
          notices:[],hsa:{cash:2000,investment:5000,rate:0.1,floor:2000,ready:true},
          retire:[],serverDate:'2026-08-30'})}); };
      w.alert=(m)=>{ w.__alert=m; }; w.confirm=()=>true; w.scrollTo=()=>{};
      w.Element.prototype.scrollIntoView=function(){}; }});
  const w=dom.window, d=w.document;
  await wait();
  w.openDetail('hsa'); w.switchHsaTab('ex');
  ok('HSA 支出页有固定支出区块', !!d.getElementById('fixedFormWrap'));

  w.showFixedForm('HSA · 医疗');
  ok('点「添加固定支出」能打开表单', !!d.getElementById('fx-name'));
  d.getElementById('fx-name').value='每月理疗';
  d.getElementById('fx-amt').value='120';
  d.getElementById('fx-freq').value='monthly';
  d.getElementById('fx-start').value=fmt(daysAgo(0));   // 用今天，保证一定已到期
  w.saveFixed_('HSA · 医疗');
  await wait();

  const saved=JSON.parse(w.localStorage.getItem('fixedExpenses_v1')||'[]');
  ok('配置存下来了', saved.length===1 && saved[0].category==='HSA · 医疗', JSON.stringify(saved));
  ok('没有弹出报错', !w.__alert, String(w.__alert||''));
  ok('表单关掉了（不再杵在那儿）', !d.getElementById('fx-name'));
  ok('列表里立刻显示出来（这就是原来「点保存没反应」的地方）',
     /每月理疗/.test(d.getElementById('hsaTabBody').textContent));
  ok('发出了 HSA 分类的自动记账',
     calls.some(c=>c.action==='autoLedger' && c.category==='HSA · 医疗' && +c.amount===120),
     calls.filter(c=>c.action==='autoLedger').map(c=>c.category).join(','));
  ok('这笔不进流动现金', w.eval("catSign('HSA · 医疗')")===0);

  // 停用 / 删除也要立刻反映
  w.toggleFixed(saved[0].id);
  ok('停用后列表立刻标注', /已停用/.test(d.getElementById('hsaTabBody').textContent));
  w.deleteFixed(saved[0].id);
  // 注意：删的是「配置」，已经记过的那笔流水按设计保留（确认框里写明了），
  // 所以明细里仍能看到「每月理疗（固定）」——不能拿整页文字来断言。
  ok('删除后配置清空', w.eval("loadFixed().length")===0);
  ok('固定支出区块里不再列出它', !/已停用/.test(d.getElementById('hsaTabBody').textContent));
  ok('已记过的那笔流水仍在（不该被连带删掉）',
     /每月理疗/.test(d.getElementById('hsaTabBody').textContent));
}

// ══════════════ 十、HSA 输入路径的往返与阻塞 ══════════════
// 现金页踩过两个坑：① 已写过的 key 每次打开重发一遍（窗口没对齐）
// ② 界面渲染排在串行写请求之后。HSA 复用了同一套机制，这里逐条盯住。
console.log('\n— HSA 输入路径 —');
{
  const fmtd=d=>`${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}`;
  const ago=n=>{const d=new Date();d.setHours(0,0,0,0);d.setDate(d.getDate()-n);return d;};
  const CFG={employee:250,employer:100,freq:'biweekly',start:fmtd(ago(56))};
  // 历史供款都已记过
  const seeded=(()=>{ const r=[]; let n=100; let d=ago(56);
    while(d<=new Date()){ const ds=fmtd(d);
      r.push({row:++n,date:ds,category:'HSA · Income · 供款',amount:250,note:'本人供款',key:'hsaee:'+ds});
      r.push({row:++n,date:ds,category:'HSA · Income · 雇主补助',amount:100,note:'雇主补助',key:'hsaer:'+ds});
      d=new Date(d); d.setDate(d.getDate()+14); }
    return r; })();
  function app(ledger,cfg){
    const calls=[];
    const dom=new JSDOM(html,{runScripts:'dangerously',url:'https://x.test/',
      beforeParse(w){ w.localStorage.clear();
        if(cfg) w.localStorage.setItem('hsaConfig_v1', JSON.stringify(cfg));
        w.fetch=(u)=>{ const a=new URL(u).searchParams.get('action');
          if(a){ calls.push(Object.fromEntries(new URL(u).searchParams));
            return new Promise(r=>setTimeout(()=>r({json:async()=>({status:'success',row:9})}),10)); }
          return Promise.resolve({json:async()=>({status:'success',stock:[],expense:{},
            ledger:ledger,monthly:{},ledgerMonths:[],cash:{balance:0,hasAnchor:false},
            savings:[],bond:[],notices:[],
            hsa:{cash:2000,investment:5000,rate:0.1,floor:2000,ready:true},
            retire:[],serverDate:fmtd(new Date())})}); };
        w.alert=()=>{}; w.confirm=()=>true; w.scrollTo=()=>{};
        w.Element.prototype.scrollIntoView=function(){}; }});
    return {w:dom.window, calls};
  }

  // ① 打开 App：供款已全部记过 → 一次写请求都不该发
  let a=app(seeded, CFG); await wait(500);
  // 打开时可能有 saveConfig（把本地配置备份到 Sheet），这里只关心补记去重
  const autoOnOpen=a.calls.filter(c=>c.action==='autoLedger');
  ok('打开 App 不重发已记过的供款', autoOnOpen.length===0, `发了 ${autoOnOpen.length} 次`);

  // ② 补记窗口必须落在后端 70 天明细之内，否则去重查不到、每次都白跑
  const dues=JSON.parse(a.w.eval("JSON.stringify(hsaPayDatesDue(loadHsaCfg()))"));
  ok('供款补记窗口被 70 天明细覆盖', dues[0]>=fmtd(ago(70)), `最早应记日 ${dues[0]}`);

  // ③ 一次同步不该把 HSA 页重绘很多次
  a=app(seeded, CFG); await wait(500);
  a.w.openDetail('hsa');
  let n=0; const orig=a.w.renderHsa;
  a.w.renderHsa=function(){ n++; return orig.apply(this,arguments); };
  await a.w.eval("syncStockFromSheet()"); await wait(500);
  ok('一次同步的 HSA 重绘次数 ≤ 2', n<=2, `${n} 次`);

  // ④ 首次配置供款：界面必须先出来，十几次串行往返不能挡在前面
  const b=app([], null); await wait(500);
  const d2=b.w.document;
  b.w.openDetail('hsa'); b.w.switchHsaTab('in'); b.w.showHsaCfgForm();
  d2.getElementById('hsa-amt').value='250';
  d2.getElementById('hsa-freq').value='biweekly';
  d2.getElementById('hsa-start').value=fmtd(ago(56));
  b.w.saveHsaCfgForm();
  const txt=d2.getElementById('hsaTabBody').textContent;
  ok('点保存后立刻显示配置', /已设/.test(txt) && /250/.test(txt), txt.slice(0,40));
  ok('此刻后台写请求一次都还没回来',
     b.calls.filter(c=>c.action==='autoLedger').length===0);
  await wait(1500);
  ok('后台补记完成后明细才出现', /本人供款/.test(d2.getElementById('hsaTabBody').textContent));
  // 表单里只填了 Employee 一项，所以每个发薪日只补 1 笔（雇主那笔没配就不发）
  const autoB=b.calls.filter(c=>c.action==='autoLedger');
  const dueB=JSON.parse(b.w.eval("JSON.stringify(hsaPayDatesDue(loadHsaCfg()))"));
  ok('每个发薪日补一笔，不多不少', autoB.length===dueB.length,
     `${autoB.length} 笔 / ${dueB.length} 个发薪日`);
  ok('只补了本人供款，没凭空造雇主补助',
     autoB.every(c=>c.key.indexOf('hsaee:')===0), autoB.map(c=>c.key).join(','));
}

// ══════════════ 十一、自动补仓频率 ══════════════
// 原本每个收入日都做一次「结息 + 扫超额」，现在按 Sweep 列的频率来：
// 每两周 / 每月 / 每季度。频率变了，Investment 的复利节奏也跟着变。
console.log('\n— 自动补仓频率 —');
{
  const HH=['Name','Anchor Date','Anchor Amount','Rate','Floor','Sweep','Balance(自动算)','Updated'];
  // 8/01 起每两周发一次薪，共 5 次：8/01 8/15 8/29 9/12 9/26
  const led=[]; let dd=new Date('2026-08-01T00:00:00');
  for(let i=0;i<5;i++){ const t=`${dd.getFullYear()}-${pad(dd.getMonth()+1)}-${pad(dd.getDate())}`;
    led.push([t,'HSA · Income · 供款',500,'','']); dd=new Date(dd); dd.setDate(dd.getDate()+14); }
  function withFreq(freq){
    const L=mkSheet('Ledger',[['Date','category','amount','note','key'],...led]);
    const H=mkSheet('HSA',[HH,['Cash','2026-08-01',2000,'',2000,freq,'',''],
                              ['Investment','2026-08-01',10000,10,'','','','']]);
    const tabs={Ledger:L,HSA:H,
      Stock:mkSheet('Stock',[['Symbol','Category','Share','Cost','Price']]),
      Anchor:mkSheet('Anchor',[['date','amount','note']]),
      Savings:mkSheet('Savings',[['ID','Name','Type','Balance','Rate','Last Post','Next update','Maturity','Status']]),
      Bond:mkSheet('Bond',[['ID','Name','Type','Start','Term','Rate','Principal','Balance','LastPost','Status']]),
      Ledger_monthly:mkSheet('Ledger_monthly',[])};
    const {doGet}=backend(tabs);
    return J(doGet({parameter:{}})).hsa;
  }
  const bi=withFreq('biweekly'), mo=withFreq('monthly'), qu=withFreq('quarterly');
  ok('频率被读出来', bi.sweep==='biweekly' && mo.sweep==='monthly' && qu.sweep==='quarterly',
     `${bi.sweep}/${mo.sweep}/${qu.sweep}`);
  // 每两周：5 次全扫 → Cash 回到 2000，Investment 拿到 5×500
  ok('每两周：每次发薪都补仓，Cash 回到下限', Math.abs(bi.cash-2000)<0.01, '$'+bi.cash);
  ok('每两周：扫入总额 = 5×500', bi.investment>10000+2400, '$'+bi.investment);
  // 每季度：只有第一次（8/01）触发，之后 3 个月内不再补仓
  ok('每季度：只补仓一次，其余留在 Cash', qu.cash>2000, '$'+qu.cash);
  ok('每季度末次补仓日仍是 8/01', qu.lastSweep==='2026-08-01', qu.lastSweep);
  ok('频率越低，留在 Cash 的越多', qu.cash>mo.cash && mo.cash>=bi.cash,
     `每两周 ${bi.cash} / 每月 ${mo.cash} / 每季度 ${qu.cash}`);
  ok('频率越低，扫进投资的越少', bi.investment>mo.investment && mo.investment>qu.investment,
     `${bi.investment} / ${mo.investment} / ${qu.investment}`);
  // 本金 = Cash 起算 2000 + Investment 起算 10000 + 供款 5×500 = 14500。
  // 每季度只在 8/01 补过一次仓、且那次距锚点 0 天，所以一分收益都没滚到，
  // 合计恰好等于本金；每两周补了 5 次，早进去的钱多滚了一段，合计更高。
  // 两者之差就是「补仓越勤，收益越多」的量化体现 —— 不是不守恒。
  const PRIN=2000+10000+5*500;
  ok('每季度：只补一次且距锚点 0 天，合计恰等于本金',
     Math.abs((qu.cash+qu.investment)-PRIN)<0.01, `${(qu.cash+qu.investment).toFixed(2)} vs ${PRIN}`);
  ok('每两周：合计高于本金，多出来的就是多滚的收益',
     (bi.cash+bi.investment)>PRIN,
     `多 $${((bi.cash+bi.investment)-PRIN).toFixed(2)}`);
  ok('留空时默认每两周（保持原行为）', withFreq('').sweep==='biweekly');
}

// ══════════════ 十二、保存配置后界面必须立刻更新 ══════════════
// 现金页 Payroll 曾把 openDetail 排在 runPayroll 之后：点了保存，界面依旧
// 显示「未设置」，要等好几次 Apps Script 往返才变 —— 看起来像没保存成功。
// 跨月场景一并盯住：上个月设的配置，进入新月份后不能消失。
console.log('\n— 保存配置后立刻生效 / 跨月保留 —');
{
  function app(slow){
    const calls=[];
    const dom=new JSDOM(html,{runScripts:'dangerously',url:'https://x.test/',
      beforeParse(w){ w.localStorage.clear();
        w.fetch=(u)=>{ const a=new URL(u).searchParams.get('action');
          if(a){ calls.push(a);
            return new Promise(r=>setTimeout(()=>r({json:async()=>({status:'success',row:9})}),slow)); }
          return Promise.resolve({json:async()=>({status:'success',stock:[],expense:{},ledger:[],
            monthly:{},ledgerMonths:['2026-09','2026-08'],
            cash:{balance:5000,anchorDate:'2026-08-01',anchorAmount:5000,hasAnchor:true,bookStart:'2026-08-01'},
            savings:[],bond:[],notices:[],
            hsa:{cash:2000,investment:5000,rate:0.1,floor:2000,sweep:'biweekly',ready:true},
            retire:[],serverDate:'2026-09-01'})}); };
        w.alert=()=>{}; w.confirm=()=>true; w.scrollTo=()=>{};
        w.Element.prototype.scrollIntoView=function(){}; }});
    return {w:dom.window, calls};
  }
  const {w,calls}=app(300); await wait(600);
  const d=w.document;
  w.openDetail('cash'); w.showPayrollForm();
  d.getElementById('pay-amt').value='2600';
  d.getElementById('pay-freq').value='biweekly';
  d.getElementById('pay-start').value='2026-08-07';
  w.savePayroll();
  ok('Payroll 保存后立刻显示「已设」（不等写请求）',
     /已设/.test(d.getElementById('cashTabBody').textContent),
     d.getElementById('cashTabBody').textContent.slice(0,40));
  // 同步/保存过程中会顺带发 saveConfig、saveProperty 这类备份请求，
  // 它们不是「补记」。这里只数 autoLedger。
  ok('此刻补记的写请求还没回来',
     calls.filter(c=>(c.action||c)==='autoLedger').length===0,
     `已完成 ${calls.length} 次：${calls.map(c=>c.action||c).join(',')}`);
  ok('配置确实写进了 localStorage', !!w.localStorage.getItem('incomeConfig_v1'));
  await wait(1800);
  ok('后台补记跑完后仍显示「已设」',
     /已设/.test(d.getElementById('cashTabBody').textContent));

  // 跨月：8 月设的配置，9/01 打开后必须还在
  const CFG={payroll:{amount:2600,start:'2026-08-07',freq:'biweekly'}};
  const FIX=[{id:'f1',name:'网费',amount:70,category:'Bill & utilities',
              freq:'monthly',start:'2026-08-05',enabled:true}];
  const HCFG={employee:250,employer:100,freq:'biweekly',start:'2026-08-07'};
  const dom2=new JSDOM(html,{runScripts:'dangerously',url:'https://x.test/',
    beforeParse(w2){ w2.localStorage.clear();
      w2.localStorage.setItem('incomeConfig_v1',JSON.stringify(CFG));
      w2.localStorage.setItem('fixedExpenses_v1',JSON.stringify(FIX));
      w2.localStorage.setItem('hsaConfig_v1',JSON.stringify(HCFG));
      w2.fetch=(u)=>{ const a=new URL(u).searchParams.get('action');
        if(a) return Promise.resolve({json:async()=>({status:'success',row:9})});
        return Promise.resolve({json:async()=>({status:'success',stock:[],expense:{},ledger:[],
          monthly:{},ledgerMonths:['2026-09','2026-08'],
          cash:{balance:5000,anchorDate:'2026-08-01',anchorAmount:5000,hasAnchor:true,bookStart:'2026-08-01'},
          savings:[],bond:[],notices:[],
          hsa:{cash:2000,investment:5000,rate:0.1,floor:2000,sweep:'biweekly',ready:true},
          retire:[],serverDate:'2026-09-01'})}); };
      w2.alert=()=>{}; w2.confirm=()=>true; w2.scrollTo=()=>{};
      w2.Element.prototype.scrollIntoView=function(){}; }});
  const v=dom2.window, dd=v.document; await wait(600);
  ok('跨月后 Payroll 配置仍在', !!v.localStorage.getItem('incomeConfig_v1'));
  ok('跨月后固定支出配置仍在', v.eval("loadFixed().length")===1);
  ok('跨月后 HSA 供款配置仍在', v.eval("(loadHsaCfg().employee||0)")===250);
  v.openDetail('cash');
  ok('新月份里 Payroll 仍显示「已设」，不是「未设置」',
     /已设/.test(dd.getElementById('cashTabBody').textContent),
     dd.getElementById('cashTabBody').textContent.slice(0,40));
  v.switchCashTab('ex'); v.openExpenseDrill('Bill & utilities');
  ok('新月份里固定支出条目仍列出（金额归零是正常的）',
     /网费/.test(dd.getElementById('cashTabBody').textContent));
  v.openDetail('hsa'); v.switchHsaTab('in');
  ok('新月份里 HSA Contribution 仍显示「已设」',
     /已设/.test(dd.getElementById('hsaTabBody').textContent));
}

// ══════════════ 十三、配置写入失败必须看得见 ══════════════
// saveIncome / saveFixed / saveRecurring 原本裸调 localStorage.setItem。
// 一旦抛异常（隐私模式、存储满），配置静默丢失、异常还会中断后续代码 ——
// 表现就是「点了保存像没保存上」，最难查的那种。
console.log('\n— 配置写入失败要报出来 —');
{
  function app(breakWrite){
    const alerts=[]; const calls=[];
    const dom=new JSDOM(html,{runScripts:'dangerously',url:'https://x.test/',
      beforeParse(w){ w.localStorage.clear();
        if(breakWrite){
          // jsdom 的 Storage 上直接赋 setItem 覆盖不掉，得把整个 localStorage 换掉
          const box={};
          const fake={
            getItem:(k)=>(k in box?box[k]:null),
            setItem:(k,v)=>{ if(/incomeConfig_v1|fixedExpenses_v1|hsaConfig_v1/.test(k))
                               throw new Error('QuotaExceededError');
                             box[k]=String(v); },
            removeItem:(k)=>{ delete box[k]; },
            clear:()=>{ for(const k in box) delete box[k]; },
            key:(i)=>Object.keys(box)[i]??null,
            get length(){ return Object.keys(box).length; }
          };
          Object.defineProperty(w,'localStorage',{get:()=>fake, configurable:true});
        }
        w.fetch=(u)=>{ const a=new URL(u).searchParams.get('action');
          if(a){ calls.push(a); return Promise.resolve({json:async()=>({status:'success',row:9})}); }
          return Promise.resolve({json:async()=>({status:'success',stock:[],expense:{},ledger:[],
            monthly:{},ledgerMonths:[],cash:{balance:0,hasAnchor:false},savings:[],bond:[],
            notices:[],hsa:{cash:0,investment:0,rate:0.1,floor:2000,sweep:'biweekly',ready:true},
            retire:[],serverDate:'2026-09-01'})}); };
        w.alert=(m)=>alerts.push(String(m)); w.confirm=()=>true; w.scrollTo=()=>{};
        w.Element.prototype.scrollIntoView=function(){}; }});
    return {w:dom.window, alerts, calls};
  }

  // 正常情况：静默成功，不该弹窗
  let a=app(false); await wait(400);
  a.w.openDetail('cash'); a.w.showPayrollForm();
  a.w.document.getElementById('pay-amt').value='2600';
  a.w.document.getElementById('pay-start').value='2026-08-07';
  a.w.savePayroll(); await wait(300);
  ok('写入正常时不弹任何提示', a.alerts.length===0, a.alerts.join('|'));
  ok('配置存下来了', !!a.w.localStorage.getItem('incomeConfig_v1'));

  // 写入被拒：必须弹窗告知，且不能继续发补记请求
  let b=app(true); await wait(400);
  b.w.openDetail('cash'); b.w.showPayrollForm();
  b.w.document.getElementById('pay-amt').value='2600';
  b.w.document.getElementById('pay-start').value='2026-08-07';
  const nBefore=b.calls.length;
  b.w.savePayroll(); await wait(300);
  ok('写入失败时明确报错，不静默', b.alerts.some(m=>/保存失败/.test(m)), b.alerts.join('|'));
  ok('报错里点名是哪项设置', b.alerts.some(m=>/Payroll/.test(m)));
  // 本地写失败不再中断：设置已经备份进 Sheet，流程该照常走完
  // 注意这一节的假 fetch 记的是 action 字符串，不是对象
  ok('本地失败时仍把设置备份到 Sheet',
     b.calls.includes('saveConfig'), b.calls.join(','));
  ok('备份成功后流程照常走完（不再因本地失败而中断）',
     b.calls.length>nBefore, `${b.calls.length-nBefore} 次请求`);

  // 固定支出同理
  b.w.openDetail('cash'); b.w.switchCashTab('ex'); b.w.openExpenseDrill('Bill & utilities');
  b.w.showFixedForm('Bill & utilities');
  b.w.document.getElementById('fx-name').value='网费';
  b.w.document.getElementById('fx-amt').value='70';
  b.w.document.getElementById('fx-start').value='2026-08-05';
  b.alerts.length=0;
  b.w.saveFixed_('Bill & utilities'); await wait(300);
  ok('固定支出写入失败也报错', b.alerts.some(m=>/保存失败/.test(m)), b.alerts.join('|'));
  ok('报错里点名「固定支出」', b.alerts.some(m=>/固定支出/.test(m)));
}

// ══════════════ 十四、配置丢失后重输，不能重复记账 ══════════════
// 去重 key 是 fix:<配置id>:<日期>。配置 id 原来用 Date.now()，
// 一旦 localStorage 被清（例如重新 add to home screen），重输就换了 id，
// 于是 60 天补记窗口内已经记过的账会被当成新账再记一遍 —— 直接多扣一笔。
console.log('\n— 重输配置不产生重复记账 —');
{
  const pastStart=fmt(daysAgo(29));          // 一个月前，已经记过、且仍在 60 天窗口内
  function app(seedLedger){
    const calls=[];
    const dom=new JSDOM(html,{runScripts:'dangerously',url:'https://x.test/',
      beforeParse(w){ w.localStorage.clear();
        w.fetch=(u)=>{ const url=new URL(u), a=url.searchParams.get('action');
          if(a){ const p=Object.fromEntries(url.searchParams); calls.push(p);
            // 假后端照真后端的规矩：key 撞了就 skipped
            const dup=seedLedger.some(r=>r.key===p.key);
            return Promise.resolve({json:async()=>({status:dup?'skipped':'success',row:9})}); }
          return Promise.resolve({json:async()=>({status:'success',stock:[],expense:{},
            ledger:seedLedger, monthly:{}, ledgerMonths:[],
            cash:{balance:0,hasAnchor:false}, savings:[], bond:[], notices:[],
            hsa:{cash:0,investment:0,rate:0.1,floor:2000,sweep:'biweekly',ready:true},
            retire:[], serverDate:fmt(daysAgo(0))})}); };
        w.alert=()=>{}; w.confirm=()=>true; w.scrollTo=()=>{};
        w.Element.prototype.scrollIntoView=function(){}; }});
    return {w:dom.window, calls};
  }
  // 第一次输入：拿到 id
  let a=app([]); await wait(400);
  a.w.openDetail('cash'); a.w.switchCashTab('ex'); a.w.openExpenseDrill('Bill & utilities');
  a.w.showFixedForm('Bill & utilities');
  a.w.document.getElementById('fx-name').value='房贷月供';
  a.w.document.getElementById('fx-amt').value='2625.84';
  a.w.document.getElementById('fx-freq').value='monthly';
  a.w.document.getElementById('fx-start').value=pastStart;
  a.w.saveFixed_('Bill & utilities'); await wait(500);
  const id1=JSON.parse(a.w.localStorage.getItem('fixedExpenses_v1'))[0].id;
  const key1=a.calls.filter(c=>c.action==='autoLedger').map(c=>c.key)[0];
  ok('首次输入产生了记账', !!key1, key1);

  // 模拟配置丢失（全新 localStorage）后原样重输，且这笔已经在 Ledger 里
  const already=[{row:2,date:pastStart,category:'Bill & utilities',
                  amount:2625.84,note:'房贷月供（固定）',key:key1}];
  let b=app(already); await wait(400);
  b.w.openDetail('cash'); b.w.switchCashTab('ex'); b.w.openExpenseDrill('Bill & utilities');
  b.w.showFixedForm('Bill & utilities');
  b.w.document.getElementById('fx-name').value='房贷月供';
  b.w.document.getElementById('fx-amt').value='2625.84';
  b.w.document.getElementById('fx-freq').value='monthly';
  b.w.document.getElementById('fx-start').value=pastStart;
  b.w.saveFixed_('Bill & utilities'); await wait(500);
  const id2=JSON.parse(b.w.localStorage.getItem('fixedExpenses_v1'))[0].id;
  ok('重输后 id 不变（由内容派生，不是时间戳）', id1===id2, `${id1} vs ${id2}`);
  const resent=b.calls.filter(c=>c.action==='autoLedger' && c.key===key1);
  ok('重输不会用新 key 再记一遍',
     b.calls.filter(c=>c.action==='autoLedger').every(c=>c.key===key1),
     b.calls.filter(c=>c.action==='autoLedger').map(c=>c.key).join(','));
  ok('即便重发，服务器也按同一个 key 挡掉', resent.length<=1);
  ok('id 里不含时间戳', !/^f\d{13}$/.test(id2), id2);

  // 同分类同名建两笔，id 不能撞
  b.w.showFixedForm('Bill & utilities');
  b.w.document.getElementById('fx-name').value='房贷月供';
  b.w.document.getElementById('fx-amt').value='100';
  b.w.document.getElementById('fx-start').value=fmt(daysAgo(0));
  b.w.saveFixed_('Bill & utilities'); await wait(300);
  const list=JSON.parse(b.w.localStorage.getItem('fixedExpenses_v1'));
  ok('同名第二笔另给 id，不覆盖第一笔',
     list.length===2 && list[0].id!==list[1].id, list.map(x=>x.id).join(','));
}

// ══════════════ 十五、总览卡片上的年化收益率 ══════════════
// 按金额加权平均。只有储蓄和国债有明确票面利率；股票是市值波动、
// 现金无息、HSA 的 10% 是自设预估，都不该显示成「收益率」。
console.log('\n— 总览卡片的年化收益率 —');
{
  const SAV=[{id:'s1',name:'CD-A',type:'CD',balance:30000,rate:0.05,lastPost:'2026-08-01',
              nextUpdate:'',maturity:'2027-06-30',status:'active'},
             {id:'s2',name:'OS-B',type:'OS',balance:10000,rate:0.03,lastPost:'2026-08-01',
              nextUpdate:'',maturity:'',status:'active'},
             {id:'s3',name:'旧CD',type:'CD',balance:99999,rate:0.99,lastPost:'2026-08-01',
              nextUpdate:'',maturity:'',status:'closed'}];   // closed 不该参与
  const BOND=[{id:'b1',name:'T-Note',type:'T-Note',start:'2026-02-15',term:2,rate:0.04,
               principal:10000,balance:10000,lastPost:'2026-08-15',maturity:'2028-02-15',status:'active'}];
  const dom=new JSDOM(html,{runScripts:'dangerously',url:'https://x.test/',
    beforeParse(w){ w.localStorage.clear();
      w.fetch=()=>Promise.resolve({json:async()=>({status:'success',stock:[],expense:{},ledger:[],
        monthly:{},ledgerMonths:[],cash:{balance:5000,hasAnchor:true,anchorDate:'2026-08-01',anchorAmount:5000},
        savings:SAV,bond:BOND,notices:[],
        hsa:{cash:2000,investment:5000,rate:0.1,floor:2000,sweep:'biweekly',ready:true},
        retire:[],serverDate:'2026-09-01'})});
      w.alert=()=>{}; w.confirm=()=>true; w.scrollTo=()=>{};
      w.Element.prototype.scrollIntoView=function(){}; }});
  const w=dom.window, d=w.document; await wait();

  // 30000×5% + 10000×3% = 1800 → 1800/40000 = 4.5%
  ok('储蓄按金额加权平均', Math.abs(w.eval("classYield('cd')")-4.5)<1e-9, w.eval("classYield('cd')"));
  ok('closed 的账户不参与（否则会被 99999×99% 拉爆）',
     w.eval("classYield('cd')")<10, w.eval("classYield('cd')"));
  ok('国债取其票面利率', Math.abs(w.eval("classYield('bond')")-4)<1e-9, w.eval("classYield('bond')"));
  ok('股票不给收益率', w.eval("classYield('stock')")===null);
  ok('现金不给收益率', w.eval("classYield('cash')")===null);
  ok('HSA 不给收益率（10% 是自设预估，不是票面）', w.eval("classYield('hsa')")===null);

  const cards={};
  [].forEach.call(d.querySelectorAll('.ccard'),function(c){
    cards[(c.querySelector('.nm')||{}).textContent]=c; });
  const savMeta=cards['储蓄'].querySelector('.meta').innerHTML;
  ok('储蓄卡片显示 +4.5%', /\+4\.5%/.test(savMeta), savMeta);
  ok('用绿色（var(--pos)）', /var\(--pos\)/.test(savMeta), savMeta);
  ok('保留原有灰字', /Marcus/.test(savMeta), savMeta);
  const bondMeta=cards['国债'].querySelector('.meta').innerHTML;
  ok('国债卡片显示 +4.0%', /\+4\.0%/.test(bondMeta), bondMeta);
  ok('格式是一位小数', !/\+4%|\+4\.00%/.test(bondMeta), bondMeta);
  ok('股票卡片不显示收益率', !/\+\d/.test(cards['股票'].innerHTML));
  ok('现金卡片不显示收益率', !/\+\d/.test(cards['现金'].innerHTML));

  // 没有数据时不该显示 +NaN%
  const dom2=new JSDOM(html,{runScripts:'dangerously',url:'https://x.test/',
    beforeParse(w2){ w2.localStorage.clear();
      w2.fetch=()=>Promise.resolve({json:async()=>({status:'success',stock:[],expense:{},ledger:[],
        monthly:{},ledgerMonths:[],cash:{balance:0,hasAnchor:false},savings:[],bond:[],notices:[],
        hsa:null,retire:[],serverDate:'2026-09-01'})});
      w2.alert=()=>{}; w2.confirm=()=>true; w2.scrollTo=()=>{};
      w2.Element.prototype.scrollIntoView=function(){}; }});
  await wait();
  ok('没有账户时不显示 +NaN%', !/NaN/.test(visibleText(dom2.window.document)));
  ok('没有账户时收益率为 null', dom2.window.eval("classYield('cd')")===null);
}

// 新增账户后收益率要跟着变（不是一次性算完就定死）
console.log('\n— 新增账户后收益率自动更新 —');
{
  let SAV=[{id:'s1',name:'CD-A',type:'CD',balance:30000,rate:0.04,lastPost:'2026-08-01',
            nextUpdate:'',maturity:'2027-06-30',status:'active'},
           {id:'s2',name:'OS-B',type:'OS',balance:10000,rate:0.03,lastPost:'2026-08-01',
            nextUpdate:'',maturity:'',status:'active'}];
  const dom=new JSDOM(html,{runScripts:'dangerously',url:'https://x.test/',
    beforeParse(w){ w.localStorage.clear();
      w.fetch=(u)=>{ const url=new URL(u), a=url.searchParams.get('action');
        const p=Object.fromEntries(url.searchParams);
        if(a==='addSaving'){ SAV.push({id:'sN',name:p.name,type:p.type,balance:+p.balance,
          rate:+p.rate/100,lastPost:'2026-09-01',nextUpdate:'',maturity:p.maturity||'',status:'active'});
          return Promise.resolve({json:async()=>({status:'success',savings:JSON.parse(JSON.stringify(SAV))})}); }
        if(a) return Promise.resolve({json:async()=>({status:'success'})});
        return Promise.resolve({json:async()=>({status:'success',stock:[],expense:{},ledger:[],
          monthly:{},ledgerMonths:[],cash:{balance:0,hasAnchor:false},
          savings:JSON.parse(JSON.stringify(SAV)),bond:[],notices:[],hsa:null,
          retire:[],serverDate:'2026-09-01'})}); };
        w.alert=()=>{}; w.confirm=()=>true; w.scrollTo=()=>{};
        w.Element.prototype.scrollIntoView=function(){}; }});
  const w=dom.window, d=w.document; await wait();
  const meta=()=>{ let r=''; [].forEach.call(d.querySelectorAll('.ccard'),c=>{
    if((c.querySelector('.nm')||{}).textContent==='储蓄') r=(c.querySelector('.meta')||{}).textContent; }); return r; };
  ok('初始 = (30000×4 + 10000×3)/40000 = 3.8%', /\+3\.8%/.test(meta()), meta());
  w.openDetail('cd'); w.showCdForm();
  d.getElementById('f-name').value='CD-NEW'; d.getElementById('f-mv').value='20000';
  d.getElementById('f-rate').value='5'; d.getElementById('f-type').value='CD';
  w.saveCd(); await wait(600);
  ok('新增账户后自动重算为 4.2%', /\+4\.2%/.test(meta()), meta());
  ok('总额也跟着变', w.eval("classTotal.cd")===60000, '$'+w.eval("classTotal.cd"));
}

// ══════════════ 十六、设置备份到 Sheet ══════════════
// 固定支出 / Payroll / 定期转账 / HSA 供款 / 401k / 房贷 原本只存 localStorage。
// iOS 重新 add to home screen 会开一个全新的存储容器，设置就全没了且无备份。
// 现在以 Sheet 的 Config 表为准：保存时同时写本地和 Sheet，同步时取回。
console.log('\n— 设置备份到 Sheet —');
{
  function app(serverCfg, seedLocal, breakLocal){
    const calls=[];
    const dom=new JSDOM(html,{runScripts:'dangerously',url:'https://x.test/',
      beforeParse(w){ w.localStorage.clear();
        if(seedLocal) Object.keys(seedLocal).forEach(k=>w.localStorage.setItem(k,seedLocal[k]));
        if(breakLocal){
          const box=Object.assign({}, seedLocal||{});
          Object.defineProperty(w,'localStorage',{configurable:true,get:()=>({
            getItem:(k)=>(k in box?box[k]:null),
            setItem:()=>{ throw new Error('QuotaExceededError'); },
            removeItem:(k)=>{ delete box[k]; }, clear:()=>{},
            key:(i)=>Object.keys(box)[i]??null, get length(){return Object.keys(box).length;} })});
        }
        w.fetch=(u)=>{ const url=new URL(u), a=url.searchParams.get('action');
          if(a){ calls.push(Object.fromEntries(url.searchParams));
                 return Promise.resolve({json:async()=>({status:'success',row:9})}); }
          return Promise.resolve({json:async()=>({status:'success',stock:[],expense:{},ledger:[],
            monthly:{},ledgerMonths:[],cash:{balance:0,hasAnchor:false},savings:[],bond:[],
            notices:[],hsa:null,retire:[],config:serverCfg||{},serverDate:'2026-09-01'})}); };
        w.alert=()=>{}; w.confirm=()=>true; w.scrollTo=()=>{};
        w.Element.prototype.scrollIntoView=function(){}; }});
    return {w:dom.window, calls};
  }

  // ① 全新设备（localStorage 空）+ Sheet 上有配置 → 取回来
  const PAY={payroll:{amount:3634,start:'2026-08-07',freq:'biweekly'}};
  const FIX=[{id:'fabc',name:'房贷月供',amount:2625.84,category:'Bill & utilities',
              freq:'monthly',start:'2026-08-03',enabled:true}];
  let a=app({incomeConfig_v1:PAY, fixedExpenses_v1:FIX}); await wait(600);
  ok('空设备能从 Sheet 取回 Payroll', a.w.eval("(loadIncome().payroll||{}).amount")===3634);
  ok('空设备能从 Sheet 取回固定支出', a.w.eval("loadFixed().length")===1);
  a.w.openDetail('cash');
  ok('界面直接显示「已设」，不需要重输',
     /已设/.test(a.w.document.getElementById('cashTabBody').textContent),
     a.w.document.getElementById('cashTabBody').textContent.slice(0,36));

  // ② 保存时要同时写 Sheet
  let b=app({}); await wait(600);
  b.w.openDetail('cash'); b.w.showPayrollForm();
  b.w.document.getElementById('pay-amt').value='3634';
  b.w.document.getElementById('pay-start').value='2026-08-07';
  b.w.savePayroll(); await wait(500);
  const wrote=b.calls.filter(c=>c.action==='saveConfig' && c.key==='incomeConfig_v1');
  ok('保存 Payroll 会写一份到 Sheet', wrote.length>0, `${wrote.length} 次`);
  ok('写进 Sheet 的内容正确',
     wrote.length>0 && JSON.parse(wrote[wrote.length-1].value).payroll.amount===3634,
     wrote.length?wrote[wrote.length-1].value:'');

  // ③ 迁移：Sheet 上没有、本地有 → 自动上传
  let c=app({}, {fixedExpenses_v1:JSON.stringify(FIX)}); await wait(600);
  const up=c.calls.filter(x=>x.action==='saveConfig' && x.key==='fixedExpenses_v1');
  ok('本地已有但 Sheet 没有的，自动上传备份', up.length>0, `${up.length} 次`);

  // ④ Sheet 上有就以 Sheet 为准，不再重复上传
  let e=app({fixedExpenses_v1:FIX}, {fixedExpenses_v1:JSON.stringify(FIX)}); await wait(600);
  ok('Sheet 已有则不重复上传',
     e.calls.filter(x=>x.action==='saveConfig' && x.key==='fixedExpenses_v1').length===0);

  // ⑤ 本地写不进去（隐私模式）也要备份到 Sheet
  let f=app({}, {}, true); await wait(600);
  f.w.openDetail('cash'); f.w.showPayrollForm();
  f.w.document.getElementById('pay-amt').value='3634';
  f.w.document.getElementById('pay-start').value='2026-08-07';
  f.w.savePayroll(); await wait(500);
  ok('本地写入失败时仍备份到 Sheet',
     f.calls.some(x=>x.action==='saveConfig' && x.key==='incomeConfig_v1'));
}

// ══════════════ 十七、房产搬进 Sheet ══════════════
// 房产参数以前只在 localStorage。现在走 Sheet 的 Property 表（可读键值表，
// 不是 JSON blob），你能直接在表里看和改，以后也方便把月供接进固定支出。
// 存的是锚点（updated 那天的余额），当前余额按整月摊销推算。
console.log('\n— 房产搬进 Sheet —');
{
  const PHEAD=['Field','Value'];
  function runP(rows){
    const P=rows?mkSheet('Property',[PHEAD,...rows]):null;
    const tabs={Ledger:mkSheet('Ledger',[['Date','category','amount','note','key']]),
      Stock:mkSheet('Stock',[['Symbol','Category','Share','Cost','Price']]),
      Anchor:mkSheet('Anchor',[['date','amount','note']]),
      Savings:mkSheet('Savings',[['ID','Name','Type','Balance','Rate','Last Post','Next update','Maturity','Status']]),
      Bond:mkSheet('Bond',[['ID','Name','Type','Start','Term','Rate','Principal','Balance','LastPost','Status']]),
      Ledger_monthly:mkSheet('Ledger_monthly',[])};
    if(P) tabs.Property=P;
    const {doGet}=backend(tabs);
    return {get:()=>J(doGet({parameter:{}})), doGet, tabs};
  }
  // 表不存在时不该崩，property 返回 null
  let a=runP(null);
  ok('没有 Property 表时返回 null，不报错', a.get().property===null);

  // 有表时读出来
  let b=runP([['value',415000],['loan',263000],['origLoan',315000],['rate',5.124],
              ['termYears',30],['payment',2025],['extra',600],['updated','2026-08-01']]);
  const pr=b.get().property;
  ok('读出房屋估值', pr.value===415000, String(pr.value));
  ok('读出锚点余额', pr.loan===263000, String(pr.loan));
  ok('读出利率（小数不被取整）', Math.abs(pr.rate-5.124)<1e-9, String(pr.rate));
  ok('读出锚点日期', pr.updated==='2026-08-01', String(pr.updated));
  ok('读出每月额外还本', pr.extra===600, String(pr.extra));

  // 写回：只改 extra，其他字段不动
  J(b.doGet({parameter:{action:'saveProperty', extra:1500}}));
  const pr2=b.get().property;
  ok('改额外还本写回 Sheet', pr2.extra===1500, String(pr2.extra));
  ok('没传的字段保持不变', pr2.value===415000 && pr2.loan===263000 && pr2.updated==='2026-08-01');

  // 表缺字段时用默认值补齐、不崩
  let c=runP([['value',400000],['loan',200000]]);
  const pr3=c.get().property;
  ok('只填了部分字段也能读出来', pr3.value===400000 && pr3.loan===200000);

  // 前端：服务器值覆盖本地，并且保存时写回 Sheet
  const SRV={value:500000,loan:250000,origLoan:315000,rate:4.5,
             termYears:30,payment:2100,extra:800,updated:'2026-09-01'};
  const calls=[];
  const dom=new JSDOM(html,{runScripts:'dangerously',url:'https://x.test/',
    beforeParse(w){ w.localStorage.clear();
      w.fetch=(u)=>{ const url=new URL(u), act=url.searchParams.get('action');
        if(act){ calls.push(Object.fromEntries(url.searchParams));
                 return Promise.resolve({json:async()=>({status:'success'})}); }
        return Promise.resolve({json:async()=>({status:'success',stock:[],expense:{},ledger:[],
          monthly:{},ledgerMonths:[],cash:{balance:0,hasAnchor:false},savings:[],bond:[],
          notices:[],hsa:null,retire:[],config:{},property:SRV,serverDate:'2026-09-01'})}); };
      w.alert=()=>{}; w.confirm=()=>true; w.scrollTo=()=>{};
      w.Element.prototype.scrollIntoView=function(){}; }});
  const w=dom.window; await wait();
  ok('前端采用 Sheet 上的估值', w.eval("PROPERTY.value")===500000, w.eval("PROPERTY.value"));
  ok('前端采用 Sheet 上的额外还本', w.eval("PROPERTY.extra")===800, w.eval("PROPERTY.extra"));
  ok('房产不再走 Config（避免两个真相源）',
     w.eval("CFG_KEYS.indexOf('property_v1')")===-1);
  w.eval("PROPERTY.extra=1200; persistProp();"); await wait(300);
  const sent=calls.filter(c=>c.action==='saveProperty');
  ok('改完写回 Sheet', sent.length>0, `${sent.length} 次`);
  ok('写回的内容带上新值', sent.length>0 && +sent[sent.length-1].extra===1200,
     sent.length?sent[sent.length-1].extra:'');

  // 摊销：锚点当月不动，跨月才走一步
  ok('锚点当月余额 = 锚点值（房贷按月还，月中不变）',
     w.eval("(function(){PROPERTY.updated='2026-09-01';PROPERTY.loan=250000;return amortizeNow();})()")===250000);
  const next=w.eval("(function(){PROPERTY.updated='2026-08-01';PROPERTY.loan=250000;return amortizeNow();})()");
  ok('跨一个月后余额下降', next<250000, '$'+next);
  ok('额外还本调高后余额降得更多',
     w.eval("(function(){PROPERTY.extra=2000;return amortizeNow();})()")<next);
}

// ══════════════ 十八、表要在读路径里建出来 + 房贷联动 ══════════════
// 上一版 Property / Config 只在写入时建表，纯打开小程序（只读）永远建不出来，
// 用户在 Sheet 里什么都看不到。而且建表时填 0 会更糟 —— readProperty 只跳过
// 空值，0 是「有效值」，会把小程序里真实的房产数字覆盖成零。
console.log('\n— 建表与房贷联动 —');
{
  function fresh(){
    const tabs={Ledger:mkSheet('Ledger',[['Date','category','amount','note','key']]),
      Stock:mkSheet('Stock',[['Symbol','Category','Share','Cost','Price']]),
      Anchor:mkSheet('Anchor',[['date','amount','note']]),
      Savings:mkSheet('Savings',[['ID','Name','Type','Balance','Rate','Last Post','Next update','Maturity','Status']]),
      Bond:mkSheet('Bond',[['ID','Name','Type','Start','Term','Rate','Principal','Balance','LastPost','Status']]),
      Ledger_monthly:mkSheet('Ledger_monthly',[])};
    const {doGet}=backend(tabs);
    return {tabs, doGet};
  }
  const f=fresh();
  const r=J(f.doGet({parameter:{}}));
  ok('只读一次就把 Property 表建出来了', !!f.tabs.Property);
  ok('也把 Config 表建出来了', !!f.tabs.Config);
  const pv=f.tabs.Property._rows.slice(1).map(x=>x[1]);
  ok('建表时字段一律留空，不填 0', pv.every(x=>x===''), JSON.stringify(pv));
  ok('留空时 readProperty 返回 null（不会拿 0 去覆盖真实数据）', r.property===null);

  // 前端：Sheet 没数据时把本地的传上去
  const calls=[];
  function app(serverProp){
    calls.length=0;
    const dom=new JSDOM(html,{runScripts:'dangerously',url:'https://x.test/',
      beforeParse(w){ w.localStorage.clear();
        w.fetch=(u)=>{ const url=new URL(u), a=url.searchParams.get('action');
          if(a){ calls.push(Object.fromEntries(url.searchParams));
                 return Promise.resolve({json:async()=>({status:'success'})}); }
          return Promise.resolve({json:async()=>({status:'success',stock:[],expense:{},ledger:[],
            monthly:{},ledgerMonths:[],cash:{balance:0,hasAnchor:false},savings:[],bond:[],
            notices:[],hsa:null,retire:[],config:{},property:serverProp,
            serverDate:'2026-09-01'})}); };
        w.alert=()=>{}; w.confirm=()=>true; w.scrollTo=()=>{};
        w.Element.prototype.scrollIntoView=function(){}; }});
    return dom.window;
  }
  // 房产已改为不留种子：全新设备 + Sheet 也没数据时就该是 0，不该编一个出来
  let w=app(null); await wait(500);
  ok('全新设备且 Sheet 没数据时显示 0，不编造数字', w.eval("PROPERTY.value")===0);
  // 本地有数据（比如老用户的缓存）才需要上传迁移
  w=app(null);
  w.eval("PROPERTY.value=415000; PROPERTY.loan=263000; PROPERTY.updated='2026-08-01';");
  await w.eval("syncStockFromSheet()"); await wait(500);
  ok('本地有数据而 Sheet 没有时，自动上传迁移',
     calls.some(c=>c.action==='saveProperty'), calls.map(c=>c.action).join(','));
  ok('迁移后本地数据没被清零', w.eval("PROPERTY.value")===415000, '$'+w.eval("PROPERTY.value"));

  // 服务器回一份全 0 视为「未填写」，不能拿它覆盖本地已有的真实数据
  w=app(null);
  w.eval("PROPERTY.value=415000;");
  await w.eval("syncStockFromSheet()"); await wait(400);
  ok('服务器没有房产数据时不覆盖本地', w.eval("PROPERTY.value")===415000,
     '$'+w.eval("PROPERTY.value"));

  // ---- 方案 B：固定支出金额跟随房贷 ----
  w=app({value:415000,loan:263000,origLoan:315000,rate:5.124,
         termYears:30,payment:2025,extra:600,updated:'2026-08-01'}); await wait(500);
  ok('房产参数已从 Sheet 取回', w.eval("PROPERTY.payment")===2025 && w.eval("PROPERTY.extra")===600);
  ok('跟随房贷的金额 = 月供 + 额外还本',
     w.eval("fixedAmount({linkProp:true})")===2625, w.eval("fixedAmount({linkProp:true})"));
  ok('没勾跟随的仍用自己填的金额',
     w.eval("fixedAmount({amount:70})")===70);

  w.eval(`saveFixed([{id:'fm',name:'房贷月供',category:'Bill & utilities',
    freq:'monthly',start:'2026-08-03',enabled:true,linkProp:true,amount:0}])`);
  w.openDetail('cash'); w.switchCashTab('ex'); w.openExpenseDrill('Bill & utilities');
  const txt=w.document.getElementById('cashTabBody').textContent;
  ok('列表里显示的是联动后的金额 $2,625', /2,625/.test(txt), txt.slice(0,80));
  ok('并标注「跟随房贷」', /跟随房贷/.test(txt));

  // 改额外还本后，金额自动跟着变
  w.eval("PROPERTY.extra=1500;");
  ok('把额外还本改成 1500 后自动变 $3,525',
     w.eval("fixedAmount({linkProp:true})")===3525, w.eval("fixedAmount({linkProp:true})"));
  w.openExpenseDrill('Bill & utilities');
  ok('列表也跟着变', /3,525/.test(w.document.getElementById('cashTabBody').textContent));
}

// ══════════════ 十九、Sheet 上要看得到「当前欠多少」 ══════════════
// Property 表里的 loan 是锚点，不会自己变。只盯着它会误以为余额没更新。
// 后端另写一行 balanceNow(自动算)，按整月摊销推算。
console.log('\n— 房贷当前余额写进 Sheet —');
{
  const PHEAD=['Field','Value'];
  function runP(rows){
    const P=mkSheet('Property',[PHEAD,...rows]);
    const tabs={Property:P, Ledger:mkSheet('Ledger',[['Date','category','amount','note','key']]),
      Stock:mkSheet('Stock',[['Symbol','Category','Share','Cost','Price']]),
      Anchor:mkSheet('Anchor',[['date','amount','note']]),
      Savings:mkSheet('Savings',[['ID','Name','Type','Balance','Rate','Last Post','Next update','Maturity','Status']]),
      Bond:mkSheet('Bond',[['ID','Name','Type','Start','Term','Rate','Principal','Balance','LastPost','Status']]),
      Ledger_monthly:mkSheet('Ledger_monthly',[])};
    const {doGet}=backend(tabs);
    return {get:()=>J(doGet({parameter:{}})).property, P};
  }
  const T=new Date(); const ym=(d)=>`${d.getFullYear()}-${pad(d.getMonth()+1)}-01`;
  const thisMonth=ym(T);
  const sixAgo=new Date(T.getFullYear(), T.getMonth()-6, 1);

  // 锚点就是本月 → 还没走过一整月，当前余额 = 锚点
  let a=runP([['value',415000],['loan',263000],['origLoan',315000],['rate',5.124],
              ['termYears',30],['payment',2025],['extra',600],['updated',thisMonth]]);
  ok('锚点当月：当前余额 = 锚点值', a.get().balanceNow===263000, String(a.get().balanceNow));

  // 锚点在 6 个月前 → 应该少了一截
  let b=runP([['value',415000],['loan',263000],['origLoan',315000],['rate',5.124],
              ['termYears',30],['payment',2025],['extra',600],['updated',ym(sixAgo)]]);
  const six=b.get().balanceNow;
  ok('锚点在 6 个月前：余额明显下降', six<263000 && six>240000, '$'+six);
  ok('写进了 Sheet 的 balanceNow 行',
     b.P._rows.some(r=>String(r[0])==='balanceNow(自动算)' && Number(r[1])===six),
     JSON.stringify(b.P._rows.find(r=>String(r[0])==='balanceNow(自动算)')));
  ok('loan 那一格仍是锚点，没被改写',
     Number(b.P._rows.find(r=>String(r[0])==='loan')[1])===263000);

  // 额外还本调高 → 当前余额更低
  let c=runP([['value',415000],['loan',263000],['origLoan',315000],['rate',5.124],
              ['termYears',30],['payment',2025],['extra',2000],['updated',ym(sixAgo)]]);
  ok('额外还本调高后当前余额更低', c.get().balanceNow<six, '$'+c.get().balanceNow);

  // 内容没变就不重复写
  const before=JSON.stringify(b.P._rows);
  b.get();
  ok('再读一次不重复写', JSON.stringify(b.P._rows)===before);

  // 月供不够付利息时不该把余额算成负数或死循环
  let d=runP([['value',415000],['loan',263000],['rate',20],['payment',10],
              ['extra',0],['updated',ym(sixAgo)]]);
  ok('月供不足以付息时安全退出，不出负数',
     d.get().balanceNow>=0 && d.get().balanceNow<=263000, String(d.get().balanceNow));

  // 前端锚点日期不能再用 UTC
  ok('前端不再用 toISOString 取锚点日期',
     !/PROPERTY\.updated\s*=\s*new Date\(\)\.toISOString/.test(html));
}

// ══════════════ 二十、房产不再留写死的种子 ══════════════
// 房产已经由 Sheet 驱动，本地就不该再留一份写死的数字 —— 断网或 Sheet
// 还没填时，那个数看起来完全合理，反而最难发现是假的。
console.log('\n— 房产不留种子 —');
{
  const SHEET_PROP={value:500000,loan:250000,origLoan:315000,rate:4.5,
                    termYears:30,payment:2100,extra:800,updated:'2026-03-01'};
  function run(serverProp, failSync){
    const calls=[];
    const dom=new JSDOM(html,{runScripts:'dangerously',url:'https://x.test/',
      beforeParse(w){ w.localStorage.clear();
        w.fetch=(u)=>{ const a=new URL(u).searchParams.get('action');
          if(a){ calls.push(Object.fromEntries(new URL(u).searchParams));
                 return Promise.resolve({json:async()=>({status:'success'})}); }
          return failSync ? Promise.reject(new Error('断网'))
            : Promise.resolve({json:async()=>({status:'success',stock:[],expense:{},ledger:[],
              monthly:{},ledgerMonths:[],cash:{balance:0,hasAnchor:false},savings:[],bond:[],
              notices:[],hsa:null,retire:[],config:{},property:serverProp,
              serverDate:'2026-09-02'})}); };
        w.alert=()=>{}; w.confirm=()=>true; w.scrollTo=()=>{};
        w.Element.prototype.scrollIntoView=function(){}; }});
    return {w:dom.window, calls};
  }
  // 清缓存 + 能连上 → 用 Sheet 的真实数据
  let a=run(SHEET_PROP,false); await wait(600);
  ok('同步后采用 Sheet 的锚点', a.w.eval("PROPERTY.loan")===250000, a.w.eval("PROPERTY.loan"));
  ok('当前余额按 Sheet 的锚点推算', a.w.eval("amortizeNow()")<250000 && a.w.eval("amortizeNow()")>200000,
     '$'+a.w.eval("amortizeNow()"));

  // 清缓存 + 断网 → 显示 0，而不是一个看着合理的假数
  let b=run(SHEET_PROP,true); await wait(600);
  ok('断网时房贷余额为 0，不冒充真实数字', b.w.eval("amortizeNow()")===0, '$'+b.w.eval("amortizeNow()"));
  ok('断网时估值也为 0', b.w.eval("PROPERTY.value")===0);
  ok('代码里不再有写死的房贷数字', !/loan:263000|value:415000/.test(html));

  // 锚点日期为空时不能算出 NaN
  ok('锚点为空时不产生 NaN', !isNaN(b.w.eval("amortizeNow()")));
  ok('页面上不出现 NaN', !/NaN/.test(visibleText(b.w.document)));

  // Sheet 没数据但本地有 → 仍会上传做迁移
  let c=run(null,false);
  c.w.eval("PROPERTY.value=415000; PROPERTY.loan=263000; PROPERTY.updated='2026-08-01';");
  await c.w.eval("syncStockFromSheet()"); await wait(500);
  ok('本地有数据而 Sheet 没有时，仍会上传迁移',
     c.calls.some(x=>x.action==='saveProperty'), c.calls.map(x=>x.action).join(','));
}

// ══════════════ 二十一、股票趋势曲线 ══════════════
// 两条线：月末市值 vs 累计投入。只看市值分不出「赚了」还是「一直在投钱」，
// 两线之间的距离才是累计收益。历史手填，当月自动维护。
console.log('\n— 股票趋势曲线 —');
{
  const SHHEAD=['年月','月初account value','本月净入金','月末account value','本月投资收益','月收益率'];
  const T=new Date(); const curYM=`${T.getFullYear()}-${pad(T.getMonth()+1)}`;
  function run(histRows, ledgerRows, stockRows){
    const SH=mkSheet('StockHistory',[SHHEAD,...histRows]);
    const tabs={StockHistory:SH,
      Ledger:mkSheet('Ledger',[['Date','category','amount','note','key'],...(ledgerRows||[])]),
      Stock:mkSheet('Stock',[['Symbol','Category','Share','Cost','Price'],...(stockRows||[])]),
      Anchor:mkSheet('Anchor',[['date','amount','note']]),
      Savings:mkSheet('Savings',[['ID','Name','Type','Balance','Rate','Last Post','Next update','Maturity','Status']]),
      Bond:mkSheet('Bond',[['ID','Name','Type','Start','Term','Rate','Principal','Balance','LastPost','Status']]),
      Ledger_monthly:mkSheet('Ledger_monthly',[])};
    const {doGet}=backend(tabs);
    return {out:()=>J(doGet({parameter:{}})).stockHistory, SH};
  }
  // 你截图里的头几行
  const hist=[['2022',0,7500,7572,72,0.0096],
              ['2023-01',7572,0,8628,1056,0.1395],
              ['2023-02',8628,0,8192,-436,-0.0505],
              ['2023-04',8192,4000,12412,-176,-0.014]];
  let a=run(hist,[],[['VOO','IndexETF',40,400,647.4]]);
  const o=a.out();
  ok('历史行都读出来了', o.length>=4, `${o.length} 条`);
  ok('「2022」这种非年月的汇总行也保留', o[0].ym==='2022', o[0].ym);
  ok('累计投入是逐月累加', o[0].invested===7500 && o[1].invested===7500 && o[3].invested===11500,
     o.slice(0,4).map(x=>x.invested).join(','));
  ok('市值取月末 account value', o[3].value===12412, String(o[3].value));

  // 当月那一行自动追加：市值 = 股票账户总值，净入金 = 当月 Transfer · 券商
  const cur=o[o.length-1];
  ok('自动补上当月一行', cur.ym===curYM, cur.ym);
  ok('当月市值 = 40股 × $647.4', Math.abs(cur.value-25896)<0.01, String(cur.value));

  let b=run(hist,[[`${curYM}-03`,'Transfer · 券商',1500,'',''],
                  [`${curYM}-09`,'Transfer · 券商',500,'',''],
                  [`${curYM}-10`,'Dining',40,'','']],
            [['VOO','IndexETF',40,400,647.4]]);
  const cb=b.out();
  const last=cb[cb.length-1];
  ok('当月净入金 = 当月券商转账合计（1500+500）', last.net===2000, String(last.net));
  ok('餐饮支出不算进净入金', last.net===2000);
  ok('累计投入含当月', last.invested===11500+2000, String(last.invested));

  // 历史行不能被改写
  const before=JSON.stringify(b.SH._rows.slice(1,5));
  b.out();
  ok('历史行一律不动', JSON.stringify(b.SH._rows.slice(1,5))===before);

  // 券商现金也要算进账户总值
  let c=run(hist,[],[['VOO','IndexETF',10,400,100],['Cash','Gold.coin.cash','','',3000]]);
  const lc=c.out()[c.out().length-1];
  ok('账户总值含券商现金（10×100 + 3000）', Math.abs(lc.value-4000)<0.01, String(lc.value));

  // 前端画图
  const HIST=[{ym:'2023-01',value:8628,invested:7500,net:0},
              {ym:'2023-06',value:13948,invested:12000,net:500},
              {ym:'2024-08',value:25896,invested:18500,net:0}];
  function app(h){
    const dom=new JSDOM(html,{runScripts:'dangerously',url:'https://x.test/',
      beforeParse(w){ w.localStorage.clear();
        w.fetch=(u)=>{ const act=new URL(u).searchParams.get('action');
          if(act) return Promise.resolve({json:async()=>({status:'success'})});
          return Promise.resolve({json:async()=>({status:'success',
            stock:[{symbol:'VOO',category:'IndexETF',shares:40,cost:400,price:647.4}],
            expense:{},ledger:[],monthly:{},ledgerMonths:[],cash:{balance:0,hasAnchor:false},
            savings:[],bond:[],notices:[],hsa:null,retire:[],config:{},property:null,
            stockHistory:h,serverDate:'2026-09-05'})}); };
        w.alert=()=>{}; w.confirm=()=>true; w.scrollTo=()=>{};
        w.Element.prototype.scrollIntoView=function(){}; }});
    return dom.window;
  }
  let w=app(HIST); await wait(500);
  w.openDetail('stock');
  const box=w.document.getElementById('stockChart');
  ok('股票页顶部出现图表', !!box && /累计收益/.test(box.textContent));
  ok('累计收益 = 市值 − 累计投入 = 7396', /7,396/.test(box.textContent), box.textContent.slice(0,60));
  ok('收益率 +40.0%', /40\.0%/.test(box.textContent));
  ok('总额那一栏在曲线上面（hero 先出现）',
     w.document.getElementById('d-body').innerHTML.indexOf('dhero') <
     w.document.getElementById('d-body').innerHTML.indexOf('stockChart'));
  ok('有 YTD / 1Y / 3Y / All 四个区间',
     ['YTD','1Y','3Y','All'].every(t=>box.textContent.includes(t)));
  ok('横轴四个刻度，不再只有首尾',
     box.querySelectorAll('div[style*="justify-content:space-between"] span').length>=4,
     String(box.querySelectorAll('div[style*="justify-content:space-between"] span').length));
  ok('两条折线都画了', box.querySelectorAll('svg path').length>=3,
     String(box.querySelectorAll('svg path').length));
  ok('图例标出市值与累计投入', /市值/.test(box.textContent) && /累计投入/.test(box.textContent));

  // 数据不足时给出可操作的提示，而不是空白或坏图
  w=app([]); await wait(500);
  w.openDetail('stock');
  const empty=w.document.getElementById('stockChart');
  ok('没有历史数据时提示去 StockHistory 填', /StockHistory/.test(empty.textContent), empty.textContent.slice(0,40));
  ok('不画出坏图', empty.querySelectorAll('svg').length===0);
}

// 区间切换：期初基准取区间前一个点，所以 1Y 收益 ≠ 累计收益
console.log('\n— 曲线的区间切换 —');
{
  const T=new Date(); const Y=T.getFullYear();
  const h=[]; let cum=0, v=10000;
  for(let yy=Y-4; yy<=Y; yy++) for(let mm=1; mm<=12; mm++){
    if(yy===Y && mm>T.getMonth()+1) break;
    cum+=100; v+=300;
    h.push({ym:`${yy}-${pad(mm)}`, value:v, invested:cum, net:100});
  }
  const dom=new JSDOM(html,{runScripts:'dangerously',url:'https://x.test/',
    beforeParse(w){ w.localStorage.clear();
      w.fetch=(u)=>{ const a=new URL(u).searchParams.get('action');
        if(a) return Promise.resolve({json:async()=>({status:'success'})});
        return Promise.resolve({json:async()=>({status:'success',
          stock:[{symbol:'VOO',category:'IndexETF',shares:10,cost:100,price:200}],
          expense:{},ledger:[],monthly:{},ledgerMonths:[],cash:{balance:0,hasAnchor:false},
          savings:[],bond:[],notices:[],hsa:null,retire:[],config:{},property:null,
          stockHistory:h,serverDate:`${Y}-${pad(T.getMonth()+1)}-05`})}); };
      w.alert=()=>{}; w.confirm=()=>true; w.scrollTo=()=>{};
      w.Element.prototype.scrollIntoView=function(){}; }});
  const w=dom.window; await wait(500);
  w.openDetail('stock');
  const head=()=>w.document.getElementById('stockChart').querySelector('div>div>div').textContent;
  w.setStockRange('all');
  ok('All 显示「累计收益」', /累计收益/.test(head()), head().slice(0,20));
  const allGain=head();
  w.setStockRange('y1');
  ok('1Y 显示「1Y 收益」', /1Y 收益/.test(head()), head().slice(0,20));
  ok('1Y 收益小于累计收益（期初基准不同）', head()!==allGain, head().slice(0,26));
  w.setStockRange('ytd');
  ok('YTD 显示「YTD 收益」', /YTD 收益/.test(head()), head().slice(0,20));
  ok('切区间后市值仍是最新一期',
     /市值 \$/.test(w.document.getElementById('stockChart').textContent));
  w.setStockRange('y3');
  ok('3Y 也能切', /3Y 收益/.test(head()), head().slice(0,20));
  ok('区间切换不产生非数字', !/NaN/.test(visibleText(w.document)));
}

// ══════════════ 二十二、RSU 单独记，并自动同步到 Stock ══════════════
// RSU 表按归属批次记，是唯一数据源；Stock 表里 ASMIY 的股数和成本由脚本
// 自动写。曲线刻意剔除 RSU —— 那部分你没掏钱，算进去会凭空多出「收益」。
console.log('\n— RSU 单独记 —');
{
  const RHEAD=['Grant','Vest Date','Shares','Vest Price','Note'];
  const T=new Date(); const past=fmt(daysAgo(200)), soon=fmt(daysAgo(-120));
  function run(lots, stockRows){
    const R=mkSheet('RSU',[RHEAD,...lots]);
    const S=mkSheet('Stock',[['Symbol','Category','Share','Cost','Price'],...stockRows]);
    const tabs={RSU:R, Stock:S,
      StockHistory:mkSheet('StockHistory',[['年月','月初account value','本月净入金','月末account value','本月投资收益','月收益率']]),
      Ledger:mkSheet('Ledger',[['Date','category','amount','note','key']]),
      Anchor:mkSheet('Anchor',[['date','amount','note']]),
      Savings:mkSheet('Savings',[['ID','Name','Type','Balance','Rate','Last Post','Next update','Maturity','Status']]),
      Bond:mkSheet('Bond',[['ID','Name','Type','Start','Term','Rate','Principal','Balance','LastPost','Status']]),
      Ledger_monthly:mkSheet('Ledger_monthly',[])};
    const {doGet}=backend(tabs);
    return {out:()=>J(doGet({parameter:{}})), S, R};
  }
  const lots=[['2023 Grant',fmt(daysAgo(400)),100,20,''],
              ['2023 Grant',past,120,30,''],
              ['2024 Grant',soon,150,'','未归属']];
  const stk=[['ASMIY','Semiconductor',0,0,50],['VOO','IndexETF',10,400,600],
             ['Cash','Gold.coin.cash','','',1000]];
  let a=run(lots,stk);
  const r=a.out().rsu;
  ok('已归属股数 = 100 + 120', r.vestedShares===220, String(r.vestedShares));
  ok('未归属股数 = 150', r.pendingShares===150, String(r.pendingShares));
  ok('成本基准 = 100×20 + 120×30 = 5600', r.costBasis===5600, String(r.costBasis));
  ok('加权平均归属价 = 5600/220', Math.abs(r.avgCost-25.4545)<0.001, String(r.avgCost));
  ok('给出下次归属日期与股数',
     r.nextVest && r.nextVest.date===soon && r.nextVest.shares===150, JSON.stringify(r.nextVest));
  ok('自动写回 Stock 表的 ASMIY 股数',
     Number(a.S._rows[1][2])===220, String(a.S._rows[1][2]));
  ok('自动写回成本', Math.abs(Number(a.S._rows[1][3])-25.4545)<0.001, String(a.S._rows[1][3]));
  ok('绝不碰 E 列的价格（GOOGLEFINANCE）', Number(a.S._rows[1][4])===50);

  // 曲线要剔除 RSU：账户总值 = VOO 10×600 + 现金 1000 = 7000，不含 220×50
  const hist=a.out().stockHistory;
  ok('曲线的账户总值不含 RSU（10×600 + 1000 = 7000）',
     Math.abs(hist[hist.length-1].value-7000)<0.01, String(hist[hist.length-1].value));

  // 幂等：没变就不重复写
  const before=JSON.stringify(a.S._rows);
  a.out();
  ok('股数没变就不重复写 Stock', JSON.stringify(a.S._rows)===before);

  // 新归属一批 → 自动跟着涨
  let b=run([...lots,['2024 Grant',fmt(daysAgo(5)),80,40,'']],stk);
  ok('新归属一批后股数自动增加', b.out().rsu.vestedShares===300, String(b.out().rsu.vestedShares));
  ok('成本基准跟着更新 = 5600 + 80×40', b.out().rsu.costBasis===8800, String(b.out().rsu.costBasis));

  // 没建 RSU 表时不能影响其他模块
  // 注意：mkSheet 直接引用传入的数组，前面的用例已经把 stk 写成 220 了，
  // 这里必须自己造一份干净的，否则测的是脏数据
  const freshStk=[['ASMIY','Semiconductor',7,3,50],['VOO','IndexETF',10,400,600],
                  ['Cash','Gold.coin.cash','','',1000]];
  const noR={Stock:mkSheet('Stock',[['Symbol','Category','Share','Cost','Price'],...freshStk]),
    StockHistory:mkSheet('StockHistory',[['年月','月初account value','本月净入金','月末account value','本月投资收益','月收益率']]),
    Ledger:mkSheet('Ledger',[['Date','category','amount','note','key']]),
    Anchor:mkSheet('Anchor',[['date','amount','note']]),
    Savings:mkSheet('Savings',[['ID','Name','Type','Balance','Rate','Last Post','Next update','Maturity','Status']]),
    Bond:mkSheet('Bond',[['ID','Name','Type','Start','Term','Rate','Principal','Balance','LastPost','Status']]),
    Ledger_monthly:mkSheet('Ledger_monthly',[])};
  const g=backend(noR);
  const rr=J(g.doGet({parameter:{}}));
  // 现在读路径会把 RSU 表建出来并返回全 0（界面才好显示空编辑器）
  ok('没有 RSU 表时自动建表，其他模块照常',
     rr.status==='success' && rr.rsu && rr.rsu.vestedShares===0, JSON.stringify(rr.rsu));
  ok('归属表为空时不改写 Stock 的 ASMIY', Number(noR.Stock._rows[1][2])===7,
     String(noR.Stock._rows[1][2]));

  // 前端展示
  const dom=new JSDOM(html,{runScripts:'dangerously',url:'https://x.test/',
    beforeParse(w){ w.localStorage.clear();
      w.fetch=(u)=>{ const act=new URL(u).searchParams.get('action');
        if(act) return Promise.resolve({json:async()=>({status:'success'})});
        return Promise.resolve({json:async()=>({status:'success',
          stock:[{symbol:'ASMIY',category:'Semiconductor',shares:220,cost:25.45,price:50},
                 {symbol:'VOO',category:'IndexETF',shares:10,cost:400,price:600}],
          expense:{},ledger:[],monthly:{},ledgerMonths:[],cash:{balance:0,hasAnchor:false},
          savings:[],bond:[],notices:[],hsa:null,retire:[],config:{},property:null,
          rsu:{symbol:'ASMIY',vestedShares:220,costBasis:5600,avgCost:25.4545,
               pendingShares:150,nextVest:{date:soon,shares:150},lots:[]},
          stockHistory:[{ym:'2026-08',value:7000,invested:5000,net:0},
                        {ym:'2026-09',value:7000,invested:5000,net:0}],
          serverDate:'2026-09-05'})}); };
      w.alert=()=>{}; w.confirm=()=>true; w.scrollTo=()=>{};
      w.Element.prototype.scrollIntoView=function(){}; }});
  const w=dom.window; await wait(500);
  w.openDetail('stock');
  // 独立的 RSU 汇总框已去掉：信息本来就重复，改到点开 ASMIY 的面板里看
  ok('不再有独立的 RSU 汇总框', !w.document.getElementById('rsuBox'));
  w.showStockPanel(0,0);
  const box=w.document.getElementById('slot-0-0');
  ok('面板小字与其他股票一致（已归属 RSU · 股数 · 均成本 · 现价）',
     /已归属 RSU/.test(box.textContent) && /均成本/.test(box.textContent) &&
     /现价/.test(box.textContent), box.textContent.slice(0,70));
  ok('面板里显示未归属股数与下次归属日',
     /150 股/.test(box.textContent) && /下次/.test(box.textContent));
  ok('注明未归属不计入总额', /不计入总额/.test(box.textContent));
  ok('曲线标明不含 RSU',
     /不含 RSU/.test(w.document.getElementById('stockChart').textContent));
}

// ══════════════ 二十三、RSU 归属表在小程序里编辑 ══════════════
// ASMIY 的下拉面板换成归属表编辑器：一行三格（日期/股数/成本）+ 加行。
// 股数与成本由这张表推导，所以不再提供增持减持 —— 那会和归属表打架。
console.log('\n— RSU 归属表编辑器 —');
{
  const RHEAD=['Grant','Vest Date','Shares','Vest Price','Note'];
  // 后端：读路径要把表建出来（Property/Config 踩过同一个坑）
  {
    const tabs={Stock:mkSheet('Stock',[['Symbol','Category','Share','Cost','Price'],
                  ['ASMIY','Semiconductor',0,0,50]]),
      StockHistory:mkSheet('StockHistory',[['年月','月初account value','本月净入金','月末account value','本月投资收益','月收益率']]),
      Ledger:mkSheet('Ledger',[['Date','category','amount','note','key']]),
      Anchor:mkSheet('Anchor',[['date','amount','note']]),
      Savings:mkSheet('Savings',[['ID','Name','Type','Balance','Rate','Last Post','Next update','Maturity','Status']]),
      Bond:mkSheet('Bond',[['ID','Name','Type','Start','Term','Rate','Principal','Balance','LastPost','Status']]),
      Ledger_monthly:mkSheet('Ledger_monthly',[])};
    const {doGet}=backend(tabs);
    const r=J(doGet({parameter:{}}));
    ok('只读一次就把 RSU 表建出来', !!tabs.RSU);
    ok('空表时返回 0 而不是 null（界面才好显示编辑器）',
       r.rsu && r.rsu.vestedShares===0, JSON.stringify(r.rsu&&r.rsu.vestedShares));

    // 整表覆盖式保存
    const lots=[{date:fmt(daysAgo(300)),shares:100,price:20},
                {date:fmt(daysAgo(30)),shares:120,price:30},
                {date:fmt(daysAgo(-200)),shares:150,price:''}];
    const sv=J(doGet({parameter:{action:'saveRsu', lots:JSON.stringify(lots)}}));
    ok('保存后已归属 = 100+120', sv.rsu.vestedShares===220, String(sv.rsu.vestedShares));
    ok('未归属 = 150', sv.rsu.pendingShares===150, String(sv.rsu.pendingShares));
    ok('未归属的成本留空不算进基准', sv.rsu.costBasis===5600, String(sv.rsu.costBasis));
    ok('写进了 RSU 表', tabs.RSU._rows.length===4, String(tabs.RSU._rows.length));
    ok('同步写回 Stock 的股数', Number(tabs.Stock._rows[1][2])===220);

    // 再保存一次（少一行）应整表覆盖，不残留旧行
    J(doGet({parameter:{action:'saveRsu', lots:JSON.stringify(lots.slice(0,1))}}));
    // 注意跳过表头：表头的 B 列写着「Vest Date」，是 truthy
    const dataRows=()=>tabs.RSU._rows.slice(1).filter(r=>r[1]);
    ok('整表覆盖，删掉的行不残留', dataRows().length===1, String(dataRows().length));
    ok('删行后 Stock 股数跟着降到 100', Number(tabs.Stock._rows[1][2])===100);
    // 没填日期的行直接丢掉
    J(doGet({parameter:{action:'saveRsu', lots:JSON.stringify([{date:'',shares:9,price:1}])}}));
    ok('没填日期的行不写入', dataRows().length===0, String(dataRows().length));
  }

  // 前端编辑器
  const RSU={symbol:'ASMIY',vestedShares:220,costBasis:5600,avgCost:25.4545,
             pendingShares:150,nextVest:{date:'2027-01-05',shares:150},
             lots:[{date:'2025-08-01',shares:100,price:20},
                   {date:'2026-02-15',shares:120,price:30},
                   {date:'2027-01-05',shares:150,price:0}]};
  const calls=[];
  const dom=new JSDOM(html,{runScripts:'dangerously',url:'https://x.test/',
    beforeParse(w){ w.localStorage.clear();
      w.fetch=(u)=>{ const a=new URL(u).searchParams.get('action');
        if(a){ calls.push(Object.fromEntries(new URL(u).searchParams));
               return Promise.resolve({json:async()=>({status:'success',rsu:RSU})}); }
        return Promise.resolve({json:async()=>({status:'success',
          stock:[{symbol:'ASMIY',category:'Semiconductor',shares:220,cost:25.45,price:50},
                 {symbol:'VOO',category:'IndexETF',shares:10,cost:400,price:600}],
          expense:{},ledger:[],monthly:{},ledgerMonths:[],cash:{balance:0,hasAnchor:false},
          savings:[],bond:[],notices:[],hsa:null,retire:[],config:{},property:null,rsu:RSU,
          stockHistory:[{ym:'2026-08',value:7000,invested:5000,net:0},
                        {ym:'2026-09',value:7000,invested:5000,net:0}],
          serverDate:'2026-09-05'})}); };
      w.alert=(m)=>{ w.__alert=m; }; w.confirm=()=>true; w.scrollTo=()=>{};
      w.Element.prototype.scrollIntoView=function(){}; }});
  const w=dom.window, d=w.document; await wait(500);
  w.openDetail('stock');
  w.showStockPanel(0,0);
  const ed=d.getElementById('rsuEditor');
  ok('点 ASMIY 打开的是归属表编辑器', !!ed);
  ok('不再出现增持/减持', !/增持|减持/.test(d.getElementById('slot-0-0').textContent));
  ok('已有三行归属记录', d.querySelectorAll('#rsuEditor input[id^=rsu-d-]').length===3,
     String(d.querySelectorAll('#rsuEditor input[id^=rsu-d-]').length));
  ok('一行三格：日期/股数/成本',
     !!d.getElementById('rsu-d-0') && !!d.getElementById('rsu-s-0') && !!d.getElementById('rsu-p-0'));
  ok('已归属/待归属有标记', /已归/.test(ed.textContent) && /待归/.test(ed.textContent));
  ok('有「添加一行」', /添加一行/.test(ed.textContent));

  w.rsuAddRow();
  ok('点加行后多一行', d.querySelectorAll('#rsuEditor input[id^=rsu-d-]').length===4);
  ok('加行不丢已填内容', d.getElementById('rsu-s-1').value==='120',
     d.getElementById('rsu-s-1').value);
  d.getElementById('rsu-d-3').value='2027-07-01';
  d.getElementById('rsu-s-3').value='80';
  w.rsuSave(); await wait(400);
  const sent=calls.filter(c=>c.action==='saveRsu');
  ok('保存发出 saveRsu', sent.length===1, `${sent.length} 次`);
  const body=JSON.parse(sent[0].lots);
  ok('提交四行', body.length===4, String(body.length));
  ok('新加的那行内容正确',
     body[3].date==='2027-07-01' && body[3].shares===80, JSON.stringify(body[3]));
  ok('未归属的成本提交为空串，不写成 0', body[2].price==='', JSON.stringify(body[2].price));

  // 删行（面板是 toggle，重开一次即可，连点两次会收起来）
  w.showStockPanel(0,0);
  w.rsuDelRow(0);
  ok('删行后少一行', d.querySelectorAll('#rsuEditor input[id^=rsu-d-]').length===2,
     String(d.querySelectorAll('#rsuEditor input[id^=rsu-d-]').length));

  // 非 RSU 的股票面板不受影响
  w.showStockPanel(1,0);
  ok('普通股票仍是增持/减持面板',
     /增持/.test(d.getElementById('slot-1-0').textContent) &&
     !d.querySelector('#slot-1-0 #rsuEditor'));
}

// ══════════════ 二十四、分类调整 / 排序 / 右滑返回 ══════════════
console.log('\n— 分类、排序、返回手势 —');
{
  const rows=[{symbol:'NVDA',category:'Semiconductor',shares:10,cost:100,price:200},   // 2000
              {symbol:'VOO',category:'IndexETF',shares:10,cost:100,price:600},          // 6000
              {symbol:'AVUV',category:'Value ETF',shares:10,cost:50,price:90},          // 900
              {symbol:'PPA',category:'Other ETF',shares:10,cost:50,price:120},          // 1200
              {symbol:'COST',category:'Tech/AI',shares:1,cost:900,price:950},           // 旧分类
              {symbol:'VEA',category:'International ETF',shares:10,cost:40,price:55}];  // 550
  const dom=new JSDOM(html,{runScripts:'dangerously',url:'https://x.test/',
    beforeParse(w){ w.localStorage.clear();
      w.fetch=(u)=>{ const a=new URL(u).searchParams.get('action');
        if(a) return Promise.resolve({json:async()=>({status:'success'})});
        return Promise.resolve({json:async()=>({status:'success',stock:rows,expense:{},ledger:[],
          monthly:{},ledgerMonths:[],cash:{balance:0,hasAnchor:false},savings:[],bond:[],
          notices:[],hsa:null,retire:[],config:{},property:null,rsu:null,
          stockHistory:[],serverDate:'2026-09-08'})}); };
      w.alert=()=>{}; w.confirm=()=>true; w.scrollTo=()=>{};
      w.Element.prototype.scrollIntoView=function(){}; }});
  const w=dom.window, d=w.document; await wait(500);
  w.openDetail('stock');
  const cats=[].map.call(d.querySelectorAll('#d-body .stkcat'),
    e=>e.childNodes[0].textContent.trim());
  ok('半导体', cats.includes('半导体'), cats.join(' | '));
  ok('美股核心仓', cats.includes('美股核心仓'));
  ok('因子 ETF', cats.includes('因子 ETF'));
  ok('美股行业/主题 ETF', cats.includes('美股行业/主题 ETF'));
  ok('国际股票', cats.includes('国际股票'));
  ok('旧分类 Tech/AI 归到美国个股，不掉进兜底类',
     w.eval("CAT_MAP[LEGACY_CAT['Tech/AI']]")==='美国个股');
  ok('按金额从高到低排（核心仓 6000 在半导体 2000 之前）',
     cats.indexOf('美股核心仓')<cats.indexOf('半导体'), cats.join(' > '));
  ok('最小的国际股票排在最后', cats[cats.length-1]==='国际股票', cats[cats.length-1]);
  const subs=[].map.call(d.querySelectorAll('#d-body .stksub'),e=>e.textContent.trim());
  ok('小字跟着改', subs.includes('美股大盘指数') && subs.includes('美股价值'), subs.join(' | '));

  // 点开个股要定位到正确的分组（排序后 gi 不能错位）
  const semiIdx=cats.indexOf('半导体');
  ok('排序后点个股仍定位正确',
     (function(){ w.openDetail('stock');
       const row=d.querySelectorAll('#d-body .stkgroup')[semiIdx].querySelector('.row');
       row.onclick(); const slot=d.querySelector('.stock-slot:not(:empty)');
       return !!slot && /NVDA/.test(slot.textContent); })());

  // 右滑返回
  ok('提供右滑返回', typeof w.goBackOne==='function');
  w.openDetail('stock');
  w.goBackOne();
  ok('右滑从详情页回到总览',
     d.getElementById('scr-overview').classList.contains('active'));
  w.goBackOne();
  ok('已在总览时再滑不出事', d.getElementById('scr-overview').classList.contains('active'));

  // 非激活屏不再被拉伸（iOS 上会让滚动区域残留上一页长度）
  const css=html.slice(html.indexOf('.screen{'), html.indexOf('.screen.active'));
  ok('非激活屏不写 bottom:0', !/inset:0/.test(css), css.slice(0,60));
  ok('min-height 只给激活屏', /\.screen\.active\{[^}]*min-height:100%/.test(html));

  // 曲线下面的 RSU 汇总框已去掉
  ok('不再有独立的 RSU 汇总框', !d.getElementById('rsuBox'));
}

// ══════════════ 二十五、转账：股票账户 → Chase ══════════════
// 从券商把钱转回 Chase：券商现金减、Chase 加，记「Redeem · 券商」——
// 加现金但不算收入（和国债到期本金同一类），净收入不会被抬高。
console.log('\n— 转账：股票账户 → Chase —');
{
  const calls=[];
  const dom=new JSDOM(html,{runScripts:'dangerously',url:'https://x.test/',
    beforeParse(w){ w.localStorage.clear();
      w.fetch=(u)=>{ const url=new URL(u), a=url.searchParams.get('action');
        if(a){ calls.push(Object.fromEntries(url.searchParams));
               return Promise.resolve({json:async()=>({status:'success',row:77})}); }
        return Promise.resolve({json:async()=>({status:'success',
          stock:[{symbol:'VOO',category:'IndexETF',shares:10,cost:400,price:600},
                 {symbol:'Cash',category:'Gold.coin.cash',shares:'',cost:'',price:5000}],
          expense:{},ledger:[],monthly:{},ledgerMonths:[],
          cash:{balance:10000,anchorDate:'2026-09-01',anchorAmount:10000,hasAnchor:true,bookStart:'2026-09-01'},
          savings:[{id:'s1',name:'OS-9250',type:'OS',balance:20000,rate:0.034,lastPost:'2026-09-01',
                    nextUpdate:'',maturity:'',status:'active'}],
          bond:[],notices:[],hsa:null,retire:[],config:{},property:null,rsu:null,
          stockHistory:[],serverDate:'2026-09-10'})}); };
      w.alert=(m)=>{ w.__alert=m; }; w.confirm=()=>true; w.scrollTo=()=>{};
      w.Element.prototype.scrollIntoView=function(){}; }});
  const w=dom.window, d=w.document; await wait(500);
  w.openTransfer();
  const srcOpts=[].map.call(d.querySelectorAll('#x-src option'),o=>o.value);
  const tgtOpts=[].map.call(d.querySelectorAll('#x-tgt option'),o=>o.value);
  ok('From 里有股票账户', srcOpts.includes('broker'), srcOpts.join(','));
  ok('To 里有 Checking - Chase', tgtOpts.includes('chase'), tgtOpts.join(','));

  const broker0=w.eval("brokerCashItem().mv"), cash0=w.eval("cashInfo.balance");
  d.getElementById('x-src').value='broker';
  d.getElementById('x-tgt').value='chase';
  d.getElementById('x-amt').value='1500';
  w.submitOneTime(); await wait(400);
  ok('券商现金减 1500', w.eval("brokerCashItem().mv")===broker0-1500, String(w.eval("brokerCashItem().mv")));
  ok('Chase 流动现金加 1500', w.eval("cashInfo.balance")===cash0+1500, String(w.eval("cashInfo.balance")));
  const led=calls.find(c=>c.action==='addLedger');
  ok('Ledger 记一笔 Redeem · 券商', led && led.category==='Redeem · 券商' && +led.amount===1500,
     JSON.stringify(led));
  ok('券商现金按增量写回', calls.some(c=>c.action==='adjustCash' && +c.delta===-1500));
  ok('不算收入（Redeem 不进 income）', w.eval("catKind('Redeem · 券商')")==='redeem');

  // 同一账户互转要拦下并提示，而不是静默没反应
  w.__alert=null; calls.length=0;
  d.getElementById('x-src').value='broker';
  d.getElementById('x-tgt').value='broker';
  d.getElementById('x-amt').value='100';
  w.submitOneTime(); await wait(200);
  ok('同账户互转被拦下并提示', /同一个账户/.test(w.__alert||''), String(w.__alert));
  ok('且不发任何写请求', calls.length===0, `${calls.length} 次`);

  // Marcus → Chase 顺带也通了
  calls.length=0;
  d.getElementById('x-src').value='marcus';
  d.getElementById('x-tgt').value='chase';
  d.getElementById('x-amt').value='800';
  w.submitOneTime(); await wait(400);
  ok('Marcus → Chase 记 Redeem · Marcus',
     calls.some(c=>c.action==='addLedger' && c.category==='Redeem · Marcus' && +c.amount===800));
  ok('并从 Marcus 扣款', calls.some(c=>c.action==='adjustSaving' && +c.delta===-800));

  // 后端：StockHistory 的净入金要扣掉从券商转出的钱
  const T=new Date(); const ym=`${T.getFullYear()}-${pad(T.getMonth()+1)}`;
  const SH=mkSheet('StockHistory',[['年月','月初account value','本月净入金','月末account value','本月投资收益','月收益率']]);
  const tabs={StockHistory:SH,
    Ledger:mkSheet('Ledger',[['Date','category','amount','note','key'],
      [`${ym}-02`,'Transfer · 券商',3000,'',''],
      [`${ym}-05`,'Redeem · 券商',1200,'',''],
      [`${ym}-06`,'Redeem · 国债',999,'','']]),       // 国债回流与券商无关
    Stock:mkSheet('Stock',[['Symbol','Category','Share','Cost','Price'],['VOO','IndexETF',10,400,600]]),
    Anchor:mkSheet('Anchor',[['date','amount','note']]),
    Savings:mkSheet('Savings',[['ID','Name','Type','Balance','Rate','Last Post','Next update','Maturity','Status']]),
    Bond:mkSheet('Bond',[['ID','Name','Type','Start','Term','Rate','Principal','Balance','LastPost','Status']]),
    Ledger_monthly:mkSheet('Ledger_monthly',[])};
  const {doGet}=backend(tabs);
  const h=J(doGet({parameter:{}})).stockHistory;
  const cur=h[h.length-1];
  ok('本月净入金 = 转入 3000 − 转出 1200 = 1800', cur.net===1800, String(cur.net));
  ok('国债回流不算进券商净入金', cur.net===1800);
}
