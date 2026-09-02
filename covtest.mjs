// 补三处覆盖盲区：70 天流水窗口、整行回写不误伤其他列、首屏本地缓存。
// 这三处是 2026-08 改动最密集的地方，原来一条断言都没有。
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
   getValue:()=>(self._rows[r-1]||[])[c-1]??"",
   setValue(v){self._cell(r,c)[c-1]=v;},setFormula(v){self._cell(r,c)[c-1]=v;},
   getValues(){const o=[];for(let i=0;i<nr;i++){const row=[];
     for(let j=0;j<nc;j++)row.push((self._rows[r-1+i]||[])[c-1+j]??"");o.push(row);}return o;},
   setValues(v){for(let i=0;i<nr;i++){const row=self._cell(r+i,c+nc-1);
     for(let j=0;j<nc;j++)row[c-1+j]=v[i][j];}}};},
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
  ok('打开 App 不重发已记过的供款', a.calls.length===0, `发了 ${a.calls.length} 次`);

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
  ok('此刻补记的写请求还没回来', calls.length===0, `已完成 ${calls.length} 次`);
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
  ok('配置没存上就不发补记请求', b.calls.length===nBefore, `多发了 ${b.calls.length-nBefore} 次`);

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
  ok('没有账户时不显示 +NaN%', !/NaN/.test(dom2.window.document.body.innerHTML));
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
