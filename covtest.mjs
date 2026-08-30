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
