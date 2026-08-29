import fs from 'fs';
const src=fs.readFileSync('AppsScript.gs','utf8');
const ok=(l,c,e='')=>console.log((c?'✅':'❌'),l,e);

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

// 用户真实数据（Rate 用百分数写法，验证兼容）
const sav=mkSheet('Savings',[
 ['ID','Name','Type','Balance','Rate','Last Post','Next update','Maturity','Status'],
 ['','CD-7597','CD',21335,4,'2026-08-01','2026-09-01','2027-06-30','Activate'],
 ['','CD-9185','CD',10274,4.10,'2026-08-01','2026-09-01','2027-02-05','Activate'],
 ['','OS-9250','OS',30057,3.40,'2026-08-01','2026-09-01','',''],
]);
const tabs={Savings:sav,Stock:mkSheet('Stock',[['Symbol','Category','Share','Cost','Price']]),
            Ledger:mkSheet('Ledger',[['Date','category','amount']])};
let NOW=new Date('2026-08-23T00:00:00Z');
globalThis.SpreadsheetApp={flush(){},getActiveSpreadsheet:()=>({getSheetByName:n=>tabs[n]||null,
  insertSheet(n){tabs[n]=mkSheet(n,[]);return tabs[n];}})};
globalThis.ContentService={MimeType:{JSON:'json'},
  createTextOutput:t=>({_t:t,setMimeType(){return this;},getContent(){return this._t;}})};
globalThis.Utilities={getUuid:()=>Math.random().toString(36).slice(2)+'0000000',
  // Apps Script 的脚本时区就是 Phoenix，内部 new Date(y,m,d) 造的就是当地时间，
  // 所以这里直接按 UTC 格式化即可（node 里 local==UTC），不能再减 7 小时
  formatDate:(d,tz,f)=>{const p=new Date(d.getTime()).toISOString();
    return f==='yyyy-MM'?p.slice(0,7):p.slice(0,10);}};
globalThis.LockService={getScriptLock:()=>({waitLock(){},releaseLock(){}})};
globalThis.Logger={log:(...a)=>console.log('  ',...a)};
globalThis.ScriptApp={getProjectTriggers:()=>[],newTrigger:()=>({timeBased:()=>({atHour:()=>({everyDays:()=>({create:()=>{}})})})})};
const RealDate=Date;
globalThis.Date=class extends RealDate{constructor(...a){if(a.length===0)super(NOW.getTime());else super(...a);}
  static now(){return NOW.getTime();}};

const api=new Function('return (function(){'+src+'\nreturn {doGet,doPost,runSavings,readSavRows,accrue};})()')();
const ss=SpreadsheetApp.getActiveSpreadsheet();
const J=r=>JSON.parse(r.getContent());

console.log('— 利率格式与 APY 换算 —');
ok('APY 换算：1年正好等于 APY', Math.abs(api.accrue(10000,0.034,365)-340)<0.01,
   '$10,000 @3.4% 一年利息 $'+api.accrue(10000,0.034,365).toFixed(2));

console.log('\n— 8/23 打开：还没到 9/1，余额不应变 —');
api.runSavings(ss); let rows=api.readSavRows(sav);
ok('8月余额不动', rows[0].bal===21335 && rows[2].bal===30057, `CD-7597 $${rows[0].bal} / OS $${rows[2].bal}`);
ok('自动补上 ID', !!rows[0].id && !!rows[2].id, rows.map(r=>r.id).join(' '));
ok('Status 规范成 active', sav._rows[1][8]==='active', sav._rows[1][8]);

console.log('\n— 跳到 9/1：应入 8/1→9/1 共 31 天的息 —');
NOW=new RealDate('2026-09-01T00:00:00Z');
api.runSavings(ss); rows=api.readSavRows(sav);
const exp7597=21335+api.accrue(21335,0.04,31);
ok('CD-7597 入息 31 天', Math.abs(rows[0].bal-exp7597)<0.02, `期望 ${exp7597.toFixed(2)}，实际 ${rows[0].bal}`);
ok('Last Post 推到 9/1', sav._rows[1][5]==='2026-09-01', sav._rows[1][5]);
ok('Next update 推到 10/1', sav._rows[1][6]==='2026-10-01', sav._rows[1][6]);
const bal9=rows[0].bal;
api.runSavings(ss);
ok('同一天重复运行不重复计息', api.readSavRows(sav)[0].bal===bal9);

console.log('\n— 跳到 2027/2/6：CD-9185 已于 2/5 到期 —');
NOW=new RealDate('2027-02-06T00:00:00Z');
const osBefore=api.readSavRows(sav).find(r=>r.name==='OS-9250').bal;
api.runSavings(ss); rows=api.readSavRows(sav);
const cd9185=api.readSavRows(sav).find(r=>r.name==='CD-9185');
const osAfter=api.readSavRows(sav).find(r=>r.name==='OS-9250');
ok('CD-9185 已关闭', cd9185.closed, sav._rows[2][8]);
ok('本息转入 OS', osAfter.bal>osBefore, `OS $${osBefore.toFixed(2)} → $${osAfter.bal.toFixed(2)}（+$${(osAfter.bal-osBefore).toFixed(2)}）`);
ok('OS 自己的利息也照常计', osAfter.bal-osBefore>cd9185.bal, `到期本息 $${cd9185.bal.toFixed(2)}，OS 实增 $${(osAfter.bal-osBefore).toFixed(2)}`);
const nt=J(api.doGet({parameter:{}})).notices;
ok('生成到期通知', nt.length>0, nt.map(n=>n.text).join(' | '));
const osB2=api.readSavRows(sav).find(r=>r.name==='OS-9250').bal;
api.runSavings(ss);
ok('到期不会重复入账', Math.abs(api.readSavRows(sav).find(r=>r.name==='OS-9250').bal-osB2)<0.01);

console.log('\n— 改利率：先按旧利率结息，再换新利率 —');
NOW=new RealDate('2027-02-20T00:00:00Z');
api.runSavings(ss);
const os=api.readSavRows(sav).find(r=>r.name==='OS-9250');
const before=os.bal, lastPost=os.post;
const days=Math.round((new RealDate('2027-02-20T00:00:00')-lastPost)/86400000);
const shouldItr=api.accrue(before,0.034,days);
J(api.doGet({parameter:{action:'updateSaving',id:os.id,rate:3.0}}));
const os2=api.readSavRows(sav).find(r=>r.name==='OS-9250');
ok('改利率时按旧利率结息', Math.abs(os2.bal-(before+shouldItr))<0.02,
   `$${before.toFixed(2)} + $${shouldItr.toFixed(2)} = $${os2.bal.toFixed(2)}`);
ok('锚点重设为今天', sav._rows[3][5]==='2027-02-20', sav._rows[3][5]);
ok('利率已换新', Math.abs(os2.rate-0.03)<1e-9, os2.rate);

console.log('\n— 手动改余额：不结息，直接采信 —');
J(api.doGet({parameter:{action:'updateSaving',id:os.id,balance:33333}}));
ok('余额=手动值', api.readSavRows(sav).find(r=>r.name==='OS-9250').bal===33333);

console.log('\n— 新开 CD：从 OS 扣款，先结息再扣 —');
NOW=new RealDate('2027-03-15T00:00:00Z');
api.runSavings(ss);
const osB3=api.readSavRows(sav).find(r=>r.name==='OS-9250').bal;
J(api.doGet({parameter:{action:'addSaving',name:'CD-NEW',type:'CD',balance:10000,
  rate:4.5,maturity:'2028-03-15',srcId:os.id}}));
const osB4=api.readSavRows(sav).find(r=>r.name==='OS-9250').bal;
const nw=api.readSavRows(sav).find(r=>r.name==='CD-NEW');
ok('新 CD 已建立', !!nw && nw.bal===10000, nw?`$${nw.bal} @${(nw.rate*100).toFixed(2)}% 到期 ${nw.mat&&nw.mat.toISOString().slice(0,10)}`:'无');
ok('OS 扣款含先结息', Math.abs(osB4-(osB3-10000))>0 && osB4>osB3-10000,
   `$${osB3.toFixed(2)} → $${osB4.toFixed(2)}（扣 10000 但先补了利息）`);
ok('新账户起始日=今天', sav._rows.at(-1)[5]==='2027-03-15', sav._rows.at(-1)[5]);
