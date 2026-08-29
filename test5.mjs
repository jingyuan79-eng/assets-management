import { JSDOM } from 'jsdom'; import fs from 'fs';
const html=fs.readFileSync('index.html','utf8');
const ok=(l,c,e='')=>console.log((c?'✅':'❌'),l,e);
const wait=(ms=350)=>new Promise(r=>setTimeout(r,ms));
const calls=[];
let cash=1326;
const dom=new JSDOM(html,{runScripts:'dangerously',url:'https://x.test/',
  beforeParse(w){ w.localStorage.clear();
    w.fetch=async(u)=>{const url=new URL(u);const a=url.searchParams.get('action');
      if(!a) return {json:async()=>({status:'success',
        stock:[{symbol:'NVDA',category:'Semiconductor',shares:32,cost:148.71,price:214.72},
               {symbol:'Cash',category:'Gold.coin.cash',shares:'',cost:'',price:cash}],
        ledger:[],expense:{},monthly:{},ledgerMonths:['2026-08','2026-07'],
        savings:[{id:'sv_c',name:'OS-9250',type:'OS',balance:30057,rate:0.034,
                  lastPost:'2026-08-01',nextUpdate:'2026-09-01',maturity:'',status:'active'}],
        notices:[]})};
      const p=Object.fromEntries(url.searchParams); calls.push(p);
      if(a==='updateCash') cash=+p.amount;
      return {json:async()=>({status:'success'})};};
    w.alert=m=>{}; w.confirm=()=>true; w.scrollTo=()=>{};
    w.Element.prototype.scrollIntoView=function(){};
  }});
const w=dom.window,d=w.document;
await wait();

console.log('【1】股票现金 → 添加分红');
w.openDetail('stock');
ok('点现金弹出的是分红表单', /showDividendForm/.test(d.getElementById('d-body').innerHTML));
w.showDividendForm();
ok('表单是分红录入（已精简文案）', /分红金额/.test(d.getElementById('cashslot').textContent)
   && !/当前券商现金/.test(d.getElementById('cashslot').textContent));
ok('没有「手动修改现金余额」', !/手动修改现金余额/.test(d.getElementById('cashslot').textContent));
const before=w.eval("brokerCashItem().mv");
d.getElementById('dv-amt').value='42.5'; d.getElementById('dv-from').value='voo';
w.applyDividend(); await wait(500);
ok('券商现金已增加', w.eval("brokerCashItem().mv")===before+43 || w.eval("brokerCashItem().mv")===before+42,
   `$${before} → $${w.eval("brokerCashItem().mv")}`);
ok('按增量写回 Sheet', calls.some(c=>c.action==='adjustCash'&&Number(c.delta)===42.5),
   'adjustCash delta='+(calls.filter(c=>c.action==='adjustCash').pop()||{}).delta);

console.log('\n【2】界面不再暴露后台细节');
const src=fs.readFileSync('index.html','utf8');
ok('删除持仓无括号说明', !/删除持仓（同时删/.test(src));
ok('储蓄不显示下次入息', !/下次入息/.test(src));
w.openDetail('cd');
ok('储蓄行只显示到期信息', !/下次入息/.test(d.getElementById('d-body').textContent),
   d.querySelector('#d-body .row .desc') ? d.querySelector('#d-body .row .desc').textContent : '');
ok('页脚不再写 Google Sheet', !/· Google Sheet/.test(d.getElementById('syncBadge').textContent),
   d.getElementById('syncBadge').textContent);

console.log('\n【3】固定支出覆盖 9 个分类');
const fc=JSON.parse(w.eval("JSON.stringify(FIXED_CATS)"));
ok('9 个分类', fc.length===9, fc.join(' / '));
ok('不含 Gifts & Families', !fc.includes('Gifts & Families'));
ok('不含 Other or unexpected', !fc.includes('Other or unexpected'));
w.openDetail('cash'); w.switchCashTab('ex');
w.openExpenseDrill('Dining');
ok('Dining 现在有固定支出区', /添加固定支出/.test(d.getElementById('cashTabBody').textContent));
w.openExpenseDrill('Gifts & Families');
ok('Gifts 仍然没有', !/添加固定支出/.test(d.getElementById('cashTabBody').textContent));

console.log('\n【4】定期转账频率 + 排版');
w.openTransfer();
ok('默认一次性模式下没有频率字段', !d.getElementById('x-freq'));
w.setXferMode('recurring');
ok('切到定期后出现频率下拉', !!d.getElementById('x-freq'),
   d.getElementById('x-freq')?[...d.getElementById('x-freq').options].map(o=>o.text).join('/'):'无');
ok('切到定期后出现首次转账日', !!d.getElementById('x-start'));
ok('输入框改用统一排版', /class="editform"/.test(d.getElementById('transfer-body').innerHTML));
ok('不再有 padding:15px 大号内联样式',
   !/padding:15px 14px[^"]*font-size:16px/.test(d.getElementById('transfer-body').innerHTML));

// 设一条每两周的定期转账，起始日在过去 → 应补执行
const chase0=w.eval("chaseItem().mv");
d.getElementById('x-src').value='chase'; d.getElementById('x-tgt').value='mortgage';
d.getElementById('x-amt').value='500'; d.getElementById('x-freq').value='biweekly';
const past=new Date(); past.setDate(past.getDate()-30);
d.getElementById('x-start').value=past.toISOString().slice(0,10);
w.submitTransfer(); await wait(400);
const rec=JSON.parse(w.eval("JSON.stringify(loadRecurring())"))[0];
ok('保存了频率与起始日', rec.freq==='biweekly' && !!rec.start, JSON.stringify({freq:rec.freq,start:rec.start,lastRun:rec.lastRun}));
ok('30 天内每两周执行 3 次', chase0-w.eval("chaseItem().mv")===1500,
   `Chase $${chase0} → $${w.eval("chaseItem().mv")}（共转 $${chase0-w.eval("chaseItem().mv")}）`);
ok('列表显示频率', /每两周/.test(d.getElementById('transfer-body').textContent));
// 再跑一次不应重复
const c2=w.eval("chaseItem().mv");
w.runRecurring();
ok('重复运行不重复转账', w.eval("chaseItem().mv")===c2);

console.log('\n【5】旧格式定期转账迁移');
w.localStorage.setItem('recurringTransfers_v1', JSON.stringify([
  {src:'chase',tgt:'broker',amt:300,startYm:2026*12+6,lastRun:2026*12+6}]));
w.runRecurring();
const mig=JSON.parse(w.localStorage.getItem('recurringTransfers_v1'))[0];
ok('旧记录已迁移成日期格式', !!mig.start && mig.startYm===undefined && mig.freq==='monthly',
   JSON.stringify(mig));

console.log('\n【防连点】');
{
  let pending; const hold=new Promise(r=>pending=r);
  let addCalls=0;
  w.fetch=async(u)=>{const url=new URL(u);const a=url.searchParams.get('action');
    if(!a) return {json:async()=>({status:'success',stock:[{symbol:'Cash',category:'Gold.coin.cash',price:1326}],
      ledger:[],expense:{},monthly:{},savings:[],notices:[],bond:[],cash:{balance:0,hasAnchor:false}})};
    if(a==='addBond'){ addCalls++; await hold; }
    return {json:async()=>({status:'success'})};};
  w.openDetail('bond'); w.showBondForm();
  d.getElementById('b-start').value='2026-08-01';
  d.getElementById('b-mv').value='10000';
  d.getElementById('b-rate').value='4.25';
  const btn=d.querySelector('#bondAddWrap .save');
  w.saveBond({target:btn});
  await wait(30);
  ok('保存中按钮变灰并显示状态', btn.disabled===true && /保存中/.test(btn.textContent), btn.textContent);
  w.saveBond({target:btn}); w.saveBond({target:btn});   // 连点两次
  await wait(30);
  ok('连点不会重复提交', addCalls===1, `发出了 ${addCalls} 次请求`);
  pending(); await wait(300);
  ok('完成后按钮恢复', btn.disabled===false && /保存/.test(btn.textContent), btn.textContent);
}
