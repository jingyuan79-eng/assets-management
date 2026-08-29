import { JSDOM } from 'jsdom'; import fs from 'fs';
const html=fs.readFileSync('index.html','utf8');
const ok=(l,c,e='')=>console.log((c?'✅':'❌'),l,e);
const wait=(ms=350)=>new Promise(r=>setTimeout(r,ms));

let savings=[
 {id:'sv_a',name:'CD-7597',type:'CD',balance:21335,rate:0.04,lastPost:'2026-08-01',nextUpdate:'2026-09-01',maturity:'2027-06-30',status:'active'},
 {id:'sv_b',name:'CD-9185',type:'CD',balance:10274,rate:0.041,lastPost:'2026-08-01',nextUpdate:'2026-09-01',maturity:'2027-02-05',status:'active'},
 {id:'sv_c',name:'OS-9250',type:'OS',balance:30057,rate:0.034,lastPost:'2026-08-01',nextUpdate:'2026-09-01',maturity:'',status:'active'},
 {id:'sv_x',name:'CD-OLD',type:'CD',balance:5000,rate:0.05,lastPost:'2026-07-01',nextUpdate:'',maturity:'2026-07-15',status:'closed'},
];
let notices=[{id:'nt_1',date:'2026-07-15',text:'CD-OLD 已于 2026-07-15 到期，本息 $5,020 已转入 OS-9250'}];
const calls=[];
const dom=new JSDOM(html,{runScripts:'dangerously',url:'https://x.test/',
  beforeParse(w){ w.localStorage.clear();
    w.fetch=async(u)=>{const url=new URL(u);const a=url.searchParams.get('action');
      if(!a) return {json:async()=>({status:'success',
        stock:[{symbol:'Cash',category:'Gold.coin.cash',price:1326}],
        ledger:[],expense:{},monthly:{},
        savings:JSON.parse(JSON.stringify(savings)),notices:JSON.parse(JSON.stringify(notices))})};
      const p=Object.fromEntries(url.searchParams); calls.push(p);
      if(a==='dismissNotice'){ notices=notices.filter(n=>n.id!==p.id); }
      if(a==='adjustSaving'){ const r=savings.find(x=>x.id===p.id); if(r) r.balance+=+p.delta; }
      if(a==='addSaving'){ savings.push({id:'sv_new',name:p.name,type:p.type,balance:+p.balance,
        rate:+p.rate/100,lastPost:'2026-08-23',nextUpdate:'2026-09-01',maturity:p.maturity,status:'active'});
        if(p.srcId){ const r=savings.find(x=>x.id===p.srcId); if(r) r.balance-=+p.balance; } }
      if(a==='deleteSaving'){ savings=savings.filter(x=>x.id!==p.id); }
      if(a==='updateSaving'){ const r=savings.find(x=>x.id===p.id);
        if(r){ if(p.name)r.name=p.name; if(p.balance)r.balance=+p.balance;
               if(p.rate)r.rate=+p.rate/100; if(p.maturity!=null)r.maturity=p.maturity; } }
      // 真后端在写操作后会带回最新列表，假后端也要照做，否则占位行的临时 id 不会被替换
      const snap=()=>JSON.parse(JSON.stringify(savings));
      if(['addSaving','updateSaving','deleteSaving','adjustSaving'].includes(a))
        return {json:async()=>({status:'success',savings:snap()})};
      return {json:async()=>({status:'success'})};};
    w.alert=m=>console.log('   [alert]',m); w.confirm=()=>true; w.scrollTo=()=>{};
    w.Element.prototype.scrollIntoView=function(){};
  }});
const w=dom.window,d=w.document;
await wait();

console.log('— 储蓄以 Sheet 为准 —');
ok('closed 的账户不显示', !w.eval("DATA.cd.groups[0].holdings.some(x=>x.sym==='CD-OLD')"),
   w.eval("DATA.cd.groups[0].holdings.map(x=>x.sym).join(',')"));
ok('储蓄总额=三个活跃账户', w.eval("classTotal.cd")===21335+10274+30057, '$'+w.eval("classTotal.cd"));
ok('小程序不再本地算息', w.eval("typeof rollCd")==='undefined', 'rollCd='+w.eval("typeof rollCd"));
w.openDetail('cd');
const body=d.getElementById('d-body').textContent;
ok('显示利率，不显示下次入息日', /4%/.test(body) && !/下次入息/.test(body));
ok('marcusItem 按 type 识别', w.eval("marcusItem().sym")==='OS-9250');

console.log('\n— 通知条 —');
ok('显示到期通知', /CD-OLD/.test(d.getElementById('noticeSlot').textContent));
w.dismissNotice('nt_1'); await wait();
ok('点 × 后消失', d.getElementById('noticeSlot').innerHTML==='');
ok('并写回 Sheet 标记已读', calls.some(c=>c.action==='dismissNotice'&&c.id==='nt_1'));

console.log('\n— 新增 / 编辑 / 删除 —');
w.openDetail('cd'); w.showCdForm();
d.getElementById('f-name').value='CD-NEW'; d.getElementById('f-mv').value='10000';
d.getElementById('f-rate').value='4.5'; d.getElementById('f-mat').value='2027-08-23';
d.getElementById('f-src').value='marcus';
w.saveCd(); await wait(500);
ok('新增走 addSaving', calls.some(c=>c.action==='addSaving'&&c.name==='CD-NEW'));
ok('从 OS 扣款', savings.find(x=>x.id==='sv_c').balance===30057-10000,
   '$'+savings.find(x=>x.id==='sv_c').balance);
ok('新账户已出现在列表', w.eval("DATA.cd.groups[0].holdings.some(x=>x.sym==='CD-NEW')"));

const idx=w.eval("DATA.cd.groups[0].holdings.findIndex(x=>x.sym==='CD-7597')");
w.editCd(idx);
d.getElementById('f-rate').value='3.8';
w.saveCd(); await wait(500);
const upd=calls.filter(c=>c.action==='updateSaving').pop();
ok('只改利率时不传 balance', upd && upd.rate==='3.8' && upd.balance===undefined,
   JSON.stringify({rate:upd&&upd.rate,balance:upd&&upd.balance}));
w.editCd(w.eval("DATA.cd.groups[0].holdings.findIndex(x=>x.sym==='CD-7597')"));
d.getElementById('f-mv').value='22000';
w.saveCd(); await wait(500);
const upd2=calls.filter(c=>c.action==='updateSaving').pop();
ok('改余额时传 balance（手动为准）', upd2 && upd2.balance==='22000', 'balance='+(upd2&&upd2.balance));

const n0=savings.length;
w.deleteCd(w.eval("DATA.cd.groups[0].holdings.findIndex(x=>x.sym==='CD-NEW')")); await wait(500);
ok('删除走 deleteSaving', savings.length===n0-1, `${n0} → ${savings.length}`);

console.log('\n— 转账页 Marcus 走 Sheet —');
const osBefore=savings.find(x=>x.id==='sv_c').balance;
w.doTransfer('marcus','broker',2000); await wait(500);
ok('转出调 adjustSaving', calls.some(c=>c.action==='adjustSaving'&&c.delta==='-2000'));
ok('OS 余额已减', savings.find(x=>x.id==='sv_c').balance===osBefore-2000,
   `$${osBefore} → $${savings.find(x=>x.id==='sv_c').balance}`);
