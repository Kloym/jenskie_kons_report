/* Run: node tests/test_dashboard_ui.cjs
   Optional equivalence check: node tests/test_dashboard_ui.cjs fixture.json old/dashboard.js
   The default fixture is synthetic and contains no patient or employee information. */
'use strict';
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const assets = path.join(__dirname, '..', 'dashboard_assets');
const UI = require(path.join(assets, 'dashboard_ui.js'));
let checks = 0;
function check(name, fn) { fn(); checks++; console.log('PASS', name); }
const synthetic = {
  meta: {start:'2026-09-01',end:'2026-10-24',dataEnd:'2026-10-23',today:'2026-10-24'},
  departments:['Test clinic'],doctors:['Test doctor'],
  services:[{code:'001983',name:'A'},{code:'001984',name:'B'},{code:'001441',name:'C'}],
  targets:['001983','001984'], records:[]
};
for (let d = new Date('2026-09-01T12:00:00Z'); d <= new Date('2026-10-23T12:00:00Z'); d.setUTCDate(d.getUTCDate()+1)) {
  const day=d.toISOString().slice(0,10);
  synthetic.records.push([day,0,0,2,1,500,1]); // Activity remains on days without target services.
  const value=day==='2026-10-23'?1:10;
  synthetic.records.push([day,0,0,0,value,value*200,1]);
}
function load(filename, data) {
  const fields=Object.fromEntries(Object.entries({dateFrom:data.meta.start,dateTo:data.meta.end,department:'',doctor:'',service:'',threshold:'0.4'}).map(([id,value])=>[id,{value}]));
  fields['report-data']={textContent:JSON.stringify(data)};
  const context={window:{DashboardUI:UI}, document:{getElementById:id=>fields[id]}, Intl, Date};
  vm.createContext(context);
  const source=fs.readFileSync(filename,'utf8');
  assert.equal((source.match(/^initialize\(\);$/gm)||[]).length,1);
  vm.runInContext(source.replace(/^initialize\(\);$/m,''),context,{filename});
  return {api:context.window.Dashboard, fields};
}
function serialize(value) {
  return JSON.stringify(value,(_,v)=>v?.constructor?.name==='Map'?[...v]:v?.constructor?.name==='Set'?[...v]:v);
}
const currentFile=path.join(assets,'dashboard.js');
const local=load(currentFile,synthetic);
check('last four active same weekdays form the baseline',()=>{
  const s=local.api.aggregate().signals.find(s=>s.day==='2026-10-23');
  assert.equal(s.baseline,10);assert.equal(s.current,1);assert.equal(s.drop,.9);
  assert.deepEqual(Array.from(s.previous),['2026-09-25','2026-10-02','2026-10-09','2026-10-16']);
});
check('unchanged scope reuses aggregation; threshold invalidates it',()=>{
  const a=local.api.aggregate();assert.equal(a,local.api.aggregate());
  local.fields.threshold.value='0.95';assert.notEqual(local.api.aggregate(),a);
  assert.equal(local.api.aggregate().signals.length,0);
});
check('focus and list navigation never alter report aggregation',()=>{
  const a=local.api.aggregate();local.api.state.focus={doctorId:0,day:'2026-10-23'};
  local.api.state.heatPage=1;local.api.state.signalLimit=80;
  assert.equal(local.api.aggregate(),a);
});
check('missing export dates are not represented as zero',()=>{
  assert.equal(local.api.chartValue(local.api.aggregate(),'2026-10-24','quantity'),null);
});
check('every light and dark heatmap colour has >= 4.5:1 contrast',()=>{
  for(const dark of [false,true]) for(let value=0;value<=10000;value++)assert.ok(UI.heatColor(value,10000,dark).contrast>=4.5);
});
check('invalid dates are rejected before the cache is read',()=>{
  local.fields.dateFrom.value='2027-01-01';assert.throws(()=>local.api.aggregate());
});
check('motion retargets from the visible position and cancels on keyboard/reduced motion',()=>{
  const media={matches:false,addEventListener(){}},events={},doc={documentElement:{dataset:{}},addEventListener:(name,fn)=>events[name]=fn};
  const context={window:{document:doc,matchMedia:()=>media},document:doc,getComputedStyle:()=>({transform:'matrix(1,0,0,1,42,0)'})};
  vm.createContext(context);vm.runInContext(fs.readFileSync(path.join(assets,'dashboard_ui.js'),'utf8'),context);
  const calls=[],element={style:{},animate:(frames,options)=>{
    const a={frames,options,cancelled:false,cancel(){this.cancelled=true;},finished:new Promise(()=>{})};calls.push(a);return a;
  }};
  const move=context.window.DashboardUI.transform;
  move(element,'translateX(100px)');move(element,'translateX(200px)');
  assert.equal(calls.length,2);assert.equal(calls[0].cancelled,true);
  assert.equal(calls[1].frames[0].transform,'matrix(1,0,0,1,42,0)');
  events.keydown();move(element,'translateX(300px)');
  assert.equal(calls.length,2);assert.equal(calls[1].cancelled,true);assert.equal(element.style.transform,'translateX(300px)');
  events.pointerdown();media.matches=true;move(element,'translateX(400px)');assert.equal(calls.length,2);
});
function clinicIndex(filename,data) {
  const context={window:{},document:{},Intl,Date};vm.createContext(context);
  vm.runInContext(fs.readFileSync(filename,'utf8'),context);
  return context.window.ClinicPace.makeIndex(data);
}
check('clinic index sums quantities for both codes within the plan month',()=>{
  const data={...synthetic,annualPlan:{asOf:'2026-10-23',departments:[{name:'Test clinic'}]}};
  data.records=[...synthetic.records,['2026-10-23',0,0,1,3,600,1]];
  const result=clinicIndex(path.join(assets,'clinic_pace.js'),data);
  assert.equal(result[0].quantities['2026-10-23'],4);
  assert.equal(result[0].quantities['2026-09-30'],undefined);
});
if(process.argv[2]&&process.argv[3]) {
  const data=JSON.parse(fs.readFileSync(process.argv[2],'utf8'));
  const before=load(path.resolve(process.argv[3]),data), after=load(currentFile,data);
  const cases=[{}, {tab:'all'}, {threshold:'0.25'}, {threshold:'0.6'}, {dateFrom:data.meta.dataEnd,dateTo:data.meta.dataEnd}];
  data.departments.forEach((_,i)=>cases.push({department:String(i)}, {department:String(i),tab:'all'}));
  data.services.slice(0,6).forEach(s=>cases.push({tab:'all',service:s.code}));
  data.doctors.slice(0,6).forEach(name=>cases.push({doctor:name}));
  for(const [i,options] of cases.entries()) check('baseline equivalence '+(i+1),()=>{
    for(const fixture of [before,after]) {
      for(const [key,value] of Object.entries({dateFrom:data.meta.start,dateTo:data.meta.end,department:'',doctor:'',service:'',threshold:'0.4',...options}))if(key!=='tab')fixture.fields[key].value=value;
      fixture.api.state.tab=options.tab||'target';
    }
    assert.equal(serialize(after.api.aggregate()),serialize(before.api.aggregate()));
  });
  check('real clinic index is unchanged',()=>{
    const input={...data,annualPlan:{asOf:data.meta.dataEnd,departments:data.departments.map(name=>({name}))}};
    assert.equal(serialize(clinicIndex(path.join(assets,'clinic_pace.js'),input)),serialize(clinicIndex(path.join(path.dirname(process.argv[3]),'clinic_pace.js'),input)));
  });
}
console.log(`${checks} checks passed.`);
