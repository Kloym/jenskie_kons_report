(() => {
"use strict";
const D = JSON.parse(document.getElementById("report-data").textContent);
const $ = id => document.getElementById(id);
const E = value => String(value ?? "").replace(/[&<>"']/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
const num = new Intl.NumberFormat("ru-RU", {maximumFractionDigits: 1});
const money = new Intl.NumberFormat("ru-RU", {minimumFractionDigits: 2,maximumFractionDigits: 2});
const formatQ = value => num.format(value);
const formatM = value => money.format(value / 100);
const pct = value => new Intl.NumberFormat("ru-RU", {maximumFractionDigits: 1}).format(value) + "%";
const dateObj = value => new Date(value + "T12:00:00Z");
const dayText = value => dateObj(value).toLocaleDateString("ru-RU",{day:"numeric",month:"short",timeZone:"UTC"});
const fullDate = value => dateObj(value).toLocaleDateString("ru-RU",{day:"numeric",month:"long",year:"numeric",timeZone:"UTC"});
const weekday = value => dateObj(value).getUTCDay();
const weekdayName = value => ["Вс","Пн","Вт","Ср","Чт","Пт","Сб"][weekday(value)];
const calendarDate = value => weekdayName(value)+", "+fullDate(value);
const sum = (rows, index) => rows.reduce((total,r) => total + r[index], 0);
const median = values => {const a=[...values].sort((x,y)=>x-y);const i=Math.floor(a.length/2);return a.length%2?a[i]:(a[i-1]+a[i])/2;};
const state = {tab:"target",detail:"doctors",sort:"quantity",desc:true,limit:25,heatPage:0,signalLimit:40,signalDoctor:null,signalDay:null,focus:null,view:null};
const UI=window.DashboardUI;
let aggregateCache=null,pickerCache=new Map();
const doctorKeys=D.doctors.map(value=>String(value).toLocaleLowerCase("ru-RU").replace(/ё/g,"е").trim());
const datesBetween = (start,end) => {
  const dates=[];const cursor=dateObj(start);
  while(cursor.toISOString().slice(0,10)<=end && dates.length<372){dates.push(cursor.toISOString().slice(0,10));cursor.setUTCDate(cursor.getUTCDate()+1);}
  return dates;
};
const normalize = value => String(value).toLocaleLowerCase("ru-RU").replace(/ё/g,"е").trim();
function shorten(name){const p=name.trim().split(/\s+/);return p.length>1?p[0]+" "+p.slice(1).map(v=>v[0]+".").join(""):name;}
function bucket(){return {quantity:0,amount:0,rows:0,days:new Set(),doctors:new Set(),codes:new Set()};}
function add(b,r){b.quantity+=r[4];b.amount+=r[5];b.rows+=r[6];if(r[4]>0){b.days.add(r[0]);b.doctors.add(r[2]);b.codes.add(r[3]);}}
function aggregate() {
  const start=$("dateFrom").value,end=$("dateTo").value,dep=$("department").value;
  const doctor=normalize($("doctor").value),search=normalize($("service").value);
  if(!start||!end||start>end||start<D.meta.start||end>D.meta.end){
    throw new Error("Выберите корректные даты внутри периода "+dayText(D.meta.start)+" – "+dayText(D.meta.end));
  }
  const cacheKey=JSON.stringify([start,end,dep,doctor,search,state.tab,$("threshold").value]);
  if(aggregateCache?.key===cacheKey)return aggregateCache.value;
  const dates=datesBetween(start,end);
  const scoped=D.records.filter(r=>r[0]>=start&&r[0]<=end&&(dep===""||r[1]===Number(dep))&&(!doctor||doctorKeys[r[2]]===doctor));
  const selected=scoped.filter(r=>state.tab==="target"?D.targets.includes(D.services[r[3]].code):
    !search||D.services[r[3]].code===search);
  const total=bucket(),doctors=new Map(),services=new Map(),days=new Map(dates.map(day=>[day,{...bucket(),day,activity:0}]));
  const active=new Map();
  for(const r of scoped){
    const key=r[2]+"|"+r[0];active.set(key,(active.get(key)||0)+r[4]);
    days.get(r[0]).activity+=r[4];
    if(!doctors.has(r[2]))doctors.set(r[2],{...bucket(),id:r[2],name:D.doctors[r[2]],daily:new Map(),departments:new Set(),signals:[]});
    doctors.get(r[2]).departments.add(D.departments[r[1]]);
  }
  for(const r of selected){
    add(total,r);add(days.get(r[0]),r);
    const b=doctors.get(r[2]);add(b,r);
    if(!b.daily.has(r[0]))b.daily.set(r[0],bucket());add(b.daily.get(r[0]),r);
    if(!services.has(r[3]))services.set(r[3],{...bucket(),id:r[3],...D.services[r[3]]});
    add(services.get(r[3]),r);
  }
  const threshold=Number($("threshold").value);
  function signalFor(series,activity,id=null){
    const signals=[],history=Array.from({length:7},()=>[]);
    for(let i=0;i<dates.length;i++){
      const day=dates[i];if(!activity(day))continue;
      const previous=[...history[weekday(day)]];
      history[weekday(day)].push(day);if(history[weekday(day)].length>4)history[weekday(day)].shift();
      if(day>=D.meta.today||previous.length<2)continue;
      const current=series(day);
      const bq=median(previous.map(x=>series(x).quantity)),ba=median(previous.map(x=>series(x).amount));
      const quantityDrop=bq>0 ? 1-current.quantity/bq : 0;
      const amountDrop=ba>0 ? 1-current.amount/ba : 0;
      if(quantityDrop>threshold+1e-9||amountDrop>threshold+1e-9){
        const metric=quantityDrop>=amountDrop?"quantity":"amount";
        signals.push({id,day,metric,baseline:metric==="quantity"?bq:ba,current:current[metric],
          drop:Math.max(quantityDrop,amountDrop),previous,quantityDrop,amountDrop});
      }
    }
    return signals;
  }
  const signals=[];
  for(const b of doctors.values()){
    b.signals=signalFor(day=>b.daily.get(day)||bucket(),day=>(active.get(b.id+"|"+day)||0)>0,b.id);
    signals.push(...b.signals);
  }
  const dailySignals=signalFor(day=>days.get(day),day=>days.get(day).activity>0);
  signals.sort((a,b)=>b.day.localeCompare(a.day)||b.drop-a.drop);
  const value={dates,start,end,scoped,selected,total,doctors,services,days,active,signals,dailySignals};
  aggregateCache={key:cacheKey,value};return value;
}
function niceStep(max,integer){
  const raw=max/4||1;const pow=10**Math.floor(Math.log10(raw));
  let step=[1,2,2.5,5,10].map(n=>n*pow).find(n=>n>=raw)||pow*10;
  return integer?Math.max(1,Math.ceil(step)):step;
}
// A zero describes services found in this snapshot, not the doctor's working schedule.
function chartValue(view,day,metric){return day<=D.meta.dataEnd?view.days.get(day)[metric]:null;}
function chart(id,view,metric,type){
  const el=$(id), width=Math.max(260,Math.round(el.getBoundingClientRect().width)),height=190;
  const left=metric==="amount"?56:40,right=12,top=12,bottom=30;
  const w=width-left-right,h=height-top-bottom,n=view.dates.length;
  const vals=view.dates.map(day=>chartValue(view,day,metric));
  const max=Math.max(0,...vals.filter(x=>x!==null)),step=niceStep(max,metric==="quantity"),ymax=Math.max(step,Math.ceil((max*1.08)/step)*step);
  const x=i=>left+(i+.5)*w/n,y=value=>top+h-value/ymax*h;
  const low=new Set(view.dailySignals.filter(s=>metric==="quantity"?s.quantityDrop>Number($("threshold").value):s.amountDrop>Number($("threshold").value)).map(s=>s.day));
  let body='<rect class="day-focus" data-focus-cursor x="0" y="'+top+'" width="'+w/n+'" height="'+h+'" visibility="hidden"/>';
  const tick=v=>metric==="quantity"?formatQ(v):(v>=100000?formatQ(v/100000)+" тыс.":formatQ(v/100));
  const count=Math.ceil(ymax/step);
  for(let i=0;i<=count;i++){const val=i*step,yy=y(val);body+='<line class="gridline" x1="'+left+'" x2="'+(width-right)+'" y1="'+yy+'" y2="'+yy+'"/><text x="'+(left-8)+'" y="'+(yy+4)+'" text-anchor="end">'+E(tick(val))+'</text>';}
  const labelStep=Math.max(1,Math.ceil(n/(width<420?4:7)));
  let points=[];
  function path(){if(points.length===1){body+='<circle cx="'+points[0][0]+'" cy="'+points[0][1]+'" r="3" fill="var(--accent)"/>';}else if(points.length>1){body+='<polyline class="mainline" points="'+points.map(p=>p.join(",")).join(" ")+'"/>';}points=[];}
  for(let i=0;i<n;i++){
    const value=vals[i],day=view.dates[i],xx=x(i);
    if(i%labelStep===0)body+='<text x="'+xx+'" y="'+(height-8)+'" text-anchor="middle">'+E(dayText(day))+'</text>';
    if(value===null){if(type==="line")path();body+='<text x="'+xx+'" y="'+(y(0)-8)+'" text-anchor="middle">·</text>';}
    else if(type==="bar"){
      const bw=Math.max(2,Math.min(22,w/n*.68));
      body+='<rect class="bar '+(low.has(day)?"low":"")+'" x="'+(xx-bw/2)+'" y="'+y(value)+'" width="'+bw+'" height="'+(y(0)-y(value))+'" rx="4"/>';
    }else{points.push([xx,y(value)]);if(low.has(day))body+='<circle cx="'+xx+'" cy="'+y(value)+'" r="4" fill="var(--negative)"/>';}
  }
  if(type==="line")path();
  for(let i=0;i<n;i++){
    const day=view.dates[i],b=view.days.get(day);
    const tip=calendarDate(day)+"\n"+(day>D.meta.dataEnd?"За пределами выгрузки: данных нет":formatQ(b.quantity)+" услуг · "+formatM(b.amount)+" ₽"+(!b.activity?"\nУслуг в выгрузке нет; график работы неизвестен":""))+(day===D.meta.today?"\nТекущий день: данные могут быть неполными":"");
    body+='<rect class="hit" tabindex="'+(i===0?'0':'-1')+'" role="button" data-chart-day="'+day+'" aria-label="'+E(tip)+'" data-tip="'+E(tip)+'" x="'+(left+i*w/n)+'" y="'+top+'" width="'+w/n+'" height="'+h+'"/>';
  }
  el.innerHTML='<svg viewBox="0 0 '+width+' '+height+'" role="group" aria-label="'+(metric==="quantity"?"Количество услуг":"Сумма услуг")+' по дням"><title>Дневная динамика. Точные значения доступны в таблице «По дням».</title>'+body+'</svg>';
  updateChartFocus(false);
}
function updateChartFocus(animate=true){
  for(const id of ['quantityChart','amountChart']){
    const chart=$(id),cursor=chart.querySelector('[data-focus-cursor]');if(!cursor)continue;
    const mark=chart.querySelector('[data-chart-day="'+state.focus?.day+'"]');
    cursor.setAttribute('visibility',mark?'visible':'hidden');if(mark)UI.transform(cursor,'translateX('+mark.getAttribute('x')+'px)',animate);
  }
}
function focusDay(day,doctorId=null,animate=true){
  state.focus={day,doctorId};$('focusContext').hidden=false;
  $('focusLabel').textContent=(doctorId===null?'Все врачи':D.doctors[doctorId])+' · '+fullDate(day);
  updateChartFocus(animate);
  for(const cell of $('heatmap').querySelectorAll('[data-heat-day]')){
    const selected=cell.dataset.heatDay===day&&(doctorId===null||Number(cell.dataset.heatDoctor)===doctorId);
    cell.classList.toggle('is-focused',selected);cell.setAttribute('aria-selected',selected);
  }
}
function renderKpis(v){
  $("quantityKpi").textContent=formatQ(v.total.quantity);
  $("amountKpi").textContent=formatM(v.total.amount);
  $("doctorsKpi").textContent=formatQ(v.total.doctors.size);
  $("dailyKpi").textContent=formatQ(v.total.days.size?v.total.quantity/v.total.days.size:0);
  $("quantityNote").textContent=state.tab==="target"?"001983 + 001984":"Различных кодов: "+v.total.codes.size;
  $("doctorNote").textContent="Из "+v.doctors.size+" врачей с записями в периоде";
  $("dailyNote").textContent="Дней с выбранными услугами: "+v.total.days.size;
  const active=[...v.days.values()].filter(b=>b.activity>0);
  const best=active.reduce((a,b)=>!a||b.quantity>a.quantity?b:a,null);
  $("quantityChartNote").textContent=best?"Максимум: "+formatQ(best.quantity)+" услуг · "+dayText(best.day):"По текущему выбору услуг нет";
  const bestM=active.reduce((a,b)=>!a||b.amount>a.amount?b:a,null);
  $("amountChartNote").textContent=bestM?"Максимум: "+formatM(bestM.amount)+" ₽ · "+dayText(bestM.day):"По текущему выбору услуг нет";
  $("amountCaption").textContent="Средняя стоимость услуги: "+(v.total.quantity?formatM(v.total.amount/v.total.quantity)+" ₽":"нет услуг")+" · Сумма рассчитана по строкам выгрузки";
  $("scopeText").textContent=($("doctor").value.trim()||"Все врачи")+" · "+($("department").value===""?D.meta.scope:D.departments[Number($("department").value)]);
}
function heatEntries(v){
  const metric=$("heatMetric").value;
  return [...v.doctors.values()].sort((a,b)=>b[metric]-a[metric]||a.name.localeCompare(b.name,"ru"));
}
function renderHeat(v){
  const metric=$("heatMetric").value,entries=heatEntries(v),size=40;
  state.heatPage=Math.max(0,Math.min(state.heatPage,Math.ceil(entries.length/size)-1));
  const shown=entries.slice(state.heatPage*size,(state.heatPage+1)*size);
  let maximum=1;for(const d of entries)for(const b of d.daily.values())maximum=Math.max(maximum,b[metric]);
  const dark=document.documentElement.dataset.theme==="dark";
  const activeKey=$("heatmap").contains(document.activeElement)?document.activeElement.dataset.heatKey:null;
  let html='<thead><tr><th scope="col">Врач</th>'+v.dates.map(day=>'<th scope="col" class="calendar-day '+([0,6].includes(weekday(day))?"weekend ":"")+(weekday(day)===1?"week-start":"")+'" aria-label="'+E(calendarDate(day))+'" title="'+E(calendarDate(day))+'"><b>'+Number(day.slice(-2))+'</b>'+weekdayName(day)+'</th>').join("")+'<th scope="col">Итого</th></tr></thead><tbody>';
  for(const [index,d] of shown.entries()){
    const flags=new Map(d.signals.map(s=>[s.day,s]));
    html+='<tr><td><button class="doctor-link" data-roving data-heat-key="'+d.id+'|name" tabindex="'+(index===0?'0':'-1')+'" data-doctor="'+d.id+'" title="'+E(d.name)+'">'+E(shorten(d.name))+'</button></td>';
    for(const day of v.dates){
      const b=d.daily.get(day)||bucket(),activity=(v.active.get(d.id+"|"+day)||0)>0,flag=flags.get(day);
      const color=UI.heatColor(b[metric],maximum,dark);
      const amountText=metric==="quantity"?formatQ(b.quantity):(b.amount>=100000?formatQ(b.amount/100000)+"к":formatQ(b.amount/100));
      const tip=d.name+"\n"+calendarDate(day)+"\n"+(activity?formatQ(b.quantity)+" услуг · "+formatM(b.amount)+" ₽":"Нет услуг врача в выгрузке; график неизвестен")+(flag?"\nСнижение "+pct(flag.drop*100)+" относительно сопоставимых дней":"");
      const style=activity?'background:'+color.background+';color:'+color.color:'';
      html+='<td role="gridcell" tabindex="-1" data-roving data-heat-key="'+d.id+'|'+day+'" data-heat-doctor="'+d.id+'" data-heat-day="'+day+'" aria-selected="false" aria-label="'+E(tip)+'" class="cell '+(!activity?"missing ":"")+(flag?"low-cell ":"")+(weekday(day)===1?"week-start":"")+'" style="'+style+'" data-tip="'+E(tip)+'">'+(activity?(flag?"↓":"")+E(amountText):"·")+'</td>';
    }
    html+='<td style="text-align:right;padding-left:10px;white-space:nowrap;font-size:12px">'+(metric==="quantity"?formatQ(d.quantity):formatM(d.amount))+'</td></tr>';
  }
  if(!entries.length)html+='<tr><td colspan="'+(v.dates.length+2)+'" class="empty">Нет врачей по выбранным фильтрам.</td></tr>';
  $("heatmap").innerHTML=html+'</tbody>';
  $("heatCount").textContent=entries.length?'Врачи '+(state.heatPage*size+1)+'–'+Math.min((state.heatPage+1)*size,entries.length)+' из '+entries.length:'Нет врачей';
  $("heatPrev").disabled=state.heatPage===0;$("heatNext").disabled=(state.heatPage+1)*size>=entries.length;
  if(activeKey){const target=[...$("heatmap").querySelectorAll("[data-heat-key]")].find(n=>n.dataset.heatKey===activeKey);target?.focus({preventScroll:true});}
  if(state.focus)focusDay(state.focus.day,state.focus.doctorId,false);
}
function signalMarkup(s,index,full=false){
  const value=s.metric==="quantity"?formatQ:formatM,unit=s.metric==="quantity"?"усл.":"₽";
  if(full)return '<button class="signal-row" data-signal-index="'+index+'"><span><strong>'+E(D.doctors[s.id])+'</strong><small>'+E([...state.view.doctors.get(s.id).departments].join(" · "))+'</small></span><span>'+E(dayText(s.day))+'<small>'+(s.metric==="quantity"?"Объём":"Сумма")+'</small></span><span>'+value(s.current)+' '+unit+'<small>База: '+value(s.baseline)+' '+unit+'</small></span><span>−'+pct(s.drop*100)+'</span></button>';
  return '<div class="signal"><div class="signal-title"><span>'+E(dayText(s.day))+' · '+(s.metric==="quantity"?"Объём":"Сумма")+'</span><strong>−'+pct(s.drop*100)+'</strong></div><button data-signal-index="'+index+'">'+E(shorten(D.doctors[s.id]))+'</button><p>'+value(s.current)+' '+unit+' при базе '+value(s.baseline)+' '+unit+'</p><p>База: '+s.previous.map(dayText).join(", ")+'</p></div>';
}
function renderSignals(v){
  $("signalCount").textContent=v.signals.length;$("jumpSignalCount").textContent=v.signals.length;
  $("allSignals").textContent="Все сигналы · "+v.signals.length;$("allSignals").disabled=!v.signals.length;
  $("signals").innerHTML=v.signals.length?v.signals.slice(0,5).map((s,i)=>signalMarkup(s,i)).join(""):'<div class="empty">Нет выявленных снижений.<p class="fine">Нужны минимум два предыдущих таких же дня недели с услугами врача.</p></div>';
  if(!$("signalQueue").hidden)renderSignalQueue();
}
function renderSignalQueue(){
  const query=normalize($("signalSearch").value),v=state.view;if(!v)return;
  const rows=v.signals.map((s,index)=>({s,index})).filter(({s})=>
    (state.signalDoctor===null||s.id===state.signalDoctor)&&(!state.signalDay||s.day===state.signalDay)&&(!query||normalize(D.doctors[s.id]).includes(query)));
  rows.sort((a,b)=>$("signalSort").value==="drop"?b.s.drop-a.s.drop||b.s.day.localeCompare(a.s.day):b.s.day.localeCompare(a.s.day)||b.s.drop-a.s.drop);
  $("signalQueueCount").textContent="Показано "+Math.min(rows.length,state.signalLimit)+" из "+rows.length;
  $("signalQueueTitle").textContent=state.signalDoctor===null?(state.signalDay?"Сигналы · "+dayText(state.signalDay):"Все сигналы"):"Сигналы · "+shorten(D.doctors[state.signalDoctor]);
  $("signalQueueRows").innerHTML=rows.slice(0,state.signalLimit).map(({s,index})=>signalMarkup(s,index,true)).join("")||'<p class="empty">Сигналов по этому выбору нет. Измените поиск или откройте «Все сигналы».</p>';
  $("moreSignals").hidden=rows.length<=state.signalLimit;
}
function showSignalQueue(doctorId=null,day=null){
  state.signalDoctor=doctorId;state.signalDay=day;state.signalLimit=40;$("signalSearch").value="";
  $("signalQueue").hidden=false;$("allSignals").setAttribute("aria-expanded","true");renderSignalQueue();
  $("signalQueueTitle").focus({preventScroll:true});$("signalQueue").scrollIntoView({block:"start"});
}
function inspectDay(doctorId,day,trigger,keyboard=false,signal=null){
  const index=heatEntries(state.view).findIndex(d=>d.id===doctorId);
  if(index<0)return;state.heatPage=Math.floor(index/40);renderHeat(state.view);focusDay(day,doctorId,!keyboard);
  openDoctorServices(doctorId,trigger,keyboard,day,signal);
}

function tableEntries(v){
  let entries;
  if(state.detail==="doctors")entries=[...v.doctors.values()].map(b=>({...b,label:b.name,subtitle:[...b.departments].join(" · "),extra:b.days.size,signals:b.signals.length}));
  else if(state.detail==="services")entries=[...v.services.values()].map(b=>({...b,label:b.code,subtitle:b.name,extra:b.doctors.size,signals:0}));
  else entries=[...v.days.values()].map(b=>({...b,label:fullDate(b.day),subtitle:b.activity?"День с записями": "Нет сведений об активности",extra:b.doctors.size,signals:v.dailySignals.filter(s=>s.day===b.day).length}));
  const query=normalize($("tableSearch").value);
  entries=entries.filter(b=>normalize(b.label+" "+b.subtitle).includes(query));
  const field=state.sort;
  entries.sort((a,b)=>{
    let compare=field==="label"?(state.detail==="days"?a.day.localeCompare(b.day):a.label.localeCompare(b.label,"ru")):a[field]-b[field];
    return state.desc?-compare:compare;
  });
  return entries;
}
function renderTable(v){
  const entries=tableEntries(v),shown=entries.slice(0,state.limit);
  const titles={doctors:["Результаты врачей","Врач","Дней с услугами"],services:["Структура услуг","Код и услуга","Врачей"],days:["Дневная детализация","Дата","Врачей"]},t=titles[state.detail];
  $("detailTitle").textContent=t[0];$("detailSubtitle").textContent=state.detail==="doctors"?"Фамилия открывает отчёт врача, количество — расшифровку услуг.":"Количество и суммы в текущем выборе";
  $("tableCount").textContent=entries.length?"Показано "+shown.length+" из "+entries.length:"Нет строк";
  const heads=[["label",t[1]],["quantity","Количество"],["amount","Сумма, ₽"],["extra",t[2]]];
  let html='<thead><tr>'+heads.map(([key,label])=>'<th scope="col" aria-sort="'+(state.sort===key?(state.desc?"descending":"ascending"):"none")+'"><button data-sort="'+key+'">'+E(label)+(state.sort===key?(state.desc?" ↓":" ↑"):"")+'</button></th>').join("")+'<th scope="col">Доля количества</th><th scope="col" aria-sort="'+(state.sort==='signals'?(state.desc?'descending':'ascending'):'none')+'"><button data-sort="signals">Сигналы'+(state.sort==='signals'?(state.desc?' ↓':' ↑'):'')+'</button></th></tr></thead><tbody>';
  for(const b of shown){
    const missing=state.detail==="days"&&!b.activity,share=v.total.quantity?b.quantity/v.total.quantity*100:0;
    const label=state.detail==="doctors"?'<button class="doctor-link" data-doctor="'+b.id+'">'+E(b.label)+'</button>':E(b.label);
    const quantity=state.detail==="doctors"?'<button class="quantity-detail" data-services-doctor="'+b.id+'" aria-haspopup="dialog" aria-label="'+E('Услуги врача '+b.label+': '+formatQ(b.quantity))+'" title="Открыть все услуги врача"><span>'+formatQ(b.quantity)+'</span><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M7 17 17 7M7 7h10v10"/></svg></button>':formatQ(b.quantity);
    html+='<tr><td>'+label+'<small>'+E(b.subtitle)+'</small>'+(state.detail==="services"?'<div class="rowbar" style="width:'+share+'%"></div>':"")+'</td><td>'+(missing?"·":quantity)+'</td><td>'+(missing?"·":formatM(b.amount))+'</td><td>'+(missing?"·":b.extra)+'</td><td>'+(missing?"·":pct(share))+'</td><td>'+(b.signals?'<button class="signal-label" '+(state.detail==='doctors'?'data-queue-doctor="'+b.id+'"':'data-queue-day="'+b.day+'"')+' aria-label="Открыть сигналы: '+E(b.label)+'">'+b.signals+' ↓</button>':'<span class="neutral">нет</span>')+'</td></tr>';
  }
  if(!shown.length)html+='<tr><td colspan="6" class="empty">По этому запросу ничего не найдено. Измените поиск или сбросьте фильтры.</td></tr>';
  const activeSort=$("detailTable").contains(document.activeElement)?document.activeElement.dataset.sort:null;
  $("detailTable").innerHTML=html+'</tbody>';
  if(activeSort)$("detailTable").querySelector('[data-sort="'+activeSort+'"]')?.focus({preventScroll:true});
  let empty=$("tableEmpty");if(!empty){empty=document.createElement("p");empty.id="tableEmpty";empty.className="table-empty";$("detailTable").closest(".table-scroll").after(empty);}
  empty.hidden=Boolean(shown.length);empty.textContent="По этому запросу ничего не найдено. Измените поиск или сбросьте фильтры.";
  $("detailTable").closest(".table-scroll").hidden=!shown.length;
  $("moreButton").hidden=entries.length<=state.limit;
}

function renderExtra(v){
  const weekdays=[1,2,3,4,5,6,0],labels=["Пн","Вт","Ср","Чт","Пт","Сб","Вс"];
  const groups=weekdays.map(day=>[...v.days.values()].filter(b=>weekday(b.day)===day&&b.activity>0));
  const averages=groups.map(rows=>rows.length?rows.reduce((a,b)=>a+b.quantity,0)/rows.length:0),max=Math.max(1,...averages);
  $("weekdayChart").innerHTML=averages.map((value,i)=>'<div class="weekday-row" title="'+groups[i].length+' дней с записями"><span>'+labels[i]+'</span><div class="weekday-track"><div class="weekday-fill" style="width:'+(value/max*100)+'%"></div></div><span>'+(groups[i].length?formatQ(value):'нет данных')+'</span></div>').join("");
  const sorted=[...v.services.values()].sort((a,b)=>b.quantity-a.quantity),items=sorted.slice(0,5);
  if(sorted.length>5)items.push({code:"Остальные",name:(sorted.length-5)+" кодов",quantity:sorted.slice(5).reduce((a,b)=>a+b.quantity,0),amount:sorted.slice(5).reduce((a,b)=>a+b.amount,0)});
  $("mixChart").innerHTML=items.length?items.map(b=>{
    const share=v.total.quantity?b.quantity/v.total.quantity*100:0;
    const name=b.code==="001983"?"18–29 лет":b.code==="001984"?"30–49 лет":b.name;
    return '<div class="mix-row"><div class="mix-label"><span title="'+E(b.name)+'">'+E(b.code)+' · '+E(name)+'</span><span>'+formatQ(b.quantity)+' · '+pct(share)+'</span></div><div class="mix-track"><div class="mix-fill" style="width:'+share+'%"></div></div><small>'+formatM(b.amount)+' ₽</small></div>';
  }).join(""):'<p class="empty">В выбранном периоде услуги не найдены.</p>';
}

function render(){
  try{
    const view=aggregate();if(view!==state.view){state.heatPage=0;state.signalLimit=40;pickerCache.clear();state.focus=null;$("focusContext").hidden=true;}state.view=view;$("summaryKpis").hidden=false;$("filterError").hidden=true;$("report").hidden=false;$("exportButton").disabled=false;
    renderKpis(view);chart("quantityChart",view,"quantity","bar");chart("amountChart",view,"amount","line");renderHeat(view);renderSignals(view);renderExtra(view);renderTable(view);syncFilters();
    $("periodLabel").textContent=fullDate(view.start)+" – "+fullDate(view.end);
  }catch(error){$("summaryKpis").hidden=true;state.view=null;aggregateCache=null;hideTooltip();$("filterError").textContent=error.message;$("filterError").hidden=false;$("report").hidden=true;$("exportButton").disabled=true;$("resetButton").disabled=false;}
}
function csvCell(value){
  let s=String(value??"");
  if(/^[\s]*[=+\-@]/.test(s))s="'"+s;
  return '"'+s.replace(/"/g,'""')+'"';
}
function exportCsv(){
  const v=state.view;if(!v)return;
  const rows=[["Дата","Подразделение","Врач","Код услуги","Наименование услуги","Количество","Сумма, руб.","Исходных строк"]];
  for(const r of v.selected)rows.push([r[0],D.departments[r[1]],D.doctors[r[2]],"'"+D.services[r[3]].code,D.services[r[3]].name,r[4],(r[5]/100).toFixed(2).replace(".",","),r[6]]);
  const text="\uFEFF"+rows.map(row=>row.map(csvCell).join(";")).join("\r\n");
  const url=URL.createObjectURL(new Blob([text],{type:"text/csv;charset=utf-8;"}));
  const a=document.createElement("a");a.href=url;a.download="services_"+state.tab+"_"+v.start+"_"+v.end+".csv";a.click();setTimeout(()=>URL.revokeObjectURL(url),1000);
}
function switchTab(tab){
  state.tab=tab;state.limit=25;window.AnnualPlan?.setVisible(tab==="target");
  for(const b of document.querySelectorAll("[data-tab]")){const active=b.dataset.tab===tab;b.setAttribute("aria-selected",active);b.tabIndex=active?0:-1;}
  $("report").setAttribute("aria-labelledby",tab==="target"?"targetTab":"allTab");
  $("codeField").hidden=tab==="target";render();
}
let tooltipTarget=null,tooltipSize=null,tooltipFrame=0,tooltipPoint=null;
function hideTooltip(){cancelAnimationFrame(tooltipFrame);tooltipFrame=0;$("tooltip").hidden=true;tooltipTarget=null;tooltipSize=null;}
function showTooltip(target,event){
  const tooltip=$("tooltip");
  if(target!==tooltipTarget){
    tooltipTarget=target;tooltip.textContent=target.dataset.tip;tooltip.hidden=false;
    tooltipSize={width:tooltip.offsetWidth,height:tooltip.offsetHeight};
  }
  if(event)tooltipPoint={x:event.clientX,y:event.clientY};
  else{const r=target.getBoundingClientRect();tooltipPoint={x:r.left+r.width/2,y:r.top};}
  if(tooltipFrame)return;
  tooltipFrame=requestAnimationFrame(()=>{
    tooltipFrame=0;if(!tooltipTarget||!tooltipSize)return;
    const x=Math.max(8,Math.min(tooltipPoint.x+12,innerWidth-tooltipSize.width-8));
    const y=Math.max(8,Math.min(tooltipPoint.y+12,innerHeight-tooltipSize.height-8));
    tooltip.style.transform='translate3d('+x+'px,'+y+'px,0)';
  });
}
function initializeRefinements(){
  const media=window.matchMedia("(max-width:640px)"),toggle=$("filtersToggle"),controls=$("filterControls");
  const sync=()=>{const expanded=!media.matches||toggle.getAttribute("aria-expanded")==="true";controls.hidden=!expanded;toggle.textContent=expanded?"Скрыть фильтры":"Настроить фильтры";};
  toggle.addEventListener("click",()=>{toggle.setAttribute("aria-expanded",toggle.getAttribute("aria-expanded")!=="true");sync();});
  media.addEventListener("change",sync);sync();
  UI.roving($("heatmap"),"[data-roving]",true);
  for(const id of ["quantityChart","amountChart"]){
    UI.roving($(id),"[data-chart-day]");
    $(id).addEventListener("focusin",e=>{if(UI.instant()&&e.target.dataset.chartDay)focusDay(e.target.dataset.chartDay,null,false);});
    $(id).addEventListener("click",e=>{if(e.target.dataset.chartDay)focusDay(e.target.dataset.chartDay,null,e.detail>0);});
    $(id).addEventListener("keydown",e=>{if(e.target.dataset.chartDay&&["Enter"," "].includes(e.key)){e.preventDefault();focusDay(e.target.dataset.chartDay,null,false);}});
  }
  for(const [id,step] of [["heatPrev",-1],["heatNext",1]])$(id).addEventListener("click",()=>{
    state.heatPage+=step;renderHeat(state.view);$("heatmap").closest(".heatmap-scroll").scrollTop=0;
    $("heatmap").querySelector("[data-roving]")?.focus({preventScroll:true});
  });
  $("clearFocus").addEventListener("click",()=>{state.focus=null;$("focusContext").hidden=true;updateChartFocus(false);renderHeat(state.view);$("heatMetric").focus({preventScroll:true});});
  $("heatmap").addEventListener("click",e=>{const cell=e.target.closest("[data-heat-day]");if(cell)inspectDay(Number(cell.dataset.heatDoctor),cell.dataset.heatDay,cell,e.detail===0);});
  $("heatmap").addEventListener("keydown",e=>{if(e.target.dataset.heatDay&&["Enter"," "].includes(e.key)){e.preventDefault();inspectDay(Number(e.target.dataset.heatDoctor),e.target.dataset.heatDay,e.target,true);}});
  $("allSignals").addEventListener("click",()=>showSignalQueue());
  $("closeSignals").addEventListener("click",()=>{$("signalQueue").hidden=true;$("allSignals").setAttribute("aria-expanded","false");$("allSignals").focus();});
  $("signalSearch").addEventListener("input",()=>{state.signalLimit=40;renderSignalQueue();});
  $("signalSort").addEventListener("change",()=>{state.signalLimit=40;renderSignalQueue();});
  $("moreSignals").addEventListener("click",()=>{state.signalLimit+=40;renderSignalQueue();});
}

function initialize(){
  for(const id of ["dateFrom","dateTo"]){$(id).min=D.meta.start;$(id).max=D.meta.end;}
  $("dateFrom").value=D.meta.start;$("dateTo").value=D.meta.end;
  $("serviceCount").textContent=D.services.length+" кодов";
  const gap=D.meta.dataEnd<D.meta.end;
  $("coverageNotice").innerHTML=(gap?'<strong>Выгрузка заканчивается '+E(fullDate(D.meta.dataEnd))+'</strong> Даты после неё не считаются нулевыми и не участвуют в поиске снижений.':'<strong>Данные по '+E(fullDate(D.meta.dataEnd))+'</strong> Текущий день исключён из поиска снижений.');
  $("coverageNotice").innerHTML='<svg viewBox="0 0 24 24" aria-hidden="true"><circle cx="12" cy="12" r="9"/><path d="M12 7v6m0 3v1"/></svg><span class="notice-copy">'+$("coverageNotice").innerHTML+'</span>';
  $("sourceLabel").textContent=D.meta.source+" · "+D.meta.sheet+" · "+formatQ(D.meta.stats.included_rows)+" строк";
  const stats=D.meta.stats;
  $("qualityDetails").innerHTML='<ul><li>Прочитано строк: '+formatQ(stats.source_rows||0)+'</li><li>Включено: '+formatQ(stats.included_rows||0)+'</li><li>Вне периода: '+formatQ(stats.outside_period||0)+'</li><li>Другие подразделения: '+formatQ(stats.excluded_departments||0)+'</li><li>Пустых строк: '+formatQ(stats.blank_rows||0)+'</li><li>Восстановлено пустых тарифов: '+formatQ(stats.restored_tariff_rows||0)+'</li><li>Нулевые тарифы при положительном количестве: '+formatQ(stats.zero_tariff_rows||0)+'</li><li>Строк без врача: '+formatQ(stats.missing_doctor_rows||0)+'</li></ul><p>Источник: '+E(D.meta.source)+'; лист «'+E(D.meta.sheet)+'». Создан: '+E(D.meta.created)+'.';
  if(stats.zero_tariff_rows||stats.restored_tariff_rows||stats.missing_doctor_rows)$("coverageNotice").innerHTML+=' <strong>Есть особенности данных: см. «Как считаем».</strong>';
  document.addEventListener("click",event=>{
    const b=event.target.closest("button");if(!b)return;
    if(b.dataset.signalIndex!==undefined){const s=state.view?.signals[Number(b.dataset.signalIndex)];if(s)inspectDay(s.id,s.day,b,event.detail===0,s);return;}
    if(b.dataset.queueDoctor!==undefined){showSignalQueue(Number(b.dataset.queueDoctor));return;}
    if(b.dataset.queueDay){showSignalQueue(null,b.dataset.queueDay);return;}
    if(b.dataset.servicesDoctor!==undefined){openDoctorServices(Number(b.dataset.servicesDoctor),b,event.detail===0);return;}
    if(b.dataset.tab){switchTab(b.dataset.tab);return;}
    if(b.dataset.doctor!==undefined){$("doctor").value=D.doctors[Number(b.dataset.doctor)];state.limit=25;render();$("scopeText").scrollIntoView({block:"nearest"});return;}
    if(b.dataset.detail){state.detail=b.dataset.detail;state.sort=state.detail==="days"?"label":"quantity";state.desc=state.detail!=="days";state.limit=25;$("tableSearch").value="";document.querySelectorAll("[data-detail]").forEach(x=>x.setAttribute("aria-pressed",x===b));renderTable(state.view);return;}
    if(b.dataset.sort){state.desc=state.sort===b.dataset.sort?!state.desc:b.dataset.sort!=="label";state.sort=b.dataset.sort;renderTable(state.view);}
  });
  document.querySelector(".tabs").addEventListener("keydown",e=>{
    if(["ArrowLeft","ArrowRight","Home","End"].includes(e.key)){e.preventDefault();const tab=e.key==="Home"?"target":e.key==="End"?"all":state.tab==="target"?"all":"target";switchTab(tab);$(tab==="target"?"targetTab":"allTab").focus();}
  });
  $("heatMetric").addEventListener("change",()=>{state.heatPage=0;if(state.view)renderHeat(state.view);});
  for(const id of ["dateFrom","dateTo","threshold"])$(id).addEventListener("change",()=>{state.limit=25;render();});
  $("tableSearch").addEventListener("input",()=>{state.limit=25;renderTable(state.view);syncFilters();});
  $("moreButton").addEventListener("click",()=>{state.limit+=50;renderTable(state.view);});
  $("resetButton").addEventListener("click",()=>{for(const id of ["department","doctor","service","tableSearch"])$(id).value="";$("dateFrom").value=D.meta.start;$("dateTo").value=D.meta.end;state.limit=25;render();});
  $("exportButton").addEventListener("click",exportCsv);$("printButton").addEventListener("click",()=>window.print());
  for(const id of ["methodButton","mobileMethod"])$(id).addEventListener("click",()=>$("methodDialog").showModal());
  $("closeMethod").addEventListener("click",()=>$("methodDialog").close());
  $("methodDialog").addEventListener("click",e=>{if(e.target===$("methodDialog")){const r=e.target.getBoundingClientRect();if(e.clientX<r.left||e.clientX>r.right||e.clientY<r.top||e.clientY>r.bottom)e.target.close();}});
  document.addEventListener("pointermove",e=>{const t=e.target.closest("[data-tip]");if(t)showTooltip(t,e);else hideTooltip();});
  document.addEventListener("focusin",e=>{if(e.target.matches("[data-tip]"))showTooltip(e.target);});
  document.addEventListener("focusout",hideTooltip);
  document.addEventListener("keydown",e=>{if(e.key==="Escape")hideTooltip();});
  window.addEventListener("blur",hideTooltip);
  document.addEventListener("scroll",hideTooltip,true);
  let resizeTimer;window.addEventListener("resize",()=>{clearTimeout(resizeTimer);resizeTimer=setTimeout(()=>{if(state.view){chart("quantityChart",state.view,"quantity","bar");chart("amountChart",state.view,"amount","line");}window.AnnualPlan?.resize();window.ClinicPace?.resize();hideTooltip();},150);});
  initializeFilters();initializeTheme();initializeDoctorServices();initializeRefinements();window.ClinicPace?.init(D);window.AnnualPlan?.init(D.annualPlan,D.meta);render();
}
// Search menus keep a draft query separate from the applied filter.
const pickers = new Map();
function pickerRows(kind){
  const cacheKey=JSON.stringify([kind,$("dateFrom").value,$("dateTo").value,$("department").value,$("doctor").value,$("service").value,state.tab]);
  if(pickerCache.has(cacheKey))return pickerCache.get(cacheKey);
  const start=$("dateFrom").value,end=$("dateTo").value,dep=$("department").value;
  const doctor=$("doctor").value,service=$("service").value;
  const rows=D.records.filter(r=>r[0]>=start&&r[0]<=end);
  const selectedCode=r=>state.tab==="target"?D.targets.includes(D.services[r[3]].code):!service||D.services[r[3]].code===service;
  const totals=new Map(),units=new Map();
  for(const r of rows){
    if(kind!=="department"&&dep!==""&&r[1]!==Number(dep))continue;
    if(kind!=="doctor"&&doctor&&D.doctors[r[2]]!==doctor)continue;
    const key=kind==="doctor"?D.doctors[r[2]]:kind==="department"?String(r[1]):D.services[r[3]].code;
    if(!totals.has(key))totals.set(key,0);
    if(kind==="service"||selectedCode(r))totals.set(key,totals.get(key)+r[4]);
    if(!units.has(key))units.set(key,new Set());units.get(key).add(D.departments[r[1]]);
  }
  const labels={doctor:"Все врачи",department:"Все подразделения",service:"Все услуги"};
  const items=[...totals].map(([value,count])=>({value,count,
    label:kind==="department"?D.departments[Number(value)]:kind==="service"?value+" · "+D.services.find(s=>s.code===value).name:value,
    note:kind==="doctor"?[...units.get(value)].join(" · "):kind==="department"?"Подразделение":"Код услуги"}));
  items.sort((a,b)=>a.label.localeCompare(b.label,"ru"));
  const result=[{value:"",label:labels[kind],note:"Снять ограничение",count:items.reduce((total,item)=>total+item.count,0)},...items];
  if(pickerCache.size>12)pickerCache.clear();pickerCache.set(cacheKey,result);return result;
}
class SearchPicker {
  constructor(kind){
    this.kind=kind;this.root=$(kind+"Picker");this.button=$(kind+"Trigger");this.menu=$(kind+"Menu");
    this.query=$(kind+"Query");this.list=$(kind+"List");this.status=$(kind+"Status");this.index=0;this.items=[];
    this.button.addEventListener("click",event=>this.menu.hidden?this.open(event.detail===0):this.close());
    this.query.addEventListener("input",()=>{this.index=0;this.populate();});
    this.query.addEventListener("keydown",event=>{
      if(event.key==="Escape"){event.preventDefault();this.close(true);return;}
      if(event.key==="Enter"){event.preventDefault();if(this.items[this.index])this.choose(this.items[this.index].value);return;}
      if(!["ArrowDown","ArrowUp","Home","End"].includes(event.key)||!this.items.length)return;
      event.preventDefault();
      if(event.key==="Home")this.index=0;
      else if(event.key==="End")this.index=this.items.length-1;
      else this.index=(this.index+(event.key==="ArrowDown"?1:-1)+this.items.length)%this.items.length;
      this.highlight();$(this.kind+"Option"+this.index)?.scrollIntoView({block:"nearest"});
    });
    this.list.addEventListener("pointerdown",event=>event.preventDefault());
    this.list.addEventListener("click",event=>{const row=event.target.closest("[data-choice]");if(row)this.choose(row.dataset.choice);});
    this.root.addEventListener("focusout",()=>setTimeout(()=>{if(!this.root.contains(document.activeElement))this.close();},0));
    document.addEventListener("pointerdown",event=>{if(!this.root.contains(event.target))this.close();});
    $(kind+"Close").addEventListener("click",()=>this.close(true));
    pickers.set(kind,this);
  }
  open(keyboard=false){
    this.menu.dataset.instant=keyboard||UI.instant();
    for(const picker of pickers.values())picker.close();
    this.menu.hidden=false;this.root.classList.add("is-open");this.button.setAttribute("aria-expanded","true");
    this.query.value="";this.index=0;this.populate();this.query.focus();
  }
  close(focus=false){
    this.menu.hidden=true;this.root.classList.remove("is-open");this.button.setAttribute("aria-expanded","false");
    this.query.removeAttribute("aria-activedescendant");if(focus)this.button.focus();
  }
  populate(){
    const query=normalize(this.query.value);this.items=pickerRows(this.kind).filter(item=>!query||normalize(item.label+" "+item.note).includes(query));
    this.index=Math.min(this.index,Math.max(0,this.items.length-1));
    this.list.innerHTML=this.items.map((item,i)=>'<li id="'+this.kind+'Option'+i+'" role="option" aria-selected="'+(item.value===$(this.kind).value)+'" data-choice="'+E(item.value)+'" class="picker-option">'+
      '<span class="option-avatar" aria-hidden="true">'+(item.value?E(this.kind==="doctor"?item.label.slice(0,2):this.kind==="department"?"ЖК":"№"):"∞")+'</span><span class="option-text"><strong>'+E(item.label)+'</strong><small>'+E(item.note)+'</small></span><span class="option-count">'+formatQ(item.count)+'<small>услуг</small></span><span class="option-check" aria-hidden="true">'+(item.value===$(this.kind).value?"✓":"")+'</span></li>').join("")||'<li class="picker-empty" role="presentation">Ничего не найдено. Попробуйте другую фамилию, код или название.</li>';
    this.status.textContent="Найдено: "+this.items.filter(item=>item.value!=="").length;
    this.highlight();
  }
  highlight(){
    [...this.list.querySelectorAll("[data-choice]")].forEach((item,i)=>item.classList.toggle("is-highlighted",i===this.index));
    if(this.items.length)this.query.setAttribute("aria-activedescendant",this.kind+"Option"+this.index);
    else this.query.removeAttribute("aria-activedescendant");
  }
  choose(value){$(this.kind).value=value;this.close(true);state.limit=25;render();}
  sync(){
    const value=$(this.kind).value;
    const labels={doctor:"Все врачи",department:"Все подразделения",service:"Все услуги"};
    let label=labels[this.kind];
    if(value)label=this.kind==="department"?D.departments[Number(value)]:this.kind==="service"?value+" · "+(D.services.find(s=>s.code===value)?.name||""):value;
    $(this.kind+"Label").textContent=label;this.button.title=label;this.button.classList.toggle("has-selection",Boolean(value));
    if(!this.menu.hidden)this.populate();
  }
}
function periodRange(preset){
  if(preset==="month")return [D.meta.start,D.meta.end];
  const end=D.meta.dataEnd<D.meta.end?D.meta.dataEnd:D.meta.end;
  const first=dateObj(end);if(preset==="week")first.setUTCDate(first.getUTCDate()-6);
  const start=first.toISOString().slice(0,10);
  return [start<D.meta.start?D.meta.start:start,end];
}
function syncFilters(){
  for(const picker of pickers.values())picker.sync();
  const chips=[];
  for(const kind of ["department","doctor","service"]){
    if($(kind).value&&(kind!=="service"||state.tab==="all"))chips.push({kind,label:$(kind+"Label").textContent});
  }
  if($("dateFrom").value!==D.meta.start||$("dateTo").value!==D.meta.end)chips.push({kind:"period",label:dayText($("dateFrom").value)+" – "+dayText($("dateTo").value)});
  $("activeFilters").innerHTML=chips.map(c=>'<button class="filter-chip" data-clear="'+c.kind+'" aria-label="Убрать фильтр: '+E(c.label)+'"><span>'+E(c.label)+'</span><span aria-hidden="true">×</span></button>').join("");
  $("filterHint").textContent=chips.length?"Выбрано условий: "+chips.length:"Весь период · все врачи";
  $("resetButton").disabled=!chips.length&&!$("tableSearch").value;
  for(const button of document.querySelectorAll("[data-range]")){
    const [from,to]=periodRange(button.dataset.range);
    button.setAttribute("aria-pressed",from===$("dateFrom").value&&to===$("dateTo").value);
  }
}
function initializeFilters(){
  for(const kind of ["department","doctor","service"])new SearchPicker(kind);
  document.addEventListener("click",event=>{
    const button=event.target.closest("button");if(!button)return;
    if(button.dataset.range){const [from,to]=periodRange(button.dataset.range);$("dateFrom").value=from;$("dateTo").value=to;state.limit=25;render();}
    if(button.dataset.clear){
      const kind=button.dataset.clear;if(kind==="period"){$("dateFrom").value=D.meta.start;$("dateTo").value=D.meta.end;}else $(kind).value="";
      state.limit=25;render();$(kind==="period"?"dateFrom":kind+"Trigger").focus();
    }
  });
}
function setTheme(dark,save=false){
  document.documentElement.dataset.theme=dark?"dark":"light";$("themeToggle").checked=dark;
  $("themeLabel").textContent=dark?"Тёмная тема":"Светлая тема";
  if(save){try{localStorage.setItem("veresaev-dashboard-theme",dark?"dark":"light");}catch{}}
  if(state.view)renderHeat(state.view);
}
function initializeTheme(){
  let saved=null;try{saved=localStorage.getItem("veresaev-dashboard-theme");}catch{}
  setTheme(saved==="dark");$("themeToggle").addEventListener("change",event=>setTheme(event.target.checked,true));
}

// The detail uses exactly the selected rows behind the clicked doctor's total.
function doctorServiceRows(view,doctorId,tab=state.tab){
  const grouped=new Map();
  if(tab==="target")for(const code of D.targets){
    const service=D.services.find(s=>s.code===code);
    grouped.set(code,{code,name:service?.name||"Услуга "+code,quantity:0,amount:0});
  }
  for(const r of view.selected){
    if(r[2]!==doctorId)continue;
    const service=D.services[r[3]];
    if(!grouped.has(service.code))grouped.set(service.code,{...service,quantity:0,amount:0});
    const row=grouped.get(service.code);row.quantity+=r[4];row.amount+=r[5];
  }
  return [...grouped.values()];
}
let serviceDetail=null,serviceTrigger=null,serviceMotion=null,shareMotion=null;
const reducedMotion=()=>window.matchMedia?.("(prefers-reduced-motion: reduce)").matches??true;
function renderDoctorServices(){
  if(!serviceDetail)return;
  const query=normalize($("doctorServicesSearch").value),sort=$("doctorServicesSort").value;
  const rows=serviceDetail.rows.filter(row=>normalize(row.code+" "+row.name).includes(query));
  rows.sort((a,b)=>sort==="code"?a.code.localeCompare(b.code):b[sort]-a[sort]||a.code.localeCompare(b.code));
  $("doctorServicesCount").textContent="Кодов: "+rows.length+" из "+serviceDetail.rows.length;
  $("doctorServicesBody").innerHTML=rows.map(row=>{
    const share=serviceDetail.quantity?row.quantity/serviceDetail.quantity*100:0;
    return '<tr><th scope="row"><span class="service-code">'+E(row.code)+'</span><span class="service-name">'+E(row.name)+'</span><div class="service-share" aria-hidden="true"><i style="width:'+share+'%"></i></div></th><td><strong>'+formatQ(row.quantity)+'</strong><small>'+pct(share)+'</small></td><td>'+formatM(row.amount)+'</td></tr>';
  }).join("")||'<tr><td colspan="3" class="empty">Услуги не найдены. Измените код или название в поиске.</td></tr>';
  $("doctorServicesSubtotalQ").textContent=formatQ(rows.reduce((n,r)=>n+r.quantity,0));
  $("doctorServicesSubtotalM").textContent=formatM(rows.reduce((n,r)=>n+r.amount,0));
}
function openDoctorServices(doctorId,trigger,keyboard=false,day=null,signal=null){
  const base=state.view,doctor=base?.doctors.get(doctorId);if(!doctor)return;
  const view=day?{...base,start:day,end:day,selected:base.selected.filter(r=>r[0]===day)}:base;
  const dialog=$("doctorServicesDialog"),rows=doctorServiceRows(view,doctorId);
  serviceMotion?.cancel();shareMotion?.cancel();
  serviceDetail={doctorId,rows,quantity:rows.reduce((n,r)=>n+r.quantity,0),amount:rows.reduce((n,r)=>n+r.amount,0)};
  serviceTrigger=trigger;
  $("doctorSignalContext").hidden=!signal;
  $("doctorSignalContext").textContent=signal?("Снижение "+pct(signal.drop*100)+". База: "+(signal.metric==="quantity"?formatQ(signal.baseline)+" усл.":formatM(signal.baseline)+" ₽")+" — медиана за "+signal.previous.map(dayText).join(", ").replace(/\.$/,"")+". Фильтры отчёта и период сравнения сохранены."):"";
  $("doctorServicesName").textContent=doctor.name;
  $("doctorServicesUnits").textContent=[...doctor.departments].join(" · ");
  $("doctorServicesScope").textContent=(state.tab==="target"?"Репродуктивное здоровье · 001983 и 001984":$("service").value?"Услуга "+$("service").value:"Все услуги")+" · "+fullDate(view.start)+" – "+fullDate(view.end);
  $("doctorServicesTotalQ").textContent=formatQ(serviceDetail.quantity);
  $("doctorServicesTotalM").textContent=formatM(serviceDetail.amount);
  $("doctorServicesSearch").value="";$("doctorServicesSort").value="quantity";
  renderDoctorServices();
  $("tooltip").hidden=true;
  dialog.classList.remove("is-closing");dialog.classList.toggle("instant-motion",keyboard||reducedMotion());
  dialog.style.opacity="1";dialog.style.transform="none";
  document.documentElement.classList.add("has-service-dialog");
  if(!dialog.open)dialog.showModal();
  $("doctorServicesScroll").scrollTop=0;
  $("doctorServicesTitle").focus({preventScroll:true});
  if(!keyboard&&!reducedMotion()&&window.anime){
    serviceMotion=window.anime.animate(dialog,{opacity:[0,1],translateY:[12,0],scale:[.975,1],duration:220,ease:"outQuint"});
    shareMotion=window.anime.animate(dialog.querySelectorAll(".service-share i"),{scaleX:[0,1],duration:240,ease:"outQuint"});
  }
}
function closeDoctorServices(keyboard=false){
  const dialog=$("doctorServicesDialog");if(!dialog.open)return;
  // Escape always closes immediately, including during an entrance or exit.
  if(dialog.classList.contains("is-closing")&&!keyboard)return;
  serviceMotion?.cancel();shareMotion?.cancel();
  const finish=()=>{dialog.close();};
  if(keyboard||reducedMotion()||!window.anime){finish();return;}
  dialog.classList.add("is-closing");
  serviceMotion=window.anime.animate(dialog,{opacity:0,translateY:8,scale:.985,duration:160,ease:"outQuad",onComplete:finish});
}
function initializeDoctorServices(){
  const dialog=$("doctorServicesDialog");
  $("doctorServicesSearch").addEventListener("input",renderDoctorServices);
  $("doctorServicesSort").addEventListener("change",renderDoctorServices);
  $("closeDoctorServices").addEventListener("click",event=>closeDoctorServices(event.detail===0));
  dialog.addEventListener("cancel",event=>{event.preventDefault();closeDoctorServices(true);});
  // Search inputs consume Escape to clear themselves; close this modal on the first press.
  dialog.addEventListener("keydown",event=>{if(event.key==="Escape"){event.preventDefault();closeDoctorServices(true);}});
  let startedOnBackdrop=false;
  const outside=event=>{const r=dialog.getBoundingClientRect();return event.clientX<r.left||event.clientX>r.right||event.clientY<r.top||event.clientY>r.bottom;};
  dialog.addEventListener("pointerdown",event=>{startedOnBackdrop=event.target===dialog&&outside(event);});
  dialog.addEventListener("click",event=>{if(startedOnBackdrop&&event.target===dialog&&outside(event))closeDoctorServices();startedOnBackdrop=false;});
  dialog.addEventListener("close",()=>{
    serviceMotion?.cancel();shareMotion?.cancel();serviceMotion=null;shareMotion=null;
    document.documentElement.classList.remove("has-service-dialog");
    dialog.classList.remove("is-closing");dialog.style.opacity="1";dialog.style.transform="none";
    const heatTrigger=state.focus&&[...$("heatmap").querySelectorAll("[data-heat-key]")].find(n=>n.dataset.heatKey===state.focus.doctorId+"|"+state.focus.day);
    const trigger=serviceTrigger?.isConnected?serviceTrigger:serviceTrigger?.dataset.heatKey?heatTrigger:document.querySelector('[data-services-doctor="'+serviceDetail?.doctorId+'"]');
    (trigger||heatTrigger||$("tableSearch")).focus({preventScroll:true});serviceDetail=null;serviceTrigger=null;
  });
}

initialize();
window.Dashboard = {data:D,state,aggregate,render,switchTab,chartValue,doctorServiceRows};
})();
