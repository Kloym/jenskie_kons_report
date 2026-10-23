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
const sum = (rows, index) => rows.reduce((total,r) => total + r[index], 0);
const median = values => {const a=[...values].sort((x,y)=>x-y);const i=Math.floor(a.length/2);return a.length%2?a[i]:(a[i-1]+a[i])/2;};
const state = {tab:"target",detail:"doctors",sort:"quantity",desc:true,limit:25,view:null};
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
    throw new Error("Выберите корректные даты внутри периода "+dayText(D.meta.start)+" – "+dayText(D.meta.end)+".");
  }
  const dates=datesBetween(start,end);
  const scoped=D.records.filter(r=>r[0]>=start&&r[0]<=end&&(dep===""||r[1]===Number(dep))&&normalize(D.doctors[r[2]]).includes(doctor));
  const selected=scoped.filter(r=>state.tab==="target"?D.targets.includes(D.services[r[3]].code):
    !search||normalize(D.services[r[3]].code+" "+D.services[r[3]].name).includes(search));
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
    const signals=[];
    for(let i=0;i<dates.length;i++){
      const day=dates[i];if(day>=D.meta.today||!activity(day))continue;
      const previous=dates.slice(0,i).filter(x=>weekday(x)===weekday(day)&&activity(x)).slice(-4);
      if(previous.length<2)continue;
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
  return {dates,start,end,scoped,selected,total,doctors,services,days,active,signals,dailySignals};
}
function niceStep(max,integer){
  const raw=max/4||1;const pow=10**Math.floor(Math.log10(raw));
  let step=[1,2,2.5,5,10].map(n=>n*pow).find(n=>n>=raw)||pow*10;
  return integer?Math.max(1,Math.ceil(step)):step;
}
function chart(id,view,metric,type){
  const el=$(id), width=Math.max(260,Math.round(el.getBoundingClientRect().width)),height=190;
  const left=metric==="amount"?56:40,right=12,top=12,bottom=30;
  const w=width-left-right,h=height-top-bottom,n=view.dates.length;
  const vals=view.dates.map(day=>{const b=view.days.get(day);return b.activity>0?b[metric]:null;});
  const max=Math.max(0,...vals.filter(x=>x!==null)),step=niceStep(max,metric==="quantity"),ymax=Math.max(step*4,Math.ceil((max*1.06)/step)*step);
  const x=i=>left+(i+.5)*w/n,y=value=>top+h-value/ymax*h;
  const low=new Set(view.dailySignals.filter(s=>metric==="quantity"?s.quantityDrop>Number($("threshold").value):s.amountDrop>Number($("threshold").value)).map(s=>s.day));
  let body="";
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
      body+='<rect class="bar '+(low.has(day)?"low":"")+'" x="'+(xx-bw/2)+'" y="'+y(value)+'" width="'+bw+'" height="'+(y(0)-y(value))+'" rx="1"/>';
    }else{points.push([xx,y(value)]);if(low.has(day))body+='<circle cx="'+xx+'" cy="'+y(value)+'" r="4" fill="var(--negative)"/>';}
  }
  if(type==="line")path();
  for(let i=0;i<n;i++){
    const day=view.dates[i],b=view.days.get(day);
    const tip=fullDate(day)+"\n"+(b.activity?formatQ(b.quantity)+" услуг · "+formatM(b.amount)+" ₽":"Нет сведений об активности")+(day===D.meta.today?"\nТекущий день: данные могут быть неполными":"");
    body+='<rect class="hit" tabindex="0" role="img" aria-label="'+E(tip)+'" data-tip="'+E(tip)+'" x="'+(left+i*w/n)+'" y="'+top+'" width="'+w/n+'" height="'+h+'"/>';
  }
  el.innerHTML='<svg viewBox="0 0 '+width+' '+height+'" role="img" aria-label="'+(metric==="quantity"?"Количество услуг":"Сумма услуг")+' по дням"><title>Дневная динамика. Точные значения доступны в таблице «По дням».</title>'+body+'</svg>';
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
function renderHeat(v){
  const metric=$("heatMetric").value,entries=[...v.doctors.values()].sort((a,b)=>b[metric]-a[metric]||a.name.localeCompare(b.name,"ru"));
  const max=Math.max(1,...entries.flatMap(d=>[...d.daily.values()].map(b=>b[metric])));
  let html='<thead><tr><th scope="col">Врач</th>'+v.dates.map(day=>'<th scope="col" class="'+([0,6].includes(weekday(day))?"weekend":"")+'" title="'+E(fullDate(day))+'">'+Number(day.slice(-2))+'</th>').join("")+'<th scope="col">Итого</th></tr></thead><tbody>';
  for(const d of entries){
    const flags=new Map(d.signals.map(s=>[s.day,s]));
    html+='<tr><td><button class="doctor-link" data-doctor="'+d.id+'" title="'+E(d.name)+'">'+E(shorten(d.name))+'</button></td>';
    for(const day of v.dates){
      const b=d.daily.get(day)||bucket(),activity=(v.active.get(d.id+"|"+day)||0)>0;
      const intensity=b[metric]>0?Math.max(15,Math.round(15+85*Math.sqrt(b[metric]/max))):0;
      const dark=intensity>58,flag=flags.get(day);
      const amountText=metric==="quantity"?formatQ(b.quantity):(b.amount>=100000?formatQ(b.amount/100000)+"к":formatQ(b.amount/100));
      const tip=d.name+"\n"+fullDate(day)+"\n"+(activity?formatQ(b.quantity)+" услуг · "+formatM(b.amount)+" ₽":"Нет услуг врача в выгрузке; график неизвестен")+(flag?"\nСнижение "+pct(flag.drop*100)+" относительно сопоставимых дней":"");
      const style=intensity?'background:color-mix(in oklch,var(--accent) '+intensity+'%,var(--surface));color:'+(dark?"var(--surface)":"var(--text)"):'';
      html+='<td class="cell '+(!activity?"missing ":"")+(flag?"low-cell":"")+'" style="'+style+'" title="'+E(tip)+'" data-tip="'+E(tip)+'">'+(activity?(flag?"↓":"")+E(amountText):"·")+'</td>';
    }
    html+='<td style="text-align:right;padding-left:10px;white-space:nowrap;font-size:12px">'+(metric==="quantity"?formatQ(d.quantity):formatM(d.amount))+'</td></tr>';
  }
  if(!entries.length)html+='<tr><td colspan="'+(v.dates.length+2)+'" class="empty">Нет врачей по выбранным фильтрам.</td></tr>';
  $("heatmap").innerHTML=html+'</tbody>';
}
function renderSignals(v){
  $("signalCount").textContent=v.signals.length;
  if(!v.signals.length){
    $("signals").innerHTML='<div class="empty">Нет выявленных снижений.<p class="fine">Для сигнала нужны минимум два предыдущих таких же дня недели с услугами врача. Короткий период может не дать базы сравнения.</p></div>';
    return;
  }
  $("signals").innerHTML=v.signals.slice(0,5).map(s=>{
    const value=s.metric==="quantity"?formatQ:formatM,unit=s.metric==="quantity"?"усл.":"₽";
    return '<div class="signal"><div class="signal-title"><span>'+E(dayText(s.day))+' · '+(s.metric==="quantity"?"Объём":"Сумма")+'</span><strong>−'+pct(s.drop*100)+'</strong></div><button data-doctor="'+s.id+'">'+E(shorten(D.doctors[s.id]))+'</button><p>'+value(s.current)+' '+unit+' при базе '+value(s.baseline)+' '+unit+'</p><p title="'+E(s.previous.map(fullDate).join(", "))+'">База: '+s.previous.map(dayText).join(", ")+'</p></div>';
  }).join("")+(v.signals.length>5?'<p class="fine">Показаны 5 последних из '+v.signals.length+'. Выберите врача для разбора.</p>':"");
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
  $("detailTitle").textContent=t[0];$("detailSubtitle").textContent=state.detail==="doctors"?"Выберите врача для персонального отчёта.":"Количество и суммы в текущем выборе";
  $("tableCount").textContent=entries.length?"Показано "+shown.length+" из "+entries.length:"Нет строк";
  const heads=[["label",t[1]],["quantity","Количество"],["amount","Сумма, ₽"],["extra",t[2]]];
  let html='<thead><tr>'+heads.map(([key,label])=>'<th scope="col" aria-sort="'+(state.sort===key?(state.desc?"descending":"ascending"):"none")+'"><button data-sort="'+key+'">'+E(label)+(state.sort===key?(state.desc?" ↓":" ↑"):"")+'</button></th>').join("")+'<th scope="col">Доля количества</th><th scope="col">Сигналы</th></tr></thead><tbody>';
  for(const b of shown){
    const missing=state.detail==="days"&&!b.activity,share=v.total.quantity?b.quantity/v.total.quantity*100:0;
    const label=state.detail==="doctors"?'<button class="doctor-link" data-doctor="'+b.id+'">'+E(b.label)+'</button>':E(b.label);
    html+='<tr><td>'+label+'<small>'+E(b.subtitle)+'</small>'+(state.detail==="services"?'<div class="rowbar" style="width:'+share+'%"></div>':"")+'</td><td>'+(missing?"·":formatQ(b.quantity))+'</td><td>'+(missing?"·":formatM(b.amount))+'</td><td>'+(missing?"·":b.extra)+'</td><td>'+(missing?"·":pct(share))+'</td><td>'+(b.signals?'<span class="signal-label">'+b.signals+' ↓</span>':'<span class="neutral">нет</span>')+'</td></tr>';
  }
  if(!shown.length)html+='<tr><td colspan="6" class="empty">По этому запросу ничего не найдено. Измените поиск или сбросьте фильтры.</td></tr>';
  $("detailTable").innerHTML=html+'</tbody>';
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
    const view=aggregate();state.view=view;$("filterError").hidden=true;$("report").hidden=false;$("exportButton").disabled=false;
    renderKpis(view);chart("quantityChart",view,"quantity","bar");chart("amountChart",view,"amount","line");renderHeat(view);renderSignals(view);renderExtra(view);renderTable(view);
    $("periodLabel").textContent=fullDate(view.start)+" – "+fullDate(view.end);
  }catch(error){$("filterError").textContent=error.message;$("filterError").hidden=false;$("report").hidden=true;$("exportButton").disabled=true;}
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
  state.tab=tab;state.limit=25;
  for(const b of document.querySelectorAll("[data-tab]")){const active=b.dataset.tab===tab;b.setAttribute("aria-selected",active);b.tabIndex=active?0:-1;}
  $("report").setAttribute("aria-labelledby",tab==="target"?"targetTab":"allTab");
  $("codeField").hidden=tab==="target";render();
}
function showTooltip(target,event){
  const tooltip=$("tooltip");tooltip.textContent=target.dataset.tip;tooltip.hidden=false;
  const rect=target.getBoundingClientRect();let x=event?.clientX??rect.left+rect.width/2,y=event?.clientY??rect.top;
  x=Math.max(8,Math.min(x+12,innerWidth-tooltip.offsetWidth-8));y=Math.max(8,Math.min(y+12,innerHeight-tooltip.offsetHeight-8));
  tooltip.style.left=x+"px";tooltip.style.top=y+"px";
}
function initialize(){
  D.departments.forEach((name,i)=>{const option=document.createElement("option");option.value=i;option.textContent=name;$("department").append(option);});
  D.doctors.forEach(name=>{const option=document.createElement("option");option.value=name;$("doctorOptions").append(option);});
  D.services.forEach(s=>{const option=document.createElement("option");option.value=s.code;option.label=s.name;$("serviceOptions").append(option);});
  for(const id of ["dateFrom","dateTo"]){$(id).min=D.meta.start;$(id).max=D.meta.end;}
  $("dateFrom").value=D.meta.start;$("dateTo").value=D.meta.end;
  $("serviceCount").textContent=D.services.length+" кодов";
  const gap=D.meta.dataEnd<D.meta.end;
  $("coverageNotice").innerHTML=(gap?'<strong>Выгрузка заканчивается '+E(fullDate(D.meta.dataEnd))+'.</strong> Даты после неё не считаются нулевыми и не участвуют в поиске снижений.':'<strong>Данные по '+E(fullDate(D.meta.dataEnd))+'.</strong> Текущий день исключён из поиска снижений.');
  $("sourceLabel").textContent=D.meta.source+" · "+D.meta.sheet+" · "+formatQ(D.meta.stats.included_rows)+" строк";
  const stats=D.meta.stats;
  $("qualityDetails").innerHTML='<ul><li>Прочитано строк: '+formatQ(stats.source_rows||0)+'</li><li>Включено: '+formatQ(stats.included_rows||0)+'</li><li>Вне периода: '+formatQ(stats.outside_period||0)+'</li><li>Другие подразделения: '+formatQ(stats.excluded_departments||0)+'</li><li>Пустых строк: '+formatQ(stats.blank_rows||0)+'</li><li>Восстановлено пустых тарифов: '+formatQ(stats.restored_tariff_rows||0)+'</li><li>Нулевые тарифы при положительном количестве: '+formatQ(stats.zero_tariff_rows||0)+'</li><li>Строк без врача: '+formatQ(stats.missing_doctor_rows||0)+'</li></ul><p>Источник: '+E(D.meta.source)+'; лист «'+E(D.meta.sheet)+'». Создан: '+E(D.meta.created)+'.</p><p class="hash">SHA-256: '+E(D.meta.sha256)+'</p>';
  if(stats.zero_tariff_rows||stats.restored_tariff_rows||stats.missing_doctor_rows)$("coverageNotice").innerHTML+=' <strong>Есть особенности данных: см. «Как считаем».</strong>';
  document.addEventListener("click",event=>{
    const b=event.target.closest("button");if(!b)return;
    if(b.dataset.tab){switchTab(b.dataset.tab);return;}
    if(b.dataset.doctor!==undefined){$("doctor").value=D.doctors[Number(b.dataset.doctor)];state.limit=25;render();$("scopeText").scrollIntoView({block:"nearest"});return;}
    if(b.dataset.detail){state.detail=b.dataset.detail;state.sort=state.detail==="days"?"label":"quantity";state.desc=state.detail!=="days";state.limit=25;$("tableSearch").value="";document.querySelectorAll("[data-detail]").forEach(x=>x.setAttribute("aria-pressed",x===b));renderTable(state.view);return;}
    if(b.dataset.sort){state.desc=state.sort===b.dataset.sort?!state.desc:b.dataset.sort!=="label";state.sort=b.dataset.sort;renderTable(state.view);}
  });
  document.querySelector(".tabs").addEventListener("keydown",e=>{
    if(["ArrowLeft","ArrowRight","Home","End"].includes(e.key)){e.preventDefault();const tab=e.key==="Home"?"target":e.key==="End"?"all":state.tab==="target"?"all":"target";switchTab(tab);$(tab==="target"?"targetTab":"allTab").focus();}
  });
  for(const id of ["department","dateFrom","dateTo","heatMetric","threshold"])$(id).addEventListener("change",()=>{state.limit=25;render();});
  let inputTimer;for(const id of ["doctor","service"])$(id).addEventListener("input",()=>{clearTimeout(inputTimer);inputTimer=setTimeout(()=>{state.limit=25;render();},120);});
  $("tableSearch").addEventListener("input",()=>{state.limit=25;renderTable(state.view);});
  $("moreButton").addEventListener("click",()=>{state.limit+=50;renderTable(state.view);});
  $("resetButton").addEventListener("click",()=>{for(const id of ["department","doctor","service","tableSearch"])$(id).value="";$("dateFrom").value=D.meta.start;$("dateTo").value=D.meta.end;state.limit=25;render();});
  $("exportButton").addEventListener("click",exportCsv);$("printButton").addEventListener("click",()=>window.print());
  $("methodButton").addEventListener("click",()=>$("methodDialog").showModal());
  $("closeMethod").addEventListener("click",()=>$("methodDialog").close());
  $("methodDialog").addEventListener("click",e=>{if(e.target===$("methodDialog")){const r=e.target.getBoundingClientRect();if(e.clientX<r.left||e.clientX>r.right||e.clientY<r.top||e.clientY>r.bottom)e.target.close();}});
  $("themeButton").addEventListener("click",()=>{const dark=document.documentElement.dataset.theme!=="dark";document.documentElement.dataset.theme=dark?"dark":"light";$("themeButton").textContent=dark?"Светлая тема":"Тёмная тема";});
  document.addEventListener("pointermove",e=>{const t=e.target.closest("[data-tip]");if(t)showTooltip(t,e);else $("tooltip").hidden=true;});
  document.addEventListener("focusin",e=>{if(e.target.matches("[data-tip]"))showTooltip(e.target);});
  document.addEventListener("focusout",()=>$("tooltip").hidden=true);
  document.addEventListener("keydown",e=>{if(e.key==="Escape")$("tooltip").hidden=true;});
  window.addEventListener("blur",()=>$("tooltip").hidden=true);
  let resizeTimer;window.addEventListener("resize",()=>{clearTimeout(resizeTimer);resizeTimer=setTimeout(render,150);});
  render();
}
initialize();
window.Dashboard = {data:D,state,aggregate,render,switchTab};
})();