(() => {
'use strict';
const $=id=>document.getElementById(id);
const E=value=>String(value??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const number=new Intl.NumberFormat('ru-RU',{maximumFractionDigits:1});
const q=value=>number.format(value);
const normalize=value=>String(value).toLocaleLowerCase('ru-RU').replace(/ё/g,'е').replace(/\s+/g,'').replace(/№/g,'');
const formatDate=value=>new Date(value+'T12:00:00Z').toLocaleDateString('ru-RU',{day:'numeric',month:'long',timeZone:'UTC'});
const weekday=value=>new Date(value+'T12:00:00Z').getUTCDay();
const dayName=value=>['Вс','Пн','Вт','Ср','Чт','Пт','Сб'][weekday(value)];
const clamp=value=>Math.max(0,Math.min(1,value));
let data,selected=-1,date='',index=[],totalQuantities={},animation=null,viewKey='',min='',max='',status='',keyboard=true;
let chartKey='',calculationCache=new Map();
function stop(){animation?.cancel();animation=null;}
function makeIndex(source){
  const month=source.annualPlan.asOf.slice(0,7);
  const ids=new Map(source.departments.map((name,id)=>[normalize(name),id]));
  const result=source.annualPlan.departments.map(row=>({row,id:ids.get(normalize(row.name))??-1,quantities:{}}));
  const byId=new Map(result.filter(unit=>unit.id>=0).map(unit=>[unit.id,unit]));
  const targetIds=new Set(source.services.flatMap((s,i)=>source.targets.includes(s.code)?[i]:[]));
  for(const r of source.records){
    const unit=byId.get(r[1]);
    if(unit&&r[0].startsWith(month)&&targetIds.has(r[3]))unit.quantities[r[0]]=(unit.quantities[r[0]]||0)+r[4];
  }
  return result;
}
function calculate(which=-1,day=date){
  if(!data?.annualPlan)return {ready:false,reason:'no-plan'};
  if(!day||day.slice(0,7)!==data.annualPlan.asOf.slice(0,7))return {ready:false,reason:'other-month'};
  const unit=which<0?null:index[which];
  if(which>=0&&!unit)return {ready:false,reason:'unknown-unit'};
  if(unit?.id===-1||(!unit&&index.some(x=>x.id===-1)))return {ready:false,reason:'missing-unit'};
  const row=unit?.row||data.annualPlan.total;
  const key=which+"|"+day;
  if(!calculationCache.has(key))calculationCache.set(key,window.DailyPaceCore.calculate({target:row.monthPlan,date:day,quantities:unit?.quantities||totalQuantities,snapshotStart:data.meta.start,snapshotEnd:data.meta.dataEnd}));
  return calculationCache.get(key);
}
function reason(result){
  return ({'no-plan':'Добавьте месячный план в окне «Годовой план…».','other-month':'Месяц плана и выгрузки различается. Обновите дату и месячный план в программе.',
    'unknown-unit':'Для этой консультации нет плана.','missing-unit':'Консультация отсутствует в выгрузке. Её факт не подставляется как ноль.',
    'incomplete-month':'Для расчёта нужна полная выгрузка с 1-го числа месяца.','outside-snapshot':'На эту дату в выгрузке ещё нет данных.'})[result.reason]||'Недостаточно данных для расчёта.';
}
function rowMarkup(which){
  const row=data.annualPlan.departments[which],v=calculate(which);
  if(!v.ready)return '<span class="clinic-row-caption">План месяца: '+q(row.monthPlan)+' · '+E(reason(v))+'</span>';
  return '<span class="clinic-row-caption">План месяца '+q(row.monthPlan)+' · '+E(formatDate(date))+'</span><span class="clinic-row-numbers"><span>Нужно <b>'+q(v.dailyTarget)+'</b></span><span>Есть <b>'+q(v.actual)+'</b></span><span class="'+(v.status==='no-target'?'':v.dailyLeft?'pace-pending':'pace-complete')+'">Осталось <b>'+q(v.dailyLeft)+'</b></span></span>';
}
function refreshRows(){
  for(const node of document.querySelectorAll('[data-clinic-pace]'))node.innerHTML=rowMarkup(Number(node.dataset.clinicPace));
}
function drawMonth(v,animate=false){
  const nextKey=selected+"|"+date.slice(0,7),rebuild=nextKey!==chartKey;
  const width=800,height=102,left=34,right=10,top=13,bottom=24,w=width-left-right,h=height-top-bottom;
  const rows=v.days,known=rows.filter(d=>d.known);
  const values=known.map(d=>({...d,goal:calculate(selected,d.date).dailyTarget}));
  const maximum=Math.max(1,...values.flatMap(d=>[d.actual,d.goal]));
  const ceiling=Math.max(1,Math.ceil(maximum/5)*5);
  const x=i=>left+(i+.5)*w/rows.length,y=value=>top+h-value/ceiling*h;
  if(rebuild){
  let body='<rect id="clinicDateFocus" class="pace-current-day" x="0" y="'+(top-5)+'" width="'+w/rows.length*.9+'" height="'+(h+bottom+3)+'" rx="6"/><line class="gridline" x1="'+left+'" x2="'+(width-right)+'" y1="'+y(0)+'" y2="'+y(0)+'"/><text x="'+(left-7)+'" y="'+(y(0)+4)+'" text-anchor="end">0</text><text x="'+(left-7)+'" y="'+(y(ceiling)+4)+'" text-anchor="end">'+q(ceiling)+'</text>';
  rows.forEach((day,i)=>{
    const xx=x(i),bw=Math.min(14,w/rows.length*.55),goal=day.known?values[i].goal:null;

    if(day.known){
      body+='<rect class="pace-day-bar" x="'+(xx-bw/2)+'" y="'+y(day.actual)+'" width="'+bw+'" height="'+(y(0)-y(day.actual))+'" rx="2"/>';
      body+='<line class="pace-goal-tick" x1="'+(xx-bw*.8)+'" x2="'+(xx+bw*.8)+'" y1="'+y(goal)+'" y2="'+y(goal)+'"/>';
    }else body+='<text x="'+xx+'" y="'+(y(0)-6)+'" text-anchor="middle">·</text>';
    if(i%2===0)body+='<text x="'+xx+'" y="'+(height-3)+'" text-anchor="middle">'+Number(day.date.slice(-2))+'</text>';
    const tip=dayName(day.date)+', '+formatDate(day.date)+'; вес '+(day.units/2)+' дня. '+(day.known?'План на день '+goal+', факт '+day.actual+', осталось '+Math.max(0,goal-day.actual):'Фактических данных ещё нет');
    body+='<g '+(day.known?'role="button" tabindex="-1" data-pace-day="'+day.date+'"':'role="img"')+' aria-label="'+E(tip)+'"><title>'+E(tip)+'</title><rect class="pace-hit" x="'+(left+i*w/rows.length)+'" y="0" width="'+w/rows.length+'" height="'+height+'"/></g>';
  });
  $('clinicPaceChart').innerHTML='<svg viewBox="0 0 '+width+' '+height+'" role="group" aria-label="Дневной план: стрелки выбирают дату, Enter подтверждает">'+body+'</svg>';
  chartKey=nextKey;
  }
  const current=rows.findIndex(day=>day.date===date);
  window.DashboardUI.transform($('clinicDateFocus'),'translateX('+(x(current)-w/rows.length*.45)+'px)',animate&&!rebuild);
  for(const node of $('clinicPaceChart').querySelectorAll('[data-pace-day]')){const active=node.dataset.paceDay===date;node.tabIndex=active?0:-1;node.setAttribute('aria-pressed',active);}
  window.DashboardUI.scrollInside($('clinicPaceChart'),$('clinicPaceChart').querySelector('[data-pace-day="'+date+'"]'));
}
function render(animate=false){
  if(!data)return;stop();refreshRows();
  const section=$('clinicPace'),v=calculate(selected),row=selected<0?data.annualPlan.total:index[selected]?.row;
  $('clinicPaceScope').textContent=row?.name||'Всё учреждение';
  $('clinicPaceDate').value=date;
  $('clinicPacePrev').disabled=min>max||!date||date<=min;$('clinicPaceNext').disabled=min>max||!date||date>=max;
  $('clinicPaceContent').hidden=!v.ready;$('clinicPaceMissing').hidden=v.ready;
  if(!v.ready){$('clinicPaceMissing').textContent=reason(v);return;}
  const today=new Intl.DateTimeFormat('sv-SE',{timeZone:'Europe/Moscow'}).format(new Date());
  $('clinicPaceStamp').textContent=(date===today?'Сегодня, ':dayName(date)+', ')+formatDate(date)+' · '+(v.dayWeight===.5?'выходной, вес 0,5':'будний день, вес 1');
  $('clinicPaceGoal').textContent=q(v.dailyTarget);$('clinicPaceActual').textContent=q(v.actual);$('clinicPaceLeft').textContent=q(v.dailyLeft);
  section.dataset.state=v.status;
  const done=v.status==='month-done'||(v.status==='day-done'&&v.monthLeft===0);
  const statusText=v.status==='no-target'?'Месячный план равен нулю':done?'Месячный план выполнен':v.status==='day-done'?'Ориентир на день достигнут':'До ориентира на день';
  $('clinicPaceStatus').textContent=statusText;
  $('clinicPaceProgress').setAttribute('aria-label','Выполнено '+v.actual+' услуг при ориентире '+v.dailyTarget);
  const ratio=done?1:v.ratio===null?0:clamp(v.ratio),fill=$('clinicPaceFill');
  const changed=viewKey!==selected+'|'+date;viewKey=selected+'|'+date;
  if(animate&&changed&&!window.matchMedia?.('(prefers-reduced-motion: reduce)').matches&&window.anime){
    animation=window.anime.animate(fill,{scaleX:ratio,duration:240,ease:'outQuint'});
  }else fill.style.transform='scaleX('+ratio+')';
  $('clinicPaceMonth').textContent=q(v.monthActual)+' из '+q(v.target)+' услуг за месяц';
  $('clinicPaceMonthLeft').textContent='До плана месяца: '+q(v.monthLeft);
  $('clinicPaceBase').textContent='Базовый темп: '+q(v.baseDaily)+' за такой день · осталось '+q(v.remainingWeight)+' условных дней, включая выбранный';
  $('clinicPaceFormula').textContent='На начало дня выполнено '+q(v.before)+'. Остаток '+q(Math.max(0,v.target-v.before))+' ÷ '+q(v.remainingWeight)+' условных дней × '+q(v.dayWeight)+' = '+q(v.dailyTarget)+' услуг после округления вверх. Факт выбранного дня не меняет его ориентир.';
  $('clinicPaceSource').textContent='Факт: 001983 + 001984 из выгрузки по '+formatDate(data.meta.dataEnd)+'. '+(date===today?'Текущий день может быть заполнен частично. ':'')+'Годовые проценты выше сохраняют методику исходного отчёта.';
  drawMonth(v,animate);
  if(changed)$('clinicPaceAnnouncement').textContent=row.name+': '+statusText+'. Выполнено '+v.actual+', осталось '+v.dailyLeft+'.';
  status=v.status;
}
function setDate(value,animate=false){
  if(!value||value<min||value>max){$('clinicPaceDate').value=date;return;}
  date=value;render(animate);
}
function select(which,animate=false){if(selected===which&&viewKey)return;selected=which;render(animate);}
function settle(){stop();render(false);}
function init(source){
  if(!source.annualPlan)return;
  data=source;index=makeIndex(source);totalQuantities={};calculationCache.clear();chartKey="";
  for(const unit of index)for(const [day,value] of Object.entries(unit.quantities))totalQuantities[day]=(totalQuantities[day]||0)+value;
  const planMonth=data.annualPlan.asOf.slice(0,7),days=window.DailyPaceCore.monthDays(planMonth+'-01');
  min=[planMonth+'-01',data.meta.start].sort().at(-1);max=[days.at(-1).date,data.meta.dataEnd,data.meta.end].sort()[0];
  date=data.meta.dataEnd.slice(0,7)===planMonth?max:data.meta.dataEnd;
  $('clinicPaceDate').min=min;$('clinicPaceDate').max=max;
  $('clinicPaceDate').disabled=min>max;
  $('clinicPaceDate').addEventListener('pointerdown',()=>{keyboard=false;});
  $('clinicPaceDate').addEventListener('keydown',()=>{keyboard=true;});
  $('clinicPaceDate').addEventListener('change',event=>setDate(event.target.value,!keyboard));
  for(const [id,step] of [['clinicPacePrev',-1],['clinicPaceNext',1]])$(id).addEventListener('click',event=>{
    const next=new Date(date+'T12:00:00Z');next.setUTCDate(next.getUTCDate()+step);setDate(next.toISOString().slice(0,10),event.detail>0);
  });
  window.DashboardUI.roving($('clinicPaceChart'),'[data-pace-day]');
  $('clinicPaceChart').addEventListener('focusin',event=>{if(window.DashboardUI.instant()&&event.target.dataset.paceDay)setDate(event.target.dataset.paceDay);});
  $('clinicPaceChart').addEventListener('click',event=>{const day=event.target.closest('[data-pace-day]');if(day)setDate(day.dataset.paceDay,event.detail>0);});
  $('clinicPaceChart').addEventListener('keydown',event=>{const day=event.target.closest('[data-pace-day]');if(day&&['Enter',' '].includes(event.key)){event.preventDefault();const chosen=day.dataset.paceDay;setDate(chosen);$('clinicPaceChart').querySelector('[data-pace-day="'+chosen+'"]').focus();}});
  render();
}
function resize(){const value=calculate(selected);if(value.ready)drawMonth(value,false);}
window.ClinicPace={init,select,calculate,makeIndex,rowMarkup,refreshRows,settle,setDate,resize};
})();
