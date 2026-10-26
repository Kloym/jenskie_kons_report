(() => {
"use strict";
const $=id=>document.getElementById(id);
const escape=value=>String(value??"").replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
const count=new Intl.NumberFormat("ru-RU",{maximumFractionDigits:0});
const decimal=new Intl.NumberFormat("ru-RU",{maximumFractionDigits:1,minimumFractionDigits:1});
const plural=new Intl.PluralRules("ru-RU");
const women=value=>count.format(value)+" "+({one:"женщина",few:"женщины",many:"женщин",other:"женщин"}[plural.select(value)]);
const change=new Intl.NumberFormat("ru-RU",{maximumFractionDigits:2,minimumFractionDigits:2,signDisplay:"exceptZero"});
const rate=value=>value===null?"Не задан":decimal.format(value)+"%";
const clamp=value=>Math.max(0,Math.min(1,value));
const quiet=()=>window.matchMedia?.("(prefers-reduced-motion: reduce)").matches??true;
let plan=null,selected=-1,expanded=false,motions=[];
let stamp="",month="",period="";
function stop(){motions.forEach(animation=>animation.cancel());motions=[];}
function move(element,property,value,animate){
  if(animate&&!quiet()&&window.anime){motions.push(window.anime.animate(element,{[property]:value,duration:240,ease:"outQuint"}));}
  else element.style.transform=property==="scaleX"?`scaleX(${value})`:`translateX(${value})`;
}
function selectedRow(){return selected<0?plan.total:plan.departments[selected];}
function render(animate=false,announce=false){
  if(!plan)return;
  stop();$("annualBody").style.opacity="1";$("annualBody").style.transform="none";const row=selectedRow();
  $("annualScope").textContent=row.name;$("annualReset").hidden=selected<0;
  $("annualSummary").textContent=row.name+" · "+women(row.yearActual);
  $("annualHeadlineRate").textContent=rate(row.yearRate);
  $("annualActual").textContent=count.format(row.yearActual);$("annualYearPlan").textContent=count.format(row.yearPlan);
  $("annualTrend").textContent=row.trend===null?"Нет базы сравнения":change.format(row.trend)+" п.п.";
  $("annualTrend").classList.toggle("is-positive",row.trend>0);$("annualTrend").classList.toggle("is-negative",row.trend<0);
  $("annualTrendPeriod").textContent="к выполнению плана на начало "+period;
  $("annualShare").textContent=rate(row.yearShare);
  $("annualPriorVolume").textContent="До "+period+": "+count.format(row.previousActual);
  $("annualMonthVolume").textContent="За месяц: "+count.format(row.monthActual);
  $("annualExpectedShare").textContent=row.yearPlan?rate(plan.yearFraction*100):"Не задан";
  const gap=row.gapToDate;
  $("annualGap").textContent=!row.yearPlan?"Введите годовой план, чтобы увидеть сравнение.":gap>0?"До плана на "+stamp+": "+women(Math.ceil(gap)):gap<0?"План на "+stamp+" превышен на "+women(Math.floor(-gap)):"План на "+stamp+" выполнен.";
  $("annualGap").classList.toggle("is-behind",row.yearPlan>0&&gap>0);
  $("annualMonthTitle").textContent=month+" · на "+stamp;
  $("annualMonthActual").textContent=women(row.monthActual);
  $("annualMonthPlan").textContent="План полного месяца: "+count.format(row.monthPlan);
  $("annualMonthRate").textContent=rate(row.monthRate);
  const fraction=row.yearPlan?clamp(row.yearActual/row.yearPlan):0,prior=row.yearPlan?clamp(row.previousActual/row.yearPlan):0;
  move($("annualFactFill"),"scaleX",fraction,animate);
  move($("annualPriorFill"),"scaleX",prior,animate);
  move($("annualFactMark"),"translateX",fraction*100+"%",animate);
  move($("annualTargetMark"),"translateX",plan.yearFraction*100+"%",false);
  $("annualFactMark").hidden=!row.yearPlan;$("annualTargetMark").hidden=!row.yearPlan;
  $("annualRuler").setAttribute("aria-label",row.name+": выполнено "+women(row.yearActual)+" из "+count.format(row.yearPlan)+" по годовому плану. "+rate(row.yearShare)+" годового плана; "+rate(row.yearRate)+" к плану на "+stamp+".");
  $("annualRuler").title="До "+period+": "+women(row.previousActual)+"; за месяц: "+women(row.monthActual)+". Суммарная доля годового плана: "+rate(row.yearShare)+".";
  for(const button of $("annualUnits").querySelectorAll("button[data-annual-unit]"))button.setAttribute("aria-pressed",Number(button.dataset.annualUnit)===selected);
  window.ClinicPace?.select(selected,animate);
  if(announce)$("annualAnnouncement").textContent=row.name+": "+rate(row.yearRate)+" к плану на дату, "+women(row.yearActual)+".";
}
function select(index,keyboard=false){
  if(!plan||!Number.isInteger(index)||index< -1||index>=plan.departments.length)return;
  selected=index;render(!keyboard,true);
}
function expand(value,keyboard=false){
  if(!plan)return;expanded=value;stop();window.ClinicPace?.settle();
  $("annualBody").hidden=!value;$("annualToggle").setAttribute("aria-expanded",value);
  $("annualToggleLabel").textContent=value?"Свернуть":"Раскрыть план";
  if(value){
    if(!keyboard&&!quiet()&&window.anime){
      // The two marks travel on one shared scale; values are always immediately readable.
      $("annualFactFill").style.transform="scaleX(0)";$("annualPriorFill").style.transform="scaleX(0)";
      $("annualFactMark").style.transform="translateX(0%)";
      render(true);
      motions.push(window.anime.animate($("annualBody"),{opacity:[0,1],translateY:[8,0],duration:220,ease:"outQuint"}));
    }else{ $("annualBody").style.opacity="1";$("annualBody").style.transform="none";render(false); }
  }
}
function table(){
  const headings=["Подразделение","План года, женщин","До месяца: факт","До месяца: %","План месяца","За месяц: факт","За месяц: % на дату","С начала года: факт","С начала года: % на дату","Тренд, п.п."];
  const html=[plan.total,...plan.departments].map(row=>'<tr><th scope="row">'+escape(row.name)+'</th>'+[count.format(row.yearPlan),count.format(row.previousActual),rate(row.priorRate),count.format(row.monthPlan),count.format(row.monthActual),rate(row.monthRate),count.format(row.yearActual),rate(row.yearRate),row.trend===null?"Нет базы":change.format(row.trend)].map(value=>'<td>'+escape(value)+'</td>').join("")+'</tr>').join("");
  $("annualSourceTable").innerHTML='<thead><tr>'+headings.map(h=>'<th scope="col">'+h+'</th>').join("")+'</tr></thead><tbody>'+html+'</tbody>';
}
function init(data,meta){
  if(!data)return;plan=data;selected=-1;
  const date=new Date(plan.asOf+"T12:00:00Z");
  stamp=date.toLocaleDateString("ru-RU",{day:"numeric",month:"long",timeZone:"UTC"});
  month=date.toLocaleDateString("ru-RU",{month:"long",timeZone:"UTC"});month=month[0].toUpperCase()+month.slice(1);
  period=["января","февраля","марта","апреля","мая","июня","июля","августа","сентября","октября","ноября","декабря"][date.getUTCMonth()];
  $("annualTitle").textContent="Годовой план ДОРЗ · "+plan.year;
  $("annualStamp").textContent="На "+stamp+" "+plan.year;
  $("annualBasis").textContent=plan.basis+" Процент на дату = факт ÷ план на дату × 100. Тренд рассчитан по неокруглённым процентам относительно начала месяца. Число женщин из годового отчёта не складывается с количеством услуг месячной выгрузки.";
  $("annualContext").textContent="Отдельный отчёт ДОРЗ на "+stamp+" "+plan.year+". Фильтры месячных услуг не изменяют этот отчёт; консультацию можно выбрать здесь."+(plan.asOf!==meta.dataEnd?" Дата плана отличается от последней даты выгрузки услуг.":"");
  $("annualWarnings").textContent=plan.warnings.join(" ");$("annualWarnings").hidden=!plan.warnings.length;
  $("annualUnits").innerHTML=plan.departments.map((row,index)=>'<button type="button" class="annual-unit" data-annual-unit="'+index+'" aria-pressed="false"><span class="annual-unit-name">'+escape(row.name)+'</span><strong>'+rate(row.yearRate)+'</strong><span class="annual-unit-track" aria-hidden="true"><i style="transform:scaleX('+clamp((row.yearRate||0)/100)+')"></i></span><small>'+women(row.yearActual)+' с начала года</small><span class="annual-unit-change">'+(row.trend===null?"Нет базы":change.format(row.trend)+" п.п.")+'</span><span class="clinic-unit-pace" data-clinic-pace="'+index+'"></span></button>').join("");
  $("annualUnits").addEventListener("click",event=>{const button=event.target.closest("button[data-annual-unit]");if(button)select(Number(button.dataset.annualUnit),event.detail===0);});
  $("annualReset").addEventListener("click",event=>select(-1,event.detail===0));
  $("annualToggle").addEventListener("click",event=>expand(!expanded,event.detail===0));
  table();render();setVisible(true);
}
function setVisible(visible){if(!plan)return;$("annualPlan").hidden=!visible;render(false);}
window.AnnualPlan={init,setVisible,select,expand};
})();
