(() => {
'use strict';
const utc=value=>new Date(value+'T12:00:00Z');
function monthDays(date){
  if(!/^\d{4}-\d{2}-\d{2}$/.test(date)||!Number.isFinite(utc(date).getTime())||utc(date).toISOString().slice(0,10)!==date)throw new Error('Некорректная дата');
  const prefix=date.slice(0,7),cursor=utc(prefix+'-01'),days=[];
  while(cursor.toISOString().slice(0,7)===prefix){
    const weekday=cursor.getUTCDay();
    days.push({date:cursor.toISOString().slice(0,10),weekday,units:weekday===0||weekday===6?1:2});
    cursor.setUTCDate(cursor.getUTCDate()+1);
  }
  return days;
}
function calculate({target,date,quantities={},snapshotStart,snapshotEnd}){
  if(!Number.isSafeInteger(target)||target<0||target>1e9)throw new Error('Месячный план должен быть целым неотрицательным числом');
  const days=monthDays(date),first=days[0].date,index=days.findIndex(d=>d.date===date);
  if(!snapshotStart||snapshotStart>first)return {ready:false,reason:'incomplete-month',first};
  if(!snapshotEnd||date>snapshotEnd)return {ready:false,reason:'outside-snapshot',first};
  const quantity=day=>{
    const value=quantities[day]??0;
    if(!Number.isSafeInteger(value)||value<0)throw new Error('Некорректное количество услуг');
    return value;
  };
  let cumulative=0;
  const totalUnits=days.reduce((s,d)=>s+d.units,0);
  const before=days.slice(0,index).reduce((s,d)=>s+quantity(d.date),0);
  const actual=quantity(date),monthActual=before+actual;
  const remainingUnits=days.slice(index).reduce((s,d)=>s+d.units,0);
  const remainingAtStart=Math.max(0,target-before);
  // Integers in half-day units avoid rounding Sat+Sun into two full days.
  const dailyTarget=Math.ceil(remainingAtStart*days[index].units/remainingUnits);
  const dailyLeft=Math.max(0,dailyTarget-actual);
  const expectedByDate=target*days.slice(0,index+1).reduce((s,d)=>s+d.units,0)/totalUnits;
  const history=days.map(day=>{
    cumulative+=day.units;
    const known=day.date<=snapshotEnd;
    return {...day,known,actual:known?quantity(day.date):null,base:target*day.units/totalUnits,expected:target*cumulative/totalUnits};
  });
  return {ready:true,date,target,days:history,before,actual,monthActual,dailyTarget,dailyLeft,
    baseDaily:target*days[index].units/totalUnits,dayWeight:days[index].units/2,
    totalWeight:totalUnits/2,remainingWeight:remainingUnits/2,expectedByDate,
    monthLeft:Math.max(0,target-monthActual),overDaily:Math.max(0,actual-dailyTarget),
    ratio:dailyTarget?actual/dailyTarget:null,monthRatio:target?monthActual/target:null,
    status:target===0?'no-target':before>=target?'month-done':dailyLeft===0?'day-done':'in-progress'};
}
window.DailyPaceCore={monthDays,calculate};
})();
