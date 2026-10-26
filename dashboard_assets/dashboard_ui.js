/* Shared interaction primitives. No network, framework or per-frame data rendering. */
(function (root) {
  'use strict';
  const ramps = {
    light: ['#eff6ff','#dbeafe','#bfdbfe','#93c5fd','#60a5fa','#3b82f6','#2563eb','#1d4ed8','#1e40af'],
    dark: ['#202b3a','#293c54','#345477','#416d98','#5088bb','#62a0d0','#7bb5e4','#93c5fd','#bfdbfe']
  };
  function luminance(hex) {
    const rgb=hex.slice(1).match(/../g).map(v=>parseInt(v,16)/255).map(v=>v<=.04045?v/12.92:((v+.055)/1.055)**2.4);
    return rgb[0]*.2126+rgb[1]*.7152+rgb[2]*.0722;
  }
  function heatColor(value,maximum,dark=false) {
    const ramp=ramps[dark?'dark':'light'];
    const level=Math.min(ramp.length-1,Math.floor(Math.sqrt(Math.max(0,value)/Math.max(1,maximum))*(ramp.length-1)));
    const background=ramp[level],lightness=luminance(background);
    const black=(lightness+.05)/.05,white=1.05/(lightness+.05);
    return {background,color:black>=white?'#000000':'#ffffff',contrast:Math.max(black,white)};
  }
  const api={heatColor,luminance,ramps};
  if(typeof module!=='undefined'&&module.exports)module.exports=api;
  if(!root?.document)return;
  const motions=new Map(),media=root.matchMedia('(prefers-reduced-motion: reduce)');
  const instant=()=>media.matches||document.documentElement.dataset.input==='keyboard';
  function transform(element,value,animate=true,duration=220) {
    if(!element)return;
    const from=getComputedStyle(element).transform;
    motions.get(element)?.cancel();
    motions.delete(element);
    element.style.transform=value;
    if(!animate||instant()||!element.animate)return;
    const animation=element.animate([{transform:from==='none'?value:from},{transform:value}],{duration,easing:'cubic-bezier(0.23,1,0.32,1)'});
    motions.set(element,animation);
    const release=()=>{if(motions.get(element)===animation)motions.delete(element);};
    animation.finished.then(release,release);
  }
  function indicator(container,target,animate=true) {
    let marker=container.querySelector(':scope > .selection-marker');
    if(!marker){marker=document.createElement('i');marker.className='selection-marker';marker.setAttribute('aria-hidden','true');container.append(marker);}
    marker.hidden=!target;if(!target)return;
    const c=container.getBoundingClientRect(),t=target.getBoundingClientRect();
    transform(marker,`translate3d(${t.left-c.left+container.scrollLeft}px,${t.bottom-c.top+container.scrollTop-2}px,0) scaleX(${t.width})`,animate);
  }
  function scrollInside(container,target) {
    if(!container||!target)return;
    const c=container.getBoundingClientRect(),t=target.getBoundingClientRect();
    if(t.left<c.left||t.right>c.right)container.scrollLeft+=t.left-c.left-(c.width-t.width)/2;
  }
  function roving(container,selector,grid=false) {
    const items=()=>[...container.querySelectorAll(selector)];
    const activate=target=>{items().forEach(n=>n.tabIndex=n===target?0:-1);};
    container.addEventListener('focusin',e=>{if(e.target.matches(selector))activate(e.target);});
    container.addEventListener('keydown',e=>{
      if(!['ArrowLeft','ArrowRight','ArrowUp','ArrowDown','Home','End'].includes(e.key)||!e.target.matches(selector))return;
      const all=items();let next=all.indexOf(e.target);
      if(grid){
        const rows=[...container.querySelectorAll('tbody tr')].map(r=>[...r.querySelectorAll(selector)]).filter(r=>r.length);
        const row=rows.findIndex(r=>r.includes(e.target)),col=rows[row]?.indexOf(e.target);if(row<0)return;
        let rr=row,cc=col;
        if(e.key==='ArrowUp')rr=Math.max(0,row-1);if(e.key==='ArrowDown')rr=Math.min(rows.length-1,row+1);
        if(e.key==='ArrowLeft')cc=Math.max(0,col-1);if(e.key==='ArrowRight')cc=Math.min(rows[rr].length-1,col+1);
        if(e.key==='Home'){cc=0;if(e.ctrlKey)rr=0;}if(e.key==='End'){if(e.ctrlKey)rr=rows.length-1;cc=rows[rr].length-1;}
        next=all.indexOf(rows[rr][Math.min(cc,rows[rr].length-1)]);
      }else next=e.key==='Home'?0:e.key==='End'?all.length-1:Math.max(0,Math.min(all.length-1,next+(['ArrowRight','ArrowDown'].includes(e.key)?1:-1)));
      if(next<0)return;e.preventDefault();activate(all[next]);all[next].focus({preventScroll:true});
      all[next].scrollIntoView({block:'nearest',inline:'nearest'});
    });
  }
  document.addEventListener('keydown',()=>document.documentElement.dataset.input='keyboard',true);
  document.addEventListener('pointerdown',()=>document.documentElement.dataset.input='pointer',true);
  media.addEventListener('change',()=>{for(const a of motions.values())a.cancel();motions.clear();});
  root.DashboardUI={...api,transform,indicator,roving,scrollInside,instant};
})(typeof window==='undefined'?null:window);
