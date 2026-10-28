/* Prototipo de planificación Cocina. El catálogo completo se conectará desde Pruebas.xlsx. */
const demoRecipes = [
  {id:'MEN-4C5E8339',name:'Albóndigas',moment:'Comida',single:true,freq:'2',body:'Receta de albóndigas.\n\nEl detalle completo se importará desde el documento Word.'},
  {id:'MEN-B09C9240',name:'Arroz con atún/calamares',moment:'Comida',single:true,freq:'4'},
  {id:'MEN-636909BB',name:'Bocadillo',moment:'Cena',single:true,freq:'1'},
  {id:'MEN-F1704A2A',name:'Boloñesa',moment:'Comida',single:true,freq:'2'},
  {id:'MEN-149627A6',name:'Callos con garbanzos',moment:'Comida',single:true,freq:'1'},
  {id:'MEN-7B8248C4',name:'Canelones',moment:'Comida',single:true,freq:'1'},
  {id:'MEN-A0E0A499',name:'Carbonara',moment:'Comida',single:true,freq:'min 2',freqExact:2,onlyWeekday:5,minGapDays:7},
  {id:'MEN-70D9AC80',name:'Champiñones',moment:'Cena',single:false,freq:'2'},
  {id:'MEN-29EAE5F6',name:'Chipirones',moment:'Comida',single:true,freq:'3'},
  {id:'MEN-26A5BDAD',name:'Chuleta con espárragos',moment:'Cena',single:true,freq:'1'},
  {id:'MEN-ECF78A19',name:'Chuletas con patatas',moment:'Cena',single:true,freq:'1'},
  {id:'MEN-97A6CA0C',name:'Cocido madrileño',moment:'Comida',single:true,freq:'3'},
  {id:'MEN-51390396',name:'Conejo',moment:'Comida',single:true,freq:'2'},
  {id:'MEN-047762CC',name:'Ensalada',moment:'Cena',single:false,freq:'2'},
  {id:'MEN-817B40F4',name:'Entraña',moment:'Cena',single:true,freq:'1'},
  {id:'MEN-D2DD3D9A',name:'Fabes',moment:'Comida',single:true,freq:'3'},
  {id:'MEN-D880097A',name:'Fideos con caballa',moment:'Comida',single:true,freq:'3'},
  {id:'MEN-03D0D732',name:'Gambas',moment:'Cena',single:false,freq:'3'},
  {id:'MEN-21A15EF2',name:'Gazpacho',moment:'Comida',single:true,freq:'2'},
  {id:'MEN-9A1E795A',name:'Hamburguesa',moment:'Cena',single:true,freq:'1'},
  {id:'MEN-E9477DA0',name:'Huevos con patatas',moment:'Cena',single:true,freq:'2'},
  {id:'MEN-DD4AE5DF',name:'Lasaña',moment:'Comida',single:true,freq:'2'},
  {id:'MEN-54451846',name:'Lentejas',moment:'Comida',single:true,freq:'3'},
  {id:'MEN-BD5B420F',name:'Nachos',moment:'Cena',single:false,freq:'2'},
  {id:'MEN-B94C0F1C',name:'Paella',moment:'Comida',single:true,freq:'3'},
  {id:'MEN-1E92375D',name:'Perritos calientes',moment:'Cena',single:true,freq:'2'},
  {id:'MEN-1F6CCD2B',name:'Pizza',moment:'Cena',single:true,freq:'1'},
  {id:'MEN-93DDBF5E',name:'Provolone',moment:'Cena',single:false,freq:'2'},
  {id:'MEN-BB0ED6D5',name:'Sepia',moment:'Cena',single:false,freq:'1'},
  {id:'MEN-7433B712',name:'Setas',moment:'Cena',single:false,freq:'2'},
  {id:'MEN-579DB7D6',name:'Sandwich',moment:'Cena',single:true,freq:'1'},
  {id:'MEN-B754598B',name:'Solomillo de cerdo',moment:'Cena',single:true,freq:'2'},
  {id:'MEN-176F21F1',name:'Sopa Brick',moment:'Cena',single:false,freq:'min 2'},
  {id:'MEN-8BF3C6A7',name:'Sopa Cebolla',moment:'Cena',single:false,freq:'min 2'},
  {id:'MEN-D782114A',name:'Tacos',moment:'Cena',single:true,freq:'2'},
  {id:'MEN-AF443EB4',name:'Tortilla francesa',moment:'Cena',single:true,freq:'2'},
  {id:'MEN-17DD28DA',name:'Tortilla patata',moment:'Cena',single:true,freq:'2'}
];
const catalog = window.COCINA_DATA || { menu: demoRecipes, recipes: [], headings: [], wordText: '' };
const menuForApp=(()=>{const rows=[...catalog.menu],nachos=rows.filter(r=>r.name.trim().toLocaleLowerCase('es')==='nachos');if(!nachos.length)return rows;let replaced=false;return rows.flatMap(r=>{if(r.name.trim().toLocaleLowerCase('es')!=='nachos')return [r];if(replaced)return [];replaced=true;return [{...r,moment:'Cena',single:false,freq:'max 2',active:nachos.some(x=>/^si$/i.test(x.active))?'Si':'No',notes:'Regla consolidada para la app: máximo 2 al mes, solo cena; automático requiere pareja. Manualmente puede quedar como plato único.'}]})})();
function normalizeHeadingTitle(value){return String(value||'').normalize('NFD').replace(/\p{M}/gu,'').toLocaleLowerCase('es').replace(/[^\p{L}\p{N}]+/gu,' ').trim().split(/\s+/).filter(word=>!['de','del','la','el','los','las'].includes(word)).join(' ')}
const familyOnlyNames=new Set(['Aceites','Ajo','Arroz','Boquerones','Carne','Marisco','Patata','Pasta','Pescados','Pollo','Salsas','Tartar'].map(normalizeHeadingTitle));
function isFamilyOnlyName(name){return familyOnlyNames.has(normalizeHeadingTitle(name))}
function wordBodyFor(name){const target=normalizeHeadingTitle(name),matches=[];catalog.headings.forEach((h,index)=>{if(normalizeHeadingTitle(h.title)===target)matches.push(index)});if(!matches.length)return '';const subtreeLength=index=>{const root=catalog.headings[index];let total=0;for(let i=index;i<catalog.headings.length;i++){const h=catalog.headings[i];if(i>index&&h.level<=root.level)break;total+=(h.body||'').length}return total};const collect=index=>{const root=catalog.headings[index],parts=[];let hasText=false;for(let i=index;i<catalog.headings.length;i++){const h=catalog.headings[i];if(i>index&&h.level<=root.level)break;if(h.body?.trim())hasText=true;if(i===index){if(h.body?.trim())parts.push(h.body.trim());continue}if(h.title?.trim())parts.push(h.title.trim());if(h.body?.trim())parts.push(h.body.trim())}return hasText?parts.join('\n').trim():''};const rootIndex=matches.sort((a,b)=>subtreeLength(b)-subtreeLength(a))[0],direct=collect(rootIndex);if(direct)return direct;const root=catalog.headings[rootIndex],parentName=root.parent;if(!parentName)return '';const parentIndex=catalog.headings.findIndex(h=>normalizeHeadingTitle(h.title)===normalizeHeadingTitle(parentName));if(parentIndex<0)return '';const sharedSection=/^(ingrediente|ingredientes|preparado|preparada|preparacion|elaboracion|procedimiento|pasos?|limpieza|montado|montaje|recomendaciones|consejos|notas?)(\s|$)/;const shared=[];for(let i=parentIndex+1;i<catalog.headings.length;i++){const h=catalog.headings[i];if(h.level<=catalog.headings[parentIndex].level)break;if(h.level===root.level&&sharedSection.test(normalizeHeadingTitle(h.title))){const section=collect(i);if(section)shared.push(section)}}return shared.join('\n').trim()}
const menuNames = new Map();
const recipes = menuForApp.map((item,index)=>{
  const occurrence=(menuNames.get(item.id)||0)+1;menuNames.set(item.id,occurrence);
  const sourceRecipe=catalog.recipes.find(r=>r.id===item.recipeId)||catalog.recipes.find(r=>r.name===item.name);
  const body=wordBodyFor(item.name);
  return { ...item, id:`${item.id||item.recipeId||item.name}#${occurrence}`, sourceId:item.id, moment:item.moment, single:item.single, familyOnly:isFamilyOnlyName(item.name)||/categor[ií]a\s*\/\s*tema|informaci[oó]n/i.test(sourceRecipe?.type||''), manualSingleAllowed:item.name.trim().toLocaleLowerCase('es')==='nachos', freq:item.freq, active:/^si$/i.test(item.active), category:sourceRecipe?.category||'', type:sourceRecipe?.type||'Pendiente de clasificación', linkStatus:sourceRecipe?.linkStatus||'', body, ...(item.name==='Carbonara'?{freq:'min 2',freqExact:2,onlyWeekday:5,minGapDays:7}:{}) };
});
const catalogRecipes=catalog.recipes.map((item,index)=>{
  const body=wordBodyFor(item.name);
  return {...item,body,familyOnly:isFamilyOnlyName(item.name)||/categor[ií]a\s*\/\s*tema|informaci[oó]n/i.test(item.type||''),id:item.id||`recipe-${index}`};
});
const weekdays=['Lun','Mar','Mié','Jue','Vie','Sáb','Dom'];
const monthNames=['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre'];
const $=s=>document.querySelector(s);
const grid=$('#calendarGrid'), alertBox=$('#alerts');
let viewed=new Date(); viewed.setDate(1);
const storeKey='cocina-calendar-v1';
const scheduleVersion=4;
let saved=JSON.parse(localStorage.getItem(storeKey)||'{}');
function migrateSavedNachos(){const row=catalog.menu.find(r=>r.name.trim().toLocaleLowerCase('es')==='nachos'),canonical=row?`${row.id}#1`:'' ,legacy=row?`${row.id}#2`:'';if(!legacy)return;for(const month of Object.values(saved)){for(const [name,manual] of [['manual',true],['schedule',false]]){const entries=month?.[name];if(!entries)continue;for(const [k,ids] of Object.entries(entries)){if(!Array.isArray(ids))continue;if(manual)entries[k]=[...new Set(ids.map(id=>id===legacy?canonical:id))];else if(ids.includes(legacy)){entries[k]=ids.filter(id=>id!==legacy);month.needsRegeneration=true}}}}}
migrateSavedNachos();
function migrateSavedFamilies(){const familyIds=new Set(recipes.filter(r=>r.familyOnly).map(r=>r.id));for(const month of Object.values(saved)){for(const field of ['manual','schedule']){const entries=month?.[field];if(!entries)continue;for(const [k,ids] of Object.entries(entries)){if(!Array.isArray(ids))continue;const cleaned=ids.filter(id=>!familyIds.has(id));if(cleaned.length!==ids.length){entries[k]=cleaned;if(field==='schedule')month.needsRegeneration=true}}}}}
migrateSavedFamilies();
function key(y,m,d,slot){return `${y}-${String(m+1).padStart(2,'0')}-${String(d).padStart(2,'0')}-${slot}`}
function parseFreq(r){const n=Number(String(r.freq).replace(/[^0-9]/g,''))||1,min=/^\s*min/i.test(r.freq);return {count:r.freqExact??n,min,max:!min}}
function allowed(r,slot){if(r.familyOnly)return false;if(slot==='lunch')return r.single && ['Comida','Ambos'].includes(r.moment);return ['Cena','Ambos'].includes(r.moment)}
function sourceWarnings(){const warnings=[],rawNachos=catalog.menu.filter(r=>r.name.trim().toLocaleLowerCase('es')==='nachos'),families=catalog.menu.filter(r=>isFamilyOnlyName(r.name));if(rawNachos.length>1)warnings.push(`El Excel mantiene ${rawNachos.length} filas de Nachos. En este prototipo se consolidan como: solo cena, máximo 2 al mes, requieren pareja en automático; el cambio manual puede dejarlos solos.`);if(families.length)warnings.push(`Filas de familias en Menu (${[...new Set(families.map(r=>r.name))].join(', ')}) se mantienen para consulta, pero se excluyen de la planificación automática porque no son platos concretos.`);return warnings}
function daysOfMonth(y,m){const a=[];for(let d=1;d<=new Date(y,m+1,0).getDate();d++){const dt=new Date(y,m,d);if(dt.getDay()>=1&&dt.getDay()<=5)a.push(dt)}return a}
function recipeKey(r){return r.name.trim().toLocaleLowerCase('es')}
function monthUsage(month){const used=new Set(),keys=new Set([...Object.keys(month?.schedule||{}),...Object.keys(month?.manual||{})]);for(const k of keys){const ids=month?.manual?.[k]??month?.schedule?.[k]??[];for(const id of ids){const r=recipes.find(x=>x.id===id);if(r)used.add(recipeKey(r))}}return used}
function hasMonthAssignments(month){return Object.values(month?.schedule||{}).some(ids=>Array.isArray(ids)&&ids.length)||Object.values(month?.manual||{}).some(ids=>Array.isArray(ids)&&ids.length)}
function monthRandom(seed){let value=2166136261;for(const char of seed){value^=char.charCodeAt(0);value=Math.imul(value,16777619)}return ()=>{value+=0x6D2B79F5;let next=value;next=Math.imul(next^next>>>15,next|1);next^=next+Math.imul(next^next>>>7,next|61);return ((next^next>>>14)>>>0)/4294967296}}
function buildScheduleFor(y,m,previousUsed=new Set(),variation=0){
 const random=monthRandom(`${catalog.generatedAt||'cocina'}|${y}-${m}|${variation}`);
 const workdays=daysOfMonth(y,m),dinners=workdays.filter(dt=>dt.getDay()!==5),schedule={},counts=new Map(),lastUsed=new Map(),weekFamilies=new Map(),warnings=[],candidates=recipes.filter(r=>r.active&&!r.familyOnly),carbonara=candidates.find(r=>r.name==='Carbonara'),unfilled={lunch:0,dinner:0};
 const monthKey=`${y}-${m+1}`,dayMs=86400000;
 const countOf=r=>counts.get(r.id)||0;
 const canUse=(r,dt)=>{const f=parseFreq(r),last=lastUsed.get(r.id),cap=f.min?Infinity:f.count;return countOf(r)<cap&&(!last||dt-last>=(r.minGapDays||6)*dayMs)};
 const weekKey=dt=>`${y}-${m}-${Math.floor((dt.getDate()-1)/7)}`;
 const scoreOf=(r,dt)=>{const family=String(r.category||'').trim(),inWeek=family?(weekFamilies.get(`${weekKey(dt)}|${family}`)||0):0;let score=random();if(previousUsed.has(recipeKey(r)))score-=5;score-=inWeek*2.5;score-=countOf(r)*1.4;return score};
 const rank=(list,dt)=>list.map(r=>({r,score:scoreOf(r,dt)})).sort((a,b)=>b.score-a.score).map(x=>x.r);
 const place=(dt,slot,items)=>{schedule[key(y,m,dt.getDate(),slot)]=items.map(r=>r.id);for(const r of items){counts.set(r.id,countOf(r)+1);lastUsed.set(r.id,dt);const family=String(r.category||'').trim();if(family){const k=`${weekKey(dt)}|${family}`;weekFamilies.set(k,(weekFamilies.get(k)||0)+1)}}};
 // Carbonara is a fixed rule: two Fridays, spread as far apart as possible.
 if(carbonara){const fridays=workdays.filter(dt=>dt.getDay()===5),target=2,chosen=[];for(let i=0;i<target&&i<fridays.length;i++){const index=target===1?0:Math.round(i*(fridays.length-1)/(target-1)),dt=fridays[index];if(!chosen.some(x=>x.getTime()===dt.getTime())&&(!chosen.length||dt-chosen[chosen.length-1]>=7*dayMs))chosen.push(dt)}chosen.forEach(dt=>place(dt,'lunch',[carbonara]));if(chosen.length<target)warnings.push(`Carbonara: solo se encontraron ${chosen.length} viernes separados al menos una semana; se requieren 2.`)}
 for(const dt of workdays){const k=key(y,m,dt.getDate(),'lunch');if(schedule[k])continue;const options=candidates.filter(r=>r!==carbonara&&allowed(r,'lunch')&&canUse(r,dt)),mandatory=options.filter(r=>parseFreq(r).min&&countOf(r)<parseFreq(r).count),chosen=rank(mandatory.length?mandatory:options,dt)[0];if(chosen)place(dt,'lunch',[chosen]);else unfilled.lunch++}
 for(const dt of dinners){const options=candidates.filter(r=>allowed(r,'dinner')&&canUse(r,dt)),mandatoryNo=options.filter(r=>!r.single&&parseFreq(r).min&&countOf(r)<parseFreq(r).count),mandatorySingle=options.filter(r=>r.single&&parseFreq(r).min&&countOf(r)<parseFreq(r).count);let bundle=null;
   if(mandatoryNo.length){const first=rank(mandatoryNo,dt)[0],partners=rank(options.filter(r=>!r.single&&r.id!==first.id),dt);if(partners.length)bundle=[first,partners[0]]}
   if(!bundle&&mandatorySingle.length)bundle=[rank(mandatorySingle,dt)[0]];
   if(!bundle){const singles=options.filter(r=>r.single).map(r=>({items:[r],score:scoreOf(r,dt)}));const pairs=[];const nonSingles=options.filter(r=>!r.single);for(let i=0;i<nonSingles.length;i++)for(let j=i+1;j<nonSingles.length;j++){const a=nonSingles[i],b=nonSingles[j];pairs.push({items:[a,b],score:scoreOf(a,dt)+scoreOf(b,dt)+0.35})}const available=[...singles,...pairs].sort((a,b)=>b.score-a.score);bundle=available[0]?.items||null}
   if(bundle)place(dt,'dinner',bundle);else unfilled.dinner++
 }
 const missed=[];for(const r of candidates){const f=parseFreq(r);if(f.min&&countOf(r)<f.count)missed.push(`${r.name} (${countOf(r)}/${f.count})`)}if(missed.length)warnings.push(`No se pudo cumplir la frecuencia mínima de: ${missed.join('; ')}.`);if(unfilled.lunch||unfilled.dinner)warnings.push(`Quedaron espacios sin plato: ${unfilled.lunch} comidas y ${unfilled.dinner} cenas. Faltan candidatos compatibles o sus límites mensuales ya se agotaron.`);return {schedule,warnings,monthKey}
}
function previousMonthUsage(y,m){const previousDate=new Date(y,m-1,1),py=previousDate.getFullYear(),pm=previousDate.getMonth(),previousKey=`${py}-${pm+1}`;let previous=saved[previousKey];if((previous?.algorithmVersion!==scheduleVersion||previous?.needsRegeneration||!hasMonthAssignments(previous))&&!previous?.deliberatelyCleared){const olderDate=new Date(py,pm-1,1),older=saved[`${olderDate.getFullYear()}-${olderDate.getMonth()+1}`],built=buildScheduleFor(py,pm,monthUsage(older),previous?.variation||0);previous={schedule:built.schedule,manual:previous?.manual||{},variation:previous?.variation||0,generatedForRotation:true,algorithmVersion:scheduleVersion};saved[previousKey]=previous}return monthUsage(previous)}
function generate(newVariation=false){const y=viewed.getFullYear(),m=viewed.getMonth(),monthKey=`${y}-${m+1}`,previous=saved[monthKey],variation=(previous?.variation||0)+(newVariation?1:0),previousUsed=previousMonthUsage(y,m),built=buildScheduleFor(y,m,previousUsed,variation);saved[monthKey]={schedule:built.schedule,manual:previous?.manual||{},variation,deliberatelyCleared:false,algorithmVersion:scheduleVersion};persist();render();showWarnings(built.warnings)}
function persist(){localStorage.setItem(storeKey,JSON.stringify(saved))}
function showWarnings(list){const combined=[...new Set(list)];alertBox.hidden=!combined.length;alertBox.innerHTML=combined.length?`<details><summary>Revisar ${combined.length} aviso${combined.length===1?'':'s'} del calendario</summary><ul>${combined.map(x=>`<li>${escapeHTML(x)}</li>`).join('')}</ul></details>`:''}
function monthData(){return saved[`${viewed.getFullYear()}-${viewed.getMonth()+1}`]||{schedule:{},manual:{}}}
function render(){const y=viewed.getFullYear(),m=viewed.getMonth();$('#monthTitle').textContent=`${monthNames[m][0].toUpperCase()+monthNames[m].slice(1)} ${y}`;const first=new Date(y,m,1),offset=(first.getDay()+6)%7,total=new Date(y,m+1,0).getDate();let html=weekdays.map(d=>`<div class="weekday">${d}</div>`).join('');const cal=monthData();
 for(let i=0;i<offset;i++)html+='<div class="day-cell empty"></div>';
 for(let d=1;d<=total;d++){const dt=new Date(y,m,d),wd=dt.getDay(),weekend=wd===0||wd===6,today=dt.toDateString()===new Date().toDateString();html+=`<div class="day-cell ${weekend?'weekend':''} ${today?'today':''}"><div class="day-number">${d}<span class="day-name">${dt.toLocaleDateString('es-ES',{weekday:'long'})}</span></div>`;
   if(!weekend){html+=slotHtml(d,'lunch','Comida',cal);if(wd===5)html+='<div class="meal-slot blocked"><span class="meal-label">Cena</span><span class="meal-dishes">Libre</span></div>';else html+=slotHtml(d,'dinner','Cena',cal)}html+='</div>'}
 const trailing=(7-((offset+total)%7))%7;for(let i=0;i<trailing;i++)html+='<div class="day-cell empty"></div>';grid.innerHTML=html;bindSlots();renderToday();}
function slotHtml(d,slot,label,cal){
 const k=key(viewed.getFullYear(),viewed.getMonth(),d,slot),ids=cal.manual?.[k]??cal.schedule?.[k]??[],names=ids.map(id=>recipes.find(r=>r.id===id)).filter(r=>r&&!r.familyOnly),manual=Object.hasOwn(cal.manual||{},k);
 return `<div class="meal-slot ${slot==='dinner'?'dinner':''} ${names.length?'':'empty-slot'}" data-day="${d}" data-slot="${slot}" aria-label="${label} del día ${d}"><span class="meal-dishes">${names.length?names.map(r=>`<button class="dish-link" data-recipe="${r.id}">${escapeHTML(r.name)}</button>`).join(' + '):'＋ Añadir plato'}${manual?'<span class="manual-badge">Cambio manual</span>':''}</span><button type="button" class="edit-slot-button" aria-label="Cambiar ${label.toLowerCase()} del día ${d}">✎</button></div>`
}
function bindSlots(){grid.querySelectorAll('.meal-slot[data-day]').forEach(el=>{el.addEventListener('click',e=>{const dish=e.target.closest('[data-recipe]');if(dish){openRecipe(dish.dataset.recipe);return}editSlot(+el.dataset.day,el.dataset.slot)})})}
function editSlot(day,slot){
 const cal=monthData(),k=key(viewed.getFullYear(),viewed.getMonth(),day,slot),original=cal.manual?.[k]??cal.schedule?.[k]??[],options=recipes.filter(r=>r.active&&allowed(r,slot)).sort((a,b)=>a.name.localeCompare(b.name,'es')),first=selectBox(options,original[0]);
 const selects=slot==='dinner'?[first,selectBox(options.filter(r=>!r.single),original[1])]:[first];
 openEditor(day,slot,selects)
}
function selectBox(opts,selected){const s=document.createElement('select');s.className='editor-select';s.innerHTML='<option value="">— Dejar pendiente —</option>'+opts.map(r=>`<option value="${r.id}" ${r.id===selected?'selected':''}>${escapeHTML(r.name)}${r.single?'':' · requiere pareja'}</option>`).join('');return s}
function openEditor(day,slot,selects){
 const dialog=$('#mealEditorDialog'),box=document.createElement('div');box.className='meal-editor-content';
 const close=document.createElement('button');close.type='button';close.className='close-button';close.setAttribute('aria-label','Cerrar');close.textContent='×';close.onclick=()=>dialog.close();
 const title=document.createElement('h2');title.id='mealEditorTitle';title.textContent=`${day} de ${monthNames[viewed.getMonth()]} · ${slot==='lunch'?'Comida':'Cena'}`;
 box.append(close,title);
 selects.forEach((select,i)=>{const label=document.createElement('label');label.textContent=i?'Segundo plato':'Plato';label.append(select);box.append(label)});
 if(slot==='dinner'){
   const refreshPair=()=>{const first=recipes.find(r=>r.id===selects[0].value);selects[1].disabled=!first||first.single;if(selects[1].disabled)selects[1].value='';for(const option of selects[1].options)option.disabled=!!option.value&&option.value===first?.id;if(selects[1].value===first?.id)selects[1].value=''};
   selects[0].addEventListener('change',refreshPair);refreshPair();
 }
 const note=document.createElement('p');note.className='meal-editor-note';note.textContent='Los cambios manuales pueden saltarse las frecuencias y quedarán marcados en el calendario.';box.append(note);
 const save=document.createElement('button');save.type='button';save.textContent='Guardar cambio';save.className='primary-button';
 save.onclick=()=>{
   const first=recipes.find(r=>r.id===selects[0].value),second=slot==='dinner'?recipes.find(r=>r.id===selects[1].value):null,values=selects.map(s=>s.value).filter(Boolean);
   if(slot==='dinner'&&first&&!first.single&&!second&&!first.manualSingleAllowed){alert('Este plato necesita un segundo plato marcado “No” para completar la cena.');return}
   if(slot==='dinner'&&second&&(!first||first.single||second.single||first.id===second.id)){alert('La pareja debe tener dos platos distintos, ambos marcados “No”.');return}
   const monthKey=`${viewed.getFullYear()}-${viewed.getMonth()+1}`,cal=saved[monthKey]||{schedule:{},manual:{}};saved[monthKey]=cal;cal.manual??={};cal.manual[key(viewed.getFullYear(),viewed.getMonth(),day,slot)]=values;
   persist();dialog.close();render();showWarnings([])
 };
 box.append(save);dialog.replaceChildren(box);dialog.showModal()
}
function renderToday(){const now=new Date(),year=now.getFullYear(),month=now.getMonth(),day=now.getDate(),date=now.toLocaleDateString('es-ES',{weekday:'long',day:'numeric',month:'long'}),cal=saved[`${year}-${month+1}`]||{schedule:{},manual:{}};$('#todayHeading').textContent='Hoy';$('#todayContent').innerHTML=`<div class="today-date-card"><span class="today-sun">☀</span><div><span class="eyebrow">TU MENÚ DE HOY</span><h2>${date}</h2></div></div><div class="today-meal-cards">${todayMealHTML(cal,year,month,day,'lunch','Comida')}${todayMealHTML(cal,year,month,day,'dinner','Cena')}</div>`}
function todayMealHTML(cal,year,month,day,slot,label){
 const k=key(year,month,day,slot),weekday=new Date(year,month,day).getDay(),free=weekday===0||weekday===6||(weekday===5&&slot==='dinner'),ids=cal.manual?.[k]??cal.schedule?.[k]??[],names=ids.map(id=>recipes.find(r=>r.id===id)).filter(r=>r&&!r.familyOnly);
 const dishes=free?'<span class="today-unassigned">Libre</span>':names.length?names.map(r=>`<button class="today-dish-link" data-recipe="${r.id}">${escapeHTML(r.name)}</button>`).join('<span class="today-plus">+</span>'):'<span class="today-unassigned">Sin asignar</span>';
 return `<article class="today-meal-card ${slot==='dinner'?'dinner':''}"><div class="today-meal-label">${label}</div><div class="today-meal-dishes">${dishes}</div></article>`
}
function escapeHTML(value){return String(value).replace(/[&<>"']/g,ch=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[ch]))}
function recipeContentHTML(body,title){const lines=String(body||'').split(/\r?\n/).map(x=>x.trim()).filter(Boolean);while(lines.length&&lines[0].toLocaleLowerCase('es')===String(title).toLocaleLowerCase('es'))lines.shift();const sectionPattern=/^(ingredientes?(?:\s+.*)?|preparad[oa](?:\s+.*)?|preparaci[oó]n(?:\s+.*)?|elaboraci[oó]n(?:\s+.*)?|procedimiento(?:\s+.*)?|pasos?(?:\s+\d+)?|montad[oa](?:\s+.*)?|montaje(?:\s+.*)?|limpieza(?:\s+.*)?|cocci[oó]n(?:\s+.*)?|marinado(?:\s+.*)?|recomendaciones?(?:\s+.*)?|consejos?(?:\s+.*)?|notas?(?:\s+.*)?|condimentos?(?:\s+.*)?|base(?:\s+.*)?|salsa(?:\s+.*)?|relleno(?:\s+.*)?|emplatado(?:\s+.*)?)\s*:?$/i;let mode='text',items=[],html='';const flush=()=>{if(!items.length)return;const cls=mode==='ingredients'?'ingredient-list':mode==='steps'?'step-list':'text-list';html+=`<${mode==='steps'?'ol':'ul'} class="${cls}">${items.map(t=>`<li>${escapeHTML(t)}</li>`).join('')}</${mode==='steps'?'ol':'ul'}>`;items=[]};for(const line of lines){const section=line.match(sectionPattern);if(section){flush();const label=line.replace(/:\s*$/,'');mode=/^(ingredientes?)/i.test(label)?'ingredients':/^(preparad[oa]|preparaci[oó]n|elaboraci[oó]n|procedimiento|pasos?|cocci[oó]n)/i.test(label)?'steps':'text';html+=`<h3 class="recipe-section-title">${escapeHTML(label.replace(/:$/,''))}</h3>`;continue}if(mode==='text'&&/^[^.!?]{2,65}:$/.test(line)){flush();html+=`<h4 class="recipe-subtitle">${escapeHTML(line.slice(0,-1))}</h4>`;continue}if(mode==='ingredients'&&/^[^.!?]{2,65}:$/.test(line)){flush();html+=`<h4 class="recipe-subtitle">${escapeHTML(line.slice(0,-1))}</h4>`;mode='ingredient-sublist';continue}if(mode==='ingredient-sublist'&&/^[^.!?]{2,65}:$/.test(line)){flush();html+=`<h4 class="recipe-subtitle">${escapeHTML(line.slice(0,-1))}</h4>`;continue}if(mode==='ingredient-sublist')mode='ingredients';if(mode==='ingredients'||mode==='steps')items.push(line);else{flush();html+=`<p class="recipe-paragraph">${escapeHTML(line)}</p>`}}flush();return html||'<p class="recipe-paragraph">No hay texto asociado.</p>'}
function openRecipe(id){const r=recipes.find(x=>x.id===id)||catalogRecipes.find(x=>x.id===id);if(!r)return;$('#recipeTitle').textContent=r.name;$('#recipeMeta').textContent=[r.category,r.type,r.moment,r.single===undefined?'':(r.single?'Plato único':'Necesita acompañamiento'),r.linkStatus].filter(Boolean).join(' · ');const missing=r.body?'':catalog.headings.some(h=>normalizeHeadingTitle(h.title)===normalizeHeadingTitle(r.name))?'En el Word aparece este encabezado, pero no he encontrado texto asociado debajo. Puede ser una ficha pendiente de completar.':'No encuentro un encabezado equivalente en el Word importado. Puede estar con otro nombre o no estar todavía en la copia de trabajo.';$('#recipeBody').innerHTML=recipeContentHTML(r.body||missing,r.name);$('#recipeDialog').showModal()}
function renderRecipeLibrary(){
 const filter=$('#categoryFilter'),category=filter.value,categories=[...new Set(catalogRecipes.map(r=>r.category).filter(Boolean))].sort((a,b)=>a.localeCompare(b,'es'));
 filter.innerHTML='<option value="">Todas las categorías</option>'+categories.map(c=>`<option value="${escapeHTML(c)}">${escapeHTML(c)}</option>`).join('');filter.value=category;
 const search=normalizeHeadingTitle($('#recipeSearch').value),found=catalogRecipes.filter(r=>(!category||r.category===category)&&(!search||normalizeHeadingTitle(`${r.name} ${r.category} ${r.type} ${r.body}`).includes(search))).sort((a,b)=>a.name.localeCompare(b.name,'es'));
 $('#recipeCount').textContent=`${found.length} entradas`;
 $('#recipeList').innerHTML=found.map(r=>`<button class="recipe-card" data-recipe="${r.id}"><span class="recipe-card-title">${escapeHTML(r.name)}</span><span class="recipe-card-meta">${escapeHTML(r.familyOnly?'Familia / información':r.category||r.type||'Sin clasificar')}</span></button>`).join('')||'<p class="empty-library">No hay resultados para esa búsqueda.</p>';
 $('#recipeList').querySelectorAll('[data-recipe]').forEach(el=>el.onclick=()=>openRecipe(el.dataset.recipe))
}
function ensureMonth(){const m=saved[`${viewed.getFullYear()}-${viewed.getMonth()+1}`];const hasAssignments=Object.values(m?.schedule||{}).some(ids=>Array.isArray(ids)&&ids.length)||Object.values(m?.manual||{}).some(ids=>Array.isArray(ids)&&ids.length);if(!m?.deliberatelyCleared&&(!m?.schedule||m?.algorithmVersion!==scheduleVersion||m?.needsRegeneration||!hasAssignments))generate();else{render();showWarnings([])}}
$('#prevMonth').onclick=()=>{viewed.setMonth(viewed.getMonth()-1);ensureMonth()};$('#nextMonth').onclick=()=>{viewed.setMonth(viewed.getMonth()+1);ensureMonth()};$('#thisMonth').onclick=()=>{viewed=new Date();viewed.setDate(1);ensureMonth()};$('#generate').onclick=()=>generate(true);
$('#clearMonth').onclick=()=>{const cal=monthData();cal.manual={};cal.schedule={};cal.deliberatelyCleared=true;cal.algorithmVersion=scheduleVersion;saved[`${viewed.getFullYear()}-${viewed.getMonth()+1}`]=cal;persist();render();showWarnings([])};
$('#recipeSearch').addEventListener('input',renderRecipeLibrary);$('#categoryFilter').addEventListener('change',renderRecipeLibrary);
const viewRoutes=['calendario','hoy','recetas'];
let currentView='';
function routeFromHash(){const route=location.hash.replace(/^#/,'').toLocaleLowerCase('es');return viewRoutes.includes(route)?route:(matchMedia('(max-width:700px)').matches?'hoy':'calendario')}
function showView(route){if(!viewRoutes.includes(route))route='calendario';currentView=route;document.querySelectorAll('.app-view').forEach(el=>{el.hidden=el.id!==({calendario:'calendarView',hoy:'todayView',recetas:'recipesView'}[route])});document.querySelectorAll('.nav-item').forEach(el=>{const active=el.hash===`#${route}`;el.classList.toggle('active',active);if(active)el.setAttribute('aria-current','page');else el.removeAttribute('aria-current')});document.title=`Cocina · ${{calendario:'Calendario',hoy:'Hoy',recetas:'Recetas'}[route]}`;if(route==='hoy'){const now=new Date();viewed=new Date(now.getFullYear(),now.getMonth(),1);ensureMonth();renderToday()}else if(route==='calendario')ensureMonth()}
function navigateTo(route){if(!viewRoutes.includes(route))return;if(route===currentView)return;history.pushState({view:route},'',`#${route}`);showView(route)}
document.querySelectorAll('.nav-item, .brand').forEach(link=>link.addEventListener('click',event=>{event.preventDefault();navigateTo(link.hash.replace(/^#/,''))}));
$('#todayContent').addEventListener('click',event=>{const dish=event.target.closest('[data-recipe]');if(dish)openRecipe(dish.dataset.recipe)});
window.addEventListener('popstate',()=>showView(routeFromHash()));
renderRecipeLibrary();const initialView=routeFromHash();history.replaceState({view:initialView},'',`#${initialView}`);showView(initialView);
