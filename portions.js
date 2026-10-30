/* Cantidades estructuradas y correcciones explícitas de sus referencias. */
(function(root){
 'use strict';
 const normalize=value=>String(value||'').normalize('NFD').replace(/\p{M}/gu,'').toLocaleLowerCase('es').replace(/[^\p{L}\p{N}]+/gu,' ').trim();
 function validatePeople(value){if(!Number.isSafeInteger(value)||value<1)throw new Error('Indica un número entero de personas mayor que cero.');return value;}
 function validateModels(models){
  if(!Array.isArray(models))throw new Error('Las raciones deben ser una lista.');
  const names=new Set();
  for(const model of models){
   if(!Array.isArray(model.names)||!model.names.length||model.names.some(name=>typeof name!=='string'||!name.trim()))throw new Error('Falta el nombre de una receta.');
   for(const name of model.names){const key=normalize(name);if(names.has(key))throw new Error(`Raciones duplicadas para ${name}.`);names.add(key);}
   if(!Number.isFinite(model.basePeople)||model.basePeople<=0)throw new Error('Falta la ración base.');
   if(!['word_per_person','owner_standard'].includes(model.sourceType))throw new Error('Origen de cantidades no reconocido.');
   if(model.sourceType==='word_per_person'&&model.basePeople!==1)throw new Error('Una referencia por persona debe tener base 1.');
   if(!['partial','complete'].includes(model.coverage))throw new Error('Cobertura de cantidades no reconocida.');
   if(!Array.isArray(model.quantities)||!model.quantities.length)throw new Error('Faltan cantidades.');
   if(model.statusText!==undefined&&typeof model.statusText!=='string')throw new Error('Nota de cantidades inválida.');
   if(model.referenceOverrides!==undefined){
    if(!Array.isArray(model.referenceOverrides)||model.referenceOverrides.some(item=>!item||typeof item.from!=='string'||!item.from.trim()||typeof item.to!=='string'||!item.to.trim()||/[\r\n]/.test(item.from+item.to)))throw new Error('Corrección de referencia inválida.');
    if(new Set(model.referenceOverrides.map(item=>item.from)).size!==model.referenceOverrides.length)throw new Error('Correcciones de referencia duplicadas.');
   }
   for(const row of model.quantities){
    if(!row.label||!row.unit?.one||!row.unit?.many)throw new Error('Ingrediente o unidad incompletos.');
    if(!Number.isFinite(row.amount)||row.amount<0||row.maxAmount!==undefined&&(!Number.isFinite(row.maxAmount)||row.maxAmount<row.amount))throw new Error('Cantidad inválida.');
    if(model.sourceType==='word_per_person'&&(!row.evidence||typeof row.evidence!=='string'))throw new Error('Falta la referencia por persona en Word.');
   }
  }
  return models;
 }
 function findModel(models,name){return models.find(model=>model.names.some(alias=>normalize(alias)===normalize(name)))||null;}
 function scale(model,body,people){
  validatePeople(people);
  if(!model)return null;
  validateModels([model]);
  if(model.sourceType==='owner_standard'&&model.review?.status!=='confirmed')return null;
  const valid=model.quantities.filter(row=>model.sourceType!=='word_per_person'||String(body||'').includes(row.evidence));
  if(!valid.length)return null;
  const factor=people/model.basePeople;
  return {people,basePeople:model.basePeople,sourceType:model.sourceType,coverage:valid.length===model.quantities.length?model.coverage:'partial',note:model.note||'',statusText:model.statusText||'',quantities:valid.map(row=>({...row,amount:row.amount*factor,...(row.maxAmount===undefined?{}:{maxAmount:row.maxAmount*factor})}))};
 }
 function displayBody(model,body){
  const original=String(body||'');
  if(!model||model.sourceType!=='owner_standard'||model.review?.status!=='confirmed'||!model.referenceOverrides?.length)return original;
  validateModels([model]);
  const replacements=new Map(model.referenceOverrides.map(item=>[item.from,item.to]));
  return original.split(/\r?\n/).map(line=>replacements.get(line.trim())??line).join('\n');
 }
 const number=value=>new Intl.NumberFormat('es-ES',{maximumFractionDigits:3}).format(value);
 function formatQuantity(row){const range=row.maxAmount!==undefined&&row.maxAmount!==row.amount,amount=range?`${number(row.amount)}–${number(row.maxAmount)}`:number(row.amount),unit=!range&&row.amount===1?row.unit.one:row.unit.many;return `${amount} ${unit}`;}
 const api={normalize,validatePeople,validateModels,findModel,scale,formatQuantity,displayBody};
 if(typeof module==='object'&&module.exports)module.exports=api;else root.CocinaPortions=api;
})(typeof window==='object'?window:globalThis);
