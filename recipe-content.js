/* Organización de recetas validada, conservando el texto de origen. */
(function(root){
 'use strict';
 const normalize=value=>String(value||'').normalize('NFD').replace(/\p{M}/gu,'').toLocaleLowerCase('es').trim();
 const source=value=>String(value||'').replace(/\r\n/g,'\n').trim();
 const escape=value=>String(value).replace(/[&<>"']/g,char=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[char]));
 function validateItems(items,depth=0){
  if(!Array.isArray(items)||!items.length||depth>5)throw new Error('Lista de receta inválida.');
  for(const item of items){
   if(typeof item==='string'){if(!item.trim())throw new Error('Ingrediente o paso vacío.');}
   else if(item&&typeof item.text==='string'&&item.text.trim())validateItems(item.children,depth+1);
   else throw new Error('Ingrediente o paso inválido.');
  }
 }
 function validateModels(models){
  if(!Array.isArray(models))throw new Error('Falta la lista de contenido revisado.');
  const names=new Set();
  for(const model of models){
   if(typeof model.name!=='string'||!model.name.trim()||names.has(normalize(model.name)))throw new Error('Nombre de receta ausente o duplicado.');
   names.add(normalize(model.name));
   if(!source(model.sourceBody)||!Array.isArray(model.blocks)||!model.blocks.length)throw new Error('Falta el texto original o su organización.');
   for(const block of model.blocks){
    if(['ingredients','steps','bullets'].includes(block.type))validateItems(block.items);
    else if(['heading','text','note','raw'].includes(block.type)){
     if(typeof block.text!=='string'||!block.text.trim())throw new Error('Bloque de texto vacío.');
     if(block.type==='heading'&&![3,4].includes(block.level))throw new Error('Nivel de título inválido.');
    }else throw new Error('Tipo de bloque desconocido.');
   }
  }
  return models;
 }
 function findModel(models,name){return models.find(model=>normalize(model.name)===normalize(name))||null;}
 function matchesSource(model,body){return !!model&&source(model.sourceBody)===source(body);}
 function list(items,type){
  const tag=type==='steps'?'ol':'ul',css=type==='ingredients'?'ingredient-list':type==='steps'?'step-list':'text-list';
  return `<${tag} class="${css}">${items.map(item=>typeof item==='string'?`<li>${escape(item)}</li>`:`<li>${escape(item.text)}${list(item.children,type==='ingredients'?'ingredients':'bullets')}</li>`).join('')}</${tag}>`;
 }
 function render(model,body,renderRaw){
  if(!matchesSource(model,body))return null;
  validateModels([model]);
  return model.blocks.map(block=>{
   if(['ingredients','steps','bullets'].includes(block.type))return list(block.items,block.type);
   if(block.type==='heading')return `<h${block.level} class="${block.level===3?'recipe-section-title':'recipe-subtitle'}">${escape(block.text)}</h${block.level}>`;
   if(block.type==='raw'&&renderRaw)return renderRaw(block.text);
   return `<p class="recipe-paragraph${block.type==='note'?' recipe-note':''}">${escape(block.text)}</p>`;
  }).join('');
 }
 const api={validateModels,findModel,matchesSource,render};
 if(typeof module==='object'&&module.exports)module.exports=api;else root.CocinaRecipeContent=api;
})(typeof window==='object'?window:globalThis);
